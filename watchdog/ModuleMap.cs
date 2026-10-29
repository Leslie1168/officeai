using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace OfficeAi.Watchdog
{
    internal sealed class ModuleEntry
    {
        public IntPtr Base;
        public uint Size;
        public string FullPath;
        public string Name;
        /// <summary>来源："EnumProcessModules" 或 "VirtualQueryEx"（反查得到的）。</summary>
        public string Source = "EnumProcessModules";

        public ulong BaseU { get { return unchecked((ulong)Base.ToInt64()); } }
        public ulong EndU { get { return BaseU + Size; } }

        public bool Contains(ulong address)
        {
            return Size > 0 && address >= BaseU && address < EndU;
        }
    }

    /// <summary>
    /// 目标进程的模块表（模块名 / 基址 / 大小）+ 地址→模块反查。
    ///
    /// 这是本次调查最重要的一个转换：把裸地址 0x183ba19 翻译成「哪个 dll + 偏移(RVA)」。
    /// 有了 RVA，即使没有符号、没有网络，也能直接和已知崩溃点比对
    /// （wwlib.dll+0x183BA19 每次相同 → 同一处代码）。
    ///
    /// ★ 两条独立通道（第一条挂掉时第二条仍然能用，实测必须）：
    ///   通道 1：EnumProcessModulesEx/EnumProcessModules + GetModuleInformation
    ///           —— 一次性拿到全部模块。但实测在"进程刚被调试启动、loader 还没跑完"时会
    ///           失败并返回 win32=299 (ERROR_PARTIAL_COPY)，模块表为空。
    ///   通道 2：VirtualQueryEx(出错地址) → MEM_IMAGE 的 AllocationBase 就是模块基址，
    ///           再用 GetMappedFileNameW 拿模块路径。
    ///           —— 只需要"出错的那一个地址"，不依赖模块表，反而更抗打。
    ///   反查时通道 2 永远可用作兜底；这样即使通道 1 全灭，"模块名 + RVA"这条链也不会断。
    /// </summary>
    internal static class ModuleMap
    {
        /// <summary>枚举目标进程已加载模块。失败返回空表（调用方降级，绝不抛）。</summary>
        public static List<ModuleEntry> Enumerate(IntPtr hProcess, WatchdogLog log)
        {
            var result = new List<ModuleEntry>();
            IntPtr[] handles = null;
            uint needed = 0;

            try
            {
                int capacity = 1024;
                bool ok = false;
                for (int attempt = 0; attempt < 4 && !ok; attempt++)
                {
                    handles = new IntPtr[capacity];
                    uint cb = (uint)(handles.Length * IntPtr.Size);
                    ok = NativeMethods.EnumProcessModulesEx(hProcess, handles, cb, out needed, NativeMethods.LIST_MODULES_ALL);
                    if (!ok)
                    {
                        // WOW64 / 权限差异时退回到不带过滤标志的版本
                        ok = NativeMethods.EnumProcessModules(hProcess, handles, cb, out needed);
                    }
                    if (!ok)
                    {
                        int err = System.Runtime.InteropServices.Marshal.GetLastWin32Error();
                        if (err == 299 /*ERROR_PARTIAL_COPY*/ && attempt < 3)
                        {
                            // 典型的"目标刚启动、loader 还没跑完"。等一下再试。
                            if (log != null && attempt == 0)
                            {
                                log.Warn("EnumProcessModules 返回 win32=299 (ERROR_PARTIAL_COPY)：目标进程的模块表此刻还读不出来"
                                         + "（被调试进程刚创建、loader 尚未初始化完）。稍后重试，并启用 VirtualQueryEx 兜底通道。");
                            }
                            System.Threading.Thread.Sleep(60 * (attempt + 1));
                            continue;
                        }
                        if (log != null)
                        {
                            log.Warn("EnumProcessModules 失败 win32=" + err + " (0x" + err.ToString("X8") + ")，"
                                     + "模块表枚举放弃；模块归属将走 VirtualQueryEx 反查兜底通道。");
                        }
                        return result;
                    }
                    if (needed <= cb) break;
                    capacity = (int)(needed / (uint)IntPtr.Size) + 64;
                }

                if (!ok || handles == null) return result;

                int count = (int)(needed / (uint)IntPtr.Size);
                if (count > handles.Length) count = handles.Length;

                for (int i = 0; i < count; i++)
                {
                    if (handles[i] == IntPtr.Zero) continue;
                    var info = new NativeMethods.MODULEINFO();
                    if (!NativeMethods.GetModuleInformation(hProcess, handles[i], out info,
                            (uint)System.Runtime.InteropServices.Marshal.SizeOf(typeof(NativeMethods.MODULEINFO))))
                        continue;

                    var sb = new StringBuilder(1024);
                    uint n = NativeMethods.GetModuleFileNameExW(hProcess, handles[i], sb, (uint)sb.Capacity);
                    string full = n > 0 ? sb.ToString() : "(未知路径)";

                    result.Add(new ModuleEntry
                    {
                        Base = info.lpBaseOfDll,
                        Size = info.SizeOfImage,
                        FullPath = full,
                        Name = FullPathToName(full),
                        Source = "EnumProcessModules"
                    });
                }
            }
            catch (Exception ex)
            {
                if (log != null) log.Exception("枚举模块表", ex);
            }

            return result;
        }

        /// <summary>在模块表里查地址；找不到返回 null。</summary>
        public static ModuleEntry Find(List<ModuleEntry> modules, ulong address)
        {
            if (modules == null || address == 0) return null;
            for (int i = 0; i < modules.Count; i++)
            {
                if (modules[i].Contains(address)) return modules[i];
            }
            return null;
        }

        /// <summary>
        /// 兜底通道：直接从地址反查模块（VirtualQueryEx + GetMappedFileNameW）。
        /// 不需要模块表，因此不受 EnumProcessModules 失败影响 —— 这是 A7/A8 能 PASS 的关键。
        /// </summary>
        public static ModuleEntry ResolveByAddress(IntPtr hProcess, ulong address, WatchdogLog log)
        {
            if (address == 0 || hProcess == IntPtr.Zero) return null;
            try
            {
                var mbi = new NativeMethods.MEMORY_BASIC_INFORMATION();
                UIntPtr got = NativeMethods.VirtualQueryEx(hProcess, (IntPtr)unchecked((long)address), out mbi,
                    (UIntPtr)(uint)System.Runtime.InteropServices.Marshal.SizeOf(typeof(NativeMethods.MEMORY_BASIC_INFORMATION)));
                if (got == UIntPtr.Zero) return null;
                if (mbi.Type != NativeMethods.MEM_IMAGE) return null;   // 只对映像映射做模块归属

                IntPtr baseAddr = mbi.AllocationBase;
                if (baseAddr == IntPtr.Zero) return null;

                var buf = new char[1024];
                uint n = NativeMethods.GetMappedFileNameW(hProcess, baseAddr, buf, (uint)buf.Length);
                string devPath = n > 0
                    ? NormalizeDevicePath(new string(buf, 0, (int)Math.Min(n, (uint)buf.Length)))
                    : null;
                string name = devPath != null ? SafeFileName(devPath) : null;

                // 用模块基址再查一次，拿到映像大小（RegionSize 可能只覆盖一个节，偏小）
                uint size = (uint)mbi.RegionSize.ToUInt64();
                try
                {
                    var mbi2 = new NativeMethods.MEMORY_BASIC_INFORMATION();
                    UIntPtr got2 = NativeMethods.VirtualQueryEx(hProcess, baseAddr, out mbi2,
                        (UIntPtr)(uint)System.Runtime.InteropServices.Marshal.SizeOf(typeof(NativeMethods.MEMORY_BASIC_INFORMATION)));
                    if (got2 != UIntPtr.Zero && mbi2.RegionSize.ToUInt64() > 0)
                    {
                        // 逐区域累加，直到地址范围不再属于同一个 AllocationBase
                        ulong total = 0;
                        ulong cur = unchecked((ulong)baseAddr.ToInt64());
                        var probe = new NativeMethods.MEMORY_BASIC_INFORMATION();
                        for (int i = 0; i < 4096; i++)
                        {
                            UIntPtr g = NativeMethods.VirtualQueryEx(hProcess, (IntPtr)unchecked((long)cur), out probe,
                                (UIntPtr)(uint)System.Runtime.InteropServices.Marshal.SizeOf(typeof(NativeMethods.MEMORY_BASIC_INFORMATION)));
                            if (g == UIntPtr.Zero) break;
                            if (probe.AllocationBase != baseAddr) break;
                            ulong rs = probe.RegionSize.ToUInt64();
                            if (rs == 0) break;
                            total += rs;
                            cur += rs;
                        }
                        if (total > 0) size = (uint)Math.Min(total, uint.MaxValue);
                    }
                }
                catch { }

                return new ModuleEntry
                {
                    Base = baseAddr,
                    Size = size,
                    FullPath = devPath ?? "(GetMappedFileName 失败)",
                    Name = name ?? "(未知模块)",
                    Source = "VirtualQueryEx"
                };
            }
            catch (Exception ex)
            {
                if (log != null) log.Warn("VirtualQueryEx 反查模块失败（忽略）：" + ex.GetType().Name);
                return null;
            }
        }

        private static char[] ToCharArray(StringBuilder sb, int size)
        {
            return new char[size];
        }

        /// <summary>把 GetMappedFileNameW 的 \Device\HarddiskVolumeN\... 转成 C:\... 便于人读。</summary>
        public static string NormalizeDevicePath(string path)
        {
            if (string.IsNullOrEmpty(path)) return path;
            if (!path.StartsWith("\\Device\\", StringComparison.OrdinalIgnoreCase)) return path;

            try
            {
                int secondSlash = path.IndexOf('\\', 8);
                if (secondSlash > 8)
                {
                    string device = path.Substring(0, secondSlash); // 例如 \Device\HarddiskVolume3
                    for (char letter = 'A'; letter <= 'Z'; letter++)
                    {
                        var buf = new char[1024];
                        uint n = NativeMethods.QueryDosDeviceW(letter + ":", buf, (uint)buf.Length);
                        if (n == 0) continue;
                        string target = new string(buf, 0, (int)Math.Min(n, (uint)buf.Length)).TrimEnd('\0');
                        if (string.Equals(target, device, StringComparison.OrdinalIgnoreCase))
                        {
                            return letter + ":" + path.Substring(secondSlash);
                        }
                    }
                }
            }
            catch { }
            return path;
        }

        /// <summary>
        /// 解析地址 → 模块，先查模块表，查不到再用 VirtualQueryEx 兜底。
        /// 返回的 ModuleEntry 若来自兜底通道，会被加进 cache 以便后续复用。
        /// </summary>
        public static ModuleEntry Resolve(IntPtr hProcess, List<ModuleEntry> modules, ulong address, WatchdogLog log)
        {
            ModuleEntry m = Find(modules, address);
            if (m != null) return m;
            m = ResolveByAddress(hProcess, address, log);
            if (m != null && modules != null)
            {
                modules.Add(m); // 缓存，避免同一个地址反复 VirtualQueryEx
                if (log != null)
                {
                    log.Info("模块归属走兜底通道（VirtualQueryEx+GetMappedFileNameW）：0x" + address.ToString("X16")
                             + " → " + m.Name + " 基址=0x" + m.BaseU.ToString("X") + " RVA=0x" + (address - m.BaseU).ToString("X")
                             + " 路径=" + m.FullPath);
                }
            }
            return m;
        }

        public static string Describe(IntPtr hProcess, List<ModuleEntry> modules, ulong address, WatchdogLog log)
        {
            if (address == 0) return "0x0（空指针）";
            ModuleEntry m = Resolve(hProcess, modules, address, log);
            if (m == null) return "unknown（地址不在任何已加载映像范围内，可能是 JIT/堆/已卸载模块/被回收的映像）";
            return string.Format("{0} 基址=0x{1:X} RVA=0x{2:X} 大小=0x{3:X} 来源={4} 路径={5}",
                m.Name, m.BaseU, address - m.BaseU, m.Size, m.Source, m.FullPath);
        }

        /// <summary>把地址翻译成「模块名+0xRVA」——这是喂给交接报告/人眼的最短有效形式。</summary>
        public static string ShortForm(IntPtr hProcess, List<ModuleEntry> modules, ulong address, WatchdogLog log)
        {
            if (address == 0) return "0x0 (空指针)";
            ModuleEntry m = Resolve(hProcess, modules, address, log);
            if (m == null) return "0x" + address.ToString("X") + " (unknown)";
            return string.Format("{0}+0x{1:X}", m.Name, address - m.BaseU);
        }

        /// <summary>保留旧名（枚举通道里用）：把模块全路径转成模块名。</summary>
        public static string FullPathToName(string path)
        {
            return SafeFileName(NormalizeDevicePath(path));
        }

        public static string SafeFileName(string path)
        {
            try { return Path.GetFileName(path); }
            catch { return path; }
        }
    }
}
