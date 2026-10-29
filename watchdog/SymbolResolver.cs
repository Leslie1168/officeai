using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

namespace OfficeAi.Watchdog
{
    /// <summary>
    /// 符号解析。定位：**纯加分项，绝不阻塞取证主链路**。
    ///
    /// 交接报告 3.7#4 记录：本机跑 cdb 的 !analyze -v 会卡在微软符号服务器下载上，
    /// 实测 180s 超时。看门狗绝不能重蹈覆辙 —— 它正卡在被调试进程的调试事件里，
    /// 它一卡，Word 就跟着冻结。
    ///
    /// 所以这里的三道闸：
    ///   1) **联网可达性预检**：attach 之前就用 2 秒超时探一次 msdl.microsoft.com。
    ///      探不通 → 只走本地缓存（`cache*<dir>`，DbgHelp 完全不发网络请求）。
    ///   2) **只解析极少数地址**：异常地址 + RIP + 栈上候选返回地址（通常 &lt; 40 个），
    ///      不做"load 全部模块"这种会触发海量下载的事。
    ///   3) **跑在独立线程上**，主调试循环不等待。就算它彻底卡死，模块名+RVA
    ///      （这才是本案真正要的那一条证据）也已经在主日志里写死了。
    ///
    /// 另外：DbgHelp 不是线程安全的，MiniDumpWriteDump 和 Sym* 共用同一份库；
    /// 所有 DbgHelp 调用统一走 DbgHelpGate 串行化。
    /// </summary>
    internal static class SymbolResolver
    {
        /// <summary>DbgHelp 全局串行闸（MiniDumpWriteDump 与 Sym* 都算）。</summary>
        public static readonly object DbgHelpGate = new object();

        public static string DefaultCacheDir
        {
            get
            {
                return Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "OfficeAi", "symbols");
            }
        }

        /// <summary>
        /// 预检符号服务器可达性。**这一步决定了后面会不会卡死**，所以必须短超时、必须吞掉所有异常。
        /// 返回 true 表示可以尝试联网下载符号。
        /// </summary>
        public static bool ProbeSymbolServer(int timeoutMs, WatchdogLog log)
        {
            string url = "https://msdl.microsoft.com/download/symbols/";
            try
            {
                var req = (System.Net.HttpWebRequest)System.Net.WebRequest.Create(url);
                req.Method = "HEAD";
                req.Timeout = timeoutMs;
                req.ReadWriteTimeout = timeoutMs;
                req.Proxy = null; // 不走系统代理：符号服务器直连更可预期，配错代理会引入新的超时源
                using (var resp = (System.Net.HttpWebResponse)req.GetResponse())
                {
                    int code = (int)resp.StatusCode;
                    log.Info("符号服务器预检：可达（HTTP " + code + "），允许联网解析符号（有预算上限）");
                    return true;
                }
            }
            catch (Exception ex)
            {
                log.Warn("符号服务器预检：不可达/超时（" + ex.GetType().Name + "：" + Trim(ex.Message, 160) + "）");
                log.Warn(" → 符号解析降级为「只用本地缓存 " + DefaultCacheDir + "」，DbgHelp 不会发起任何网络请求，看门狗不会因此卡住。");
                log.Warn(" → 位置证据不受影响：模块名 + 基址 + RVA 已足够把 0x183ba19 翻译成「哪个 dll 的哪个偏移」。");
                return false;
            }
        }

        /// <summary>
        /// 解析一组地址。**必须在后台线程调用**，耗时不可控（取决于缓存命中和网络）。
        /// </summary>
        public static void ResolveAll(
            uint pid, IntPtr hProcess, List<ModuleEntry> modules, List<ulong> addresses,
            string cacheDir, bool allowNetwork, WatchdogLog log)
        {
            if (addresses == null || addresses.Count == 0) return;

            lock (DbgHelpGate)
            {
                IntPtr hSym = IntPtr.Zero;
                try
                {
                    Directory.CreateDirectory(cacheDir);

                    string searchPath = allowNetwork
                        ? "srv*" + cacheDir + "*https://msdl.microsoft.com/download/symbols"
                        : "cache*" + cacheDir;

                    uint opts = NativeMethods.SYMOPT_DEFERRED_LOADS | NativeMethods.SYMOPT_UNDNAME
                              | NativeMethods.SYMOPT_CASE_INSENSITIVE | NativeMethods.SYMOPT_LOAD_LINES;
                    NativeMethods.SymSetOptions(opts);

                    // SymInitialize 的 hProcess 必须是唯一值；用调试目标句柄，DbgHelp 才能读目标进程
                    // 内存里的 PE 头来做符号匹配。
                    if (!NativeMethods.SymInitialize(hProcess, searchPath, false))
                    {
                        log.Warn("SymInitialize 失败 win32=" + Marshal.GetLastWin32Error() + "，跳过符号解析（模块名+RVA 不受影响）");
                        return;
                    }
                    hSym = hProcess;

                    log.Info("[SYM] 符号搜索路径 = " + searchPath);
                    log.Info("[SYM] 开始解析 " + addresses.Count + " 个地址（后台线程，主流程不等它）");

                    // 只加载"确实要用到的"模块，避免 DbgHelp 去下整个进程的符号
                    var wanted = new Dictionary<ulong, ModuleEntry>();
                    for (int i = 0; i < addresses.Count; i++)
                    {
                        ModuleEntry m = ModuleMap.Find(modules, addresses[i]);
                        if (m != null && !wanted.ContainsKey(m.BaseU)) wanted[m.BaseU] = m;
                    }

                    int loaded = 0, loadFailed = 0;
                    foreach (var kv in wanted)
                    {
                        ModuleEntry m = kv.Value;
                        ulong got = NativeMethods.SymLoadModuleEx(hProcess, IntPtr.Zero, m.FullPath, null,
                            m.BaseU, m.Size, IntPtr.Zero, 0);
                        if (got == 0) loadFailed++; else loaded++;
                    }
                    log.Info("[SYM] 模块加载：成功 " + loaded + " / 失败 " + loadFailed + "（失败通常是该模块没有公开符号，正常）");

                    // 去重后逐个解析；同一个地址只解一次
                    var seen = new HashSet<ulong>();
                    int hit = 0, miss = 0;
                    for (int i = 0; i < addresses.Count; i++)
                    {
                        ulong addr = addresses[i];
                        if (!seen.Add(addr)) continue;

                        string text = SymbolizeOne(hProcess, modules, addr, log);
                        if (text != null) { hit++; log.Info("[SYM] " + text); }
                        else miss++;
                    }
                    log.Info("[SYM] 解析完成：命中 " + hit + " / 未命中 " + miss
                             + "（未命中就退回「模块名+0xRVA」，这不影响定位）");
                }
                catch (Exception ex)
                {
                    log.Warn("符号解析线程异常（已忽略，不影响 dump 与模块证据）：" + ex.GetType().Name + " " + Trim(ex.Message, 200));
                }
                finally
                {
                    try { if (hSym != IntPtr.Zero) NativeMethods.SymCleanup(hSym); } catch { }
                }
            }
        }

        private static string SymbolizeOne(IntPtr hProcess, List<ModuleEntry> modules, ulong addr, WatchdogLog log)
        {
            string where = ModuleMap.ShortForm(hProcess, modules, addr, null);

            var si = new NativeMethods.SYMBOL_INFO();
            si.Name = new char[NativeMethods.SYMBOL_NAME_CHARS];
            si.MaxNameLen = NativeMethods.SYMBOL_NAME_CHARS;
            si.SizeOfStruct = (uint)Marshal.SizeOf(typeof(NativeMethods.SYMBOL_INFO));

            ulong disp = 0;
            bool ok;
            try
            {
                ok = NativeMethods.SymFromAddr(hProcess, addr, out disp, ref si);
            }
            catch (Exception ex)
            {
                log.Warn("[SYM] SymFromAddr(0x" + addr.ToString("X") + ") 抛异常：" + ex.GetType().Name);
                return null;
            }

            string name = null;
            if (ok && si.Name != null && si.NameLen > 0)
            {
                int len = (int)Math.Min((uint)si.NameLen, (uint)si.Name.Length);
                name = new string(si.Name, 0, len).TrimEnd('\0', ' ');
                if (name.Length == 0) name = null;
            }

            string lineInfo = "";
            try
            {
                var line = new NativeMethods.IMAGEHLP_LINE64();
                line.SizeOfStruct = (uint)Marshal.SizeOf(typeof(NativeMethods.IMAGEHLP_LINE64));
                uint lineDisp;
                if (NativeMethods.SymGetLineFromAddr64(hProcess, addr, out lineDisp, out line) && line.FileName != IntPtr.Zero)
                {
                    string file = Marshal.PtrToStringAnsi(line.FileName);
                    if (!string.IsNullOrEmpty(file)) lineInfo = "  文件=" + file + ":" + line.LineNumber;
                }
            }
            catch { }

            if (name == null && lineInfo.Length == 0) return null;

            string sym = name == null ? "(仅行号信息)" : name + "+0x" + disp.ToString("X");
            return string.Format("0x{0:X16} = {1}   {2}{3}", addr, where, sym, lineInfo);
        }

        private static string Trim(string s, int max)
        {
            if (string.IsNullOrEmpty(s)) return "";
            s = s.Replace("\r", " ").Replace("\n", " ");
            return s.Length <= max ? s : s.Substring(0, max) + "…";
        }
    }
}
