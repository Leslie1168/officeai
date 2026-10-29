using System;
using System.IO;
using System.Runtime.InteropServices;

namespace OfficeAi.Watchdog
{
    internal sealed class DumpAttempt
    {
        public string Kind;          // "NULL(ExceptionParam=NULL)" 或 "EXC(带异常信息)"
        public bool ReportedOk;      // MiniDumpWriteDump 的返回值
        public int Win32Error;
        public string Path;
        public long Bytes;
        public uint FlagsUsed;
        public bool Accepted;        // 判据：返回值 true **且** 文件 > 最小可分析尺寸
    }

    internal sealed class DumpResult
    {
        public bool Success;         // 是否至少产出一份被接受的可分析 dump
        public string Path;          // 被接受的那份（优先 EXC，其次 NULL）
        public long Bytes;
        public uint FlagsUsed;
        public string Error;
        public DumpAttempt NullAttempt;
        public DumpAttempt ExcAttempt;
    }

    /// <summary>
    /// Minidump 写入。
    ///
    /// ★★ 本机实测结论（2026-10-15，Lead 独立实验 + 本工具 --selftest/__DumpProbe 双向复现）：
    ///
    ///   事实 1：`MiniDumpWriteDump(..., ExceptionParam = NULL, ...)` —— **稳定可用**
    ///           自 dump 81MB / notepad 全内存 526MB，都能正常落盘。
    ///   事实 2：`ExceptionParam != NULL` —— **本机 System32\dbghelp.dll 上必然失败**，
    ///           返回 0x800703E6（= HRESULT_FROM_WIN32(998) ERROR_NOACCESS），
    ///           而且**会让调用进程直接崩**（probe 的 T3.dmp 是 0 字节，进程随后消失）。
    ///           把 ClientPointers 取反、把 ContextRecord 置 NULL、换真实 GetThreadContext、
    ///           只给 ContextRecord、换标志组合 —— 六种组合全部同样失败。
    ///           在自己的进程上自 dump 也一样失败 → **不是跨进程指针问题，是这条通道本身坏了**。
    ///   事实 3：`ClientPointers` 的语义在网上有两种互斥说法，本机两种都试过，无差别。
    ///
    /// 所以本文件采用 **dump + sidecar 双产物** 策略：
    ///   ① 首选：ExceptionParam=NULL 的全内存 dump —— 无条件先落盘，这是不能丢的产物。
    ///      全内存 dump 里**本来就包含**出事线程的 CONTEXT（在 ThreadListStream 的
    ///      ThreadContext 里），只是没有 ExceptionStream 把它们串成"异常指针"。
    ///      分析时用 sidecar 里给的线程 ID，`cdb -z <dump> -c "~~[tid]s; r; k"` 就能落到现场。
    ///   ② sidecar：`<dump同名>.exception.txt`，把我们自己抓到的异常上下文（异常码、
    ///      异常地址、模块名+基址+RVA、访问类型、线程 ID、寄存器、栈候选）落成配对文件。
    ///      **模块名+RVA 这一行才是本案最值钱的信息，而它完全不依赖 dbghelp。**
    ///   ③ 可选：`--try-exception-param` 时会**在①成功落盘之后**再试一次带异常信息的版本
    ///      （这样即使 dbghelp 把那一次调用搞崩，我们已经拿到成品了）。换台 dbghelp 正常的
    ///      机器就能拿到更完整的 dump。
    ///
    /// 判据：**看文件是否真的 > 1MB**，不看返回值。本机已经出现过"返回失败但文件其实是好的"
    /// 与"文件 0 字节"两种情形，返回值不可信。
    /// </summary>
    internal static class DumpWriter
    {
        /// <summary>可分析的最小尺寸：小于这个值说明 dump 是残缺的（本机实测全内存 dump 至少几十 MB）。</summary>
        public const long MinAnalyzableBytes = 1024 * 1024;

        public static DumpResult Write(
            IntPtr hProcess, uint pid, NativeMethods.EXCEPTION_RECORD record, uint threadId,
            IntPtr contextBuffer, int contextSize, string outDir, string fileTag, bool clientPointers,
            bool tryExceptionParam, WatchdogLog log)
        {
            var result = new DumpResult();

            uint primary = NativeMethods.MiniDumpWithFullMemory
                         | NativeMethods.MiniDumpWithThreadInfo
                         | NativeMethods.MiniDumpWithUnloadedModules
                         | NativeMethods.MiniDumpWithHandleData
                         | NativeMethods.MiniDumpWithFullMemoryInfo
                         | NativeMethods.MiniDumpWithProcessThreadData;

            // ---------- ① 首选：ExceptionParam = NULL ----------
            string nullPath = BuildFileName(outDir, fileTag, pid, record, null);
            log.Info("【dump/①】首选写法：ExceptionParam = NULL（本机 dbghelp 唯一可靠的通道）→ " + Path.GetFileName(nullPath));
            result.NullAttempt = Attempt(hProcess, pid, nullPath, primary, IntPtr.Zero, 0, log, "NULL(ExceptionParam=NULL)");

            // 失败则降级标志组合再试一次（去掉可能不被支持的扩展标志）
            if (!result.NullAttempt.Accepted)
            {
                uint fallback = NativeMethods.MiniDumpWithFullMemory | NativeMethods.MiniDumpWithThreadInfo;
                log.Warn("【dump/①】首次未产出可分析文件，降级为 WithFullMemory|WithThreadInfo 再试一次（只改 DumpType，"
                         + "因为本机问题出在 ExceptionParam 通道，与标志无关）");
                TryDelete(nullPath);
                var a2 = Attempt(hProcess, pid, nullPath, fallback, IntPtr.Zero, 0, log, "NULL(降级标志)");
                if (a2.Accepted || a2.Bytes > result.NullAttempt.Bytes) result.NullAttempt = a2;
            }

            if (result.NullAttempt.Accepted)
            {
                result.Success = true;
                result.Path = result.NullAttempt.Path;
                result.Bytes = result.NullAttempt.Bytes;
                result.FlagsUsed = result.NullAttempt.FlagsUsed;
            }

            // ---------- ③ 可选：带异常信息的版本（仅在①已经拿到成品之后） ----------
            if (tryExceptionParam && result.Success)
            {
                log.Info("【dump/③】--try-exception-param 已开启：现在才尝试带异常信息的版本。"
                         + "之所以排在①之后，是因为本机 dbghelp 在这条路径上会让调用进程崩溃 —— "
                         + "先保住①，再赌③，即使③把看门狗搞死，成品已经在磁盘上了。");

                IntPtr recBuf = IntPtr.Zero, ctxBuf = IntPtr.Zero, epBuf = IntPtr.Zero;
                try
                {
                    int recSize = NativeMethods.EXCEPTION_RECORD.NativeSize;
                    recBuf = Marshal.AllocHGlobal(recSize);
                    ZeroMemory(recBuf, recSize);
                    Marshal.StructureToPtr(record, recBuf, false);

                    bool haveContext = contextBuffer != IntPtr.Zero && contextSize > 0;
                    if (haveContext)
                    {
                        ctxBuf = Marshal.AllocHGlobal(contextSize);
                        NativeMethods.RtlMoveMemory(ctxBuf, contextBuffer, (UIntPtr)(uint)contextSize);
                    }

                    epBuf = Marshal.AllocHGlobal(Marshal.SizeOf(typeof(NativeMethods.EXCEPTION_POINTERS)));
                    var ep = new NativeMethods.EXCEPTION_POINTERS
                    {
                        ExceptionRecord = recBuf,
                        ContextRecord = haveContext ? ctxBuf : IntPtr.Zero
                    };
                    Marshal.StructureToPtr(ep, epBuf, false);

                    var mei = new NativeMethods.MINIDUMP_EXCEPTION_INFORMATION
                    {
                        ThreadId = threadId,
                        ExceptionPointers = epBuf,
                        ClientPointers = clientPointers ? 1 : 0
                    };

                    string excPath = BuildFileName(outDir, fileTag, pid, record, "_exc");
                    result.ExcAttempt = AttemptRef(hProcess, pid, excPath, primary, ref mei, log,
                        "EXC(ClientPointers=" + (clientPointers ? "1" : "0") + ")");

                    if (result.ExcAttempt.Accepted)
                    {
                        log.Banner("【dump/③】带异常信息的 dump 也成功产出（本机罕见）→ 采用它作为主产物：" + result.ExcAttempt.Path);
                        // 保留两份：EXC 为主（含 ExceptionStream，cdb 可直接 .ecxr），NULL 作为独立证据
                        result.Path = result.ExcAttempt.Path;
                        result.Bytes = result.ExcAttempt.Bytes;
                        result.FlagsUsed = result.ExcAttempt.FlagsUsed;
                    }
                    else
                    {
                        log.Warn("【dump/③】带异常信息的版本不可用（" + Describe(result.ExcAttempt) + "）→ 删除它，保留①的 NULL 版本。"
                                 + "这不是我们的 bug：本机 dbghelp 的 ExceptionParam 通道是坏的。");
                        TryDelete(excPath);
                    }
                }
                catch (Exception ex)
                {
                    log.Warn("【dump/③】构造异常信息时抛异常（忽略，①的成品不受影响）：" + ex.GetType().Name + " " + ex.Message);
                }
                finally
                {
                    if (epBuf != IntPtr.Zero) Marshal.FreeHGlobal(epBuf);
                    if (ctxBuf != IntPtr.Zero) Marshal.FreeHGlobal(ctxBuf);
                    if (recBuf != IntPtr.Zero) Marshal.FreeHGlobal(recBuf);
                }
            }
            else if (tryExceptionParam && !result.Success)
            {
                log.Warn("【dump/③】①没有拿到可分析成品，跳过带异常信息的尝试（避免在没有任何产物的情况下让 dbghelp 崩掉看门狗）。");
            }

            if (!result.Success)
            {
                result.Error = "NULL 通道=" + Describe(result.NullAttempt);
                log.Error("★ dump 完全失败：①②③ 都没能产出 > " + (MinAnalyzableBytes / 1024 / 1024) + "MB 的 dump。"
                          + "模块名+RVA 与寄存器仍已写入 sidecar 与主日志，取证链没有断。");
            }

            return result;
        }

        // ------------------------------------------------------------------ 单次尝试

        private static DumpAttempt Attempt(IntPtr hProcess, uint pid, string path, uint flags,
            IntPtr excParam, int unused, WatchdogLog log, string kind)
        {
            var a = new DumpAttempt { Kind = kind, Path = path, FlagsUsed = flags };
            FileStream fs = null;
            try
            {
                fs = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None, 64 * 1024);
                a.ReportedOk = NativeMethods.MiniDumpWriteDump(hProcess, pid, fs.SafeFileHandle.DangerousGetHandle(),
                    flags, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero);
                a.Win32Error = Marshal.GetLastWin32Error();
                try { fs.Flush(); } catch { }
            }
            catch (Exception ex)
            {
                a.ReportedOk = false;
                a.Win32Error = -1;
                log.Warn("【dump】" + kind + " 调用抛异常：" + ex.GetType().Name + " " + ex.Message);
            }
            finally
            {
                try { if (fs != null) fs.Dispose(); } catch { }
            }

            a.Bytes = FileLength(path);
            a.Accepted = a.Bytes > MinAnalyzableBytes;   // ★ 判据是文件大小，不是返回值
            log.Info("【dump】" + kind + " → 返回值=" + (a.ReportedOk ? "true" : "false")
                     + " win32=" + a.Win32Error + " (0x" + (a.Win32Error & 0xFFFFFFFFL).ToString("X8") + ")"
                     + " 文件=" + a.Bytes + " 字节 → " + (a.Accepted ? "接受" : "拒绝"));
            if (!a.ReportedOk && a.Bytes > 0 && !a.Accepted)
            {
                log.Warn("【dump】" + kind + " 产生了不完整文件（" + a.Bytes + " 字节），已删除以免误导分析");
                TryDelete(path);
            }
            return a;
        }

        private static DumpAttempt AttemptRef(IntPtr hProcess, uint pid, string path, uint flags,
            ref NativeMethods.MINIDUMP_EXCEPTION_INFORMATION mei, WatchdogLog log, string kind)
        {
            var a = new DumpAttempt { Kind = kind, Path = path, FlagsUsed = flags };
            FileStream fs = null;
            try
            {
                fs = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None, 64 * 1024);
                a.ReportedOk = NativeMethods.MiniDumpWriteDump(hProcess, pid, fs.SafeFileHandle.DangerousGetHandle(),
                    flags, ref mei, IntPtr.Zero, IntPtr.Zero);
                a.Win32Error = Marshal.GetLastWin32Error();
                try { fs.Flush(); } catch { }
            }
            catch (Exception ex)
            {
                a.ReportedOk = false;
                a.Win32Error = -1;
                log.Warn("【dump】" + kind + " 调用抛异常（本机已知现象）：" + ex.GetType().Name + " " + ex.Message);
            }
            finally
            {
                try { if (fs != null) fs.Dispose(); } catch { }
            }

            a.Bytes = FileLength(path);
            a.Accepted = a.Bytes > MinAnalyzableBytes;
            log.Info("【dump】" + kind + " → 返回值=" + (a.ReportedOk ? "true" : "false")
                     + " win32=" + a.Win32Error + " (0x" + (a.Win32Error & 0xFFFFFFFFL).ToString("X8") + ")"
                     + " 文件=" + a.Bytes + " 字节 → " + (a.Accepted ? "接受" : "拒绝"));
            if (!a.Accepted)
            {
                if (a.Bytes > 0) log.Warn("【dump】" + kind + " 产生了不完整文件（" + a.Bytes + " 字节），已删除以免误导分析");
                TryDelete(path);
            }
            return a;
        }

        private static string Describe(DumpAttempt a)
        {
            if (a == null) return "(未尝试)";
            return a.Kind + " 返回值=" + (a.ReportedOk ? "true" : "false")
                 + " win32=0x" + (a.Win32Error & 0xFFFFFFFFL).ToString("X8")
                 + " 文件=" + a.Bytes + " 字节";
        }

        private static long FileLength(string path)
        {
            try
            {
                var fi = new FileInfo(path);
                return fi.Exists ? fi.Length : 0;
            }
            catch { return 0; }
        }

        private static void TryDelete(string path)
        {
            try { if (File.Exists(path)) File.Delete(path); } catch { }
        }

        private static void ZeroMemory(IntPtr p, int size)
        {
            for (int i = 0; i < size; i++) Marshal.WriteByte(p, i, 0);
        }

        public static string BuildFileName(string outDir, string fileTag, uint pid,
            NativeMethods.EXCEPTION_RECORD record, string suffix)
        {
            string proc = string.IsNullOrEmpty(fileTag) ? "target" : Sanitize(fileTag);
            string name = string.Format("{0}_{1}_0x{2:X8}_0x{3:X16}_{4}{5}.dmp",
                proc, pid, record.ExceptionCode, unchecked((ulong)record.ExceptionAddress.ToInt64()),
                DateTime.Now.ToString("yyyyMMdd-HHmmss.fff"), suffix ?? "");
            return System.IO.Path.Combine(outDir, name);
        }

        private static string Sanitize(string s)
        {
            char[] invalid = System.IO.Path.GetInvalidFileNameChars();
            char[] chars = s.ToCharArray();
            for (int i = 0; i < chars.Length; i++)
            {
                for (int j = 0; j < invalid.Length; j++)
                {
                    if (chars[i] == invalid[j]) { chars[i] = '_'; break; }
                }
            }
            return new string(chars);
        }
    }
}
