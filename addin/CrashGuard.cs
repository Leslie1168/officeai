using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

namespace OfficeAi
{
    /// <summary>
    /// 进程内崩溃取证（v2）。本机 WER 被全局禁用
    /// （HKLM\SOFTWARE\Microsoft\Windows\Windows Error Reporting\Disabled=1），
    /// Word 闪退（wwlib.dll 0xc0000005 @0x183ba19）拿不到任何转储，只能自己抓。
    ///
    /// ===== 为什么 v1 抓不到东西（两个真 bug）=====
    ///   1) 用 MiniDumpNormal(0x0)：只写线程/模块元数据，**不含线程栈内存**，
    ///      产物几十 KB，栈回溯全是问号，等于没抓。
    ///   2) 只装了 SetUnhandledExceptionFilter：它是进程级「后装者赢」——
    ///      Office / CLR / WebView2 任何一方在我们之后再装一个过滤器就把我们顶掉了；
    ///      而 __fastfail / 栈溢出这类异常天然绕过 SEH 过滤器。
    ///      → 所以 14:57 崩溃时 crashes 目录根本没被创建。
    ///
    /// ===== v2 的通道设计 =====
    ///   主通道：AddVectoredExceptionHandler（VEH，first=1 先入队，**不可被覆盖**，
    ///           且 first-chance 也能收到）。过滤 ACCESS_VIOLATION 与 STACK_OVERFLOW。
    ///   兜底：  SetUnhandledExceptionFilter（保留，但返回 EXCEPTION_CONTINUE_SEARCH）。
    ///   托管：  AppDomain.CurrentDomain.UnhandledException。
    ///   进程外：StartWatchdog() 拉起 T3 的看门狗（只负责 best-effort 启动，绝不阻塞）。
    ///
    /// ===== 关键纪律（写这个文件时不能违反）=====
    ///   ★ 出事的线程什么都不做：只把异常上下文指针抄进一个**预分配**的请求结构，然后
    ///     Set 一个事件并**限时**等待。所有分配/字符串/文件 IO/dbghelp 调用都交给
    ///     进程启动时就创建好的 dump 线程（健康线程）。
    ///     原因：AV 经常发生在堆锁/加载器锁里面，在出事线程上再分配内存或 LoadLibrary
    ///     会自死锁 —— 那就不是「闪退」，而是「Word 永久卡死」，比崩了还糟。
    ///   ★ 等待有硬上限（见 WaitForDumpCompletion）：绝不无限阻塞宿主。
    ///   ★ 任何失败都不能影响加载项：全部 try/catch 吞掉，只记日志。
    ///   ★ 绝不吞异常：所有通道都返回 EXCEPTION_CONTINUE_SEARCH(0)，
    ///     让 Word 走它原本的终止/恢复路径（见 UnhandledFilter 里的注释）。
    ///   ★ delegate 一律静态持有，否则被 GC 回收后原生回调进已释放 thunk = 崩溃。
    /// </summary>
    internal static class CrashGuard
    {
        // ===================== 常量 =====================

        /// <summary>每进程最多写几个 dump。正常致命崩溃只会有 1 个（写完进程就没了）；
        /// 上限是为了防「宿主自行处理掉的 AV」造成 dump 风暴（每个几百 MB）。</summary>
        private const int MaxDumpsPerProcess = 3;

        /// <summary>出事线程的第一段等待：3 秒。绝大多数情况（只记现场 / 小进程 dump）在这里结束。</summary>
        private const int SoftWaitMs = 3000;

        /// <summary>有实际进展（dump 文件在长大）时允许延长到的绝对上限，绝不无限阻塞。</summary>
        private const int HardWaitCapMs = 30000;

        private const int WaitSliceMs = 1000;

        /// <summary>可选的「栈文本」收集最多等多久（超时就放弃；dump 里本来就有全部线程栈）。</summary>
        private const int StackJoinTimeoutMs = 1500;

        /// <summary>full-memory dump 的标志组合（可 grep：MiniDumpWithFullMemory）。</summary>
        internal const uint DumpFlags =
            NativeProbe.MiniDumpWithFullMemory |        // 0x2 ★ 全内存：含所有线程栈
            NativeProbe.MiniDumpWithFullMemoryInfo |    // 内存区域属性（VirtualQuery 结果）
            NativeProbe.MiniDumpWithThreadInfo |        // 线程时间/TEB
            NativeProbe.MiniDumpWithHandleData |        // 句柄表（查 GDI/COM 句柄泄漏）
            NativeProbe.MiniDumpWithUnloadedModules |   // 已卸载模块（use-after-free 关键证据）
            NativeProbe.MiniDumpWithProcessThreadData |
            NativeProbe.MiniDumpWithIndirectlyReferencedMemory |
            NativeProbe.MiniDumpIgnoreInaccessibleMemory;   // 坏内存跳过，别让整次 dump 失败

        // ===================== 状态 =====================

        private static bool _installed;
        private static uint _pid;
        private static string _procName = "proc";
        private static string _exeName = "WINWORD.EXE";
        private static bool _dumpAllAv;             // settings.crashDumpAllAv（Install 时读一次）
        private static IntPtr _vehHandle;

        // delegate 必须静态持有（防 GC 回收原生回调）
        private static NativeProbe.VectoredExceptionHandler _vehDelegate;
        private static NativeProbe.UnhandledExceptionFilter _uefDelegate;
        private static readonly NativeProbe.ReadProcessMemoryRoutine64 ReadMemRoutine = NativeProbe.ReadMemoryCallback;
        private static readonly NativeProbe.FunctionTableAccessRoutine64 FuncTableRoutine = NativeProbe.FunctionTableCallback;
        private static readonly NativeProbe.GetModuleBaseRoutine64 ModuleBaseRoutine = NativeProbe.ModuleBaseCallback;

        // ---- 预分配的请求槽（崩溃路径不允许 new） ----
        private sealed class DumpRequest
        {
            public IntPtr ExceptionInfo;     // EXCEPTION_POINTERS*（出事线程还在等我们，指针一直有效）
            public ulong FaultAddress;       // ExceptionAddress：出错**指令**地址（= WER 的"错误偏移"基准）
            public ulong AccessedAddress;    // ExceptionInformation[1]：AV 真正读/写的**内存**地址
            public uint ExceptionCode;
            public ulong AccessKind;
            public uint ThreadId;
            public string Stage = "";        // 只在健康线程赋常量字符串
            public bool LooksLikeHandledProbe;
            public bool SkipDump;            // 只记现场不写 dump
        }

        private static readonly DumpRequest Request = new DumpRequest();
        private static int _requestBusy;                 // 0/1：同时只处理一次请求（dump 期间的新异常一律丢弃）
        private static int _inHandler;                   // VEH 重入保护
        private static readonly AutoResetEvent DumpSignal = new AutoResetEvent(false);
        private static readonly AutoResetEvent DumpDone = new AutoResetEvent(false);

        /// <summary>当前 dump 文件路径的原生副本（出事线程读文件大小用，零托管分配）。
        /// 故意**不释放**：出事线程可能正在读它，FreeHGlobal 会造成 use-after-free。
        /// 每次崩溃最多泄漏几百字节，进程反正要退出了。</summary>
        private static IntPtr _dumpPathNative = IntPtr.Zero;

        // ---- 去重/限流（无锁、无分配） ----
        private static int _dumpCount;
        private static readonly ulong[] DumpAddresses = new ulong[8];

        // ---- WebView2 进程账本（环形缓冲，最新 64 条） ----
        private const int JournalCapacity = 64;
        private static readonly string[] Journal = new string[JournalCapacity];
        private static int _journalNext;      // 全局序号（Interlocked）
        private static int _webViewSeen;
        private static int _webViewExited;
        private static int _watcherStop;

        // ===================== 安装 =====================

        internal static void Install()
        {
            if (_installed) return;
            _installed = true;
            try
            {
                _pid = NativeProbe.GetCurrentProcessId();
                _procName = SafeProcessName();
                _exeName = SafeExeName();
                _dumpAllAv = Settings.GetBool("crashDumpAllAv", false);

                CrashLog.EnsureDirs();

                // 预加载 dbghelp：崩溃路径里再 LoadLibrary 会去抢加载器锁，
                // 而 AV 很可能就发生在持锁线程上 → 自死锁。这里先加载好。
                try { NativeProbe.LoadLibraryW("dbghelp.dll"); } catch { }

                WriteSessionHeader();

                // dump 线程必须比任何 VEH 更早存在，否则第一条异常就没有健康线程可用
                StartDumpThread();

                // 主通道：first=1（插到队首），不可被后装者覆盖
                _vehDelegate = VectoredHandler;
                _vehHandle = NativeProbe.AddVectoredExceptionHandler(
                    1, Marshal.GetFunctionPointerForDelegate(_vehDelegate));

                // 兜底通道：保留，但语义已修正（返回 CONTINUE_SEARCH，不吞异常）
                _uefDelegate = UnhandledFilter;
                NativeProbe.SetUnhandledExceptionFilter(Marshal.GetFunctionPointerForDelegate(_uefDelegate));

                AppDomain.CurrentDomain.UnhandledException += OnManagedUnhandled;
                AppDomain.CurrentDomain.ProcessExit += OnProcessExit;

                CrashLog.Timeline("CrashGuard 已安装：VEH(first)=" + CrashLog.Hex((ulong)_vehHandle.ToInt64()) +
                                  " + UEF 兜底 + CLR + 全内存 dump(0x" + DumpFlags.ToString("X") + ")" +
                                  " + 看门狗 + WebView2 监视；crashDumpAllAv=" + _dumpAllAv);

                StartWebView2Watcher();
                StartWatchdog();
            }
            catch (Exception e)
            {
                CrashLog.Log("CrashGuard.Install 异常（忽略）: " + e.GetType().Name + ": " + e.Message);
            }
        }

        /// <summary>会话头写进 crash.log：每次崩溃都能对上「这是哪一次 Word 会话」。</summary>
        /// <summary>
        /// 【测试钩子，生产路径无调用】把 VEH 处理函数直接暴露给同程序集的测试程序，
        /// 用**合成的** EXCEPTION_POINTERS 走一遍完整取证链路（写现场 + full-memory dump），
        /// 但不会让进程真的崩掉 → 可以断言 crash.log / dump 的内容。
        /// 为什么要留这个钩子：VEH 的参数语义（EXCEPTION_POINTERS* 而不是 EXCEPTION_RECORD*）
        /// 曾经写错并**静默失效**（异常码读成垃圾值，每次都在第一道过滤返回），
        /// 只有"喂一个真结构进去"才能回归住。真正的 OS 调用链路由独立测试程序的真 AV 覆盖。
        /// </summary>
        internal static int InvokeVectoredHandlerForTest(IntPtr exceptionPointers)
        {
            return VectoredHandler(exceptionPointers);
        }

        private static void WriteSessionHeader()
        {
            try
            {
                var sb = new StringBuilder();
                sb.AppendLine("===== 会话开始 " + CrashLog.TimestampMillis() + "  pid=" + _pid +
                              "  proc=" + _procName + " =====");
                sb.AppendLine("exe         = " + SafeMainModulePath());
                sb.AppendLine("addin       = " + SafeAssemblyPath() + "  v" + SafeAssemblyVersion());
                sb.AppendLine("clr         = " + Environment.Version + "   os = " + Environment.OSVersion +
                              "   x64 = " + Environment.Is64BitProcess);
                sb.AppendLine("crashesDir  = " + CrashLog.CrashesDir);
                sb.AppendLine("dumpFlags   = 0x" + DumpFlags.ToString("X") +
                              " (含 MiniDumpWithFullMemory=0x2 → 线程栈内存会被抓进 dump)");
                sb.AppendLine("werDisabled = 1（本机 WER 全局禁用，所以必须进程内自抓）");
                sb.AppendLine("heartbeat   = [hb] 行 = 本地毫秒时间戳 + t+自会话开始毫秒 + tag；");
                sb.AppendLine("              前端 JS 探针可用 Date.now() 与本行时间戳对齐（判定崩在渲染中/渲染后）");
                sb.AppendLine("waitPolicy  = 出事线程最多等 " + SoftWaitMs + "ms；若 dump 文件仍在增长则续等，" +
                              "绝对上限 " + HardWaitCapMs + "ms（绝不无限阻塞宿主）");
                CrashLog.AppendDurable(sb.ToString().TrimEnd());
            }
            catch { }
        }

        // ===================== 通道 1：VEH（主） =====================

        /// <summary>
        /// 向量化异常处理器：**first-chance 也会被调用**。
        /// 这里只做三件事：过滤异常码 → 抄写上下文到预分配槽 → 通知 dump 线程并限时等待。
        /// 不做任何托管分配、不拼字符串、不开文件（见类注释的纪律）。
        /// </summary>
        private static int VectoredHandler(IntPtr exceptionInfo)
        {
            bool waiting = false;
            bool dumpExpected = false;
            try
            {
                // 重入保护：如果异常是我们自己（dump 线程/处理器）触发的，直接放行，
                // 否则会无限递归 —— 这正是旧代码没有、而 VEH 必须有的保护。
                if (Interlocked.CompareExchange(ref _inHandler, 1, 0) != 0)
                    return NativeProbe.EXCEPTION_CONTINUE_SEARCH;

                try
                {
                    if (exceptionInfo == IntPtr.Zero) return NativeProbe.EXCEPTION_CONTINUE_SEARCH;

                    // ★★ 这里踩过一个致命坑（实测才发现的）：
                    //   VEH 收到的是 **EXCEPTION_POINTERS***（一个装着 ExceptionRecord 和
                    //   ContextRecord 两个指针的容器），**不是** EXCEPTION_RECORD*。
                    //   第一版直接把 exceptionInfo 当 EXCEPTION_RECORD 读，于是
                    //   er.ExceptionCode 读到的是「ExceptionRecord 指针的低 32 位」，
                    //   永远不等于 0xC0000005 → VEH 每次都在第一道过滤就返回，
                    //   表现和 14:57 那次一模一样：Word 闪退、crashes 目录空空如也。
                    //   靠单元测试（偏移/尺寸）是发现不了的，必须真的造一个 AV 出来打。
                    var ep = Marshal.PtrToStructure<NativeProbe.ExceptionPointers>(exceptionInfo);
                    if (ep.ExceptionRecord == IntPtr.Zero) return NativeProbe.EXCEPTION_CONTINUE_SEARCH;
                    var er = Marshal.PtrToStructure<NativeProbe.ExceptionRecord>(ep.ExceptionRecord);
                    if (er.ExceptionCode != NativeProbe.STATUS_ACCESS_VIOLATION &&
                        er.ExceptionCode != NativeProbe.STATUS_STACK_OVERFLOW)
                        return NativeProbe.EXCEPTION_CONTINUE_SEARCH;

                    ulong kind = er.Info0 & 0xFF;   // AV: 0=读 1=写 8=执行
                    bool isAv = er.ExceptionCode == NativeProbe.STATUS_ACCESS_VIOLATION;
                    if (isAv && kind != 0 && kind != 1 && kind != 8)
                        return NativeProbe.EXCEPTION_CONTINUE_SEARCH;

                    ulong fault = unchecked((ulong)er.ExceptionAddress.ToInt64());

                    // 同一地址只抓一次 + 每进程上限 → 防 dump 风暴
                    if (IsDuplicateOrOverLimit(fault)) return NativeProbe.EXCEPTION_CONTINUE_SEARCH;

                    // ★ dump 线程自身可重入保护：dump 期间又来的 AV 一律丢弃
                    //   （_requestBusy 在 dump 开始前就置 1，dump 全部结束才清 0）
                    if (Interlocked.CompareExchange(ref _requestBusy, 1, 0) != 0)
                        return NativeProbe.EXCEPTION_CONTINUE_SEARCH;

                    bool probe = false;
                    // ★ 关键：判断"是不是宿主自己会处理的探测"，要看 AV **真正访问的那个地址**
                    //   （ExceptionInformation[1]），不是出错指令地址（ExceptionAddress）。
                    //   实测：Marshal.ReadInt32(PAGE_NOACCESS 页) 时 ExceptionAddress 落在 clr.dll
                    //   的读桩上（怎么看都像"正常代码"，会被误判成真崩溃），
                    //   而 Info1 才是那个 PAGE_NOACCESS 页。写/执行类 AV 不看 Info1。
                    try { probe = isAv && LooksLikeHandledProbe(er.Info1, kind); } catch { }

                    Request.ExceptionInfo = exceptionInfo;
                    Request.FaultAddress = fault;
                    Request.AccessedAddress = er.Info1;
                    Request.ExceptionCode = er.ExceptionCode;
                    Request.AccessKind = kind;
                    Request.ThreadId = NativeProbe.GetCurrentThreadId();
                    Request.Stage = "VEH(first-chance)";
                    Request.LooksLikeHandledProbe = probe;
                    Request.SkipDump = probe;
                    Thread.MemoryBarrier();
                    DumpSignal.Set();
                    waiting = true;
                    dumpExpected = !probe;
                }
                finally
                {
                    Volatile.Write(ref _inHandler, 0);
                }
            }
            catch
            {
                // 取证代码自己出问题也绝不能改变宿主行为
            }

            if (waiting) WaitForDumpCompletion(dumpExpected);

            // ★ 绝不吞异常：继续搜索其它处理器，让 Word 按原有路径退出/恢复。
            return NativeProbe.EXCEPTION_CONTINUE_SEARCH;
        }

        /// <summary>
        /// 出事线程的等待策略（**绝不无限阻塞宿主**）：
        ///   第一段：无条件最多等 3 秒（lead 的硬要求：3 秒没写完就放行）。
        ///   第二段：只有当 dump 文件**确实在增长**（说明不是死锁，是 I/O 慢）时才每 1 秒续一次，
        ///          并且绝对上限 30 秒。
        /// 为什么要第二段：Word 一个 full-memory dump 有几亿字节，dbghelp 写 3 秒常常只写一半；
        /// 一到 3 秒就放行会让进程立刻终止、dump 被截断 —— 那等于白抓（这次崩溃就这一次机会）。
        /// 只在「有进展」时延长，所以真死锁（文件不再变大）时仍然是 3 秒放行，不会卡死 Word。
        /// </summary>
        private static void WaitForDumpCompletion(bool dumpExpected)
        {
            try
            {
                if (DumpDone.WaitOne(SoftWaitMs)) return;
                if (!dumpExpected) return;

                long last = ReadCurrentDumpSize();
                int waited = SoftWaitMs;
                while (waited < HardWaitCapMs)
                {
                    if (DumpDone.WaitOne(WaitSliceMs)) return;
                    waited += WaitSliceMs;
                    long cur = ReadCurrentDumpSize();
                    if (cur <= last) return;      // 停止增长 = 卡住/失败 → 立即放行
                    last = cur;
                }
            }
            catch { }
        }

        /// <summary>零托管分配地读当前 dump 文件大小（出事线程专用）。</summary>
        private static long ReadCurrentDumpSize()
        {
            try
            {
                IntPtr p = _dumpPathNative;
                if (p == IntPtr.Zero) return 0;
                NativeProbe.Win32FileAttributeData data;
                if (!NativeProbe.GetFileAttributesExW(p, 0, out data)) return 0;
                return ((long)data.FileSizeHigh << 32) | (uint)data.FileSizeLow;
            }
            catch { return 0; }
        }

        /// <summary>
        /// 探测型 AV 识别：典型 __try{读不可访问页}__except 探测（IsBadReadPtr 那一类），
        /// 宿主会自己接住并继续跑。这种 AV 如果也写 full-memory dump，
        /// 会在**正常使用中**白写几百 MB 并把 Word 卡住几秒。
        /// 判据（accessedAddress 必须是 ExceptionInformation[1]，即真正被访问的地址）：
        ///   读操作 + 目标是已提交页但 PAGE_NOACCESS/PAGE_GUARD。
        /// 拿不准（查不到区域/地址很野）时一律按真崩溃处理 —— 宁可多抓，不可漏抓。
        /// 设置里 crashDumpAllAv=true 可以关掉这个过滤，全抓。
        /// </summary>
        private static bool LooksLikeHandledProbe(ulong accessedAddress, ulong kind)
        {
            if (_dumpAllAv) return false;
            if (kind != 0) return false;            // 写失败/执行失败不是探测
            if (accessedAddress < 0x10000) return false;   // 空指针解引用：真崩溃
            NativeProbe.MemoryBasicInformation mbi;
            if (!NativeProbe.TryQueryRegion(accessedAddress, out mbi)) return false;
            if (mbi.State == NativeProbe.MEM_COMMIT &&
                (mbi.Protect & (NativeProbe.PAGE_NOACCESS | NativeProbe.PAGE_GUARD)) != 0)
                return true;
            return false;
        }

        private static bool IsDuplicateOrOverLimit(ulong fault)
        {
            int n = Volatile.Read(ref _dumpCount);
            if (n >= MaxDumpsPerProcess) return true;
            int cap = n < DumpAddresses.Length ? n : DumpAddresses.Length;
            for (int i = 0; i < cap; i++)
                if (DumpAddresses[i] == fault) return true;
            int slot = Interlocked.Increment(ref _dumpCount) - 1;
            if (slot < DumpAddresses.Length) DumpAddresses[slot] = fault;
            return false;
        }

        // ===================== 通道 2：SetUnhandledExceptionFilter（兜底） =====================

        /// <summary>
        /// ★ 语义修正（旧代码这里返回 1，是错的）：
        ///   EXCEPTION_EXECUTE_HANDLER(1) == 「我已经处理了，别再往下传」——
        ///   它会让 WER / 宿主自己的崩溃处理完全看不到现场，等于我们自己把
        ///   终止流程接管了；而文档要求这个回调返回 EXCEPTION_CONTINUE_SEARCH(0)
        ///   才是「看过了，交给系统继续走原来的路径」。
        ///   v2 里 UEF 只是兜底（VEH 才是主通道），所以照样返回 0。
        /// </summary>
        private static int UnhandledFilter(IntPtr exceptionInfo)
        {
            bool waiting = false;
            try
            {
                if (exceptionInfo == IntPtr.Zero) return NativeProbe.EXCEPTION_CONTINUE_SEARCH;
                // 同 VEH：UEF 收到的也是 EXCEPTION_POINTERS*，不是 EXCEPTION_RECORD*
                var ep = Marshal.PtrToStructure<NativeProbe.ExceptionPointers>(exceptionInfo);
                if (ep.ExceptionRecord == IntPtr.Zero) return NativeProbe.EXCEPTION_CONTINUE_SEARCH;
                var er = Marshal.PtrToStructure<NativeProbe.ExceptionRecord>(ep.ExceptionRecord);
                ulong kind = er.Info0 & 0xFF;
                bool interesting = (er.ExceptionCode == NativeProbe.STATUS_ACCESS_VIOLATION &&
                                    (kind == 0 || kind == 1 || kind == 8))
                                   || er.ExceptionCode == NativeProbe.STATUS_STACK_OVERFLOW;

                if (!interesting)
                {
                    // 非目标异常：只记一行（这里已经是健康状态，允许分配字符串）
                    CrashLog.Log("UEF 兜底：非目标异常 code=" + CrashLog.Hex(er.ExceptionCode) +
                                 " addr=" + CrashLog.Hex(unchecked((ulong)er.ExceptionAddress.ToInt64())));
                    return NativeProbe.EXCEPTION_CONTINUE_SEARCH;
                }

                ulong fault = unchecked((ulong)er.ExceptionAddress.ToInt64());
                if (IsDuplicateOrOverLimit(fault)) return NativeProbe.EXCEPTION_CONTINUE_SEARCH;
                if (Interlocked.CompareExchange(ref _requestBusy, 1, 0) != 0)
                    return NativeProbe.EXCEPTION_CONTINUE_SEARCH;

                Request.ExceptionInfo = exceptionInfo;
                Request.FaultAddress = fault;
                Request.AccessedAddress = er.Info1;
                Request.ExceptionCode = er.ExceptionCode;
                Request.AccessKind = kind;
                Request.ThreadId = NativeProbe.GetCurrentThreadId();
                Request.Stage = "UEF(second-chance)";
                Request.LooksLikeHandledProbe = false;
                Request.SkipDump = false;      // 走到这里已经是「没人处理」的第二次机会，必须写 dump
                Thread.MemoryBarrier();
                DumpSignal.Set();
                waiting = true;
            }
            catch { }
            if (waiting) WaitForDumpCompletion(true);
            return NativeProbe.EXCEPTION_CONTINUE_SEARCH;
        }

        // ===================== 通道 3：托管未处理异常 =====================

        private static void OnManagedUnhandled(object sender, UnhandledExceptionEventArgs e)
        {
            try
            {
                var ex = e.ExceptionObject as Exception;
                CrashLog.AppendDurable(
                    "===== CLR 未处理异常 " + CrashLog.TimestampMillis() +
                    "  terminating=" + e.IsTerminating + "  pid=" + _pid + " =====");
                CrashLog.AppendDurable(ex != null ? ex.ToString() : Convert.ToString(e.ExceptionObject));
                CrashLog.AppendDurable("===== CLR 未处理异常结束 =====");

                // 只有确定要终止时才写 dump（非终止的托管异常进程还会活着，写几百 MB 没意义）
                if (!e.IsTerminating) return;
                if (Interlocked.CompareExchange(ref _requestBusy, 1, 0) != 0) return;
                Request.ExceptionInfo = IntPtr.Zero;         // 没有原生异常记录，dump 照样有效
                Request.FaultAddress = 0;
                Request.AccessedAddress = 0;
                Request.ExceptionCode = 0xE0434352;          // CLR 异常包装码
                Request.AccessKind = 0;
                Request.ThreadId = NativeProbe.GetCurrentThreadId();
                Request.Stage = "CLR-unhandled";
                Request.LooksLikeHandledProbe = false;
                Request.SkipDump = false;
                Thread.MemoryBarrier();
                DumpSignal.Set();
                WaitForDumpCompletion(true);
            }
            catch { }
        }

        private static void OnProcessExit(object sender, EventArgs e)
        {
            // 正常退出（非崩溃）时的最后一条账本，用来证明「这次宿主没有崩」
            NotifyHostShutdown("ProcessExit");
        }

        // ===================== dump 线程（所有重活都在这里） =====================

        private static void StartDumpThread()
        {
            var t = new Thread(DumpThreadLoop)
            {
                IsBackground = true,
                Name = "OfficeAi.CrashDump"
            };
            t.Start();
        }

        private static void DumpThreadLoop()
        {
            while (true)
            {
                try
                {
                    DumpSignal.WaitOne();
                    try
                    {
                        HandleDumpRequest();
                    }
                    catch (Exception e)
                    {
                        CrashLog.AppendDurable("dump 线程异常: " + e.GetType().Name + ": " + e.Message);
                    }
                    finally
                    {
                        Volatile.Write(ref _requestBusy, 0);
                        DumpDone.Set();          // 叫醒等着的出事线程（它可能已经超时走了，Set 无害）
                    }
                }
                catch
                {
                    // 线程绝不能死
                }
            }
        }

        /// <summary>
        /// 崩溃现场处理：先固化「上下文 + 模块归属 + 寄存器」，再读旧日志尾巴 + WebView2 账本，
        /// 再写 full-memory dump，最后做栈扫描/栈回溯。
        /// 顺序是有讲究的（每一步都可能失败）：最要紧的证据最先落盘。
        /// </summary>
        private static void HandleDumpRequest()
        {
            DumpRequest req = Request;
            // CLR 通道没有原生异常记录（ExceptionInfo=0）：整块原生上下文跳过，
            // dump 依然有效（MiniDumpWriteDump 的 exceptionParam 传 0）。
            NativeProbe.ExceptionPointers ep;
            if (req.ExceptionInfo != IntPtr.Zero)
                ep = Marshal.PtrToStructure<NativeProbe.ExceptionPointers>(req.ExceptionInfo);
            else
                ep = default(NativeProbe.ExceptionPointers);
            NativeProbe.ExceptionRecord er;
            if (ep.ExceptionRecord != IntPtr.Zero)
                er = Marshal.PtrToStructure<NativeProbe.ExceptionRecord>(ep.ExceptionRecord);
            else
                er = default(NativeProbe.ExceptionRecord);

            var ctx = new NativeProbe.Context64();
            bool hasCtx = false;
            if (ep.ContextRecord != IntPtr.Zero)
            {
                ctx = Marshal.PtrToStructure<NativeProbe.Context64>(ep.ContextRecord);
                hasCtx = true;
            }

            ulong fault = req.FaultAddress;
            uint code = req.ExceptionCode != 0 ? req.ExceptionCode : er.ExceptionCode;
            ulong kind = req.AccessKind;
            uint tid = req.ThreadId;
            string stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss_fff", CultureInfo.InvariantCulture);

            // ---------- ① 上下文（最重要，先落盘） ----------
            var sb = new StringBuilder();
            sb.AppendLine();
            sb.AppendLine("=====================================================================");
            sb.AppendLine("===== 崩溃现场 " + CrashLog.TimestampMillis() + "  pid=" + _pid +
                          "  tid=" + tid + "  stage=" + req.Stage + " =====");
            sb.AppendLine("=====================================================================");
            sb.AppendLine("exception.code        = " + CrashLog.Hex(code) + "  (" + CrashLog.ExceptionCodeName(code) + ")");
            sb.AppendLine("exception.address     = " + CrashLog.Hex(fault));
            sb.AppendLine("exception.access      = " + CrashLog.AccessKindName(kind) +
                          "   [ExceptionInformation[0]=" + kind + "]");
            sb.AppendLine("exception.info1       = " + CrashLog.Hex(er.Info1) +
                          "   (写操作时=目标地址；执行时=0)");
            sb.AppendLine("exception.flags       = " + CrashLog.Hex(er.ExceptionFlags) +
                          "   numberParameters=" + er.NumberParameters);
            sb.AppendLine("exception.nested      = " + CrashLog.Hex(unchecked((ulong)er.ExceptionRecordPtr.ToInt64())));
            sb.AppendLine("thread.id             = " + tid + "   (faulting thread)");
            sb.AppendLine("process               = " + _procName + ".exe  pid=" + _pid +
                          "   启动于 " + SafeStartTime());
            if (req.LooksLikeHandledProbe && req.SkipDump)
            {
                sb.AppendLine("★ 注意：该 AV 看起来是宿主自己用 SEH 处理掉的「探测型读取」");
                sb.AppendLine("  （已提交页但 PAGE_NOACCESS/PAGE_GUARD）。本次只记现场、不写 dump。");
                sb.AppendLine("  若进程随后继续正常运行，说明判断正确；若进程随后退出，请把");
                sb.AppendLine("  settings.json 里的 crashDumpAllAv 设为 true 后复现一次。");
            }
            sb.Append(FormatModuleAttribution("fault.address", fault));
            if (hasCtx)
                sb.Append(FormatModuleAttribution("rip(出错指令)", ctx.Rip));
            // ★ 被访问的地址：定位"为什么崩"的关键（野指针 / 已释放堆 / PAGE_NOACCESS 探测）
            if (kind == 0 || kind == 1)
            {
                sb.Append("accessed.address     = " + CrashLog.Hex(req.AccessedAddress) +
                          "   (ExceptionInformation[1]：AV 真正读/写的那个内存地址)\n");
                if (req.AccessedAddress != 0)
                {
                    sb.Append(FormatModuleAttribution("accessed.module", req.AccessedAddress));
                    sb.Append(FormatRegion("accessed.region", req.AccessedAddress));
                }
            }
            else
            {
                sb.Append("accessed.address     = " + CrashLog.Hex(req.AccessedAddress) +
                          "   (执行类 AV：ExceptionInformation[1] 无意义)\n");
            }
            sb.Append(FormatFaultRegion(fault));
            if (hasCtx) sb.Append(FormatRegisters(ctx));
            CrashLog.AppendDurable(sb.ToString().TrimEnd());

            // ---------- ② 崩溃前的最后动作（addin.log 尾巴） ----------
            // 独立落盘，确保「最后一个面包屑」和现场在同一个文件里，不用再对时间戳。
            CrashLog.AppendDurable("---- addin.log 末尾 4KB（崩溃前最后动作）----");
            CrashLog.AppendDurable(CrashLog.ReadAddinLogTail(4096));

            // ---------- ③ WebView2 进程账本（谁先死？） ----------
            CrashLog.AppendDurable(JournalSummary());

            // ---------- ④ full-memory dump ----------
            string dumpPath = Path.Combine(CrashLog.CrashesDir,
                _procName + "_" + (code == NativeProbe.STATUS_ACCESS_VIOLATION ? "AV" : "EXC_" + code.ToString("X")) +
                "_" + fault.ToString("X16", CultureInfo.InvariantCulture) +
                "_" + stamp + "_p" + _pid + "_t" + tid + ".dmp");

            if (req.SkipDump)
            {
                CrashLog.AppendDurable("dump.skipped          = 疑似宿主自行处理的探测型 AV（settings.crashDumpAllAv=true 可强制全抓）");
            }
            else
            {
                // 先把路径的原生副本挂出来：出事线程靠它零分配地看 dump 是否在增长
                try { _dumpPathNative = Marshal.StringToHGlobalUni(dumpPath); } catch { }

                CrashLog.AppendDurable("dump.start            = " + dumpPath);
                CrashLog.AppendDurable("dump.flags            = 0x" + DumpFlags.ToString("X"));
                long freeMb = SafeFreeDiskMb(CrashLog.CrashesDir);
                if (freeMb >= 0 && freeMb < 1500)
                    CrashLog.AppendDurable("dump.warning          = 磁盘剩余空间仅 " + freeMb + " MB，dump 可能不完整");

                bool ok = false;
                int err = 0;
                long ms;
                var sw = Stopwatch.StartNew();
                // ★ ExceptionParam 必须是 MINIDUMP_EXCEPTION_INFORMATION*（见 NativeProbe 里的注释），
                //   传错会得到一个 0 字节 dump + GetLastError=0x800700E6。
                // ★ 实测还要再加一道保险：dbghelp 会拿着这个 EXCEPTION_POINTERS 里的 CONTEXT
                //   去核对/展开栈，CONTEXT 一旦不可用（例如栈已损坏、或合成数据），
                //   MiniDumpWriteDump 会立刻失败（实测 GetLastError=0x800703E6 / ERROR_NOACCESS 998，
                //   dump 0 字节）。而**不带** ExceptionParam 时同样的标志组合 100% 成功
                //   （实测：102MB / 80ms）。
                //   所以策略是：先带异常流试一次（分析体验最好），失败就丢掉 ExceptionParam 重试一次。
                //   两次都失败才认输。核心目标只有一个：**这一次崩溃必须留下可分析的现场**。
                IntPtr exParam = IntPtr.Zero;
                IntPtr exParamOwned = IntPtr.Zero;
                int attempt = 0;
                try
                {
                    if (req.ExceptionInfo != IntPtr.Zero)
                    {
                        var mei = new NativeProbe.MiniDumpExceptionInformation
                        {
                            ThreadId = req.ThreadId,
                            ExceptionPointers = req.ExceptionInfo,
                            ClientPointers = 0     // 本进程自抓：指针在调用者地址空间里
                        };
                        exParamOwned = Marshal.AllocHGlobal(Marshal.SizeOf(typeof(NativeProbe.MiniDumpExceptionInformation)));
                        Marshal.StructureToPtr(mei, exParamOwned, false);
                        exParam = exParamOwned;
                    }

                    while (attempt < 2)
                    {
                        attempt++;
                        int attemptErr;
                        ok = TryWriteDumpOnce(dumpPath, DumpFlags, exParam, out attemptErr);
                        if (!ok) err = attemptErr;
                        if (ok) break;
                        CrashLog.AppendDurable("dump.attempt" + attempt + "        = 失败 err=" + err +
                                               " (0x" + ((uint)err).ToString("X8") + ")" +
                                               (attempt == 1 && exParam != IntPtr.Zero
                                                    ? "；改为「不带 ExceptionParam」重试" : "；放弃"));
                        if (exParam == IntPtr.Zero) break;
                        exParam = IntPtr.Zero;      // 第二次：只要内存快照，不要异常流
                    }
                }
                catch (Exception e)
                {
                    err = -1;
                    CrashLog.AppendDurable("dump.exception        = " + e.GetType().Name + ": " + e.Message);
                }
                finally
                {
                    if (exParamOwned != IntPtr.Zero) { try { Marshal.FreeHGlobal(exParamOwned); } catch { } }
                }
                sw.Stop();
                ms = sw.ElapsedMilliseconds;

                long size = SafeFileSize(dumpPath);
                CrashLog.AppendDurable("dump.ok               = " + ok + (ok ? "" : "   GetLastError=" + err));
                CrashLog.AppendDurable("dump.sizeBytes        = " + size + "   (" + (size / 1048576.0).ToString("F1") + " MB)");
                CrashLog.AppendDurable("dump.elapsedMs        = " + ms);
                CrashLog.AppendDurable("dump.path             = " + dumpPath);
                if (ok && size < 200 * 1024)
                    CrashLog.AppendDurable("dump.warning          = 体积异常小（<200KB），full-memory 可能没生效，请检查 dbghelp 版本");
            }

            // ---------- ⑤ 栈（bonus：dump 里也有，但这里直接可读，省掉 cdb 一步） ----------
            // 每一步都单独兜底 + 打时间戳：栈收集只是加分项，绝不允许它把「现场已落盘」
            // 这件事本身搞坏（也不能因为它在 dump 线程上挂住而让进程退不出去）。
            CrashLog.AppendDurable("stack.begin           = " + CrashLog.TimestampMillis() + "  hasCtx=" + hasCtx);
            try
            {
                string walk = null;
                string scan = null;
                // ★ 栈收集只是"加分项"，但它依赖 dbghelp。实测这条路在自己的进程里**可能卡住**
                //   （dump 线程进了 StackWalk64 就不出来；同样参数在独立进程里却秒回）。
                //   所以把它放到独立线程 + 限时 join：超时就放弃，绝不让它拖住 dump 线程/进程退出。
                //   而且 dump 本身是 full-memory，所有线程栈都在里面，分析不受影响。
                var stackThread = new Thread(() =>
                {
                    try { walk = hasCtx ? FormatStackWalk(ctx) : "(没有 CONTEXT，跳过)"; }
                    catch (Exception e) { walk = "StackWalk 异常: " + e.GetType().Name + ": " + e.Message; }
                    try { scan = hasCtx ? FormatStackScan(ctx.Rsp) : ""; }
                    catch (Exception e) { scan = "栈扫描异常: " + e.GetType().Name + ": " + e.Message; }
                });
                stackThread.IsBackground = true;
                stackThread.Name = "OfficeAi.StackWalk";
                stackThread.Start();

                bool finished = stackThread.Join(StackJoinTimeoutMs);
                CrashLog.AppendDurable("stack.join            = " + CrashLog.TimestampMillis() +
                                       "  完成=" + finished +
                                       (finished ? "" : "（超时 " + StackJoinTimeoutMs + "ms，放弃栈文本；dump 内仍有完整线程栈）"));
                if (finished)
                {
                    CrashLog.AppendDurable("---- 栈信息（模块内地址 = 候选返回地址）----");
                    if (!string.IsNullOrEmpty(walk)) CrashLog.AppendDurable(walk.TrimEnd());
                    if (!string.IsNullOrEmpty(scan)) CrashLog.AppendDurable(scan.TrimEnd());
                }
            }
            catch (Exception e)
            {
                CrashLog.AppendDurable("stack 收集失败: " + e.GetType().Name + ": " + e.Message);
            }

            CrashLog.AppendDurable("===== 现场结束 " + CrashLog.TimestampMillis() + " =====");
        }

        /// <summary>写一次 dump（每次都用新文件）。失败时把 Win32 错误码带出来，
        /// 好让上层决定要不要丢开 ExceptionParam 重试。</summary>
        private static bool TryWriteDumpOnce(string path, uint flags, IntPtr exParam, out int error)
        {
            error = 0;
            try
            {
                using (var fs = new FileStream(path, FileMode.Create, FileAccess.Write,
                                               FileShare.None, 1, FileOptions.None))
                {
                    bool ok = NativeProbe.MiniDumpWriteDump(NativeProbe.GetCurrentProcess(), _pid,
                        fs.SafeFileHandle.DangerousGetHandle(), flags, exParam, IntPtr.Zero, IntPtr.Zero);
                    if (!ok) error = Marshal.GetLastWin32Error();
                    return ok;
                }
            }
            catch (Exception e)
            {
                error = -1;
                CrashLog.AppendDurable("dump.exception        = " + e.GetType().Name + ": " + e.Message);
                return false;
            }
        }

        // ===================== 现场格式化 =====================
        /// <summary>把一个地址翻译成「模块全路径 + 基址 + RVA」，也就是把 0x183ba19 变成
        /// wwlib.dll+0x183ba19 的那一步。查不到（JIT/堆/匿名内存）就写 unknown。</summary>
        private static string FormatModuleAttribution(string label, ulong address)
        {
            var nameBuf = new char[1024];
            int len;
            ulong moduleBase, rva;
            if (NativeProbe.TryResolveModule(address, nameBuf, out len, out moduleBase, out rva))
            {
                return label.PadRight(21) + "= " + CrashLog.Hex(address) + "\n" +
                       "".PadRight(21) + "  module = " + NativeProbe.FullName(nameBuf, len) + "\n" +
                       "".PadRight(21) + "  base   = " + CrashLog.Hex(moduleBase) + "\n" +
                       "".PadRight(21) + "  RVA    = " + CrashLog.Hex(rva) +
                       "   →   " + CrashLog.ModuleOffset(nameBuf, len, rva) + "\n";
            }
            return label.PadRight(21) + "= " + CrashLog.Hex(address) +
                   "  module = unknown（不在任何模块内：JIT/堆/匿名内存/已卸载模块）\n";
        }

        private static string FormatFaultRegion(ulong fault)
        {
            return FormatRegion("fault.region", fault);
        }

        /// <summary>把一个地址所在内存区域（state/type/protect）打印出来。前缀可换，
        /// 因为「出错指令地址」和「被访问地址」都要看区域属性。</summary>
        private static string FormatRegion(string prefix, ulong fault)
        {
            NativeProbe.MemoryBasicInformation mbi;
            if (!NativeProbe.TryQueryRegion(fault, out mbi))
                return prefix + ".          = VirtualQuery 失败（地址完全无效）\n";
            return prefix + ".base     = " + CrashLog.Hex(unchecked((ulong)mbi.BaseAddress.ToInt64())) + "\n" +
                   prefix + ".size     = " + CrashLog.Hex(unchecked((ulong)mbi.RegionSize.ToInt64())) + "\n" +
                   prefix + ".state    = " + CrashLog.StateName(mbi.State) + "\n" +
                   prefix + ".type     = " + CrashLog.TypeName(mbi.Type) + "\n" +
                   prefix + ".protect  = " + CrashLog.ProtectName(mbi.Protect) + "\n" +
                   prefix + ".alloc    = " + CrashLog.Hex(unchecked((ulong)mbi.AllocationBase.ToInt64())) +
                   "  allocProtect=" + CrashLog.ProtectName(mbi.AllocationProtect) + "\n";
        }

        private static string FormatRegisters(NativeProbe.Context64 c)
        {
            var sb = new StringBuilder();
            sb.AppendLine("---- 寄存器（x64 CONTEXT）----");
            sb.AppendLine("  rip=" + CrashLog.Hex(c.Rip) + "  rsp=" + CrashLog.Hex(c.Rsp) + "  rbp=" + CrashLog.Hex(c.Rbp));
            sb.AppendLine("  rax=" + CrashLog.Hex(c.Rax) + "  rbx=" + CrashLog.Hex(c.Rbx) +
                          "  rcx=" + CrashLog.Hex(c.Rcx) + "  rdx=" + CrashLog.Hex(c.Rdx));
            sb.AppendLine("  rsi=" + CrashLog.Hex(c.Rsi) + "  rdi=" + CrashLog.Hex(c.Rdi) +
                          "  r8 =" + CrashLog.Hex(c.R8) + "  r9 =" + CrashLog.Hex(c.R9));
            sb.AppendLine("  r10=" + CrashLog.Hex(c.R10) + "  r11=" + CrashLog.Hex(c.R11) +
                          "  r12=" + CrashLog.Hex(c.R12) + "  r13=" + CrashLog.Hex(c.R13));
            sb.AppendLine("  r14=" + CrashLog.Hex(c.R14) + "  r15=" + CrashLog.Hex(c.R15));
            sb.AppendLine("  eflags=0x" + c.EFlags.ToString("X") + "  cs=0x" + c.SegCs.ToString("X") +
                          "  ss=0x" + c.SegSs.ToString("X") + "  mxcsr=0x" + c.MxCsr.ToString("X"));
            sb.AppendLine("  dr0=" + CrashLog.Hex(c.Dr0) + "  dr1=" + CrashLog.Hex(c.Dr1) +
                          "  dr2=" + CrashLog.Hex(c.Dr2) + "  dr3=" + CrashLog.Hex(c.Dr3) +
                          "  dr6=" + CrashLog.Hex(c.Dr6) + "  dr7=" + CrashLog.Hex(c.Dr7));
            return sb.ToString();
        }

        /// <summary>用 dbghelp!StackWalk64 走一遍调用栈。不需要符号（不联网、不下载 PDB），
        /// 只要「模块+偏移」——这足够定位 wwlib.dll+0x183ba19 这种签名。</summary>
        private static string FormatStackWalk(NativeProbe.Context64 ctx)
        {
            var sb = new StringBuilder();
            sb.AppendLine("---- 调用栈（StackWalk64，无符号，模块+RVA）----");
            var frame = new NativeProbe.StackFrame64();
            frame.AddrPC = ctx.Rip;
            frame.AddrStack = ctx.Rsp;
            frame.AddrFrame = ctx.Rbp;
            var ctxCopy = ctx;
            var nameBuf = new char[1024];
            int shown = 0;
            for (int i = 0; i < 48; i++)
            {
                bool ok;
                try
                {
                    ok = NativeProbe.StackWalk64(NativeProbe.IMAGE_FILE_MACHINE_AMD64,
                        NativeProbe.GetCurrentProcess(), IntPtr.Zero, ref frame, ref ctxCopy,
                        ReadMemRoutine, FuncTableRoutine, ModuleBaseRoutine, IntPtr.Zero);
                }
                catch { break; }
                if (!ok || frame.AddrPC == 0) break;

                int len;
                ulong moduleBase, rva;
                string text;
                if (NativeProbe.TryResolveModule(frame.AddrPC, nameBuf, out len, out moduleBase, out rva))
                    text = NativeProbe.FullName(nameBuf, len) + " + " + CrashLog.Hex(rva) +
                           "   (base " + CrashLog.Hex(moduleBase) + ")";
                else
                    text = "(不在模块内：JIT/匿名内存)";
                sb.Append("  #").Append(i.ToString("D2")).Append("  ")
                  .Append(CrashLog.Hex(frame.AddrPC)).Append("  ").AppendLine(text);
                shown++;
            }
            if (shown == 0) sb.AppendLine("  (StackWalk64 未能展开，改用下面的栈扫描)");
            return sb.ToString();
        }

        /// <summary>栈扫描（等价 WinDbg 的 dps @rsp）：把栈上落在模块里的值列出来。
        /// StackWalk64 因为栈损坏失败时，这一招仍能给出「崩在谁的代码附近」。</summary>
        private static string FormatStackScan(ulong rsp)
        {
            var sb = new StringBuilder();
            sb.AppendLine("---- 栈扫描 dps @rsp（找到 " + CrashLog.Hex(rsp) + "）----");
            if (rsp == 0) return sb.ToString();
            var hProcess = NativeProbe.GetCurrentProcess();
            var buf = new byte[8];
            var nameBuf = new char[1024];
            int shown = 0;
            for (int i = 0; i < 512 && shown < 48; i++)
            {
                ulong addr = rsp + (ulong)(i * 8);
                IntPtr read;
                try
                {
                    if (!NativeProbe.ReadProcessMemory(hProcess, new IntPtr(unchecked((long)addr)),
                            buf, new IntPtr(8), out read) || read.ToInt64() != 8)
                        continue;
                }
                catch { continue; }
                ulong v = BitConverter.ToUInt64(buf, 0);
                if (v < 0x10000) continue;
                int len;
                ulong moduleBase, rva;
                if (!NativeProbe.TryResolveModule(v, nameBuf, out len, out moduleBase, out rva)) continue;
                sb.Append("  rsp+0x").Append((i * 8).ToString("X3", CultureInfo.InvariantCulture))
                  .Append("  ").Append(CrashLog.Hex(v)).Append("  ")
                  .Append(NativeProbe.ShortName(nameBuf, len)).Append("+0x")
                  .Append(rva.ToString("X", CultureInfo.InvariantCulture).ToLowerInvariant())
                  .AppendLine();
                shown++;
            }
            if (shown == 0) sb.AppendLine("  (栈上没找到模块内地址)");
            return sb.ToString();
        }

        // ===================== 进程外看门狗（T3 提供 exe，这里只负责拉起） =====================

        /// <summary>
        /// 拉起进程外看门狗：宿主整个进程都没了（fail-fast / 被强杀 / dump 写不出来）时，
        /// 还能靠它留下记录。best-effort：找不到 exe、启动失败都只记日志，绝不影响加载项。
        /// ★ 绝不 WaitForExit：看门狗要活到宿主死去之后。
        /// </summary>
        internal static void StartWatchdog()
        {
            try
            {
                if (!Settings.GetBool("watchdog", false))
                {
                    CrashLog.Timeline("看门狗：settings.watchdog=false，不拉起");
                    return;
                }
                string root = InstallRoot();
                if (string.IsNullOrEmpty(root))
                {
                    CrashLog.Timeline("看门狗：拿不到 InstallRoot，跳过");
                    return;
                }
                string exe = Path.Combine(root, "OfficeAi.Watchdog.exe");
                if (!File.Exists(exe)) exe = Path.Combine(root, "tools", "OfficeAi.Watchdog.exe");
                if (!File.Exists(exe))
                {
                    CrashLog.Timeline("看门狗：未找到 OfficeAi.Watchdog.exe（查过 <root>\\ 与 <root>\\tools\\），跳过");
                    return;
                }

                var psi = new ProcessStartInfo
                {
                    FileName = exe,
                    Arguments = "--pid " + _pid + " --out \"" + CrashLog.CrashesDir + "\"",
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    WorkingDirectory = root
                };
                // using 保证句柄释放：Process 对象持有子进程句柄，不浪费也不干扰子进程运行
                using (Process p = Process.Start(psi))
                {
                    CrashLog.Timeline("看门狗已拉起 exe=" + exe + " childPid=" +
                                      (p != null ? p.Id.ToString(CultureInfo.InvariantCulture) : "?"));
                }
            }
            catch (Exception e)
            {
                CrashLog.Log("看门狗拉起失败（忽略，不影响加载项）: " + e.GetType().Name + ": " + e.Message);
            }
        }

        private static string InstallRoot()
        {
            // 不调 TaskPaneHost.InstallRoot()：取证层必须能在 UI 类之前独立工作（也可独立测试），
            // 结果完全一致（都是本程序集所在目录）。
            try { return Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location); }
            catch { return ""; }
        }

        // ===================== WebView2 进程账本（谁先死？） =====================

        private sealed class WebView2Child
        {
            public Process Proc;
            public uint Pid;
            public string Type = "?";
            public DateTime Start;
        }

        /// <summary>
        /// 监视本进程拉起来的 msedgewebview2 子进程的启动/退出（含退出码、类型、存活时长）。
        /// 为什么是第一顺位证据：崩溃签名固定是 wwlib@0x183ba19，必须先回答
        /// 「是 WebView2 渲染进程先死（浏览器侧），还是 Word 宿主先死（宿主侧）」。
        /// 判据（两条都用，因为渲染进程的父 PID 往往指向 WebView2 自己的代理进程而不是 WINWORD）：
        ///   a) 父 PID == 本次宿主 pid；
        ///   b) 命令行含 `--webview-exe-name=<宿主 exe>` 或 `--user-data-dir=...\<宿主 exe>.WebView2\...`
        /// 性能纪律：只用 Process API（**不用 WMI/CIM**，那会拉 WMI 服务且有重入风险）；
        /// 循环整体 try/catch，任何一次失败只记日志、绝不停止；宿主退出时立即结束线程。
        /// </summary>
        private static void StartWebView2Watcher()
        {
            var t = new Thread(WebView2WatchLoop)
            {
                IsBackground = true,
                Name = "OfficeAi.WebView2Watch"
            };
            t.Start();
            CrashLog.Timeline("WebView2 子进程监视已启动（2s 轮询，只看 pid=" + _pid +
                              " 的子进程 / --webview-exe-name=" + _exeName + "）");
        }

        private static void WebView2WatchLoop()
        {
            var known = new Dictionary<uint, WebView2Child>();
            var foreign = new HashSet<uint>();
            while (Volatile.Read(ref _watcherStop) == 0)
            {
                try { PollWebView2Children(known, foreign); }
                catch (Exception e)
                {
                    // 绝不因为一次失败停止监视
                    CrashLog.Log("WebView2 监视异常（继续）: " + e.GetType().Name + ": " + e.Message);
                }
                try { Thread.Sleep(2000); } catch { }
            }
            CrashLog.Timeline("WebView2 监视线程结束（宿主退出通知已收到）");
        }

        private static void PollWebView2Children(Dictionary<uint, WebView2Child> known, HashSet<uint> foreign)
        {
            Process[] all = Process.GetProcessesByName("msedgewebview2");
            foreach (Process p in all)
            {
                bool keep = false;
                try
                {
                    uint childPid = unchecked((uint)p.Id);
                    if (childPid == 0 || known.ContainsKey(childPid) || foreign.Contains(childPid))
                    {
                        keep = false;
                    }
                    else if (known.Count >= 64)
                    {
                        keep = false;                       // 上限保护：异常情况下不无限增长
                    }
                    else
                    {
                        string cmd = null;
                        try { NativeProbe.TryGetCommandLine(p.Handle, out cmd); } catch { }
                        uint parent = 0;
                        try { parent = NativeProbe.TryGetParentProcessId(p.Handle); } catch { }

                        bool ours = parent == _pid;
                        if (!ours && !string.IsNullOrEmpty(cmd))
                        {
                            if (cmd.IndexOf("--webview-exe-name=" + _exeName, StringComparison.OrdinalIgnoreCase) >= 0)
                                ours = true;
                            else if (cmd.IndexOf(_exeName + ".WebView2", StringComparison.OrdinalIgnoreCase) >= 0)
                                ours = true;
                        }

                        if (!ours)
                        {
                            foreign.Add(childPid);          // 别的程序（SearchHost/tiez-app…）的 WebView2
                        }
                        else
                        {
                            DateTime start;
                            try { start = p.StartTime; } catch { start = DateTime.Now; }
                            known[childPid] = new WebView2Child
                            {
                                Proc = p,
                                Pid = childPid,
                                Type = ExtractArg(cmd, "--type="),
                                Start = start
                            };
                            keep = true;
                            Interlocked.Increment(ref _webViewSeen);
                            string line = "T+" + UptimeSeconds() + "s  WebView2 子进程启动 pid=" + childPid +
                                          " type=" + known[childPid].Type +
                                          " parent=" + parent +
                                          " start=" + start.ToString("HH:mm:ss.fff", CultureInfo.InvariantCulture);
                            CrashLog.Timeline(line);
                            JournalAdd(line);
                        }
                    }
                }
                catch { }
                finally
                {
                    if (!keep) { try { p.Dispose(); } catch { } }
                }
            }

            List<uint> gone = null;
            foreach (var kv in known)
            {
                bool exited;
                try { exited = kv.Value.Proc.HasExited; }
                catch { exited = true; }
                if (!exited) continue;
                int exitCode = -1;
                try { exitCode = kv.Value.Proc.ExitCode; } catch { }
                if (gone == null) gone = new List<uint>();
                gone.Add(kv.Key);

                double lifeSec = (DateTime.Now - kv.Value.Start).TotalSeconds;
                uint uc = unchecked((uint)exitCode);
                Interlocked.Increment(ref _webViewExited);
                string line = "T+" + UptimeSeconds() + "s  ★WebView2 子进程退出 pid=" + kv.Key +
                              " type=" + kv.Value.Type +
                              " exitCode=" + CrashLog.Hex(uc) + " (" + CrashLog.ExceptionCodeName(uc) + ")" +
                              " 存活=" + lifeSec.ToString("F1", CultureInfo.InvariantCulture) + "s" +
                              "  ★该进程先于宿主退出（宿主此刻仍在运行）";
                CrashLog.Timeline(line);
                JournalAdd(line);
            }
            if (gone != null)
            {
                foreach (uint id in gone)
                {
                    try { known[id].Proc.Dispose(); } catch { }
                    known.Remove(id);
                }
            }
            if (foreign.Count > 256) foreign.Clear();
        }

        private static void JournalAdd(string line)
        {
            try
            {
                int index = Interlocked.Increment(ref _journalNext) - 1;
                Journal[index % JournalCapacity] = line;
            }
            catch { }
        }

        /// <summary>WebView2 进程账本总结：崩溃现场与宿主正常退出时都写一份。</summary>
        private static string JournalSummary()
        {
            var sb = new StringBuilder();
            int total = Volatile.Read(ref _journalNext);
            int seen = Volatile.Read(ref _webViewSeen);
            int exited = Volatile.Read(ref _webViewExited);
            sb.AppendLine("---- WebView2 进程账本（回答「谁先死」）----");
            sb.AppendLine("统计：曾识别到本宿主的 WebView2 子进程 " + seen + " 个，其中已退出 " + exited + " 个");
            if (total == 0)
            {
                sb.AppendLine("  (本次会话没有 WebView2 子进程事件：窗格未创建或 WebView2 未启动)");
                sb.AppendLine("结论：崩溃时没有 WebView2 进程参与 → 方向在宿主侧 / 加载项自身代码");
                return sb.ToString();
            }
            int n = total < JournalCapacity ? total : JournalCapacity;
            for (int i = 0; i < n; i++)
            {
                string line = Journal[(total - n + i) % JournalCapacity];
                if (line != null) sb.AppendLine("  " + line);
            }
            if (exited == 0)
                sb.AppendLine("结论：★崩溃/退出时所有 WebView2 子进程都还活着 → 不是浏览器侧先死，方向在宿主侧");
            else
                sb.AppendLine("结论：★有 " + exited + " 个 WebView2 子进程在宿主结束前就退出了 → " +
                              "请对比上面的退出码与 addin.log 时间戳，判断是否浏览器侧先崩");
            return sb.ToString();
        }

        /// <summary>宿主正常关闭时调用（Connect 的 OnBeginShutdown/OnDisconnection、以及 ProcessExit）。</summary>
        internal static void NotifyHostShutdown(string reason)
        {
            try
            {
                Volatile.Write(ref _watcherStop, 1);
                CrashLog.Timeline("宿主关闭通知（" + reason + "）");
                CrashLog.AppendDurable(JournalSummary());
            }
            catch { }
        }

        private static string ExtractArg(string cmdLine, string prefix)
        {
            try
            {
                if (string.IsNullOrEmpty(cmdLine)) return "?";
                int i = cmdLine.IndexOf(prefix, StringComparison.OrdinalIgnoreCase);
                if (i < 0) return "?";
                int start = i + prefix.Length;
                int end = cmdLine.IndexOf(' ', start);
                if (end < 0) end = cmdLine.Length;
                string v = cmdLine.Substring(start, end - start).Trim('"');
                return string.IsNullOrEmpty(v) ? "?" : v;
            }
            catch { return "?"; }
        }

        private static string UptimeSeconds()
        {
            try { return (CrashLog.UptimeMs / 1000.0).ToString("F1", CultureInfo.InvariantCulture); }
            catch { return "?"; }
        }

        // ===================== 小工具 =====================

        private static string SafeProcessName()
        {
            try
            {
                string name = Process.GetCurrentProcess().ProcessName;
                if (string.IsNullOrEmpty(name)) return "proc";
                var sb = new StringBuilder(name.Length);
                foreach (char c in name)
                    if (char.IsLetterOrDigit(c) || c == '_' || c == '-' || c == '.') sb.Append(c);
                return sb.Length > 0 ? sb.ToString() : "proc";
            }
            catch { return "proc"; }
        }

        private static string SafeExeName()
        {
            try
            {
                string path = Process.GetCurrentProcess().MainModule.FileName;
                return Path.GetFileName(path);
            }
            catch { return "WINWORD.EXE"; }
        }

        private static string SafeMainModulePath()
        {
            try { return Process.GetCurrentProcess().MainModule.FileName; }
            catch { return "?"; }
        }

        private static string SafeAssemblyPath()
        {
            try { return Assembly.GetExecutingAssembly().Location; }
            catch { return "?"; }
        }

        private static string SafeAssemblyVersion()
        {
            try { return Assembly.GetExecutingAssembly().GetName().Version.ToString(); }
            catch { return "?"; }
        }

        private static string SafeStartTime()
        {
            try { return Process.GetCurrentProcess().StartTime.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture); }
            catch { return "?"; }
        }

        private static long SafeFileSize(string path)
        {
            try
            {
                var fi = new FileInfo(path);
                return fi.Exists ? fi.Length : 0;
            }
            catch { return 0; }
        }

        private static long SafeFreeDiskMb(string dir)
        {
            try
            {
                string root = Path.GetPathRoot(Path.GetFullPath(dir));
                return new DriveInfo(root).AvailableFreeSpace / 1048576;
            }
            catch { return -1; }
        }
    }
}
