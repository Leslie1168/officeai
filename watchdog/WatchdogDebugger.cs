using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

namespace OfficeAi.Watchdog
{
    /// <summary>
    /// 进程外看门狗的核心：一个最小但正确的 Win32 调试器。
    ///
    /// 为什么必须是"进程外"：
    ///   · 本机 WER 全局禁用（HKLM\...\Windows Error Reporting\Disabled=1），拿不到系统 dump；
    ///   · 进程内的 SetUnhandledExceptionFilter 是"后装者赢"，Office / CLR / WebView2 任何一方
    ///     后来装过滤器都会顶掉它（交接报告 3.6，14:57 崩溃就是这么丢掉现场的）；
    ///   · VEH 虽然不可被覆盖，但它仍然在目标进程内部，崩溃发生在"过滤器跑不了"的状态时就没了。
    ///   调试器（DebugActiveProcess）走的是内核通道：异常在分发到进程内任何处理器之前，
    ///   先以 EXCEPTION_DEBUG_EVENT 送到这里。这是唯一"先于一切处理器"的位置。
    ///
    /// 行为准则：**只取证，不改变目标进程行为**。
    ///   所有真实异常一律以 DBG_EXCEPTION_NOT_HANDLED 放行 → Word 完全按它原本的路径走
    ///   （该处理处理、该崩崩）。唯一用 DBG_CONTINUE 的是 attach 后的初始断点。
    /// </summary>
    internal sealed class WatchdogDebugger
    {
        private readonly Options _o;
        private readonly WatchdogLog _log;

        private IntPtr _hProcess = IntPtr.Zero;
        private uint _pid;
        private string _procName = "target";
        private string _procPath;
        private bool _is32Bit;
        private List<ModuleEntry> _modules;

        private int _dumpWritten;                       // 0/1，保证一次运行只写一个 dump
        private string _dumpPath;
        private readonly HashSet<uint> _loggedCodes = new HashSet<uint>();
        private readonly HashSet<uint> _closedThreadIds = new HashSet<uint>();
        private readonly object _handleBook = new object();
        private int _handlesClosed, _handlesSkipped;
        private int _selfHandlesAtStart;
        // 符号解析要用的地址（异常现场收集，等 dump 写完后异步解析）
        private readonly List<ulong> _symbolAddresses = new List<ulong>();
        private readonly object _symbolGate = new object();
        private Thread _symbolThread;
        private IntPtr _hProcessForSymbols = IntPtr.Zero;

        public WatchdogDebugger(Options options, WatchdogLog log)
        {
            _o = options;
            _log = log;
        }

        public string DumpPath { get { return _dumpPath; } }

        // ------------------------------------------------------------------ 入口

        public int Run()
        {
            if (_o.SelfTest) return new Selftest(_o, _log).Run();
            if (_o.WorkerMode != null) return Worker.Run(_o, _log);

            _selfHandlesAtStart = SelfHandleCount();

            int attachResult = _o.ExePath != null ? AttachByLaunch() : AttachByPid();
            if (attachResult != 0) return attachResult;

            _log.Info("看门狗自身 PID=" + NativeMethods.GetCurrentProcessId() + "，句柄数(attach 后)=" + SelfHandleCount());

            int rc = DebugLoop();

            int selfEnd = SelfHandleCount();
            _log.Info("句柄清点：看门狗自身句柄数 起始=" + _selfHandlesAtStart + " 结束=" + selfEnd
                      + "；关闭线程句柄 " + _handlesClosed + " 个，跳过重复 " + _handlesSkipped + " 个");

            // 符号线程：给一小段宽限期收尾，但绝不无限等
            WaitForSymbolThread(_o.SymbolWaitSeconds);

            CleanupHandles();
            StartupGuard.ExitedDebugLoop();
            return rc;
        }

        private static int SelfHandleCount()
        {
            try
            {
                uint n;
                return NativeMethods.GetProcessHandleCount(NativeMethods.GetCurrentProcess(), out n) ? (int)n : 0;
            }
            catch { return 0; }
        }

        // ------------------------------------------------------------------ attach

        private int AttachByPid()
        {
            _pid = (uint)_o.Pid;
            StartupGuard.SetPhase("OpenProcess(pid=" + _pid + ", PROCESS_ALL_ACCESS)");

            _hProcess = NativeMethods.OpenProcess(NativeMethods.PROCESS_ALL_ACCESS, false, _o.Pid);
            if (_hProcess == IntPtr.Zero)
            {
                int err = Marshal.GetLastWin32Error();
                _log.Error("attach 失败：OpenProcess(pid=" + _o.Pid + ") 返回 NULL，win32=" + err + " (" + Win32Text(err) + ")");
                _log.Error("attach 失败诊断：" + DiagnoseOpenProcessFailure(err, _o.Pid));
                return ExitCodes.AttachFailed;
            }
            _log.Info("OpenProcess 成功：hProcess=0x" + _hProcess.ToInt64().ToString("X"));

            return DoDebugActiveProcess();
        }

        private int AttachByLaunch()
        {
            StartupGuard.SetPhase("CreateProcess(DEBUG_ONLY_THIS_PROCESS) " + _o.ExePath);
            var si = new NativeMethods.STARTUPINFO();
            si.cb = (uint)Marshal.SizeOf(typeof(NativeMethods.STARTUPINFO));

            string cmd = Quote(_o.ExePath) + (_o.ChildArgs != null ? " " + _o.ChildArgs : "");
            NativeMethods.PROCESS_INFORMATION pi;
            bool ok = NativeMethods.CreateProcessW(_o.ExePath, cmd, IntPtr.Zero, IntPtr.Zero, false,
                NativeMethods.DEBUG_ONLY_THIS_PROCESS | NativeMethods.CREATE_NEW_CONSOLE,
                IntPtr.Zero, null, ref si, out pi);
            if (!ok)
            {
                int err = Marshal.GetLastWin32Error();
                _log.Error("attach 失败：CreateProcess 失败 win32=" + err + " (" + Win32Text(err) + ") cmd=" + cmd);
                return ExitCodes.LaunchFailed;
            }

            // 这两个句柄要立刻关掉：调试会话自己会再给我们一份（CREATE_PROCESS_DEBUG_EVENT）
            uint pid = pi.dwProcessId;
            NativeMethods.CloseHandle(pi.hThread);
            NativeMethods.CloseHandle(pi.hProcess);

            _pid = pid;
            _log.Info("已以调试方式启动：" + _o.ExePath + " (pid=" + pid + ") 命令行=" + cmd);
            _log.Info("注意：CreateProcess 的 hProcess/hThread 已立即 CloseHandle（否则 3 个句柄泄漏一组）；"
                      + "调试事件的 CREATE_PROCESS_DEBUG_EVENT 会再给一份，那份由调试循环管理。");
            StartupGuard.SetPhase("等待 CREATE_PROCESS_DEBUG_EVENT（pid=" + pid + "）");
            return 0;
        }

        private int DoDebugActiveProcess()
        {
            // attach 前先探一下目标是否还活着，给出更准确的失败原因（而不是让 DebugActiveProcess 报个含糊的错）
            string name = TryGetProcessName(_pid);
            if (name != null) _procName = name;
            _log.Info("目标进程：" + _procName + " (pid=" + _pid + ")");

            StartupGuard.SetPhase("DebugActiveProcess(pid=" + _pid + ")");
            if (!NativeMethods.DebugActiveProcess(_o.Pid))
            {
                int err = Marshal.GetLastWin32Error();
                _log.Error("attach 失败：DebugActiveProcess(pid=" + _o.Pid + ") 返回 FALSE，win32=" + err + " (" + Win32Text(err) + ")");
                _log.Error("attach 失败诊断：" + DiagnoseDebugActiveFailure(err));
                return ExitCodes.AttachFailed;
            }
            _log.Info("DebugActiveProcess 成功：已附加到 " + _procName + " (pid=" + _pid + ")，此后目标每次异常都会先送到本看门狗");

            SetKillOnExitFalse();
            DetectArchitecture();
            EnumerateInitialThreads();
            _modules = ModuleMap.Enumerate(_hProcess, _log);
            _log.Info("模块表已枚举：" + _modules.Count + " 个模块");

            StartupGuard.SetPhase("进入 WaitForDebugEvent 循环");
            return 0;
        }

        private void SetKillOnExitFalse()
        {
            StartupGuard.SetPhase("DebugSetProcessKillOnExit(FALSE)");

            // ★★ 这是本程序最重要的一条不变式 ★★
            // 默认语义（TRUE）是"调试器进程一退出，内核就把被调试进程一起杀掉"。
            // 看门狗是个可以随时被 taskkill / 关机 / 崩溃带走的独立进程，一旦默认生效，
            // 用户干掉看门狗，陪葬的是他正在编辑的 Word 文档 —— 对用户是灾难。
            bool ok = NativeMethods.DebugSetProcessKillOnExit(false);
            int err = Marshal.GetLastWin32Error();
            if (ok)
            {
                _log.Banner("DebugSetProcessKillOnExit(FALSE) 已确认生效：看门狗退出时不会带走 " + _procName + " (pid=" + _pid + ")");
            }
            else
            {
                _log.Error("DebugSetProcessKillOnExit(FALSE) 失败！win32=" + err + " (" + Win32Text(err) + ")");
                _log.Error("★ 风险：看门狗若异常退出，可能连带终止被调试进程（Word）。"
                           + "请勿在此期间手动终止看门狗进程；如需停止取证，请先正常关闭 Word。");
            }
        }

        private void DetectArchitecture()
        {
            _is32Bit = false;
            try
            {
                ushort procMachine, nativeMachine;
                if (NativeMethods.IsWow64Process2(_hProcess, out procMachine, out nativeMachine))
                {
                    // IMAGE_FILE_MACHINE_UNKNOWN(0) 表示"不是 WOW64"，即与调试器同为 x64
                    _is32Bit = procMachine != 0;
                    _log.Info("目标架构：processMachine=0x" + procMachine.ToString("X") + " nativeMachine=0x" + nativeMachine.ToString("X")
                              + " → " + (_is32Bit ? "32 位 (WOW64)" : "64 位"));
                }
                else
                {
                    bool wow;
                    if (NativeMethods.IsWow64Process(_hProcess, out wow)) _is32Bit = wow;
                    _log.Info("目标架构：IsWow64Process2 不可用（老系统），回退 IsWow64Process → " + (_is32Bit ? "32 位" : "64 位"));
                }
            }
            catch (Exception ex)
            {
                _log.Warn("架构探测失败，按 64 位处理：" + ex.GetType().Name);
            }
            if (_is32Bit)
            {
                _log.Warn("目标是 32 位进程，而看门狗是 64 位 —— DebugActiveProcess 通常会失败；"
                          + "若竟然成功了，CONTEXT 大小按 716 处理。");
            }
        }

        /// <summary>
        /// attach 到已在运行的进程时，系统会把"当前已存在的每个线程"都作为 CREATE_THREAD_DEBUG_EVENT
        /// 排队送给我们。那些 hThread 每来一个都要 CloseHandle，否则 Word（几十个线程）会让我们
        /// 瞬间泄漏几十个句柄。这里先枚举一遍线程 ID 记入"已知"，配合事件流做去重，避免重复关闭。
        /// </summary>
        private void EnumerateInitialThreads()
        {
            try
            {
                IntPtr snap = NativeMethods.CreateToolhelp32Snapshot(NativeMethods.TH32CS_SNAPTHREAD, 0);
                if (snap == IntPtr.Zero || snap == NativeMethods.INVALID_HANDLE_VALUE)
                {
                    _log.Warn("CreateToolhelp32Snapshot 失败，无法预枚举线程（不影响取证）");
                    return;
                }
                try
                {
                    var te = new NativeMethods.THREADENTRY32();
                    te.dwSize = (uint)Marshal.SizeOf(typeof(NativeMethods.THREADENTRY32));
                    int count = 0;
                    if (NativeMethods.Thread32First(snap, ref te))
                    {
                        do
                        {
                            if (te.th32OwnerProcessID == _pid) count++;
                            te.dwSize = (uint)Marshal.SizeOf(typeof(NativeMethods.THREADENTRY32));
                        } while (NativeMethods.Thread32Next(snap, ref te));
                    }
                    _log.Info("目标进程当前线程数：" + count + "（attach 后每个线程都会来一个 CREATE_THREAD_DEBUG_EVENT，其 hThread 会被关闭）");
                }
                finally { NativeMethods.CloseHandle(snap); }
            }
            catch (Exception ex)
            {
                _log.Warn("预枚举线程异常（忽略）：" + ex.GetType().Name);
            }
        }

        // ------------------------------------------------------------------ 主循环

        private int DebugLoop()
        {
            StartupGuard.EnteredDebugLoop();
            _log.Info("=== 进入 WaitForDebugEvent 循环（目标 pid=" + _pid + "，out=" + _o.OutDir + "）===");

            var de = new NativeMethods.DEBUG_EVENT();
            long events = 0;
            var startedAt = DateTime.Now;

            while (true)
            {
                bool ok;
                try
                {
                    ok = NativeMethods.WaitForDebugEvent(out de, 1000);
                }
                catch (Exception ex)
                {
                    _log.Error("WaitForDebugEvent 抛异常（P/Invoke 层）：" + ex.GetType().Name + " " + ex.Message);
                    break;
                }

                if (!ok)
                {
                    int err = Marshal.GetLastWin32Error();
                    // 121 = ERROR_SEM_TIMEOUT，官方语义就是"这一秒没有事件"，属正常
                    if (err == 121 || err == 0) continue;
                    _log.Error("WaitForDebugEvent 失败 win32=" + err + " (" + Win32Text(err) + ")，退出调试循环");
                    break;
                }

                events++;
                int continueStatus = NativeMethods.DBG_EXCEPTION_NOT_HANDLED;
                bool stop = false;

                try
                {
                    switch (de.dwDebugEventCode)
                    {
                        case NativeMethods.CREATE_PROCESS_DEBUG_EVENT:
                            continueStatus = OnCreateProcess(de);
                            break;

                        case NativeMethods.CREATE_THREAD_DEBUG_EVENT:
                            OnCreateThread(de);
                            break;

                        case NativeMethods.EXIT_THREAD_DEBUG_EVENT:
                            OnExitThread(de);
                            break;

                        case NativeMethods.EXCEPTION_DEBUG_EVENT:
                            continueStatus = OnException(de);
                            break;

                        case NativeMethods.EXIT_PROCESS_DEBUG_EVENT:
                            OnExitProcess(de, events, startedAt);
                            stop = true;
                            break;

                        case NativeMethods.LOAD_DLL_DEBUG_EVENT:
                            // 不刷屏：只计数（Word 有 200+ 模块，逐条写会把日志淹掉）
                            _loadDllCount++;
                            break;

                        case NativeMethods.UNLOAD_DLL_DEBUG_EVENT:
                            _unloadDllCount++;
                            break;

                        case NativeMethods.OUTPUT_DEBUG_STRING_EVENT:
                            _outputDebugStringCount++;
                            break;

                        case NativeMethods.RIP_EVENT:
                            _log.Warn("RIP_EVENT：系统调试信息丢失，系统错误=" + de.u.RipInfo.dwError);
                            break;

                        default:
                            _log.Warn("未知调试事件码 " + de.dwDebugEventCode);
                            break;
                    }
                }
                catch (Exception ex)
                {
                    // 任何事件处理异常都不能让看门狗死掉：记下来，放行，继续
                    _log.Exception("处理调试事件 " + de.dwDebugEventCode + " 时抛异常（已放行继续）", ex);
                }

                if (!NativeMethods.ContinueDebugEvent(de.dwProcessId, de.dwThreadId, continueStatus))
                {
                    int err = Marshal.GetLastWin32Error();
                    if (err != 0)
                    {
                        _log.Warn("ContinueDebugEvent 返回 FALSE win32=" + err + " (" + Win32Text(err) + ") status=0x"
                                  + continueStatus.ToString("X") + "（事件=" + de.dwDebugEventCode + "）");
                    }
                }

                if (stop) break;
            }

            _log.Info("调试循环结束：共处理调试事件 " + events + " 个"
                      + "（LOAD_DLL " + _loadDllCount + "，UNLOAD_DLL " + _unloadDllCount
                      + "，OUTPUT_DEBUG_STRING " + _outputDebugStringCount + "）");
            return ExitCodes.Ok;
        }

        private int _loadDllCount, _unloadDllCount, _outputDebugStringCount;

        private int OnCreateProcess(NativeMethods.DEBUG_EVENT de)
        {
            var info = de.u.CreateProcessInfo;
            _hProcess = info.hProcess;
            _pid = de.dwProcessId;

            _log.Info("CREATE_PROCESS：pid=" + _pid + " 映像基址=0x" + info.lpBaseOfImage.ToInt64().ToString("X")
                      + " hProcess=0x" + info.hProcess.ToInt64().ToString("X"));

            // ★ 必须关 hFile：这是系统给调试器的"映像文件句柄"，不关就一路泄漏
            CloseThreadOrFileHandle(info.hFile, "CREATE_PROCESS.hFile");

            // 初始线程句柄同样要关（记住它的线程 ID 以便后续 CREATE_THREAD 去重）
            if (info.hThread != IntPtr.Zero)
            {
                CloseThreadHandle(info.hThread, de.dwThreadId, "CREATE_PROCESS.hThread");
            }

            // ★ 刻意**不在这里**枚举模块表。
            // 实测（--selftest 场景 A）：进程刚被调试启动时，它停在 loader 里的调试断点上，
            // 此刻模块表还没建好，EnumProcessModules 会返回 win32=299 (ERROR_PARTIAL_COPY)，
            // 枚举结果为空 —— 于是"异常地址→模块名+RVA"这条最关键的链在唯一的真实崩溃场景里失效。
            // 改为**延迟到异常发生的那一刻再枚举**（那时进程早已跑过 loader），并且带兜底通道。
            _modules = null;
            _log.Info("模块表改为延迟枚举（首次异常时再读）—— 避开进程刚启动时 win32=299 的窗口期");

            string n = TryGetProcessName(_pid);
            if (n != null) _procName = n;

            _procPath = TryGetProcessPath(_pid);

            // --exe 模式没有单独调过 SetKillOnExit，这里补上
            SetKillOnExitFalse();
            DetectArchitecture();
            StartupGuard.EnteredDebugLoop();
            return NativeMethods.DBG_CONTINUE;
        }

        private void OnCreateThread(NativeMethods.DEBUG_EVENT de)
        {
            // 不刷屏：只有需要确认句柄纪律时才记一行
            CloseThreadHandle(de.u.CreateThread.hThread, de.dwThreadId, "CREATE_THREAD.hThread");
        }

        private void OnExitThread(NativeMethods.DEBUG_EVENT de)
        {
            // 线程 ID 会被 Windows 复用；退出时从"已关闭"表里移除，
            // 这样复用的新线程不会再被误判成"已关闭"而漏关句柄。
            lock (_handleBook) { _closedThreadIds.Remove(de.dwThreadId); }
        }

        private void OnExitProcess(NativeMethods.DEBUG_EVENT de, long events, DateTime startedAt)
        {
            uint code = de.u.ExitProcess.dwExitCode;
            double secs = (DateTime.Now - startedAt).TotalSeconds;
            string exitLine = "EXIT_PROCESS：pid=" + _pid + " (" + _procName + ") 退出，退出码=0x" + code.ToString("X8")
                              + " (" + code + ") 运行时长=" + secs.ToString("F1") + "s 调试事件总数=" + events;

            _log.Banner(exitLine);
            _log.Banner(DescribeExitCode(code));

            if (_dumpPath != null)
            {
                _log.Info("本次会话已捕获 dump：" + _dumpPath);
            }
            else
            {
                _log.Warn("本次会话未写任何 dump：目标进程在 attach 期间没有出现符合条件的异常。");
                _log.Warn("（若用户确实观察到闪退，请检查看门狗是否在崩溃前就已经退出、或 attach 是否失败 —— 往上翻本日志的 attach 段。）");
            }

            // 把退出信息补进 sidecar：进程退出发生在 dump 之后，sidecar 是追加式证据文件
            ExceptionSidecar.AppendExit(_sidecarPath, exitLine + Environment.NewLine + "  " + DescribeExitCode(code), _log);
        }

        /// <summary>
        /// 模块表延迟枚举（只在需要时读一次，成功后缓存）。
        /// 提到异常发生时再枚举，是为了避开"进程刚被调试启动、loader 未完成"那个 win32=299 的窗口期。
        /// </summary>
        private void EnsureModules()
        {
            if (_modules != null) return;
            _modules = ModuleMap.Enumerate(_hProcess, _log);
            if (_modules.Count > 0)
            {
                _log.Info("模块表已枚举（延迟枚举，异常发生时读取）：" + _modules.Count + " 个模块");
            }
            else
            {
                _log.Warn("模块表枚举结果为空 —— 模块归属将完全依赖 VirtualQueryEx+GetMappedFileNameW 兜底通道"
                          + "（该通道只需要出错的那个地址，实测可用）。");
            }
        }

        // ------------------------------------------------------------------ 异常处理

        private int OnException(NativeMethods.DEBUG_EVENT de)
        {
            NativeMethods.EXCEPTION_RECORD rec = de.u.Exception.ExceptionRecord;
            bool firstChance = de.u.Exception.dwFirstChance != 0;
            uint code = rec.ExceptionCode;
            uint tid = de.dwThreadId;
            ulong addr = unchecked((ulong)rec.ExceptionAddress.ToInt64());
            _firstChanceOfCurrent = firstChance;

            // ---- 1) 初始/加载器断点：这是 attach 后系统插的 int3，必须 DBG_CONTINUE 放行，
            //         否则被调试进程会因为一个没人处理的断点直接死掉。
            if (code == NativeMethods.STATUS_BREAKPOINT && firstChance)
            {
                _log.Info("初始断点 STATUS_BREAKPOINT (first-chance, tid=" + tid + ") → DBG_CONTINUE 放行（这是 attach 的正常握手，不是崩溃）");
                return NativeMethods.DBG_CONTINUE;
            }

            if (code == NativeMethods.STATUS_SINGLE_STEP)
            {
                return NativeMethods.DBG_CONTINUE;
            }

            // ---- 2) 完整证据行 ----
            ulong access = rec.ExceptionInformation0;
            ulong faultAddr = rec.ExceptionInformation1;
            string accessText = DescribeAccess(access);

            // 模块表延迟枚举：异常发生时才读，这时进程早已跑过 loader，win32=299 不会再来。
            EnsureModules();
            if (ModuleMap.Find(_modules, addr) == null)
            {
                // 地址不在快照里 —— 可能是快照之后才加载的模块（Word 会动态加载），重读一次
                _log.Info("异常地址不在当前模块表快照里，重新枚举一次模块表（可能是在快照之后才加载的模块）");
                _modules = null;
                EnsureModules();
            }
            string moduleText = ModuleMap.Describe(_hProcess, _modules, addr, _log);

            _log.Info("──────────────────────────────────────────────────────────────");
            _log.Info("EXCEPTION code=0x" + code.ToString("X8") + " (" + ExceptionText(code) + ") "
                      + (firstChance ? "first-chance" : "★second-chance(进程即将终止)") + " tid=" + tid);
            _log.Info("  异常地址 : 0x" + addr.ToString("X16") + "  →  " + ModuleMap.ShortForm(_hProcess, _modules, addr, _log));
            _log.Info("  访问类型 : " + accessText + "  (ExceptionInformation[0]=0x" + access.ToString("X") + ")");
            _log.Info("  出错地址 : 0x" + faultAddr.ToString("X16")
                      + (faultAddr == 0 ? "  ← 空指针！" : (faultAddr < 0x10000 ? "  ← 极低地址，基本可判定为野指针/空指针偏移访问" : "")));
            _log.Info("  模块归属 : " + moduleText);
            _log.Info("  NumberParameters=" + rec.NumberParameters);

            // ---- 3) 取寄存器上下文 + 栈扫描（这一步是"没有符号也能看路径"的关键）----
            IntPtr ctxBuf = IntPtr.Zero;
            int ctxSize = _is32Bit ? NativeMethods.CONTEXT_SIZE_X86 : NativeMethods.CONTEXT_SIZE_X64;
            string regs = null;
            List<ulong> stack = new List<ulong>();
            try
            {
                ctxBuf = CaptureThreadContext(tid, ctxSize, out regs);
                if (ctxBuf != IntPtr.Zero)
                {
                    _log.Info("  寄存器   : " + regs);
                    stack = ScanStack(tid, ctxBuf);
                    if (stack.Count > 0)
                    {
                        _log.Info("  栈候选返回地址（启发式扫描 RSP 起 4KB，只列出落在模块映像范围内的值）共 " + stack.Count + " 个：");
                        for (int i = 0; i < stack.Count; i++)
                        {
                            _log.Info("    #" + i.ToString("D2") + " 0x" + stack[i].ToString("X16") + "  "
                                      + ModuleMap.ShortForm(_hProcess, _modules, stack[i], null));
                        }
                    }
                    else
                    {
                        _log.Info("  栈扫描   : RSP 起 4KB 内没有落在模块范围内的值（可能是栈损坏 / 栈指针异常）");
                    }
                }
            }
            catch (Exception ex)
            {
                _log.Warn("采集寄存器/栈失败（不影响 dump）：" + ex.GetType().Name + " " + ex.Message);
            }

            // 收集给符号线程的地址
            lock (_symbolGate)
            {
                if (addr != 0) _symbolAddresses.Add(addr);
                if (ctxBuf != IntPtr.Zero)
                {
                    ulong rip = ReadPtrReg(ctxBuf, NativeMethods.X64_OFF_Rip);
                    if (rip != 0) _symbolAddresses.Add(rip);
                }
                for (int i = 0; i < stack.Count && i < 40; i++) _symbolAddresses.Add(stack[i]);
            }

            // ---- 4) 是否取证 ----
            string reason;
            bool dumpWorthy = ShouldDump(rec, firstChance, out reason);

            if (dumpWorthy && Interlocked.CompareExchange(ref _dumpWritten, 1, 0) == 0)
            {
                _log.Banner("★★★ 判定为可取证异常 → 写全内存 minidump + sidecar（原因：" + reason + "）");

                // 先把现场整理成一个对象：**dump 可能失败，但这些字段一定写得出来**
                var c = new CrashContext
                {
                    Pid = _pid,
                    ProcName = _procName,
                    ProcPath = _procPath,
                    ExceptionCode = code,
                    FirstChance = firstChance,
                    ExceptionAddress = addr,
                    Info0 = access,
                    Info1 = faultAddr,
                    ThreadId = tid,
                    NumberParameters = rec.NumberParameters,
                    ExceptionName = ExceptionText(code),
                    AccessText = accessText,
                    Registers = regs,
                    StackCandidates = stack,
                    HaveContext = ctxBuf != IntPtr.Zero,
                    ContextFlags = ctxBuf != IntPtr.Zero ? ReadContextFlags(ctxBuf) : 0
                };

                // 异常地址所在模块（含 VirtualQueryEx 兜底通道）
                ModuleEntry m = ModuleMap.Resolve(_hProcess, _modules, addr, _log);
                if (m != null)
                {
                    c.ModuleName = m.Name; c.ModuleBase = m.BaseU; c.ModuleSize = m.Size;
                    c.ModulePath = m.FullPath; c.ModuleSource = m.Source; c.ModuleRva = addr - m.BaseU;
                }
                ModuleEntry fm = faultAddr != 0 ? ModuleMap.Resolve(_hProcess, _modules, faultAddr, null) : null;
                if (fm != null)
                {
                    c.FaultModuleName = fm.Name; c.FaultModuleBase = fm.BaseU; c.FaultModuleRva = faultAddr - fm.BaseU;
                }

                for (int i = 0; i < stack.Count; i++)
                {
                    c.StackCandidatesText.Add("#" + i.ToString("D2") + " 0x" + stack[i].ToString("X16") + "  "
                        + ModuleMap.ShortForm(_hProcess, _modules, stack[i], null));
                }

                WriteDump(c, rec, tid, ctxBuf, ctxSize, faultAddr, reason);
            }
            else if (dumpWorthy)
            {
                _log.Warn("该异常同样符合取证条件（" + reason + "），但本次运行已写过 dump → 只记录不重复写（防风暴）");
            }
            else
            {
                _log.Info("不写 dump（" + reason + "）。异常已按 DBG_EXCEPTION_NOT_HANDLED 原样放行给被调试进程自己的处理器。");
            }

            if (ctxBuf != IntPtr.Zero) Marshal.FreeHGlobal(ctxBuf);

            // ★ 放行：绝不吞异常。first-chance 交给 Word 自己的 SEH/CLR；
            //   second-chance 交回内核走原本的终止路径。我们只取证，不改变行为。
            _log.Info("→ ContinueDebugEvent(DBG_EXCEPTION_NOT_HANDLED)：把异常原样放行给 " + _procName + " 自己的处理路径");
            return NativeMethods.DBG_EXCEPTION_NOT_HANDLED;
        }

        /// <summary>
        /// 取证判定。原则：宁可多抓一次，不可漏掉真崩溃；但明确的"正常控制流"要排除，
        /// 否则一次运行的唯一 dump 名额会被噪声用掉（Word/CLR 启动期会有大量正常 AV 探测）。
        /// </summary>
        private bool ShouldDump(NativeMethods.EXCEPTION_RECORD rec, bool firstChance, out string reason)
        {
            uint code = rec.ExceptionCode;
            ulong access = rec.ExceptionInformation0;
            ulong faultAddr = rec.ExceptionInformation1;

            if (code == NativeMethods.STATUS_ACCESS_VIOLATION)
            {
                // PAGE_GUARD 位（0x100）：CLR/CRT 用 guard page 做栈增长探测，是正常控制流。
                // 这类 first-chance AV 在 Word 启动期会成片出现，必须排除。
                if ((access & 0x100) != 0)
                {
                    reason = "PAGE_GUARD 探测（ExceptionInformation[0] 含 0x100），属正常控制流，不是崩溃";
                    return false;
                }
                reason = (access == NativeMethods.EXCEPTION_EXECUTE_FAULT)
                    ? "执行了不可执行地址（坏函数指针 / DEP 违规）"
                    : "非 guard 的访问违例" + (faultAddr < 0x10000 ? "，且出错地址极低（" + "0x" + faultAddr.ToString("X") + "），几乎必然是野指针" : "");
                return true;
            }

            if (code == NativeMethods.STATUS_STACK_OVERFLOW)
            {
                reason = "栈溢出（STATUS_STACK_OVERFLOW）";
                return true;
            }
            if (code == NativeMethods.STATUS_FAIL_FAST_EXCEPTION)
            {
                reason = "fail-fast（STATUS_FAIL_FAST_EXCEPTION，CRT/CLR 主动终止）";
                return true;
            }
            if (code == 0xC0000374) // STATUS_HEAP_CORRUPTION
            {
                reason = "堆损坏（STATUS_HEAP_CORRUPTION）";
                return true;
            }

            // 任何 second-chance 异常都意味着"没有处理器接手，进程就要死了" —— 值得留一份现场
            if (!firstChance)
            {
                reason = "second-chance 异常，进程即将终止，任何异常码都值得留现场";
                return true;
            }

            // 其余 first-chance（C++ 异常 0xE06D7363、CLR 异常、断点等）都是正常控制流，只记日志
            reason = "first-chance 且非访问违例/栈溢出/fail-fast（C++/CLR 异常属正常控制流），只记录";
            return false;
        }

        private void WriteDump(CrashContext c, NativeMethods.EXCEPTION_RECORD rec, uint tid, IntPtr ctxBuf, int ctxSize,
            ulong faultAddr, string reason)
        {
            // 磁盘空间预警（Word 全内存 dump 可能 1~2GB）
            WarnIfLowDiskSpace();

            var sw = System.Diagnostics.Stopwatch.StartNew();
            _log.Info("开始写 dump：目标 pid=" + _pid + " 全内存标志，期间被调试进程保持冻结（它此刻停在调试停点上，"
                      + "不冻结就取不到一致的现场）；大进程可能需要数十秒。");

            DumpResult r;
            lock (SymbolResolver.DbgHelpGate)
            {
                r = DumpWriter.Write(_hProcess, _pid, rec, tid, ctxBuf, ctxSize, _o.OutDir, _procName,
                    _o.ClientPointers, _o.TryExceptionParam, _log);
            }
            sw.Stop();

            if (r.Success)
            {
                _dumpPath = r.Path;
                c.DumpPath = r.Path;
                c.DumpBytes = r.Bytes;
                c.DumpKind = (r.ExcAttempt != null && r.ExcAttempt.Accepted) ? "ExceptionParam 版本（含 ExceptionStream，可 .ecxr）"
                                                                            : "ExceptionParam=NULL 版本（无 ExceptionStream，用 ~~[tid]s 代替 .ecxr）";
                _log.Banner("dump 已落盘：" + r.Path);
                _log.Banner("  · 大小      : " + FormatBytes(r.Bytes));
                _log.Banner("  · 耗时      : " + sw.Elapsed.TotalSeconds.ToString("F1") + "s");
                _log.Banner("  · 标志      : 0x" + r.FlagsUsed.ToString("X8"));
            }
            else
            {
                _dumpPath = null;
                c.DumpPath = null;
                _log.Error("dump 未能落盘：" + r.Error);
            }

            // ---- sidecar：dump 成功与否都要写。它承载的是"结论"，而 dump 只是内存快照 ----
            string sidecar = ExceptionSidecar.Write(c, _o.OutDir, _log);
            _sidecarPath = sidecar;
            if (sidecar != null)
            {
                _log.Banner("异常现场 sidecar 已落盘：" + sidecar);
                _log.Banner("  （内容 = 异常码 / 异常地址 / 模块名+基址+RVA / 访问类型 / 出错线程 ID / 寄存器 / 栈候选 + 可直接复制的 cdb 命令）");
            }

            // ★ 给下一个看日志的人一句话摘要：这一行就是本次调查要的答案。
            //   注意即使 dump 失败也要写这一行 —— 定位 wwlib+0x183BA19 靠的是它，不是 dump。
            _log.Banner("[KEY] pid=" + _pid + " proc=" + _procName
                        + " code=0x" + rec.ExceptionCode.ToString("X8")
                        + " addr=0x" + unchecked((ulong)rec.ExceptionAddress.ToInt64()).ToString("X16")
                        + " module=" + (c.ModuleName != null ? c.ModuleName + "+0x" + c.ModuleRva.ToString("X") : "unknown")
                        + " modulebase=0x" + c.ModuleBase.ToString("X")
                        + " modulesource=" + (c.ModuleSource ?? "unknown")
                        + " faultaddr=0x" + faultAddr.ToString("X")
                        + " access=" + DescribeAccessShort(c.Info0)
                        + " tid=" + tid
                        + " chance=" + (_firstChanceOfCurrent ? "first" : "second")
                        + " reason=" + reason
                        + " dump=" + (r.Success ? r.Path : "(失败)")
                        + " sidecar=" + (sidecar ?? "(失败)"));

            QueueSymbolResolution();
        }

        private string _sidecarPath;
        private bool _firstChanceOfCurrent;

        private void WarnIfLowDiskSpace()
        {
            try
            {
                string root = Path.GetPathRoot(Path.GetFullPath(_o.OutDir));
                if (string.IsNullOrEmpty(root)) return;
                var di = new DriveInfo(root);
                long freeMb = di.AvailableFreeSpace / (1024 * 1024);
                _log.Info("输出盘剩余空间：" + freeMb + " MB（" + root + "）");
                if (freeMb < 3000)
                {
                    _log.Warn("★ 剩余空间不足 3GB：Word 的全内存 dump 可能有 1~2GB，可能写到一半失败。"
                              + "建议清理磁盘或把 --out 指到空间更大的盘。");
                }
            }
            catch (Exception ex)
            {
                _log.Warn("磁盘空间检查失败（忽略）：" + ex.GetType().Name);
            }
        }

        // ------------------------------------------------------------------ 上下文 / 栈

        /// <summary>
        /// 取异常线程的 CONTEXT。此刻该线程已被系统因调试事件而挂起，
        /// 所以这里读到的就是异常发生瞬间的寄存器 —— 注意不要额外 SuspendThread，
        /// 那会多一层挂起计数，ResumeThread 配平不回来就把 Word 冻死。
        /// </summary>
        private IntPtr CaptureThreadContext(uint tid, int ctxSize, out string regs)
        {
            regs = null;
            IntPtr buf = Marshal.AllocHGlobal(ctxSize);
            for (int i = 0; i < ctxSize; i += 8) Marshal.WriteInt64(buf, i, 0);

            if (_is32Bit) Marshal.WriteInt32(buf, 0, unchecked((int)NativeMethods.CONTEXT_ALL_X86));
            else Marshal.WriteInt32(buf, 48, unchecked((int)NativeMethods.CONTEXT_ALL_X64)); // ContextFlags @0x30

            IntPtr hThread = NativeMethods.OpenThread(
                NativeMethods.THREAD_GET_CONTEXT | NativeMethods.THREAD_QUERY_INFORMATION, false, tid);
            if (hThread == IntPtr.Zero)
            {
                int err = Marshal.GetLastWin32Error();
                _log.Warn("OpenThread(tid=" + tid + ") 失败 win32=" + err + " (" + Win32Text(err) + ")，无法取寄存器");
                Marshal.FreeHGlobal(buf);
                return IntPtr.Zero;
            }
            try
            {
                if (!NativeMethods.GetThreadContext(hThread, buf))
                {
                    int err = Marshal.GetLastWin32Error();
                    _log.Warn("GetThreadContext(tid=" + tid + ") 失败 win32=" + err + " (" + Win32Text(err) + ")");
                    Marshal.FreeHGlobal(buf);
                    return IntPtr.Zero;
                }
            }
            finally
            {
                NativeMethods.CloseHandle(hThread);
            }

            if (_is32Bit)
            {
                regs = string.Format("EIP=0x{0:X8} ESP=0x{1:X8} EBP=0x{2:X8} EAX=0x{3:X8} EBX=0x{4:X8} ECX=0x{5:X8} EDX=0x{6:X8} ESI=0x{7:X8} EDI=0x{8:X8} EFLAGS=0x{9:X8}",
                    (uint)Marshal.ReadInt32(buf, 184), (uint)Marshal.ReadInt32(buf, 196), (uint)Marshal.ReadInt32(buf, 180),
                    (uint)Marshal.ReadInt32(buf, 176), (uint)Marshal.ReadInt32(buf, 164), (uint)Marshal.ReadInt32(buf, 172),
                    (uint)Marshal.ReadInt32(buf, 168), (uint)Marshal.ReadInt32(buf, 160), (uint)Marshal.ReadInt32(buf, 156),
                    (uint)Marshal.ReadInt32(buf, 192));
            }
            else
            {
                regs = string.Format(
                    "RIP=0x{0:X16} RSP=0x{1:X16} RBP=0x{2:X16} RAX=0x{3:X16} RBX=0x{4:X16} RCX=0x{5:X16} RDX=0x{6:X16}\r\n"
                    + "             RSI=0x{7:X16} RDI=0x{8:X16} R8=0x{9:X16} R9=0x{10:X16} R10=0x{11:X16} R11=0x{12:X16}\r\n"
                    + "             R12=0x{13:X16} R13=0x{14:X16} R14=0x{15:X16} R15=0x{16:X16} EFLAGS=0x{17:X8}",
                    ReadPtrReg(buf, NativeMethods.X64_OFF_Rip), ReadPtrReg(buf, NativeMethods.X64_OFF_Rsp),
                    ReadPtrReg(buf, NativeMethods.X64_OFF_Rbp), ReadPtrReg(buf, NativeMethods.X64_OFF_Rax),
                    ReadPtrReg(buf, NativeMethods.X64_OFF_Rbx), ReadPtrReg(buf, NativeMethods.X64_OFF_Rcx),
                    ReadPtrReg(buf, NativeMethods.X64_OFF_Rdx), ReadPtrReg(buf, NativeMethods.X64_OFF_Rsi),
                    ReadPtrReg(buf, NativeMethods.X64_OFF_Rdi), ReadPtrReg(buf, NativeMethods.X64_OFF_R8),
                    ReadPtrReg(buf, NativeMethods.X64_OFF_R9), ReadPtrReg(buf, NativeMethods.X64_OFF_R10),
                    ReadPtrReg(buf, NativeMethods.X64_OFF_R11), ReadPtrReg(buf, NativeMethods.X64_OFF_R12),
                    ReadPtrReg(buf, NativeMethods.X64_OFF_R13), ReadPtrReg(buf, NativeMethods.X64_OFF_R14),
                    ReadPtrReg(buf, NativeMethods.X64_OFF_R15), (uint)Marshal.ReadInt32(buf, NativeMethods.X64_OFF_EFlags));
            }
            return buf;
        }

        private static ulong ReadPtrReg(IntPtr buf, int offset)
        {
            return unchecked((ulong)Marshal.ReadInt64(buf, offset));
        }

        private uint ReadContextFlags(IntPtr buf)
        {
            try { return (uint)Marshal.ReadInt32(buf, _is32Bit ? 0 : 48); }
            catch { return 0; }
        }

        /// <summary>
        /// 从 RSP 起读 4KB，把"落在某个模块映像范围内"的 8 字节值当作候选返回地址。
        ///
        /// 这是**启发式**，不是严格的栈回溯（严格回溯需要 unwind 信息；release 构建普遍省略
        /// 帧指针，RBP 链走不通）。但对"0x183ba19 到底在哪条路径上"这个问题来说，
        /// 一份"调用点候选地址 + 它们各自的模块/RVA"列表已经足够把人引到正确的地方，
        /// 而且它不依赖符号、不依赖网络、当秒就有。
        /// </summary>
        private List<ulong> ScanStack(uint tid, IntPtr ctxBuf)
        {
            var found = new List<ulong>();
            try
            {
                ulong rsp = _is32Bit ? (uint)Marshal.ReadInt32(ctxBuf, 196) : ReadPtrReg(ctxBuf, NativeMethods.X64_OFF_Rsp);
                int span = 4096;
                ulong rip = _is32Bit ? (uint)Marshal.ReadInt32(ctxBuf, 184) : ReadPtrReg(ctxBuf, NativeMethods.X64_OFF_Rip);

                // RIP 本身作为第 0 帧
                if (rip != 0) found.Add(rip);

                byte[] buf = new byte[span];
                IntPtr read;
                if (!NativeMethods.ReadProcessMemory(_hProcess, (IntPtr)unchecked((long)rsp), buf, (IntPtr)span, out read))
                {
                    _log.Warn("ReadProcessMemory(RSP=0x" + rsp.ToString("X") + ") 失败，跳过栈扫描");
                    return found;
                }

                int n = (int)read.ToInt64();
                int step = _is32Bit ? 4 : 8;
                ulong prev = ulong.MaxValue;
                for (int off = 0; off + step <= n; off += step)
                {
                    ulong v = _is32Bit
                        ? BitConverter.ToUInt32(buf, off)
                        : BitConverter.ToUInt64(buf, off);
                    if (v < 0x10000) continue;                       // 明显不是代码地址
                    if (v == prev) continue;                          // 同一地址连续重复，去掉噪声
                    // 用 Resolve：模块表命中就查表，没命中再走 VirtualQueryEx 兜底（不写日志免得刷屏）
                    ModuleEntry m = ModuleMap.Resolve(_hProcess, _modules, v, null);
                    if (m == null) continue;                          // 不落在任何映像里 → 大概率是数据
                    prev = v;
                    found.Add(v);
                    if (found.Count >= 60) break;
                }
            }
            catch (Exception ex)
            {
                _log.Warn("栈扫描异常（忽略）：" + ex.GetType().Name + " " + ex.Message);
            }
            return found;
        }

        // ------------------------------------------------------------------ 符号异步收尾

        private void QueueSymbolResolution()
        {
            if (_o.SymbolBudgetSeconds <= 0)
            {
                _log.Info("[SYM] 已禁用符号解析（--symbol-budget 0）。模块名+RVA 已在上面，足够定位。");
                return;
            }

            List<ulong> addresses;
            lock (_symbolGate)
            {
                addresses = new List<ulong>(_symbolAddresses);
            }
            if (addresses.Count == 0) return;

            // 调试会话的目标句柄在进程存活期间一直有效，符号线程直接复用（不再 OpenProcess，
            // 少一个句柄、也少一处需要清理的失败路径）。
            _hProcessForSymbols = _hProcess;
            IntPtr hForSym = _hProcess;
            uint pid = _pid;
            List<ModuleEntry> mods = _modules;
            string cache = _o.SymbolCacheDir;
            bool net = _o.SymbolServerReachable;
            WatchdogLog log = _log;

            _symbolThread = new Thread(() =>
            {
                try { SymbolResolver.ResolveAll(pid, hForSym, mods, addresses, cache, net, log); }
                catch (Exception ex) { log.Warn("[SYM] 线程顶层异常（忽略）：" + ex.GetType().Name + " " + ex.Message); }
            }, 512 * 1024);
            _symbolThread.IsBackground = true;
            _symbolThread.Name = "SymbolResolver";
            _symbolThread.Start();

            _log.Info("[SYM] 符号解析已在后台线程启动（预算 " + _o.SymbolBudgetSeconds + "s）；"
                      + "它不阻塞调试循环，拿不到符号也不影响取证。");
        }

        private void WaitForSymbolThread(int seconds)
        {
            Thread t = _symbolThread;
            if (t == null || !t.IsAlive) return;

            _log.Info("[SYM] 等待符号线程收尾（最多 " + seconds + "s，超时就放弃 —— 模块名+RVA 已经写在主日志里了）");
            bool done = t.Join(TimeSpan.FromSeconds(seconds));
            if (done) _log.Info("[SYM] 符号线程已结束");
            else _log.Warn("[SYM] 符号线程超预算未结束，直接退出（后台线程随进程结束）。位置证据不受影响。");
        }

        // ------------------------------------------------------------------ 句柄纪律

        private void CloseThreadHandle(IntPtr h, uint tid, string what)
        {
            if (h == IntPtr.Zero) return;
            lock (_handleBook)
            {
                if (!_closedThreadIds.Add(tid))
                {
                    _handlesSkipped++;
                    _log.Info(what + "：tid=" + tid + " 的句柄已关闭过（CREATE_PROCESS 与 CREATE_THREAD 会重复投递同一初始线程），跳过以防二次关闭");
                    return;
                }
            }
            // hThread 只在"本进程即将使用"时才需要。我们不用它执行任何操作，
            // 所以立刻关闭 —— Word 有几十个线程，不关就是几十个句柄泄漏。
            if (NativeMethods.CloseHandle(h)) _handlesClosed++;
        }

        private void CloseThreadOrFileHandle(IntPtr h, string what)
        {
            if (h == IntPtr.Zero) return;
            if (NativeMethods.CloseHandle(h)) _handlesClosed++;
            else _log.Warn(what + " CloseHandle 失败 win32=" + Marshal.GetLastWin32Error());
        }

        private void CleanupHandles()
        {
            try
            {
                uint selfPid = _pid;
                if (!NativeMethods.DebugActiveProcessStop((int)selfPid))
                {
                    // 目标已经退出时会失败，属正常
                }
            }
            catch { }

            if (_hProcess != IntPtr.Zero)
            {
                NativeMethods.CloseHandle(_hProcess);
                _hProcess = IntPtr.Zero;
            }
        }

        // ------------------------------------------------------------------ 文本助手

        private static string DescribeAccess(ulong info0)
        {
            var sb = new StringBuilder();
            ulong kind = info0 & 0xFF;
            if (kind == NativeMethods.EXCEPTION_READ_FAULT) sb.Append("读取(read)");
            else if (kind == NativeMethods.EXCEPTION_WRITE_FAULT) sb.Append("写入(write)");
            else if (kind == NativeMethods.EXCEPTION_EXECUTE_FAULT) sb.Append("执行(execute)");
            else sb.Append("未知类型 " + kind);

            if ((info0 & 0x100) != 0) sb.Append(" + PAGE_GUARD");
            if ((info0 & 0x200) != 0) sb.Append(" + PAGE_COPY_ON_WRITE");
            if ((info0 & 0x400) != 0) sb.Append(" + PAGE_EXECUTE");
            if ((info0 & 0x800) != 0) sb.Append(" + PAGE_NOACCESS");
            return sb.ToString();
        }

        private static string DescribeAccessShort(ulong info0)
        {
            ulong kind = info0 & 0xFF;
            if (kind == 0) return "read";
            if (kind == 1) return "write";
            if (kind == 8) return "execute";
            return "unknown(" + kind + ")";
        }

        private static string ExceptionText(uint code)
        {
            switch (code)
            {
                case 0xC0000005: return "STATUS_ACCESS_VIOLATION 访问违例";
                case 0xC00000FD: return "STATUS_STACK_OVERFLOW 栈溢出";
                case 0xC0000409: return "STATUS_FAIL_FAST_EXCEPTION 快速失败";
                case 0xC0000374: return "STATUS_HEAP_CORRUPTION 堆损坏";
                case 0xC000001D: return "STATUS_ILLEGAL_INSTRUCTION 非法指令";
                case 0xC0000094: return "STATUS_INTEGER_DIVIDE_BY_ZERO 整数除零";
                case 0x80000003: return "STATUS_BREAKPOINT 断点";
                case 0x80000004: return "STATUS_SINGLE_STEP 单步";
                case 0xE06D7363: return "C++ 异常 (msc)";
                case 0xE0434352: return "CLR 托管异常";
                case 0xC0000135: return "DLL 未找到";
                case 0xC0000142: return "DLL 初始化失败";
                case 0x40000015: return "STATUS_FATAL_APP_EXIT";
                default: return "未知异常码";
            }
        }

        private static string DescribeExitCode(uint code)
        {
            switch (code)
            {
                case 0: return "退出码 0 = 正常退出，本次会话目标进程没有崩溃。";
                case 0xC0000005: return "退出码 0xC0000005 = 访问违例终止 —— 与 wwlib.dll 0xc0000005 崩溃签名一致。";
                case 0xC00000FD: return "退出码 0xC00000FD = 栈溢出终止。";
                case 0xC0000409: return "退出码 0xC0000409 = fail-fast 终止。";
                case 0xE0434352: return "退出码 0xE0434352 = .NET 托管异常未处理终止（CLR 层，与 wwlib 原生 AV 是两类问题）。";
                case 0x80131506: return "退出码 0x80131506 = CLR 内部错误（交接报告第五节第 1 条记录过：COM AutoDual 会导致这个）。";
                default: return "退出码 0x" + code.ToString("X8") + "（未归类）。";
            }
        }

        private static string Win32Text(int err)
        {
            try { return new System.ComponentModel.Win32Exception(err).Message; }
            catch { return "?"; }
        }

        private string DiagnoseOpenProcessFailure(int err, int pid)
        {
            if (err == 87) // ERROR_INVALID_PARAMETER
                return "win32=87 参数错误 → 通常表示 pid " + pid + " 不存在（进程已退出）。"
                       + "请确认 Word 还开着，且 addin 传进来的是当前进程 ID。";
            if (err == 5)
                return "win32=5 拒绝访问 → 权限不足。若 Word 以管理员运行而看门狗不是，会走到这里。"
                       + "解决办法：让看门狗与 Word 处于同一完整性级别（看门狗清单是 asInvoker，会继承 Word 的权限）。";
            return "win32=" + err + "（" + Win32Text(err) + "）。";
        }

        private static string DiagnoseDebugActiveFailure(int err)
        {
            if (err == 5) return "win32=5 拒绝访问 → 权限不足 / 已在被另一个调试器调试 / 目标是受保护进程。";
            if (err == 87) return "win32=87 参数错误 → 进程不存在或已退出。";
            if (err == 2) return "win32=2 找不到文件 → 进程不存在。";
            return "win32=" + err + "（" + Win32Text(err) + "）。";
        }

        private string TryGetProcessName(uint pid)
        {
            try
            {
                using (var p = System.Diagnostics.Process.GetProcessById((int)pid))
                {
                    return p.ProcessName + ".exe";
                }
            }
            catch { return null; }
        }

        private string TryGetProcessPath(uint pid)
        {
            try
            {
                using (var p = System.Diagnostics.Process.GetProcessById((int)pid))
                {
                    return p.MainModule != null ? p.MainModule.FileName : null;
                }
            }
            catch { return null; }
        }

        private static string Quote(string s)
        {
            return s.IndexOf(' ') >= 0 ? "\"" + s + "\"" : s;
        }

        private static string FormatBytes(long bytes)
        {
            if (bytes < 1024) return bytes + " B";
            if (bytes < 1024 * 1024) return (bytes / 1024.0).ToString("F1") + " KB";
            if (bytes < 1024L * 1024 * 1024) return (bytes / 1024.0 / 1024).ToString("F1") + " MB";
            return (bytes / 1024.0 / 1024 / 1024).ToString("F2") + " GB";
        }
    }

    internal static class ExitCodes
    {
        public const int Ok = 0;
        public const int BadArguments = 2;
        public const int LogUnavailable = 3;
        public const int AttachFailed = 4;
        public const int LaunchFailed = 5;
        public const int LayoutHardFail = 6;
        public const int DebugLoopError = 7;
    }
}
