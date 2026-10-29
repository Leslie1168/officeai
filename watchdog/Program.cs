using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

namespace OfficeAi.Watchdog
{
    /// <summary>命令行选项。契约见 PrintHelp()。</summary>
    internal sealed class Options
    {
        public int Pid;
        public string ExePath;
        public string ChildArgs;
        public string OutDir;

        public bool SelfTest;
        public bool HelpRequested;
        public string WorkerMode;      // crash / crash-exec / crash-fastfail / sleep / noop
        public int WorkerDelayMs = 1500;

        /// <summary>
        /// MINIDUMP_EXCEPTION_INFORMATION.ClientPointers。
        /// 注意：本机 dbghelp 的 ExceptionParam 通道整体是坏的（见 DumpWriter 注释），
        /// 所以这个值只对 --try-exception-param 那一次"赌一把"的尝试有意义，默认路径根本不传异常信息。
        /// </summary>
        public bool ClientPointers = true;

        /// <summary>
        /// 是否额外尝试带 ExceptionParam 的 dump。
        /// 默认 **false**：本机实测该路径必然失败，而且会让调用进程直接崩
        /// （probe 里 T3.dmp = 0 字节、进程随后消失）。看门狗死掉就意味着丢掉整个现场，
        /// 所以默认不碰；换台 dbghelp 正常的机器可以打开。
        /// 打开时也一定是在"安全的那份 dump 已经落盘之后"才试。
        /// </summary>
        public bool TryExceptionParam;

        public bool SymbolServerReachable;
        public string SymbolCacheDir = SymbolResolver.DefaultCacheDir;
        public int SymbolBudgetSeconds = 20;
        public int SymbolWaitSeconds = 15;
        public int ProbeTimeoutMs = 2000;

        public string LogPath { get { return Path.Combine(OutDir, "watchdog.log"); } }

        public static string DefaultOutDir
        {
            get
            {
                return Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "OfficeAi", "crashes");
            }
        }
    }

    internal static class Program
    {
        private static int Main(string[] args)
        {
            // ★ 第一原则（Lead 明确要求）：看门狗的价值在于"一定能跑起来"。
            //   任何抛到这里的异常都必须先被写进日志文件，而不是只留在 stderr ——
            //   因为宿主是 CreateNoWindow 拉起的，stderr 根本没人看。
            WatchdogLog log = null;
            Options o = null;
            try
            {
                HookUnhandled();

                string parseError;
                o = ParseArgs(args, out parseError);

                if (o == null)
                {
                    // 参数错误时连 --out 都可能拿不到，退回默认目录写日志，尽量留痕
                    string dir = ExtractOutDir(args) ?? Options.DefaultOutDir;
                    log = WatchdogLog.TryOpen(Path.Combine(dir, "watchdog.log"));
                    if (log != null)
                    {
                        log.Error("参数解析失败：" + (parseError ?? "未知"));
                        log.Error("完整命令行：" + JoinArgs(args));
                        log.Error("用法见 --help；正式契约：OfficeAi.Watchdog.exe --pid <pid> --out <dir>");
                    }
                    SafeStderr("参数解析失败：" + parseError + Environment.NewLine + HelpText());
                    return ExitCodes.BadArguments;
                }

                if (o.HelpRequested) { SafeStderr(HelpText()); return ExitCodes.Ok; }

                log = WatchdogLog.TryOpen(o.LogPath);
                if (log == null)
                {
                    // 日志都写不了，看门狗就没有存在意义 —— 但仍然把原因吐到 stderr 并给出非零退出码
                    SafeStderr("无法创建日志文件：" + o.LogPath);
                    return ExitCodes.LogUnavailable;
                }

                StartupGuard.Attach(log);
                StartupGuard.SetPhase("程序入口，参数解析完成");

                log.Banner("=========== OfficeAi.Watchdog 启动 ===========");
                log.Info("版本       : " + typeof(Program).Assembly.GetName().Version);
                log.Info("自身 PID   : " + NativeMethods.GetCurrentProcessId());
                log.Info("命令行     : " + JoinArgs(args));
                log.Info("输出目录   : " + o.OutDir);
                log.Info("日志文件   : " + o.LogPath);
                log.Info("64 位进程  : " + (IntPtr.Size == 8) + "（看门狗必须与被调试的 Word 同位数）");
                log.Info("ClientPointers = " + o.ClientPointers + "（MINIDUMP_EXCEPTION_INFORMATION 语义，见 DumpWriter 注释）");

                // ---- 结构体布局自检：只记录，不作为启动门槛 ----
                StartupGuard.SetPhase("结构体布局自检");
                var layout = NativeMethods.CheckLayouts();
                if (layout.Ok)
                {
                    log.Info("结构体布局自检：全部通过（EXCEPTION_RECORD=152 DEBUG_EVENT=176 CONTEXT=1232 SYMBOL_INFOW=4088 等）");
                }
                else
                {
                    foreach (string s in layout.Hard) log.Error("结构体布局【硬性错位】" + s);
                    foreach (string s in layout.Soft) log.Warn("结构体布局【软性错位】" + s);
                    if (!layout.CanDump)
                    {
                        log.Error("★ 存在硬性布局错位：dump 内容可能是错的假证据。");
                        log.Error("★ 依 Lead 要求（看门狗必须能跑起来）：进程继续运行，但会把 dump 写入标记为不可信，");
                        log.Error("★ 并在下面每一步照常记录模块名/RVA —— 那部分不依赖这些结构体。");
                    }
                }

                if (o.WorkerMode == null && !o.SelfTest)
                {
                    StartupGuard.SetPhase("符号服务器可达性预检（决定后面会不会卡住）");
                    o.SymbolServerReachable = o.SymbolBudgetSeconds > 0
                        && SymbolResolver.ProbeSymbolServer(o.ProbeTimeoutMs, log);
                }

                StartupGuard.SetPhase("构造调试器");
                var dbg = new WatchdogDebugger(o, log);
                int rc = dbg.Run();

                log.Banner("=========== OfficeAi.Watchdog 结束，退出码 " + rc + " ===========");
                if (dbg.DumpPath != null) log.Banner("本次 dump：" + dbg.DumpPath);
                log.Dispose();
                return rc;
            }
            catch (Exception ex)
            {
                // 逐字段手写异常信息：TypeLoadException 这类会让 ex.ToString() 自己再炸一次，
                // 那样就只剩下一句"由于 Exception.ToString() 失败" —— 正是之前查不出原因的原因。
                try
                {
                    string detail = DescribeException(ex);
                    if (log != null) log.Error("顶层未处理异常：" + detail);
                    if (log != null) log.Error("启动阶段卡在：" + StartupGuard.Phase);
                    if (log != null) log.Dispose();
                    SafeStderr("看门狗顶层异常：" + detail + Environment.NewLine + "启动阶段：" + StartupGuard.Phase);
                }
                catch { }
                return ExitCodes.DebugLoopError;
            }
        }

        private static string DescribeException(Exception ex)
        {
            var sb = new StringBuilder();
            int depth = 0;
            while (ex != null && depth < 5)
            {
                string type, msg, st;
                try { type = ex.GetType().FullName; } catch (Exception e) { type = "<类型加载失败: " + e.GetType().Name + ">"; }
                try { msg = ex.Message; } catch (Exception e) { msg = "<消息取值失败: " + e.GetType().Name + ">"; }
                try { st = ex.StackTrace; } catch { st = null; }

                sb.Append((depth == 0 ? "" : "  → 由 ") + type + ": " + msg);
                if (!string.IsNullOrEmpty(st))
                {
                    sb.Append(Environment.NewLine + "      堆栈: " + st.Replace("\r", "").Replace("\n", " | "));
                }
                sb.AppendLine();
                try { ex = ex.InnerException; } catch { ex = null; }
                depth++;
            }
            return sb.ToString();
        }

        private static void HookUnhandled()
        {
            // 兜底：非 Main 线程上的未处理异常也试图留痕（日志句柄是静态可达的，见下）
            AppDomain.CurrentDomain.UnhandledException += (s, e) =>
            {
                try { SafeStderr("看门狗后台线程未处理异常：" + DescribeException(e.ExceptionObject as Exception)); }
                catch { }
            };
        }

        private static void SafeStderr(string text)
        {
            try { Console.Error.WriteLine(text); } catch { }
        }

        // ------------------------------------------------------------------ 参数解析

        private static Options ParseArgs(string[] args, out string error)
        {
            error = null;
            var o = new Options();
            o.OutDir = Options.DefaultOutDir;

            for (int i = 0; i < args.Length; i++)
            {
                string a = args[i];
                string val;
                switch (a.ToLowerInvariant())
                {
                    case "--pid":
                        if (!TryValue(args, ref i, out val)) { error = "--pid 缺少值"; return null; }
                        int pid;
                        if (!int.TryParse(val, out pid) || pid <= 0) { error = "--pid 值非法：" + val; return null; }
                        o.Pid = pid;
                        break;

                    case "--exe":
                        if (!TryValue(args, ref i, out val)) { error = "--exe 缺少值"; return null; }
                        o.ExePath = val;
                        break;

                    case "--args":
                        if (!TryValue(args, ref i, out val)) { error = "--args 缺少值"; return null; }
                        o.ChildArgs = val;
                        break;

                    case "--out":
                        if (!TryValue(args, ref i, out val)) { error = "--out 缺少值"; return null; }
                        o.OutDir = val;
                        break;

                    case "--client-pointers":
                        if (!TryValue(args, ref i, out val)) { error = "--client-pointers 缺少值"; return null; }
                        o.ClientPointers = !(val == "0" || val.Equals("false", StringComparison.OrdinalIgnoreCase));
                        break;

                    case "--try-exception-param":
                        o.TryExceptionParam = true;
                        break;

                    case "--symbol-cache":
                        if (!TryValue(args, ref i, out val)) { error = "--symbol-cache 缺少值"; return null; }
                        o.SymbolCacheDir = val;
                        break;

                    case "--symbol-budget":
                        if (!TryValue(args, ref i, out val)) { error = "--symbol-budget 缺少值"; return null; }
                        int b;
                        if (!int.TryParse(val, out b)) { error = "--symbol-budget 值非法：" + val; return null; }
                        o.SymbolBudgetSeconds = b;
                        break;

                    case "--symbol-wait":
                        if (!TryValue(args, ref i, out val)) { error = "--symbol-wait 缺少值"; return null; }
                        int w;
                        if (!int.TryParse(val, out w)) { error = "--symbol-wait 值非法：" + val; return null; }
                        o.SymbolWaitSeconds = w;
                        break;

                    case "--probe-timeout":
                        if (!TryValue(args, ref i, out val)) { error = "--probe-timeout 缺少值"; return null; }
                        int p;
                        if (!int.TryParse(val, out p)) { error = "--probe-timeout 值非法：" + val; return null; }
                        o.ProbeTimeoutMs = p;
                        break;

                    case "--worker":
                        if (!TryValue(args, ref i, out val)) { error = "--worker 缺少值"; return null; }
                        o.WorkerMode = val;
                        break;

                    case "--delay":
                        if (!TryValue(args, ref i, out val)) { error = "--delay 缺少值"; return null; }
                        int d;
                        if (!int.TryParse(val, out d)) { error = "--delay 值非法：" + val; return null; }
                        o.WorkerDelayMs = d;
                        break;

                    case "--selftest":
                        o.SelfTest = true;
                        break;

                    case "--help":
                    case "-h":
                    case "/?":
                        o.HelpRequested = true;
                        return o;

                    default:
                        error = "未知参数：" + a;
                        return null;
                }
            }

            if (!o.SelfTest && o.WorkerMode == null && o.Pid <= 0 && string.IsNullOrEmpty(o.ExePath))
            {
                error = "必须指定 --pid <pid> 或 --exe <path>（或 --selftest / --worker）";
                return null;
            }
            if (o.Pid > 0 && !string.IsNullOrEmpty(o.ExePath))
            {
                error = "--pid 与 --exe 互斥，只能给一个";
                return null;
            }
            return o;
        }

        private static bool TryValue(string[] args, ref int i, out string value)
        {
            if (i + 1 >= args.Length) { value = null; return false; }
            value = args[++i];
            return true;
        }

        private static string ExtractOutDir(string[] args)
        {
            for (int i = 0; i < args.Length - 1; i++)
            {
                if (args[i].Equals("--out", StringComparison.OrdinalIgnoreCase)) return args[i + 1];
            }
            return null;
        }

        private static string JoinArgs(string[] args)
        {
            var sb = new StringBuilder();
            for (int i = 0; i < args.Length; i++)
            {
                if (i > 0) sb.Append(' ');
                if (args[i].IndexOf(' ') >= 0) sb.Append('"').Append(args[i]).Append('"');
                else sb.Append(args[i]);
            }
            return sb.ToString();
        }

        public static string HelpText()
        {
            return string.Join(Environment.NewLine, new[]
            {
                "OfficeAi.Watchdog —— Office AI 助手插件的进程外崩溃看门狗",
                "",
                "主用（addin 拉起本程序时用的就是这个契约，不要改）：",
                "  OfficeAi.Watchdog.exe --pid <PID> --out <dump目录>",
                "",
                "备用（以调试方式启动一个进程，无需 addin）：",
                "  OfficeAi.Watchdog.exe --exe <可执行文件> [--args \"<子进程参数>\"] [--out <目录>]",
                "",
                "自检（不需要 Word，自己制造崩溃并验证整条取证链）：",
                "  OfficeAi.Watchdog.exe --selftest [--out <目录>]",
                "",
                "可选参数：",
                "  --try-exception-param           额外尝试带 ExceptionParam 的 dump（默认关）。",
                "                                  ★ 本机 System32\\dbghelp.dll 在这条路径上必然失败，并会让调用进程崩溃",
                "                                    （实测 T3 产出 0 字节 dump 后进程消失），所以默认不碰；",
                "                                    开的时候也只在安全的 NULL 版本 dump 落盘**之后**才试。",
                "  --client-pointers <true|false>  仅对 --try-exception-param 那次尝试有意义（默认 true）",
                "  --symbol-budget <秒>            符号解析预算，0 = 完全跳过符号（默认 20）",
                "  --symbol-wait <秒>              退出前等待符号线程的上限（默认 15）",
                "  --symbol-cache <目录>           符号缓存目录（默认 %LOCALAPPDATA%\\OfficeAi\\symbols）",
                "  --probe-timeout <毫秒>          符号服务器可达性预检超时（默认 2000）",
                "  --help                         显示本帮助",
                "",
                "产物（每次取证两份，互相配对）：",
                "  <out>\\<proc>_<pid>_0x<码>_0x<地址>_<时间>.dmp            全内存 minidump（ExceptionParam=NULL 写法）",
                "  <out>\\<同名>.exception.txt                              异常现场 sidecar：异常码/异常地址/模块名+基址+RVA/",
                "                                                         访问类型/出错线程ID/寄存器/栈候选 + cdb 命令",
                "  <out>\\watchdog.log                                      完整时间线日志（无缓冲逐行落盘）",
                "  ★ 关键：即使 dump 写不出来，sidecar 与日志里的「模块名+RVA」依然有效 —— 定位 wwlib+0x183BA19 靠的是它。",
                "",
                "行为承诺：",
                "  · 只取证，不改变被调试进程行为：所有异常都以 DBG_EXCEPTION_NOT_HANDLED 原样放行。",
                "  · DebugSetProcessKillOnExit(FALSE)：看门狗退出绝不会连带杀掉 Word。",
                "  · 日志写到 <out>\\watchdog.log（无缓冲逐行落盘，因为宿主是无窗口启动的，stdout 无人接收）。",
                "  · 符号服务器不可达时自动降级为本地缓存，绝不因为下载符号卡死。",
                "",
                "退出码：0=正常 2=参数错误 3=日志不可写 4=attach 失败 5=启动目标失败 6=保留 7=顶层异常",
            });
        }
    }
}
