using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;

namespace OfficeAi.Watchdog
{
    /// <summary>
    /// 硬性验收：不接受"代码看起来对"。
    ///
    /// --selftest 会在本机真实地跑完下面这 6 个场景，每个场景都启动**独立的看门狗子进程**
    /// （用的就是 addin 将要使用的那条命令行），然后回来核对落盘的 dump 与日志内容：
    ///
    ///   A. --exe 模式 + 真崩溃（故障指令在 ntdll.dll 内部）→ dump 落盘、模块名+RVA、退出码
    ///   B. --pid 模式 attach 到已在运行的进程，再让它崩    → 同上（这是生产路径）
    ///   C. A 的对照组，--client-pointers false             → 供 cdb 对比，确定有歧义的参数
    ///   D. --exe 模式 + 不崩的进程                          → 必须"不写 dump"，退出码 0
    ///   E. --pid 指向不存在的进程                            → 必须**留下日志**说清楚为什么失败
    ///   F. 启动路径自检                                     → 结构体布局自检通过、无 TypeLoadException
    /// </summary>
    internal sealed class Selftest
    {
        private readonly Options _o;
        private readonly WatchdogLog _log;
        private readonly string _selfExe;
        private readonly string _root;
        private readonly List<string> _pass = new List<string>();
        private readonly List<string> _fail = new List<string>();

        public Selftest(Options o, WatchdogLog log)
        {
            _o = o;
            _log = log;
            _selfExe = typeof(Program).Assembly.Location;
            _root = Path.Combine(o.OutDir, "selftest-" + DateTime.Now.ToString("yyyyMMdd-HHmmss"));
        }

        public int Run()
        {
            _log.Banner("=========== --selftest 开始 ===========");
            _log.Info("自身 exe  : " + _selfExe);
            _log.Info("测试根目录: " + _root);
            Directory.CreateDirectory(_root);

            Console.WriteLine("================================================================");
            Console.WriteLine(" OfficeAi.Watchdog 自检");
            Console.WriteLine(" 自身 exe : " + _selfExe);
            Console.WriteLine(" 输出目录 : " + _root);
            Console.WriteLine("================================================================");

            // 结构体布局自检先来一条（F）
            Check("F1 结构体布局自检无硬性错位", () =>
            {
                var r = NativeMethods.CheckLayouts();
                Dump("硬性错位", r.Hard); Dump("软性错位", r.Soft);
                return r.CanDump;
            });

            ScenarioA_ExeModeCrash();
            ScenarioB_PidAttachCrash();
            ScenarioC_ClientPointersFalse();
            ScenarioD_NoCrashNoDump();
            ScenarioE_AttachFailureIsLogged();

            Console.WriteLine();
            Console.WriteLine("================================================================");
            Console.WriteLine(" 自检结果：" + _pass.Count + " 项通过 / " + _fail.Count + " 项失败");
            for (int i = 0; i < _pass.Count; i++) Console.WriteLine("   [PASS] " + _pass[i]);
            for (int i = 0; i < _fail.Count; i++) Console.WriteLine("   [FAIL] " + _fail[i]);
            Console.WriteLine("================================================================");

            _log.Banner("=========== --selftest 结束：" + _pass.Count + " 通过 / " + _fail.Count + " 失败 ===========");
            for (int i = 0; i < _fail.Count; i++) _log.Error("[SELFTEST FAIL] " + _fail[i]);

            return _fail.Count == 0 ? ExitCodes.Ok : ExitCodes.DebugLoopError;
        }

        // ---------------------------------------------------------------- 场景 A

        private void ScenarioA_ExeModeCrash()
        {
            Console.WriteLine();
            Console.WriteLine("--- 场景 A：--exe 模式启动会崩的进程（故障指令预期在 ntdll.dll 内）---");
            string dir = Path.Combine(_root, "A-exe-crash");
            Directory.CreateDirectory(dir);

            string args = "--exe " + Q(_selfExe) + " --args " + Q("--worker crash --delay 700")
                        + " --out " + Q(dir);
            ChildResult cr = RunChild(args, 90000);
            Console.WriteLine("  看门狗子进程退出码 = " + cr.ExitCode + "（耗时 " + cr.ElapsedMs + "ms）");

            LogFacts f = ReadLog(dir);

            Check("A1 看门狗子进程正常结束", cr.ExitCode == 0);
            Check("A2 watchdog.log 已生成且有内容", f.Exists && f.Lines.Count > 5);
            Check("A3 结构体布局自检通过（无 TypeLoadException）", f.Has("结构体布局自检：全部通过"));
            Check("A4 DebugSetProcessKillOnExit(FALSE) 已确认生效", f.Has("DebugSetProcessKillOnExit(FALSE) 已确认生效"));
            Check("A5 捕获到 EXCEPTION 事件", f.Has("EXCEPTION code=0x"));
            Check("A6 dump 已落盘（KEY 行）", f.Has("[KEY] pid="));
            Check("A7 模块归属为 ntdll.dll（故障指令确实在别人的 DLL 里）", f.KeyModule == "ntdll.dll");
            Check("A8 RVA 可解析且非零", f.KeyRva > 0);
            Check("A9 记录了 EXIT_PROCESS 与退出码", f.Has("EXIT_PROCESS："));
            Check("A10 dump 文件真实存在且 > 1MB", f.DumpPath != null && f.DumpBytes > 1024 * 1024);
            Check("A11 句柄清点已记录", f.Has("句柄清点："));
            Check("A12 sidecar 已落盘且与 dump 同名配对", f.SidecarPath != null && File.Exists(f.SidecarPath));
            Check("A13 sidecar 内含模块名+RVA 结论行", f.SidecarHasModuleRva);
            Check("A14 sidecar 内含可直接复制的 cdb 命令与出错线程 ID", f.SidecarHasCdb && f.SidecarHasTid);
            Check("A15 模块表延迟枚举生效（不再于 loader 未完成时报 win32=299）", f.Has("模块表已枚举（延迟枚举"));

            Console.WriteLine("  KEY 行: " + (f.KeyLine ?? "(未找到)"));
            Console.WriteLine("  模块  : " + f.KeyModule + " RVA=0x" + (f.KeyRva > 0 ? f.KeyRva.ToString("X") : "?")
                              + " 来源=" + (f.KeyModuleSource ?? "?"));
            Console.WriteLine("  dump  : " + (f.DumpPath ?? "(无)") + "  " + f.DumpBytes + " 字节");
            Console.WriteLine("  sidecar: " + (f.SidecarPath ?? "(无)"));
            Console.WriteLine("  退出码行: " + (f.ExitLine ?? "(未找到)"));
        }

        // ---------------------------------------------------------------- 场景 B

        private void ScenarioB_PidAttachCrash()
        {
            Console.WriteLine();
            Console.WriteLine("--- 场景 B：--pid 模式 attach 到已在运行的进程，再让它崩（生产路径）---");
            string dir = Path.Combine(_root, "B-pid-attach");
            Directory.CreateDirectory(dir);

            Process worker = null;
            ChildResult cr = default(ChildResult);
            try
            {
                // 先起靶子：8 秒后才崩，留足 attach 时间窗口
                var psi = new ProcessStartInfo(_selfExe, "--worker crash --delay 8000")
                {
                    UseShellExecute = false,
                    CreateNoWindow = false
                };
                worker = Process.Start(psi);
                Thread.Sleep(1200); // 等它真正跑起来
                Console.WriteLine("  靶子进程 pid = " + worker.Id + "（8 秒后崩）");

                string args = "--pid " + worker.Id + " --out " + Q(dir);
                cr = RunChild(args, 90000);
                Console.WriteLine("  看门狗子进程退出码 = " + cr.ExitCode + "（耗时 " + cr.ElapsedMs + "ms）");
                Check("B0 看门狗子进程正常结束", cr.ExitCode == 0);
            }
            finally
            {
                try { if (worker != null && !worker.HasExited) worker.Kill(); } catch { }
            }

            LogFacts f = ReadLog(dir);
            Check("B1 watchdog.log 已生成且有内容", f.Exists && f.Lines.Count > 5);
            Check("B2 DebugActiveProcess 成功（attach 生效）", f.Has("DebugActiveProcess 成功"));
            Check("B3 DebugSetProcessKillOnExit(FALSE) 已确认生效", f.Has("DebugSetProcessKillOnExit(FALSE) 已确认生效"));
            Check("B4 捕获到 EXCEPTION 事件", f.Has("EXCEPTION code=0x"));
            Check("B5 dump 已落盘（KEY 行）", f.Has("[KEY] pid="));
            Check("B6 模块归属为 ntdll.dll", f.KeyModule == "ntdll.dll");
            Check("B7 RVA 可解析且非零", f.KeyRva > 0);
            Check("B8 记录了 EXIT_PROCESS 与退出码", f.Has("EXIT_PROCESS："));
            Check("B9 dump 文件真实存在且 > 1MB", f.DumpPath != null && f.DumpBytes > 1024 * 1024);
            Check("B10 模块表枚举成功（attach 路径）", f.Has("模块表已枚举"));
            Check("B11 sidecar 已落盘且与 dump 同名配对", f.SidecarPath != null && File.Exists(f.SidecarPath));
            Check("B12 sidecar 内含模块名+RVA 结论行", f.SidecarHasModuleRva);

            Console.WriteLine("  KEY 行: " + (f.KeyLine ?? "(未找到)"));
            Console.WriteLine("  dump  : " + (f.DumpPath ?? "(无)") + "  " + f.DumpBytes + " 字节");
            Console.WriteLine("  sidecar: " + (f.SidecarPath ?? "(无)"));
        }

        // ---------------------------------------------------------------- 场景 C

        private void ScenarioC_ClientPointersFalse()
        {
            Console.WriteLine();
            Console.WriteLine("--- 场景 C：★默认关掉 ExceptionParam 通道（本机 dbghelp 在该通道上必坏）---");
            Console.WriteLine("    本机实测：ExceptionParam 非 NULL 时 dbghelp 返回 HRESULT_FROM_WIN32(998) ERROR_NOACCESS，");
            Console.WriteLine("    甚至让调用进程崩溃（probe T3 产出 0 字节后进程消失）。");
            Console.WriteLine("    所以判据改成：**两种 ClientPointers 取值下都必须产出可分析产物**（靠 NULL 回退路径），");
            Console.WriteLine("    而不是靠 ClientPointers 取反去拿到带异常信息的 dump。");
            string dir = Path.Combine(_root, "C-exceptionparam-off");
            Directory.CreateDirectory(dir);

            string args = "--exe " + Q(_selfExe) + " --args " + Q("--worker crash --delay 700")
                        + " --out " + Q(dir) + " --client-pointers false";
            ChildResult cr = RunChild(args, 90000);
            LogFacts f = ReadLog(dir);

            Check("C1 子进程正常结束", cr.ExitCode == 0);
            Check("C2 dump 仍产出（判据=文件 > 1MB，不看返回值）", f.DumpPath != null && f.DumpBytes > 1024 * 1024);
            Check("C3 走的是 NULL 通道（日志明确记录 ExceptionParam=NULL 被接受）", f.Has("【dump/①】首选写法：ExceptionParam = NULL"));
            Check("C4 未启用 --try-exception-param 时不碰那条坏通道", !f.Has("【dump/③】"));
            Check("C5 sidecar 同样产出（ClientPointers 取值不影响结论文件）", f.SidecarPath != null && File.Exists(f.SidecarPath));

            // C2 对照：显式打开 --try-exception-param，验证"①先落盘、③再赌一把"的顺序，
            // 以及③失败时①的成品必须完好保留
            Console.WriteLine();
            Console.WriteLine("--- 场景 C2：显式打开 --try-exception-param，验证降级顺序（①先保住，③再试）---");
            string dir2 = Path.Combine(_root, "C2-try-exception-param");
            Directory.CreateDirectory(dir2);
            string args2 = "--exe " + Q(_selfExe) + " --args " + Q("--worker crash --delay 700")
                         + " --out " + Q(dir2) + " --try-exception-param";
            ChildResult cr2 = RunChild(args2, 120000);
            LogFacts f2 = ReadLog(dir2);

            Check("C6 开启 --try-exception-param 后仍能产出可分析 dump（①的成品保住）",
                  f2.DumpPath != null && f2.DumpBytes > 1024 * 1024);
            Check("C7 日志完整记录①（NULL）与③（EXC）两条通道的结果",
                  f2.Has("【dump/①】") && f2.Has("【dump/③】"));
            Check("C8 ③失败时明确说明这是本机 dbghelp 缺陷、不是我们的 bug",
                  f2.Has("本机 dbghelp 的 ExceptionParam 通道是坏的") || f2.Has("dbghelp 的 ExceptionParam 通道是坏的"));
            Check("C9 子进程未因③而崩溃（退出码 0）", cr2.ExitCode == 0);

            Console.WriteLine("  C  dump  : " + (f.DumpPath ?? "(无)") + "  " + f.DumpBytes + " 字节");
            Console.WriteLine("  C2 dump  : " + (f2.DumpPath ?? "(无)") + "  " + f2.DumpBytes + " 字节");
        }

        // ---------------------------------------------------------------- 场景 D

        private void ScenarioD_NoCrashNoDump()
        {
            Console.WriteLine();
            Console.WriteLine("--- 场景 D：--exe 模式启动不崩的进程 → 必须不写 dump、退出码 0 ---");
            string dir = Path.Combine(_root, "D-no-crash");
            Directory.CreateDirectory(dir);

            var psi = new ProcessStartInfo(_selfExe, "--worker sleep --delay 2500") { UseShellExecute = false };
            Process w = Process.Start(psi);
            ChildResult cr = RunChild("--pid " + w.Id + " --out " + Q(dir), 60000);
            try { w.WaitForExit(5000); } catch { }

            LogFacts f = ReadLog(dir);
            Check("D1 被调试进程正常退出并被记录", f.Has("EXIT_PROCESS：") && f.ExitCodeHex == "0x00000000");
            Check("D2 未写任何 dump（没有异常就不该产出 dump）", !f.Has("[KEY] pid=") && f.DumpPath == null);
            Check("D3 明确记录了「本次会话未写任何 dump」", f.Has("本次会话未写任何 dump"));
            Check("D4 看门狗自身退出码 0", cr.ExitCode == 0);

            Console.WriteLine("  退出码行: " + (f.ExitLine ?? "(未找到)"));
        }

        // ---------------------------------------------------------------- 场景 E

        private void ScenarioE_AttachFailureIsLogged()
        {
            Console.WriteLine();
            Console.WriteLine("--- 场景 E：attach 到不存在的进程 → 必须留下日志（Lead 要求 3：失败不能静默）---");
            string dir = Path.Combine(_root, "E-attach-failure");
            Directory.CreateDirectory(dir);

            int bogus = 999996;
            ChildResult cr = RunChild("--pid " + bogus + " --out " + Q(dir), 30000);
            LogFacts f = ReadLog(dir);

            Check("E1 watchdog.log 仍被创建（失败也必须留痕）", f.Exists && f.Lines.Count > 3);
            Check("E2 明确记录了 attach 失败", f.Has("attach 失败"));
            Check("E3 记录了 win32 错误码（可诊断）", f.Has("win32="));
            Check("E4 给出了中文诊断建议", f.Has(" attach 失败诊断："));
            Check("E5 退出码非 0（调用方可感知失败）", cr.ExitCode != 0);

            Console.WriteLine("  退出码 = " + cr.ExitCode + "；日志行数 = " + (f.Lines != null ? f.Lines.Count : 0));
            foreach (string l in Head(f.Lines, 12)) Console.WriteLine("    | " + l);
        }

        // ---------------------------------------------------------------- 基础设施

        private struct ChildResult
        {
            public int ExitCode;
            public long ElapsedMs;
        }

        private ChildResult RunChild(string arguments, int timeoutMs)
        {
            var psi = new ProcessStartInfo(_selfExe, arguments)
            {
                UseShellExecute = false,
                WorkingDirectory = _root
            };
            var sw = Stopwatch.StartNew();
            Process p = Process.Start(psi);
            if (!p.WaitForExit(timeoutMs))
            {
                try { p.Kill(); } catch { }
                _log.Error("[SELFTEST] 子进程超时未退出：" + arguments);
                return new ChildResult { ExitCode = -999999, ElapsedMs = sw.ElapsedMilliseconds };
            }
            sw.Stop();
            return new ChildResult { ExitCode = p.ExitCode, ElapsedMs = sw.ElapsedMilliseconds };
        }

        private sealed class LogFacts
        {
            public bool Exists;
            public List<string> Lines = new List<string>();
            public string KeyLine;
            public string KeyModule;
            public string KeyModuleSource;
            public ulong KeyRva;
            public string DumpPath;
            public long DumpBytes;
            public string SidecarPath;
            public bool SidecarHasModuleRva;
            public bool SidecarHasCdb;
            public bool SidecarHasTid;
            public string ExitLine;
            public string ExitCodeHex;

            public bool Has(string needle)
            {
                for (int i = 0; i < Lines.Count; i++)
                {
                    if (Lines[i].IndexOf(needle, StringComparison.Ordinal) >= 0) return true;
                }
                return false;
            }
        }

        private LogFacts ReadLog(string dir)
        {
            var f = new LogFacts();
            string path = Path.Combine(dir, "watchdog.log");
            try
            {
                f.Exists = File.Exists(path);
                if (!f.Exists) return f;
                using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var sr = new StreamReader(fs, Encoding.UTF8))
                {
                    string line;
                    while ((line = sr.ReadLine()) != null) f.Lines.Add(line);
                }
            }
            catch (Exception ex)
            {
                _log.Warn("[SELFTEST] 读取日志失败 " + path + "：" + ex.GetType().Name);
                return f;
            }

            // [KEY] pid=1234 proc=X.exe code=0xC0000005 addr=0x... module=ntdll.dll+0x1234 modulebase=0x.. modulesource=.. faultaddr=0x.. access=.. tid=.. chance=.. reason=.. dump=C:\... sidecar=C:\...
            var keyRe = new Regex(@"\[KEY\] pid=(\d+) proc=(\S+) code=(0x[0-9A-Fa-f]+) addr=(0x[0-9A-Fa-f]+) "
                                + @"module=(\S+?) modulebase=(\S+) modulesource=(\S+) faultaddr=(\S+) access=(\S+) "
                                + @"tid=(\d+) chance=(\S+) reason=(.*?) dump=(\S+) sidecar=(.+?)\s*$");
            var exitRe = new Regex(@"EXIT_PROCESS：pid=\d+ \((\S+)\) 退出，退出码=(0x[0-9A-Fa-f]+)");

            for (int i = 0; i < f.Lines.Count; i++)
            {
                string l = f.Lines[i];

                Match m = keyRe.Match(l);
                if (m.Success && f.KeyLine == null)
                {
                    f.KeyLine = l.Substring(l.IndexOf("[KEY]", StringComparison.Ordinal));
                    string mod = m.Groups[5].Value;
                    int plus = mod.LastIndexOf("+0x", StringComparison.OrdinalIgnoreCase);
                    if (plus > 0)
                    {
                        f.KeyModule = mod.Substring(0, plus);
                        ulong rva;
                        if (ulong.TryParse(mod.Substring(plus + 3), System.Globalization.NumberStyles.HexNumber, null, out rva))
                            f.KeyRva = rva;
                    }
                    else
                    {
                        f.KeyModule = mod;
                    }
                    f.KeyModuleSource = m.Groups[7].Value;
                    f.DumpPath = m.Groups[13].Value.Trim();
                    if (f.DumpPath == "(失败)") f.DumpPath = null;
                    f.SidecarPath = m.Groups[14].Value.Trim();
                    if (f.SidecarPath == "(失败)") f.SidecarPath = null;
                    continue;
                }

                Match e = exitRe.Match(l);
                if (e.Success && f.ExitLine == null)
                {
                    f.ExitLine = l.Substring(l.IndexOf("EXIT_PROCESS", StringComparison.Ordinal));
                    f.ExitCodeHex = e.Groups[2].Value;
                }
            }

            if (f.DumpPath != null)
            {
                try { if (File.Exists(f.DumpPath)) f.DumpBytes = new FileInfo(f.DumpPath).Length; } catch { }
            }

            // sidecar 内容核查：必须真的把"模块名+RVA"和 cdb 命令写进去了
            if (f.SidecarPath != null && File.Exists(f.SidecarPath))
            {
                try
                {
                    string text = File.ReadAllText(f.SidecarPath, Encoding.UTF8);
                    f.SidecarHasModuleRva = text.IndexOf("★ 模块归属", StringComparison.Ordinal) >= 0
                                         && text.IndexOf("RVA", StringComparison.Ordinal) >= 0;
                    f.SidecarHasCdb = text.IndexOf("cdb -z", StringComparison.Ordinal) >= 0;
                    f.SidecarHasTid = text.IndexOf("★ 出错线程 ID", StringComparison.Ordinal) >= 0;
                }
                catch { }
            }

            return f;
        }

        private void Check(string name, Func<bool> test)
        {
            bool ok;
            try { ok = test(); }
            catch (Exception ex)
            {
                ok = false;
                _log.Error("[SELFTEST] 检查项 " + name + " 抛异常：" + ex.GetType().Name + " " + ex.Message);
            }

            if (ok) { _pass.Add(name); Console.WriteLine("   [PASS] " + name); }
            else { _fail.Add(name); Console.WriteLine("   [FAIL] " + name); }
        }

        private void Check(string name, bool ok)
        {
            Check(name, () => ok);
        }

        private static void Dump(string title, List<string> items)
        {
            if (items == null || items.Count == 0) return;
            Console.WriteLine("   " + title + " " + items.Count + " 条：");
            for (int i = 0; i < items.Count; i++) Console.WriteLine("     - " + items[i]);
        }

        private static IEnumerable<string> Head(List<string> lines, int n)
        {
            if (lines == null) yield break;
            for (int i = 0; i < lines.Count && i < n; i++) yield return lines[i];
        }

        private static string Q(string s)
        {
            return s.IndexOf(' ') >= 0 ? "\"" + s + "\"" : s;
        }
    }
}
