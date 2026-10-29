using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace OfficeAi.Watchdog
{
    /// <summary>
    /// 崩溃现场的所有"由我们自己抓到"的信息。dump 只是内存快照，**这些才是结论**。
    /// </summary>
    internal sealed class CrashContext
    {
        public uint Pid;
        public string ProcName;
        public string ProcPath;
        public uint ExceptionCode;
        public bool FirstChance;
        public ulong ExceptionAddress;
        public ulong Info0;             // 访问类型（0 读 / 1 写 / 8 执行，可带 0x100 PAGE_GUARD）
        public ulong Info1;             // 出错地址
        public uint ThreadId;
        public uint NumberParameters;

        public string ModuleName;       // 异常地址所在模块
        public ulong ModuleBase;
        public ulong ModuleRva;
        public uint ModuleSize;
        public string ModulePath;
        public string ModuleSource;

        public string FaultModuleName;  // 出错地址（Info1）所在模块（如果有）
        public ulong FaultModuleBase;
        public ulong FaultModuleRva;

        public bool HaveContext;
        public uint ContextFlags;
        public string Registers;

        public List<ulong> StackCandidates = new List<ulong>();
        public List<string> StackCandidatesText = new List<string>();

        public string DumpPath;
        public long DumpBytes;
        public string DumpKind;

        public string ExceptionName;
        public string AccessText;
    }

    /// <summary>
    /// sidecar 文件：`&lt;dump同名&gt;.exception.txt`
    ///
    /// 为什么必须有它：本机 dbghelp 的 ExceptionParam 通道是坏的（见 DumpWriter 注释），
    /// 所以 dump 里没有 ExceptionStream，`.ecxr` 不可用。但"0x183ba19 是哪个 dll 的哪个偏移"
    /// 这个真正要回答的问题，**根本不需要 dbghelp**——我们在异常发生的那一微秒就从
    /// EXCEPTION_RECORD + VirtualQueryEx 拿到了。
    /// 把这些写成一个与 dump 同名配对的独立文件，分析的人（或下一个 agent）打开 dump 的同一秒
    /// 就能看到结论，并且拿到可直接复制的 cdb 命令。
    /// </summary>
    internal static class ExceptionSidecar
    {
        public static string PathFor(string dumpPath)
        {
            if (string.IsNullOrEmpty(dumpPath)) return null;
            string dir = System.IO.Path.GetDirectoryName(dumpPath);
            string name = System.IO.Path.GetFileNameWithoutExtension(dumpPath);
            return System.IO.Path.Combine(dir, name + ".exception.txt");
        }

        /// <summary>
        /// 写 sidecar。**即使 dump 失败也照写**（那时它更是唯一的产物）。
        /// 返回实际写入的路径；失败返回 null（绝不抛）。
        /// </summary>
        public static string Write(CrashContext c, string outDir, WatchdogLog log)
        {
            string path = c.DumpPath != null
                ? PathFor(c.DumpPath)
                : System.IO.Path.Combine(outDir, BuildFallbackName(c));

            try
            {
                Directory.CreateDirectory(outDir);
                var sb = new StringBuilder();

                sb.AppendLine("================================================================================");
                sb.AppendLine(" OfficeAi.Watchdog —— 崩溃现场 sidecar（与 minidump 配对）");
                sb.AppendLine(" 生成时间 : " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"));
                sb.AppendLine("================================================================================");
                sb.AppendLine();
                sb.AppendLine("【为什么会有这个文件】");
                sb.AppendLine("  本机 C:\\Windows\\System32\\dbghelp.dll 在 MiniDumpWriteDump 传非 NULL ExceptionParam 时");
                sb.AppendLine("  会失败（0x800700E6/0x800703E6 一类 HRESULT_FROM_WIN32 内存位置错误），甚至把调用进程直接搞崩。");
                sb.AppendLine("  因此 dump 采用 ExceptionParam=NULL 的写法（全内存 + 线程信息）。");
                sb.AppendLine("  那种 dump 里**仍然包含出事线程的 CONTEXT**（ThreadListStream 的 ThreadContext），");
                sb.AppendLine("  只是没有 ExceptionStream 把异常串起来 → cdb 里 .ecxr 不可用。");
                sb.AppendLine("  用下面的「出错线程 ID」+ cdb 的 ~~[tid]s 即可精确落到现场；");
                sb.AppendLine("  而「异常地址 → 模块 + RVA」这条最关键的结论，本文件已经直接给出，不依赖符号、不依赖网络。");
                sb.AppendLine();
                sb.AppendLine("--------------------------------------------------------------------------------");
                sb.AppendLine(" 一、结论（这一节就是本案要的答案）");
                sb.AppendLine("--------------------------------------------------------------------------------");
                sb.AppendLine("  进程          : " + (c.ProcName ?? "(未知)") + " (pid=" + c.Pid + ")");
                if (!string.IsNullOrEmpty(c.ProcPath)) sb.AppendLine("  映像路径      : " + c.ProcPath);
                sb.AppendLine("  异常码        : 0x" + c.ExceptionCode.ToString("X8") + "  (" + (c.ExceptionName ?? "?") + ")");
                sb.AppendLine("  时机          : " + (c.FirstChance ? "first-chance（第一次机会，进程还有自己的处理器机会）"
                                                                   : "★second-chance（第二次机会，进程即将终止）"));
                sb.AppendLine("  异常地址      : 0x" + c.ExceptionAddress.ToString("X16"));
                sb.AppendLine("  ★ 模块归属    : " + (c.ModuleName != null
                    ? c.ModuleName + " + 0x" + c.ModuleRva.ToString("X")
                      + "   (基址=0x" + c.ModuleBase.ToString("X") + ", 映像大小=0x" + c.ModuleSize.ToString("X")
                      + ", 来源=" + c.ModuleSource + ")"
                    : "unknown（地址不落在任何映像里）"));
                if (!string.IsNullOrEmpty(c.ModulePath)) sb.AppendLine("  模块路径      : " + c.ModulePath);
                if (c.FaultModuleName != null)
                {
                    sb.AppendLine("  出错地址模块  : " + c.FaultModuleName + " + 0x" + c.FaultModuleRva.ToString("X")
                                  + "   (基址=0x" + c.FaultModuleBase.ToString("X") + ")");
                }
                sb.AppendLine("  访问类型      : " + (c.AccessText ?? "?"));
                sb.AppendLine("  ExceptionInformation[0] = 0x" + c.Info0.ToString("X") + "   （0=读 1=写 8=执行，0x100=PAGE_GUARD）");
                sb.AppendLine("  ExceptionInformation[1] = 0x" + c.Info1.ToString("X16") + "   ← 出错的内存地址"
                                + (c.Info1 == 0 ? "（空指针）" : (c.Info1 < 0x10000 ? "（极低地址，野指针）" : "")));
                sb.AppendLine("  NumberParameters = " + c.NumberParameters);
                sb.AppendLine();
                sb.AppendLine("--------------------------------------------------------------------------------");
                sb.AppendLine(" 二、出事线程与寄存器");
                sb.AppendLine("--------------------------------------------------------------------------------");
                // ★ 注意 WinDbg/cdb 的 ~~[tid] 用**十六进制**线程 ID。
                //   只写十进制会让分析的人切到错误的线程（甚至报 Illegal thread error），
                //   所以两种进制都给出来，cdb 命令里直接用十六进制那个。
                string tidHex = c.ThreadId.ToString("X");
                sb.AppendLine("  ★ 出错线程 ID : " + c.ThreadId + "  (十六进制 0x" + tidHex + ")");
                sb.AppendLine("       → cdb 里切换线程用十六进制：" + "~~[" + tidHex + "]s   （写成十进制会被当成 0x" + c.ThreadId.ToString("X") + " 之外的线程）");
                sb.AppendLine("  上下文可用    : " + (c.HaveContext ? "是，ContextFlags=0x" + c.ContextFlags.ToString("X8") : "否（GetThreadContext 失败）"));
                if (c.HaveContext && !string.IsNullOrEmpty(c.Registers))
                {
                    sb.AppendLine("  寄存器        :");
                    foreach (string line in c.Registers.Split('\n'))
                    {
                        string t = line.TrimEnd('\r');
                        if (t.Length > 0) sb.AppendLine("      " + t);
                    }
                }
                sb.AppendLine();
                sb.AppendLine("--------------------------------------------------------------------------------");
                sb.AppendLine(" 三、栈上候选返回地址（启发式扫描，未做 unwind 验证）");
                sb.AppendLine("--------------------------------------------------------------------------------");
                if (c.StackCandidatesText.Count == 0)
                {
                    sb.AppendLine("  （无：RSP 起 4KB 内没有落在任何模块映像范围内的值，可能栈已损坏）");
                }
                else
                {
                    foreach (string t in c.StackCandidatesText) sb.AppendLine("  " + t);
                    sb.AppendLine();
                    sb.AppendLine("  说明：release 构建普遍省略帧指针，RBP 链走不通，所以这里用");
                    sb.AppendLine("        「RSP 起 4KB 内、落在某个模块映像范围内的 8 字节值」作为候选返回地址。");
                    sb.AppendLine("        它是启发式，可能有误报/漏报；但配合上面的模块+RVA，足以判断崩溃发生在哪条调用路径上。");
                }
                sb.AppendLine();
                sb.AppendLine("--------------------------------------------------------------------------------");
                sb.AppendLine(" 四、dump 文件");
                sb.AppendLine("--------------------------------------------------------------------------------");
                if (c.DumpPath != null)
                {
                    sb.AppendLine("  路径 : " + c.DumpPath);
                    sb.AppendLine("  大小 : " + c.DumpBytes + " 字节 (" + (c.DumpBytes / 1024.0 / 1024.0).ToString("F1") + " MB)");
                    sb.AppendLine("  写法 : " + (c.DumpKind ?? "?"));
                }
                else
                {
                    sb.AppendLine("  （本次没有产出 dump —— 见 watchdog.log；但上面的模块+RVA/寄存器依然有效）");
                }
                sb.AppendLine();
                sb.AppendLine("--------------------------------------------------------------------------------");
                sb.AppendLine(" 五、怎么用 cdb 分析（★ 不要用 !analyze -v，会卡在符号服务器，实测 180s 超时）");
                sb.AppendLine("--------------------------------------------------------------------------------");
                string dumpArg = c.DumpPath != null ? "\"" + c.DumpPath + "\"" : "\"<dump路径>\"";
                sb.AppendLine("  # 1) 切到出错线程，看寄存器和调用栈（不需要符号也能看到 模块+偏移）");
                sb.AppendLine("  cdb -z " + dumpArg + " -c \"~~[" + tidHex + "]s; r; k; q\"");
                sb.AppendLine();
                sb.AppendLine("  # 2) 想看全部线程的栈");
                sb.AppendLine("  cdb -z " + dumpArg + " -c \"~*k; q\"");
                sb.AppendLine();
                sb.AppendLine("  # 3) 配符号（用本地缓存优先；msdl 不可达时别加 -y srv*，会卡）");
                sb.AppendLine("  cdb -z " + dumpArg + " -y \"cache*" + SymbolResolver.DefaultCacheDir + "\" -c \"~~[" + tidHex + "]s; k; q\"");
                sb.AppendLine();
                sb.AppendLine("  # 4) 只关心异常地址属于哪个模块 + 偏移（本文件第一节已经给了，这里是交叉验证）");
                sb.AppendLine("  cdb -z " + dumpArg + " -c \"ln 0x" + c.ExceptionAddress.ToString("X") + "; q\"");
                sb.AppendLine();
                sb.AppendLine("  注意：本 dump 没有 ExceptionStream，所以 .ecxr 会报 \"Minidump doesn't have an exception context\"——");
                sb.AppendLine("        这是本机 dbghelp 缺陷导致的已知现象，不是 dump 坏了。用 ~~[tid]s 达到同样目的。");
                sb.AppendLine();
                sb.AppendLine("（进程退出信息会在 dump 之后由看门狗追加到本文件末尾 —— 见「后续更新」段。）");
                sb.AppendLine();
                sb.AppendLine("======== sidecar 结束 ========");

                File.WriteAllText(path, sb.ToString(), new UTF8Encoding(false));
                return path;
            }
            catch (Exception ex)
            {
                if (log != null) log.Error("写 sidecar 失败（不影响 dump）：" + ex.GetType().Name + " " + ex.Message);
                return null;
            }
        }

        private static string BuildFallbackName(CrashContext c)
        {
            return string.Format("{0}_{1}_0x{2:X8}_0x{3:X16}_{4}.exception.txt",
                string.IsNullOrEmpty(c.ProcName) ? "target" : c.ProcName, c.Pid, c.ExceptionCode,
                c.ExceptionAddress, DateTime.Now.ToString("yyyyMMdd-HHmmss.fff"));
        }

        /// <summary>
        /// 追加"进程退出"这一节（进程退出发生在 dump 之后，所以单独补写）。
        /// sidecar 是追加式的证据文件，允许二次写入。
        /// </summary>
        public static void AppendExit(string sidecarPath, string line, WatchdogLog log)
        {
            if (string.IsNullOrEmpty(sidecarPath)) return;
            try
            {
                if (!File.Exists(sidecarPath)) return;
                File.AppendAllText(sidecarPath,
                    Environment.NewLine + "---- 后续更新 " + DateTime.Now.ToString("HH:mm:ss.fff") + " ----"
                    + Environment.NewLine + line + Environment.NewLine, new UTF8Encoding(false));
            }
            catch (Exception ex)
            {
                if (log != null) log.Warn("补写 sidecar 退出信息失败（忽略）：" + ex.GetType().Name);
            }
        }
    }
}
