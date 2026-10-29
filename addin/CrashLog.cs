using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Text;
using System.Threading;

namespace OfficeAi
{
    /// <summary>
    /// 崩溃日志与时间线的落盘层。两条通道，职责分开：
    ///
    ///   1) addin.log（%LOCALAPPDATA%\OfficeAi\addin.log）—— 全量运行日志。
    ///      走 **无缓冲 FileStream**（不再用 File.AppendAllText 的 StreamWriter）：
    ///      AppendAllText 每次都要 open/close，且数据先落在托管缓冲里；
    ///      崩溃时进程被杀，托管缓冲里的内容直接丢。无缓冲 + FILE_APPEND_DATA
    ///      意味着每行都是一次 WriteFile，数据立刻进内核文件缓存——进程崩溃（Word 闪退）
    ///      后一定读得到。这里不做 FlushFileBuffers：那是为断电准备的，代价是每行一次磁盘同步，
    ///      会把 Word 的按钮动作/面包屑日志拖慢（交接报告明确说过日志洪水会拖慢 Word）。
    ///
    ///   2) crash.log（%LOCALAPPDATA%\OfficeAi\crash.log）—— 只写「崩溃现场」与「崩溃时间线」。
    ///      这一条走 FileStream + **Flush(true)**（FlushFileBuffers，真正落盘）：
    ///      现场信息量小、写入次数极少（每次崩溃十几行），必须保证即使紧接着是断电/蓝屏也在盘上。
    ///      用 FileMode.Append 保持只追加：多进程/多次崩溃都往同一个文件里累，方便对比时间线。
    ///
    /// 本文件所有写入都是 best-effort：任何失败都吞掉，绝不因为「日志写不进去」影响加载项。
    /// </summary>
    internal static class CrashLog
    {
        private static readonly object Sync = new object();

        internal static string BaseDir
        {
            get
            {
                return Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "OfficeAi");
            }
        }

        internal static string CrashesDir { get { return Path.Combine(BaseDir, "crashes"); } }
        internal static string CrashLogPath { get { return Path.Combine(BaseDir, "crash.log"); } }
        internal static string AddinLogPath { get { return Path.Combine(BaseDir, "addin.log"); } }

        private const long AddinLogMaxBytes = 2 * 1024 * 1024;

        internal static void EnsureDirs()
        {
            try { Directory.CreateDirectory(CrashesDir); } catch { }
        }

        // ===================== addin.log =====================

        /// <summary>
        /// 与旧 Connect.Log 完全一样的格式：HH:mm:ss.fff [t线程号]  消息。
        /// 差别只在落盘方式（见类注释）。全项目共用这一条通道。
        /// </summary>
        internal static void Log(string msg)
        {
            try
            {
                Directory.CreateDirectory(BaseDir);
                string line = DateTime.Now.ToString("HH:mm:ss.fff", CultureInfo.InvariantCulture) +
                              " [t" + Thread.CurrentThread.ManagedThreadId + "]  " + msg +
                              Environment.NewLine;
                AppendRaw(AddinLogPath, line, false);
            }
            catch { }
        }

        // ===================== crash.log =====================

        /// <summary>
        /// 崩溃时间线：一条事件同时进 addin.log（人能顺着读）与 crash.log（Flush(true) 保险）。
        /// 只在低频、关键的生命周期事件上调用（任务窗格/WebView2/看门狗/监听器），
        /// 不要用在鼠标级高频日志上。
        /// </summary>
        internal static void Timeline(string msg)
        {
            Log("[T] " + msg);
            AppendDurable("[timeline] " + TimestampMillis() + "  " + msg);
        }

        /// <summary>
        /// 无缓冲 + FlushFileBuffers 的追加写入。崩溃现场专用：
        /// 只要磁盘还在，写进去就一定读得到。
        /// </summary>
        internal static void AppendDurable(string text)
        {
            try
            {
                Directory.CreateDirectory(BaseDir);
                AppendRaw(CrashLogPath, text + Environment.NewLine, true);
            }
            catch { }
        }

        // ===================== 进程侧存活心跳（lead 验收用） =====================

        private static readonly Stopwatch Uptime = Stopwatch.StartNew();

        /// <summary>自本类首次被触达（≈ 加载项 Install）以来的毫秒数。用于把 JS 侧 Date.now()
        /// 与进程侧时间戳对齐。</summary>
        internal static long UptimeMs { get { return Uptime.ElapsedMilliseconds; } }

        /// <summary>
        /// 进程侧存活心跳：一行「本地毫秒时间戳 + t+自会话开始毫秒 + tag」，
        /// 直接追加进 crash.log 并 Flush(true)。
        /// 为什么单独一条通道：前端 JS 探针只能证明「JS 还在跑」，进程侧心跳证明
        /// 「native 侧还活着」；两边时间戳对齐后，就能判定崩溃发生在渲染中还是渲染后。
        /// O(1)：一次 WriteFile + 一次 FlushFileBuffers，无缓冲、不分配大对象。
        /// </summary>
        internal static void Heartbeat(string tag)
        {
            try
            {
                AppendDurable("[hb] " + TimestampMillis() + "  t+=" + Uptime.ElapsedMilliseconds +
                              "ms  tag=" + (tag ?? ""));
            }
            catch { }
        }

        private static void AppendRaw(string path, string text, bool hardFlush)
        {
            // bufferSize:1 = 关掉 FileStream 自己的缓冲（直接 WriteFile）。
            // FileMode.Append 在 Win32 层是 FILE_APPEND_DATA，每次写自动定位到文件尾，
            // 多线程并发追加不会互相覆盖（每条日志是一次 WriteFile）。
            byte[] bytes = Encoding.UTF8.GetBytes(text);
            using (var fs = new FileStream(path, FileMode.Append, FileAccess.Write,
                                           FileShare.ReadWrite, 1, FileOptions.None))
            {
                fs.Write(bytes, 0, bytes.Length);
                if (hardFlush) fs.Flush(true);   // FlushFileBuffers：真正写进磁盘
            }
            if (!hardFlush) TrimAddinLog(path);
        }

        private static void TrimAddinLog(string path)
        {
            // 旧的滚动策略保留：超过 2MB 就删掉重来（下次写会重新建）。
            // 只对 addin.log 做；crash.log 永不滚动（翻旧账要用）。
            try
            {
                var fi = new FileInfo(path);
                if (fi.Exists && fi.Length > AddinLogMaxBytes) File.Delete(path);
            }
            catch { }
        }

        // ===================== 给崩溃现场用的「读旧日志」 =====================

        /// <summary>
        /// 取 addin.log 末尾若干字节，嵌进 crash.log。
        /// 作用：把「崩溃前最后几条动作（含前端面包屑）」固化进 crash.log。
        /// 用 FileShare.ReadWrite 打开——Word/另一个线程可能正持有它。
        /// </summary>
        internal static string ReadAddinLogTail(int maxBytes)
        {
            try
            {
                if (!File.Exists(AddinLogPath)) return "(addin.log 不存在)";
                using (var fs = new FileStream(AddinLogPath, FileMode.Open, FileAccess.Read,
                                               FileShare.ReadWrite | FileShare.Delete))
                {
                    long len = fs.Length;
                    long start = len > maxBytes ? len - maxBytes : 0;
                    fs.Seek(start, SeekOrigin.Begin);
                    var buf = new byte[len - start];
                    int read = fs.Read(buf, 0, buf.Length);
                    string s = Encoding.UTF8.GetString(buf, 0, read);
                    // 从半行开始的话，丢掉第一行残片
                    if (start > 0)
                    {
                        int nl = s.IndexOf('\n');
                        if (nl >= 0 && nl + 1 < s.Length) s = s.Substring(nl + 1);
                    }
                    return s.TrimEnd('\r', '\n');
                }
            }
            catch (Exception e)
            {
                return "(读 addin.log 失败: " + e.GetType().Name + ")";
            }
        }

        // ===================== 格式化小工具 =====================
        // 说明：这些会在 dump 线程（健康线程）里用，允许分配字符串；
        // 真正的崩溃热路径（VEH）只用 NativeProbe + 预分配缓冲，不调这里。

        internal static string Hex(ulong value)
        {
            return "0x" + value.ToString("X16", CultureInfo.InvariantCulture);
        }

        internal static string Hex(uint value)
        {
            return "0x" + value.ToString("X8", CultureInfo.InvariantCulture);
        }

        internal static string TimestampMillis()
        {
            DateTime now = DateTime.Now;
            return now.ToString("yyyy-MM-dd HH:mm:ss.fff", CultureInfo.InvariantCulture);
        }

        /// <summary>把异常码翻译成人话（WER 里那条 0xc0000005）。</summary>
        internal static string ExceptionCodeName(uint code)
        {
            switch (code)
            {
                case NativeProbe.STATUS_ACCESS_VIOLATION: return "EXCEPTION_ACCESS_VIOLATION";
                case NativeProbe.STATUS_STACK_OVERFLOW: return "EXCEPTION_STACK_OVERFLOW";
                case 0x80000003: return "EXCEPTION_BREAKPOINT";
                case 0xC000001D: return "EXCEPTION_ILLEGAL_INSTRUCTION";
                case 0xC0000094: return "EXCEPTION_INT_DIVIDE_BY_ZERO";
                case 0xC0000096: return "EXCEPTION_PRIV_INSTRUCTION";
                // 注意：0xC00000FD 已由 NativeProbe.STATUS_STACK_OVERFLOW 覆盖，
                // 不能在这里再写一遍（重复 case 是编译错误）。0xC0000409 是
                // STATUS_STACK_BUFFER_OVERRUN(/GS + __fastfail)，与栈溢出不是一回事。
                case 0xC0000374: return "STATUS_HEAP_CORRUPTION";
                case 0xC0000409: return "STATUS_STACK_BUFFER_OVERRUN(fail-fast)";
                case 0xE0434352: return "CLR 托管异常(0xE0434352)";
                default: return "未知异常码";
            }
        }

        /// <summary>ACCESS_VIOLATION 的 ExceptionInformation[0]：0=读 1=写 8=执行。</summary>
        internal static string AccessKindName(ulong kind)
        {
            switch (kind)
            {
                case 0: return "读(read)";
                case 1: return "写(write)";
                case 8: return "执行(execute/dep)";
                default: return "?(" + kind + ")";
            }
        }

        internal static string StateName(uint state)
        {
            switch (state)
            {
                case NativeProbe.MEM_COMMIT: return "MEM_COMMIT(已提交)";
                case NativeProbe.MEM_RESERVE: return "MEM_RESERVE(仅保留)";
                case NativeProbe.MEM_FREE: return "MEM_FREE(空闲/已释放)";
                default: return "0x" + state.ToString("X");
            }
        }

        internal static string TypeName(uint type)
        {
            switch (type)
            {
                case NativeProbe.MEM_IMAGE: return "MEM_IMAGE(模块映像)";
                case NativeProbe.MEM_MAPPED: return "MEM_MAPPED(文件映射)";
                case NativeProbe.MEM_PRIVATE: return "MEM_PRIVATE(私有堆/栈)";
                case 0: return "(无)";
                default: return "0x" + type.ToString("X");
            }
        }

        internal static string ProtectName(uint protect)
        {
            if (protect == 0) return "(无)";
            var sb = new StringBuilder();
            switch (protect & 0xFF)
            {
                case NativeProbe.PAGE_NOACCESS: sb.Append("PAGE_NOACCESS"); break;
                case NativeProbe.PAGE_READONLY: sb.Append("PAGE_READONLY"); break;
                case NativeProbe.PAGE_READWRITE: sb.Append("PAGE_READWRITE"); break;
                case 0x08: sb.Append("PAGE_WRITECOPY"); break;
                case NativeProbe.PAGE_EXECUTE: sb.Append("PAGE_EXECUTE"); break;
                case NativeProbe.PAGE_EXECUTE_READ: sb.Append("PAGE_EXECUTE_READ"); break;
                case NativeProbe.PAGE_EXECUTE_READWRITE: sb.Append("PAGE_EXECUTE_READWRITE"); break;
                case 0x80: sb.Append("PAGE_EXECUTE_WRITECOPY"); break;
                default: sb.Append("0x" + (protect & 0xFF).ToString("X2")); break;
            }
            if ((protect & NativeProbe.PAGE_GUARD) != 0) sb.Append("|PAGE_GUARD(守卫页)");
            if ((protect & 0x400) != 0) sb.Append("|PAGE_NOCACHE");
            if ((protect & 0x100) != 0 && (protect & NativeProbe.PAGE_GUARD) == 0) sb.Append("|0x100");
            return sb.ToString();
        }

        /// <summary>把「模块+偏移」格式化成与 WinDbg/WER 一致的样子：wwlib.dll+0x183ba19。</summary>
        internal static string ModuleOffset(char[] buffer, int length, ulong rva)
        {
            if (length <= 0) return "unknown";
            return NativeProbe.ShortName(buffer, length) + "+0x" +
                   rva.ToString("X", CultureInfo.InvariantCulture).ToLowerInvariant();
        }
    }
}
