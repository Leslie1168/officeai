using System;
using System.IO;
using System.Text;
using System.Threading;

namespace OfficeAi.Watchdog
{
    /// <summary>
    /// 看门狗日志。
    ///
    /// 为什么不写 stdout：看门狗是被 Word（无窗口 COM 宿主）用 CreateNoWindow=true 拉起的，
    /// stdout/stderr 没有接收者，写出去等于丢掉。**日志文件是唯一通道**。
    ///
    /// 为什么 FileStream 而不是 File.AppendAllText：
    /// AppendAllText 每次重开文件+缓冲，崩溃瞬间的写入极易丢；而且并发写会互相截断。
    /// 这里用 bufferSize=1（.NET Framework 中 &lt;=1 表示完全关闭缓冲）+ WriteThrough，
    /// 每次 Write 直接落盘，再叠加 lock 串行化，保证"写进去的一定在磁盘上"。
    /// </summary>
    internal sealed class WatchdogLog : IDisposable
    {
        private readonly object _gate = new object();
        private readonly string _path;
        private readonly DateTime _startedAt;
        private FileStream _fs;
        private bool _closed;

        public string Path { get { return _path; } }
        public bool IsOpen { get { return _fs != null && !_closed; } }

        private WatchdogLog(string path, FileStream fs)
        {
            _path = path;
            _fs = fs;
            _startedAt = DateTime.Now;
        }

        /// <summary>打开日志。失败时返回 null（调用方必须处理 —— 绝不静默）。</summary>
        public static WatchdogLog TryOpen(string path)
        {
            try
            {
                string dir = System.IO.Path.GetDirectoryName(path);
                if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);
                // bufferSize=1 → 关闭缓冲；WriteThrough → 每次写穿过系统缓存
                var fs = new FileStream(path, FileMode.Append, FileAccess.Write, FileShare.ReadWrite, 1, FileOptions.WriteThrough);
                return new WatchdogLog(path, fs);
            }
            catch
            {
                return null;
            }
        }

        public void Info(string message) { WriteLine("INFO ", message); }
        public void Warn(string message) { WriteLine("WARN ", message); }
        public void Error(string message) { WriteLine("ERROR", message); }

        /// <summary>醒目行：崩溃捕获这种结论用它，grep 好找。</summary>
        public void Banner(string message) { WriteLine("BANG ", message); }

        public void WriteLine(string level, string message)
        {
            string now = DateTime.Now.ToString("HH:mm:ss.fff");
            double elapsed = (DateTime.Now - _startedAt).TotalSeconds;
            string line = string.Format("{0} [T{1:D2}] [+{2,8:F3}s] {3} {4}\r\n",
                now, Thread.CurrentThread.ManagedThreadId, elapsed, level, message);

            try
            {
                byte[] bytes = Encoding.UTF8.GetBytes(line);
                lock (_gate)
                {
                    if (_closed || _fs == null) return;
                    _fs.Write(bytes, 0, bytes.Length);
                    _fs.Flush();
                }
            }
            catch
            {
                // 日志写失败不能再抛：日志是取证的最后手段，但它本身不能成为崩溃源。
            }
        }

        /// <summary>
        /// 记异常。**不要用 ex.ToString()**：如果异常类型本身加载失败（TypeLoadException），
        /// ToString() 也会失败，运行时只会打印"由于 Exception.ToString() 失败，因此无法打印异常字符串"，
        /// 反而把最关键的类型名/消息吞掉。这里逐字段单独 try，任何一段拿不到都不影响其余部分。
        /// </summary>
        public void Exception(string context, Exception ex)
        {
            Error(context + " 异常类型=" + Safe(() => ex.GetType().FullName));
            Error(context + " 消息=" + Safe(() => ex.Message));
            string inner = Safe(() => ex.InnerException == null ? "(无)" : ex.InnerException.GetType().FullName + ": " + ex.InnerException.Message);
            Error(context + " 内层=" + inner);
            string st = Safe(() => ex.StackTrace);
            if (!string.IsNullOrEmpty(st))
            {
                foreach (string line in st.Split('\n'))
                {
                    string t = line.Trim();
                    if (t.Length > 0) Error(context + "   at " + t);
                }
            }
        }

        private static string Safe(Func<string> getter)
        {
            try { return getter() ?? "(null)"; }
            catch (Exception e2) { return "<取值失败: " + e2.GetType().Name + ">"; }
        }

        public void Dispose()
        {
            lock (_gate)
            {
                if (_closed) return;
                _closed = true;
                try { if (_fs != null) { _fs.Flush(); _fs.Dispose(); } } catch { }
                _fs = null;
            }
        }
    }
}
