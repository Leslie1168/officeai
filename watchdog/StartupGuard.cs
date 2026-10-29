using System;
using System.Threading;

namespace OfficeAi.Watchdog
{
    /// <summary>
    /// 启动路径自检（Lead 要求 c）。
    ///
    /// 问题：看门狗是"被 Word 静默拉起"的，如果它卡在启动路径（attach 失败、等待锁、
    /// 语法/类型加载失败），外部完全看不出来 —— 只能看到"没有日志"。
    /// 之前的 TypeLoadException 事故就是如此：进程起来了、崩了，但日志里一条都没有。
    ///
    /// 方案：主线程每走一步就 SetPhase()，另起一个后台守护线程定期检查：
    ///   · 超过 3 秒还没进入 WaitForDebugEvent 循环 → WARN，打出当前阶段
    ///   · 超过 15 秒还没进入 → ERROR，打出当前阶段（一眼看出卡在哪一步）
    /// 这样"没有日志"这件事本身也变成了有信息量的信号。
    /// </summary>
    internal static class StartupGuard
    {
        private static volatile string _phase = "程序入口（尚未开始解析参数）";
        private static volatile bool _loopEntered;
        private static WatchdogLog _log;

        public static string Phase { get { return _phase; } }
        public static bool LoopEntered { get { return _loopEntered; } }

        public static void SetPhase(string phase)
        {
            _phase = phase;
            if (_log != null) _log.Info("[阶段] " + phase);
        }

        public static void EnteredDebugLoop()
        {
            _loopEntered = true;
        }

        public static void ExitedDebugLoop()
        {
            _phase = "已退出调试循环";
        }

        public static void Attach(WatchdogLog log)
        {
            _log = log;
            var t = new Thread(Watch, 64 * 1024);
            t.IsBackground = true;
            t.Name = "StartupGuard";
            t.Start();
        }

        private static void Watch()
        {
            try
            {
                for (int sec = 1; sec <= 60; sec++)
                {
                    Thread.Sleep(1000);
                    if (_loopEntered) return;

                    if (sec == 3)
                    {
                        _log.Warn("启动自检：3 秒内尚未进入 WaitForDebugEvent 循环，当前阶段=" + _phase);
                    }
                    else if (sec == 15)
                    {
                        _log.Error("启动自检：15 秒仍未进入调试循环 —— 启动路径卡死在「" + _phase + "」。");
                        _log.Error("启动自检：若此处之后无任何日志，说明卡点就在上面这一行所述的步骤内（常见：OpenProcess 权限不足、DebugActiveProcess 被拒、符号/网络阻塞）。");
                    }
                    else if (sec == 60)
                    {
                        _log.Error("启动自检：60 秒仍未进入调试循环，放弃等待。当前阶段=" + _phase);
                        return;
                    }
                }
            }
            catch
            {
                // 守护线程绝不能因为自身异常影响主流程
            }
        }
    }
}
