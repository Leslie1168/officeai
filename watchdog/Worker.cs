using System;
using System.Runtime.InteropServices;
using System.Threading;

namespace OfficeAi.Watchdog
{
    /// <summary>
    /// 自检用的"故意崩溃"工作进程。
    ///
    /// 为什么要有它：硬性验收要求"不接受代码看起来对"——必须有真实的崩溃进程来验证
    /// dump 落盘、模块名+RVA 正确、退出码记录正确。用 Word 来做这个测试既不可控也不负责，
    /// 所以让同一个 exe 支持一个 --worker 模式来当靶子。
    ///
    /// 三种崩溃方式刻意选择了不同的故障类型，用来覆盖看门狗的不同分支：
    ///   crash        —— ntdll!RtlMoveMemory 写到非法地址：**故障指令在 ntdll.dll 里**，
    ///                   异常地址必须被正确归属到 ntdll.dll + 某个 RVA。
    ///                   这是最贴近真实案件的一种：崩溃点在别人的 DLL 深处
    ///                   （wwlib.dll+0x183BA19 就是这个形状）。
    ///   crash-exec   —— 跳到本模块映像里的非可执行地址：验证"执行"型访问违例
    ///                   （ExceptionInformation[0]=8）与 DEP 路径。
    ///   crash-fastfail —— Environment.FailFast：验证 fail-fast 期间看门狗能收到什么。
    /// </summary>
    internal static class Worker
    {
        /// <summary>worker 用到的 ntdll 导入：故意不做任何参数校验，就是要它崩。</summary>
        [DllImport("ntdll.dll", EntryPoint = "RtlMoveMemory")]
        private static extern void RtlMoveMemoryUnchecked(IntPtr destination, IntPtr source, UIntPtr length);

        public static int Run(Options o, WatchdogLog log)
        {
            string mode = o.WorkerMode;
            log.Info("[WORKER] 模式=" + mode + " 延迟=" + o.WorkerDelayMs + "ms pid=" + NativeMethods.GetCurrentProcessId());

            if (mode == "noop")
            {
                log.Info("[WORKER] noop：什么都不做，直接正常退出（用来验证 EXIT_PROCESS 记录与「无 dump」路径）");
                return 0;
            }

            if (mode == "sleep")
            {
                log.Info("[WORKER] sleep：" + o.WorkerDelayMs + "ms 内保持存活，供 attach / detach 测试");
                Thread.Sleep(Math.Max(100, o.WorkerDelayMs));
                log.Info("[WORKER] sleep 结束，正常退出");
                return 0;
            }

            int delay = Math.Max(50, o.WorkerDelayMs);
            log.Info("[WORKER] " + delay + "ms 后制造崩溃（给 attach 留出时间窗口）");
            Thread.Sleep(delay);

            switch (mode)
            {
                case "crash":
                    return CrashInNtdll(log);

                case "crash-exec":
                    return CrashExecute(log);

                case "crash-fastfail":
                    log.Info("[WORKER] 调用 Environment.FailFast —— 预期异常码 0xC0000409");
                    Console.Out.Flush();
                    Environment.FailFast("OfficeAi.Watchdog selftest: intentional fail-fast");
                    return 1; // 不会走到

                case "crash-marshal":
                    log.Info("[WORKER] Marshal.WriteInt32((IntPtr)1, 0) —— 预期在 JIT 代码里产生 AV（模块归属可能是 unknown，正好验证降级路径）");
                    Marshal.WriteInt32(new IntPtr(1), 0);
                    return 1;

                default:
                    log.Error("[WORKER] 未知模式：" + mode);
                    return ExitCodes.BadArguments;
            }
        }

        /// <summary>
        /// 在 ntdll.dll 内部制造访问违例。
        /// 目标地址选 0x1（几乎必然不可写），长度 64 字节 → 写操作必然落到没映射的页上。
        /// 故障指令位于 ntdll!memmove 内部，所以 EXCEPTION_RECORD.ExceptionAddress 应当
        /// 归属到 ntdll.dll，ExceptionInformation[0] 应当是写入(1) 或读取(0)。
        /// </summary>
        private static int CrashInNtdll(WatchdogLog log)
        {
            log.Info("[WORKER] 即将调用 ntdll!RtlMoveMemory(dest=0x1, src=自身, len=64)：故障指令预期位于 ntdll.dll 内");
            Console.Out.Flush();

            byte[] src = new byte[64];
            IntPtr pSrc = Marshal.AllocHGlobal(src.Length);
            Marshal.Copy(src, 0, pSrc, src.Length);
            try
            {
                RtlMoveMemoryUnchecked(new IntPtr(1), pSrc, (UIntPtr)64);
            }
            finally
            {
                Marshal.FreeHGlobal(pSrc);
            }
            log.Info("[WORKER] 竟然没有崩？RtlMoveMemory 正常返回了");
            return 0;
        }

        /// <summary>
        /// 跳到"本模块映像内、但不可执行"的地址（映像头 +0x1000，通常落在 .rdata/只读页）。
        /// 目的是让异常地址落在**我们自己的 PE 映像范围内**，从而能精确验证
        /// 「模块名 + 基址 + RVA」这条计算链：期望 RVA 正是 0x1000。
        /// </summary>
        private static int CrashExecute(WatchdogLog log)
        {
            IntPtr baseAddr = NativeMethods.GetModuleHandleW(null);
            IntPtr target = new IntPtr(baseAddr.ToInt64() + 0x1000);
            log.Info("[WORKER] 模块基址=0x" + baseAddr.ToInt64().ToString("X")
                     + "，即将在 0x" + target.ToInt64().ToString("X") + "（基址+0x1000，映像头，不可执行）上执行代码");
            log.Info("[WORKER] 期望：异常码 0xC0000005，ExceptionInformation[0]=8(execute)，模块=RVA 0x1000");
            Console.Out.Flush();

            var del = (Action)Marshal.GetDelegateForFunctionPointer(target, typeof(Action));
            del();
            log.Info("[WORKER] 竟然没有崩？执行 0x" + target.ToInt64().ToString("X") + " 正常返回了");
            return 0;
        }
    }
}
