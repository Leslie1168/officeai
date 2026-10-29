using System;
using System.Runtime.InteropServices;
using System.Text;

namespace OfficeAi
{
    /// <summary>
    /// 原生互操作层：只放 P/Invoke 声明 + 进程内存/模块查询，不放任何业务逻辑。
    ///
    /// 为什么单独成文件：
    ///   1) 崩溃取证路径要求「可审计、极简」——所有进入崩溃热路径的系统调用都集中在这里，
    ///      一眼能看完，避免在业务代码里散落 DllImport。
    ///   2) 全部走 kernel32 / ntdll / dbghelp 的官方导出，**不引入任何 NuGet 依赖**
    ///      （dbghelp 直接 DllImport，见 MiniDumpWriteDump）。
    ///
    /// 这里所有查询函数都设计成「不抛异常、可失败」：崩溃现场任何一次失败都不能让
    /// 取证代码自己变成新的崩溃源。
    /// </summary>
    internal static class NativeProbe
    {
        // ===================== 异常相关常量 =====================

        /// <summary>EXCEPTION_CONTINUE_SEARCH：交给下一个处理器（我们永远返回这个）。</summary>
        internal const int EXCEPTION_CONTINUE_SEARCH = 0;

        /// <summary>EXCEPTION_EXECUTE_HANDLER：★ 会把异常「吃掉」——绝不能返回给 OS。</summary>
        internal const int EXCEPTION_EXECUTE_HANDLER = 1;

        internal const uint STATUS_ACCESS_VIOLATION = 0xC0000005;
        internal const uint STATUS_STACK_OVERFLOW = 0xC00000FD;

        // ===================== 模块查询常量 =====================

        internal const uint GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT = 0x00000002;
        internal const uint GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS = 0x00000004;

        // ===================== VirtualQuery 常量 =====================

        internal const uint MEM_COMMIT = 0x00001000;
        internal const uint MEM_RESERVE = 0x00002000;
        internal const uint MEM_FREE = 0x00010000;
        internal const uint MEM_PRIVATE = 0x00020000;
        internal const uint MEM_MAPPED = 0x00040000;
        internal const uint MEM_IMAGE = 0x01000000;

        internal const uint PAGE_NOACCESS = 0x001;
        internal const uint PAGE_GUARD = 0x100;
        internal const uint PAGE_EXECUTE = 0x010;
        internal const uint PAGE_EXECUTE_READ = 0x020;
        internal const uint PAGE_EXECUTE_READWRITE = 0x040;
        internal const uint PAGE_READONLY = 0x002;
        internal const uint PAGE_READWRITE = 0x004;

        // ===================== MiniDumpWriteDump 类型标志 =====================

        /// <summary>0x0：只有线程/模块元数据，**不含线程栈内存** —— 旧 CrashGuard 的致命 bug。</summary>
        internal const uint MiniDumpNormal = 0x00000000;
        internal const uint MiniDumpWithDataSegs = 0x00000001;
        /// <summary>0x2：抓取整个进程的已提交内存（含所有线程栈），这是能分析调用栈的最低要求。</summary>
        internal const uint MiniDumpWithFullMemory = 0x00000002;
        internal const uint MiniDumpWithHandleData = 0x00000004;
        internal const uint MiniDumpWithUnloadedModules = 0x00000020;
        internal const uint MiniDumpWithIndirectlyReferencedMemory = 0x00000040;
        internal const uint MiniDumpWithProcessThreadData = 0x00000100;
        internal const uint MiniDumpWithFullMemoryInfo = 0x00000800;
        internal const uint MiniDumpWithThreadInfo = 0x00001000;
        /// <summary>0x20000：读不到的内存跳过而不是让整次 dump 失败（崩溃现场内存可能已损坏）。</summary>
        internal const uint MiniDumpIgnoreInaccessibleMemory = 0x00020000;

        // ===================== 结构体 =====================

        /// <summary>EXCEPTION_RECORD（x64 布局，与 winnt.h 一致）。</summary>
        [StructLayout(LayoutKind.Sequential)]
        internal struct ExceptionRecord
        {
            public uint ExceptionCode;          // +0x00
            public uint ExceptionFlags;         // +0x04
            public IntPtr ExceptionRecordPtr;   // +0x08 嵌套异常
            public IntPtr ExceptionAddress;     // +0x10 ★ 出错地址（WER 的「错误偏移」基准）
            public uint NumberParameters;       // +0x18
            public uint Padding;                // +0x1C 显式补齐：x64 下 ULONG_PTR 数组按 8 字节对齐
            public ulong Info0;                 // +0x20 ExceptionInformation[0]
            public ulong Info1;                 // +0x28 ExceptionInformation[1]
        }

        /// <summary>EXCEPTION_POINTERS。</summary>
        [StructLayout(LayoutKind.Sequential)]
        internal struct ExceptionPointers
        {
            public IntPtr ExceptionRecord;
            public IntPtr ContextRecord;
        }

        /// <summary>
        /// ★ MiniDumpWriteDump 的 ExceptionParam 要的是这个结构，**不是** EXCEPTION_POINTERS*：
        ///   typedef struct _MINIDUMP_EXCEPTION_INFORMATION {
        ///       DWORD ThreadId; PEXCEPTION_POINTERS ExceptionPointers; BOOL ClientPointers;
        ///   }
        /// 直接把 EXCEPTION_POINTERS* 传进去的话，dbghelp 会把 ExceptionRecord 指针的低 32 位
        /// 当 ThreadId、把 ContextRecord 指针当 ExceptionPointers、
        /// 把 ContextRecord 的低 32 位当 ClientPointers=TRUE，于是去"跨进程"读一堆垃圾指针
        /// → MiniDumpWriteDump 立刻失败（实测 GetLastError=0x800700E6，dump 0 字节）。
        /// 这个坑同样只有真写一次 dump 才能发现。
        /// </summary>
        [StructLayout(LayoutKind.Sequential)]
        internal struct MiniDumpExceptionInformation
        {
            public uint ThreadId;
            public IntPtr ExceptionPointers;
            public int ClientPointers;    // BOOL：本进程自抓用 0（指针就在调用者地址空间里）
        }

        /// <summary>CONTEXT 的 x64 尺寸（winnt.h：sizeof(CONTEXT) == 0x4D0）。</summary>
        internal const int ContextSize = 0x4D0;

        /// <summary>
        /// x64 CONTEXT。只声明取证需要的字段，但 **Size 显式写成完整 0x4D0**，
        /// 这样 Marshal.PtrToStructure 会按真实大小复制，不会读到半截。
        /// 偏移全部对着 winnt.h 的 CONTEXT 手抄并逐个核对（DumpRegisters 会打印实测值）。
        /// </summary>
        [StructLayout(LayoutKind.Explicit, Size = ContextSize)]
        internal struct Context64
        {
            [FieldOffset(0x30)] public uint ContextFlags;
            [FieldOffset(0x34)] public uint MxCsr;
            [FieldOffset(0x38)] public ushort SegCs;
            [FieldOffset(0x3A)] public ushort SegDs;
            [FieldOffset(0x3C)] public ushort SegEs;
            [FieldOffset(0x3E)] public ushort SegFs;
            [FieldOffset(0x40)] public ushort SegGs;
            [FieldOffset(0x42)] public ushort SegSs;
            [FieldOffset(0x44)] public uint EFlags;
            [FieldOffset(0x48)] public ulong Dr0;
            [FieldOffset(0x50)] public ulong Dr1;
            [FieldOffset(0x58)] public ulong Dr2;
            [FieldOffset(0x60)] public ulong Dr3;
            [FieldOffset(0x68)] public ulong Dr6;
            [FieldOffset(0x70)] public ulong Dr7;
            [FieldOffset(0x78)] public ulong Rax;
            [FieldOffset(0x80)] public ulong Rcx;
            [FieldOffset(0x88)] public ulong Rdx;
            [FieldOffset(0x90)] public ulong Rbx;
            [FieldOffset(0x98)] public ulong Rsp;
            [FieldOffset(0xA0)] public ulong Rbp;
            [FieldOffset(0xA8)] public ulong Rsi;
            [FieldOffset(0xB0)] public ulong Rdi;
            [FieldOffset(0xB8)] public ulong R8;
            [FieldOffset(0xC0)] public ulong R9;
            [FieldOffset(0xC8)] public ulong R10;
            [FieldOffset(0xD0)] public ulong R11;
            [FieldOffset(0xD8)] public ulong R12;
            [FieldOffset(0xE0)] public ulong R13;
            [FieldOffset(0xE8)] public ulong R14;
            [FieldOffset(0xF0)] public ulong R15;
            [FieldOffset(0xF8)] public ulong Rip;   // ★ 出错指令
        }

        /// <summary>MEMORY_BASIC_INFORMATION（x64 有两个对齐填充字段）。</summary>
        [StructLayout(LayoutKind.Sequential)]
        internal struct MemoryBasicInformation
        {
            public IntPtr BaseAddress;       // +0x00
            public IntPtr AllocationBase;    // +0x08
            public uint AllocationProtect;   // +0x10
            public uint Alignment1;          // +0x14
            public IntPtr RegionSize;        // +0x18
            public uint State;               // +0x20
            public uint Protect;             // +0x24
            public uint Type;                // +0x28
            public uint Alignment2;          // +0x2C
        }

        /// <summary>PROCESS_BASIC_INFORMATION（x64），只用来拿父进程 PID。</summary>
        [StructLayout(LayoutKind.Explicit, Size = 0x30)]
        internal struct ProcessBasicInformation
        {
            [FieldOffset(0x00)] public int ExitStatus;
            [FieldOffset(0x08)] public IntPtr PebBaseAddress;
            [FieldOffset(0x10)] public IntPtr AffinityMask;
            [FieldOffset(0x18)] public int BasePriority;
            [FieldOffset(0x20)] public IntPtr UniqueProcessId;
            [FieldOffset(0x28)] public IntPtr InheritedFromUniqueProcessId;
        }

        /// <summary>STACKFRAME64：注意 dbghelp 的 Reserved[16] 改成 16 个独立字段，
        /// 避免用 ByValArray 触发托管数组分配（崩溃路径不允许）。</summary>
        [StructLayout(LayoutKind.Sequential)]
        internal struct StackFrame64
        {
            public ulong AddrPC;
            public ulong AddrReturn;
            public ulong AddrFrame;
            public ulong AddrStack;
            public ulong FuncTableEntry;
            public ulong Params0;
            public ulong Params1;
            public ulong Params2;
            public ulong Params3;
            public uint Far;
            public uint Virtual;
            public ulong Reserved0, Reserved1, Reserved2, Reserved3, Reserved4, Reserved5, Reserved6, Reserved7;
            public ulong Reserved8, Reserved9, Reserved10, Reserved11, Reserved12, Reserved13, Reserved14, Reserved15;
        }

        // ===================== 委托（必须静态持有！） =====================

        [UnmanagedFunctionPointer(CallingConvention.Winapi)]
        internal delegate int VectoredExceptionHandler(IntPtr exceptionPointers);

        [UnmanagedFunctionPointer(CallingConvention.Winapi)]
        internal delegate int UnhandledExceptionFilter(IntPtr exceptionPointers);

        [UnmanagedFunctionPointer(CallingConvention.Winapi)]
        internal delegate bool ReadProcessMemoryRoutine64(IntPtr hProcess, ulong baseAddress,
            IntPtr buffer, uint size, out uint bytesRead);

        [UnmanagedFunctionPointer(CallingConvention.Winapi)]
        internal delegate IntPtr FunctionTableAccessRoutine64(IntPtr hProcess, ulong addrBase);

        [UnmanagedFunctionPointer(CallingConvention.Winapi)]
        internal delegate ulong GetModuleBaseRoutine64(IntPtr hProcess, ulong addrBase);

        // ===================== kernel32 =====================

        [DllImport("kernel32.dll", SetLastError = true)]
        internal static extern IntPtr AddVectoredExceptionHandler(uint first, IntPtr handler);

        [DllImport("kernel32.dll", SetLastError = true)]
        internal static extern uint RemoveVectoredExceptionHandler(IntPtr handle);

        [DllImport("kernel32.dll", SetLastError = true)]
        internal static extern IntPtr SetUnhandledExceptionFilter(IntPtr filter);

        [DllImport("kernel32.dll")]
        internal static extern IntPtr GetCurrentProcess();

        [DllImport("kernel32.dll")]
        internal static extern uint GetCurrentProcessId();

        [DllImport("kernel32.dll")]
        internal static extern uint GetCurrentThreadId();

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        internal static extern IntPtr GetModuleHandleExW(uint flags, IntPtr moduleName, out IntPtr module);

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        internal static extern int GetModuleFileNameW(IntPtr module, [Out] char[] fileName, int size);

        [DllImport("kernel32.dll", SetLastError = true)]
        internal static extern IntPtr VirtualQuery(IntPtr address, out MemoryBasicInformation buffer, IntPtr length);

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        internal static extern IntPtr LoadLibraryW(string fileName);

        /// <summary>Win32 文件属性（只用大小字段，做 dump 进度判断）。</summary>
        [StructLayout(LayoutKind.Sequential)]
        internal struct Win32FileAttributeData
        {
            public uint FileAttributes;
            public long CreationTime;
            public long LastAccessTime;
            public long LastWriteTime;
            public uint FileSizeHigh;
            public uint FileSizeLow;
        }

        /// <summary>
        /// 取文件属性。★ lpFileName 用 IntPtr（原生已有的宽字符串指针），
        /// 不用 string —— string 参数会走 marshaller 分配原生缓冲，
        /// 而调用方是「出事的线程」，那次分配可能撞上它自己持有的堆锁 → 自死锁。
        /// </summary>
        [DllImport("kernel32.dll", SetLastError = true)]
        internal static extern bool GetFileAttributesExW(IntPtr lpFileName, int infoLevelId,
            out Win32FileAttributeData fileInformation);

        internal const int GetFileExInfoStandard = 0;

        [DllImport("kernel32.dll", SetLastError = true)]
        internal static extern bool ReadProcessMemory(IntPtr hProcess, IntPtr baseAddress,
            [Out] byte[] buffer, IntPtr size, out IntPtr bytesRead);

        /// <summary>x64 的 SEH 表查询（dbghelp 的 StackWalk64 需要它来展开栈）。</summary>
        [DllImport("kernel32.dll", SetLastError = true)]
        internal static extern IntPtr RtlLookupFunctionEntry(ulong controlPc, out ulong imageBase, IntPtr historyTable);

        // ===================== dbghelp =====================

        /// <summary>
        /// ★ 唯一的 dump 写入入口。dumpType 必须含 MiniDumpWithFullMemory，否则只写元数据、
        /// 没有线程栈内存，拿到手也分析不出调用栈（旧代码的 bug）。
        /// </summary>
        [DllImport("dbghelp.dll", SetLastError = true)]
        internal static extern bool MiniDumpWriteDump(IntPtr hProcess, uint processId, IntPtr hFile,
            uint dumpType, IntPtr exceptionParam, IntPtr userStreamParam, IntPtr callbackParam);

        /// <summary>栈回溯。不调用 SymInitialize —— 我们只要「模块+偏移」，不下载任何符号。</summary>
        [DllImport("dbghelp.dll", SetLastError = true)]
        internal static extern bool StackWalk64(uint machineType, IntPtr hProcess, IntPtr hThread,
            ref StackFrame64 stackFrame, ref Context64 contextRecord,
            ReadProcessMemoryRoutine64 readMemoryRoutine,
            FunctionTableAccessRoutine64 functionTableAccessRoutine,
            GetModuleBaseRoutine64 getModuleBaseRoutine,
            IntPtr translateAddress);

        // ===================== ntdll =====================

        [DllImport("ntdll.dll")]
        internal static extern int NtQueryInformationProcess(IntPtr processHandle, int infoClass,
            ref ProcessBasicInformation info, int infoLength, out int returnLength);

        internal const uint IMAGE_FILE_MACHINE_AMD64 = 0x8664;

        // ===================== 查询辅助（全部返回 bool，不抛异常） =====================

        /// <summary>
        /// 把一个地址反查成「哪个模块 + 模块内偏移(RVA)」。
        /// 这是把 wwlib.dll 0xc0000005 @0x183ba19 翻译成人话的核心：
        ///   module = GetModuleHandleEx(FROM_ADDRESS) → GetModuleFileName
        ///   rva    = address - moduleBase
        /// 地址不在任何模块里（JIT/堆/匿名内存）返回 false。
        /// </summary>
        internal static bool TryResolveModule(ulong address, char[] nameBuffer,
            out int nameLength, out ulong moduleBase, out ulong rva)
        {
            nameLength = 0;
            moduleBase = 0;
            rva = 0;
            if (address == 0 || nameBuffer == null || nameBuffer.Length == 0) return false;
            try
            {
                IntPtr hModule;
                IntPtr ok = GetModuleHandleExW(
                    GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS | GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
                    new IntPtr(unchecked((long)address)), out hModule);
                if (ok == IntPtr.Zero || hModule == IntPtr.Zero) return false;
                moduleBase = unchecked((ulong)hModule.ToInt64());
                int n = GetModuleFileNameW(hModule, nameBuffer, nameBuffer.Length);
                if (n <= 0) return false;
                if (n >= nameBuffer.Length) n = nameBuffer.Length - 1;   // 被截断
                nameLength = n;
                rva = address - moduleBase;
                return true;
            }
            catch { return false; }
        }

        /// <summary>VirtualQuery 包装：查地址所在内存区域的 state/type/protect。</summary>
        internal static bool TryQueryRegion(ulong address, out MemoryBasicInformation info)
        {
            info = default(MemoryBasicInformation);
            try
            {
                return VirtualQuery(new IntPtr(unchecked((long)address)), out info,
                    new IntPtr(Marshal.SizeOf(typeof(MemoryBasicInformation)))) != IntPtr.Zero;
            }
            catch { return false; }
        }

        /// <summary>取父进程 PID（用于只监视 Word 自己拉起来的 msedgewebview2 子进程）。</summary>
        internal static uint TryGetParentProcessId(IntPtr processHandle)
        {
            try
            {
                var info = new ProcessBasicInformation();
                int retLen;
                int status = NtQueryInformationProcess(processHandle, 0, ref info,
                    Marshal.SizeOf(typeof(ProcessBasicInformation)), out retLen);
                if (status != 0) return 0;
                return unchecked((uint)info.InheritedFromUniqueProcessId.ToInt64());
            }
            catch { return 0; }
        }

        // ===================== 命令行读取（PEB 走读，不碰 WMI） =====================

        /// <summary>
        /// 读目标进程的命令行。**刻意不用 WMI/CIM**（会拉 WMI 服务，代价大且有重入风险），
        /// 走经典 PEB 走读：
        ///   NtQueryInformationProcess(ProcessBasicInformation) → PEB.ProcessParameters(+0x20)
        ///   → RTL_USER_PROCESS_PARAMETERS.CommandLine(+0x70, UNICODE_STRING)
        ///   → ReadProcessMemory 取字符串。
        /// 只有 2s 轮询的监视线程会用它（那里允许分配）。
        /// </summary>
        internal static bool TryGetCommandLine(IntPtr processHandle, out string commandLine)
        {
            commandLine = null;
            try
            {
                var pbi = new ProcessBasicInformation();
                int retLen;
                if (NtQueryInformationProcess(processHandle, 0, ref pbi,
                        Marshal.SizeOf(typeof(ProcessBasicInformation)), out retLen) != 0)
                    return false;
                if (pbi.PebBaseAddress == IntPtr.Zero) return false;

                // x64 PEB：ProcessParameters 在 +0x20
                IntPtr parameters = ReadPointer(processHandle, pbi.PebBaseAddress + 0x20);
                if (parameters == IntPtr.Zero) return false;

                // x64 RTL_USER_PROCESS_PARAMETERS：CommandLine(UNICODE_STRING) 在 +0x70
                byte[] us = ReadBytes(processHandle, parameters + 0x70, 16);
                if (us == null) return false;
                int byteLength = BitConverter.ToUInt16(us, 0);
                IntPtr buffer = new IntPtr(BitConverter.ToInt64(us, 8));
                if (byteLength <= 0 || byteLength > 32768 || buffer == IntPtr.Zero) return false;

                byte[] raw = ReadBytes(processHandle, buffer, byteLength);
                if (raw == null) return false;
                commandLine = Encoding.Unicode.GetString(raw).TrimEnd('\0');
                return commandLine.Length > 0;
            }
            catch { return false; }
        }

        private static IntPtr ReadPointer(IntPtr hProcess, IntPtr address)
        {
            byte[] b = ReadBytes(hProcess, address, 8);
            if (b == null) return IntPtr.Zero;
            return new IntPtr(BitConverter.ToInt64(b, 0));
        }

        private static byte[] ReadBytes(IntPtr hProcess, IntPtr address, int count)
        {
            try
            {
                var buf = new byte[count];
                IntPtr read;
                if (!ReadProcessMemory(hProcess, address, buf, new IntPtr(count), out read)) return null;
                if (read.ToInt64() != count) return null;
                return buf;
            }
            catch { return null; }
        }

        /// <summary>模块名（短名，不含路径）—— 用于日志里紧凑显示，如 wwlib.dll。</summary>
        internal static string ShortName(char[] fullPath, int length)
        {
            if (fullPath == null || length <= 0) return "?";
            int start = 0;
            for (int i = length - 1; i >= 0; i--)
            {
                if (fullPath[i] == '\\' || fullPath[i] == '/') { start = i + 1; break; }
            }
            return new string(fullPath, start, length - start);
        }

        /// <summary>模块全路径字符串。</summary>
        internal static string FullName(char[] fullPath, int length)
        {
            if (fullPath == null || length <= 0) return "";
            return new string(fullPath, 0, length);
        }

        /// <summary>DbgHelp 的 StackWalk64 回调：模块基址。</summary>
        internal static ulong ModuleBaseCallback(IntPtr hProcess, ulong addrBase)
        {
            try
            {
                IntPtr hModule;
                if (GetModuleHandleExW(
                        GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS | GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
                        new IntPtr(unchecked((long)addrBase)), out hModule) != IntPtr.Zero && hModule != IntPtr.Zero)
                    return unchecked((ulong)hModule.ToInt64());
            }
            catch { }
            return 0;
        }

        /// <summary>DbgHelp 的 StackWalk64 回调：RUNTIME_FUNCTION 表项（x64 靠 .pdata 展开栈）。</summary>
        internal static IntPtr FunctionTableCallback(IntPtr hProcess, ulong addrBase)
        {
            try
            {
                ulong imageBase;
                return RtlLookupFunctionEntry(addrBase, out imageBase, IntPtr.Zero);
            }
            catch { return IntPtr.Zero; }
        }

        /// <summary>DbgHelp 的 StackWalk64 回调：读目标进程内存（用 ReadProcessMemory，
        /// 读不到就返回 false，**不会**像 Marshal.Read* 那样在坏地址上再抛一个 AV）。</summary>
        internal static bool ReadMemoryCallback(IntPtr hProcess, ulong baseAddress,
            IntPtr buffer, uint size, out uint bytesRead)
        {
            bytesRead = 0;
            try
            {
                var tmp = new byte[size];
                IntPtr read;
                if (!ReadProcessMemory(hProcess, new IntPtr(unchecked((long)baseAddress)), tmp, new IntPtr(size), out read))
                    return false;
                Marshal.Copy(tmp, 0, buffer, (int)read);
                bytesRead = (uint)read;
                return read.ToInt64() == size;
            }
            catch { return false; }
        }
    }
}
