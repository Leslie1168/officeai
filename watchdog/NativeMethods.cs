using System;
using System.Runtime.InteropServices;

namespace OfficeAi.Watchdog
{
    /// <summary>
    /// Win32 / DbgHelp / psapi 的 P/Invoke 声明与调试事件结构体。
    /// 全部手写、零依赖：这个工具必须在崩溃当天"一定能跑"，不能有任何还原前提。
    ///
    /// ★ 结构体布局是本文件的核心风险点。全部使用 LayoutKind.Sequential + 显式 Size，
    ///   并在启动时用 AssertLayouts() 断言关键结构的 Marshal.SizeOf 与 Win32 头文件一致。
    ///   布局错了不会崩，只会静默读到垃圾数据 —— 那比崩溃更糟（会写出假证据）。
    /// </summary>
    internal static class NativeMethods
    {
        // ---------------------------------------------------------------- 常量

        public const int PROCESS_ALL_ACCESS = 0x001F0FFF;

        public const int DBG_CONTINUE = 0x00010002;
        public const int DBG_EXCEPTION_NOT_HANDLED = unchecked((int)0x80010001);
        public const int DBG_TERMINATE_PROCESS = unchecked((int)0x40010004);

        public const int EXCEPTION_DEBUG_EVENT = 1;
        public const int CREATE_THREAD_DEBUG_EVENT = 2;
        public const int CREATE_PROCESS_DEBUG_EVENT = 3;
        public const int EXIT_THREAD_DEBUG_EVENT = 4;
        public const int EXIT_PROCESS_DEBUG_EVENT = 5;
        public const int LOAD_DLL_DEBUG_EVENT = 6;
        public const int UNLOAD_DLL_DEBUG_EVENT = 7;
        public const int OUTPUT_DEBUG_STRING_EVENT = 8;
        public const int RIP_EVENT = 9;

        public const uint STATUS_BREAKPOINT = 0x80000003;
        public const uint STATUS_SINGLE_STEP = 0x80000004;

        public const uint STATUS_ACCESS_VIOLATION = 0xC0000005;
        public const uint STATUS_ILLEGAL_INSTRUCTION = 0xC000001D;
        public const uint STATUS_STACK_OVERFLOW = 0xC00000FD;
        public const uint STATUS_FAIL_FAST_EXCEPTION = 0xC0000409;
        public const uint STATUS_ASSERTION_FAILURE = 0xC0000420;

        // EXCEPTION_RECORD.ExceptionInformation[0] 的取值（这是"读/写/执行"的唯一来源）
        public const ulong EXCEPTION_READ_FAULT = 0;      // 读取地址时访问违例
        public const ulong EXCEPTION_WRITE_FAULT = 1;     // 写入地址时访问违例
        public const ulong EXCEPTION_EXECUTE_FAULT = 8;   // 执行地址时访问违例（DEP/坏函数指针）

        // CONTEXT 标志（MiniDump 会原样带上，cdb 的 .ecxr 依赖它）
        public const uint CONTEXT_AMD64 = 0x00100000;
        public const uint CONTEXT_I386 = 0x00010000;
        public const uint CONTEXT_CONTROL = CONTEXT_AMD64 | 0x1;
        public const uint CONTEXT_INTEGER = CONTEXT_AMD64 | 0x2;
        public const uint CONTEXT_FLOATING_POINT = CONTEXT_AMD64 | 0x8;
        public const uint CONTEXT_FULL = CONTEXT_CONTROL | CONTEXT_INTEGER | CONTEXT_FLOATING_POINT;
        /// <summary>CONTEXT_ALL：CONTEXT_AMD64 | 0x3F。取寄存器给 dump 用就取全量，别省。</summary>
        public const uint CONTEXT_ALL_X64 = CONTEXT_AMD64 | 0x3F;
        public const uint CONTEXT_ALL_X86 = CONTEXT_I386 | 0x3F;

        /// <summary>sizeof(CONTEXT)：x64 = 1232，x86 = 716。</summary>
        public const int CONTEXT_SIZE_X64 = 1232;
        public const int CONTEXT_SIZE_X86 = 716;

        // GetThreadContext 需要的线程权限
        public const int THREAD_GET_CONTEXT = 0x0008;
        public const int THREAD_SUSPEND_RESUME = 0x0002;
        public const int THREAD_QUERY_INFORMATION = 0x0040;

        // x64 CONTEXT 里寄存器的字节偏移（与 CONTEXT_X64 的 FieldOffset 一一对应，
        // 单独列成常量是为了在"按原始偏移读取"这条路径上也不依赖结构体布局自检）
        public const int X64_OFF_Rax = 120;
        public const int X64_OFF_Rcx = 128;
        public const int X64_OFF_Rdx = 136;
        public const int X64_OFF_Rbx = 144;
        public const int X64_OFF_Rsp = 152;
        public const int X64_OFF_Rbp = 160;
        public const int X64_OFF_Rsi = 168;
        public const int X64_OFF_Rdi = 176;
        public const int X64_OFF_R8 = 184;
        public const int X64_OFF_R9 = 192;
        public const int X64_OFF_R10 = 200;
        public const int X64_OFF_R11 = 208;
        public const int X64_OFF_R12 = 216;
        public const int X64_OFF_R13 = 224;
        public const int X64_OFF_R14 = 232;
        public const int X64_OFF_R15 = 240;
        public const int X64_OFF_Rip = 248;
        public const int X64_OFF_EFlags = 68;
        public const int X64_OFF_SegCs = 56;

        // MINIDUMP_TYPE
        public const uint MiniDumpNormal = 0x00000000;
        public const uint MiniDumpWithDataSegs = 0x00000001;
        public const uint MiniDumpWithFullMemory = 0x00000002;
        public const uint MiniDumpWithHandleData = 0x00000004;
        public const uint MiniDumpWithUnloadedModules = 0x00000020;
        public const uint MiniDumpWithProcessThreadData = 0x00000100;
        public const uint MiniDumpWithFullMemoryInfo = 0x00000800;
        public const uint MiniDumpWithThreadInfo = 0x00001000;
        public const uint MiniDumpWithTokenInformation = 0x00004000;

        // SymSetOptions
        public const uint SYMOPT_CASE_INSENSITIVE = 0x00000001;
        public const uint SYMOPT_UNDNAME = 0x00000002;
        public const uint SYMOPT_DEFERRED_LOADS = 0x00000004;
        public const uint SYMOPT_LOAD_LINES = 0x00000010;
        public const uint SYMOPT_DEBUG = 0x80000000;

        // psapi
        public const uint LIST_MODULES_ALL = 0x03;

        public static readonly IntPtr INVALID_HANDLE_VALUE = new IntPtr(-1);

        // ---------------------------------------------------------------- 结构体

        /// <summary>EXCEPTION_DEBUG_INFO：ExceptionRecord 内联在结构体里（不是指针）。</summary>
        [StructLayout(LayoutKind.Explicit, Size = 160)]
        public struct EXCEPTION_DEBUG_INFO
        {
            [FieldOffset(0)] public EXCEPTION_RECORD ExceptionRecord;
            [FieldOffset(152)] public uint dwFirstChance;
            [FieldOffset(156)] public uint __pad; // sizeof 补齐到 160（8 字节对齐）
        }

        /// <summary>
        /// EXCEPTION_RECORD（x64 布局，winnt.h 原文）：
        ///   +0x000 ExceptionCode        +0x004 ExceptionFlags
        ///   +0x008 ExceptionRecord      +0x010 ExceptionAddress
        ///   +0x018 NumberParameters     +0x01c __unusedAlignment (仅 _WIN64)
        ///   +0x020 ExceptionInformation[15]      → sizeof = 152
        ///
        /// ★ ExceptionInformation 只声明前 2 个元素（[0]=读/写/执行，[1]=出错地址），
        ///   其余 13 个不读。这么做的原因是：绝不能把 ulong[] 放进 union——
        ///   托管数组是引用类型，显式布局 union 里一旦有引用字段，CLR 会直接
        ///   抛 TypeLoadException（"contains an object field ... overlapped"）。
        ///   全部成员保持 blittable，union 才合法。
        /// </summary>
        [StructLayout(LayoutKind.Explicit, Size = 152)]
        public struct EXCEPTION_RECORD
        {
            [FieldOffset(0)] public uint ExceptionCode;
            [FieldOffset(4)] public uint ExceptionFlags;
            [FieldOffset(8)] public IntPtr ExceptionRecord;
            [FieldOffset(16)] public IntPtr ExceptionAddress;
            [FieldOffset(24)] public uint NumberParameters;
            [FieldOffset(28)] public uint __unusedAlignment;
            [FieldOffset(32)] public ulong ExceptionInformation0;
            [FieldOffset(40)] public ulong ExceptionInformation1;

            /// <summary>EXCEPTION_RECORD 在 x64 上是 152 字节。</summary>
            public const int NativeSize = 152;
        }

        /// <summary>EXCEPTION_POINTERS：MiniDumpWriteDump 的错误参数就长这样。</summary>
        [StructLayout(LayoutKind.Sequential)]
        public struct EXCEPTION_POINTERS
        {
            public IntPtr ExceptionRecord;
            public IntPtr ContextRecord;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct MINIDUMP_EXCEPTION_INFORMATION
        {
            public uint ThreadId;
            public IntPtr ExceptionPointers;
            public int ClientPointers; // 0 = ExceptionPointers 指向"本进程(调试器)"内存
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct CREATE_PROCESS_DEBUG_INFO
        {
            public IntPtr hFile;
            public IntPtr hProcess;
            public IntPtr hThread;
            public IntPtr lpBaseOfImage;
            public uint dwDebugInfoFileOffset;
            public uint nDebugInfoSize;
            public IntPtr lpThreadLocalBase;
            public IntPtr lpStartAddress;
            public IntPtr lpImageName;
            public ushort fUnicode;
            public ushort __pad;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct CREATE_THREAD_DEBUG_INFO
        {
            public IntPtr hThread;
            public IntPtr lpThreadLocalBase;
            public IntPtr lpStartAddress;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct EXIT_THREAD_DEBUG_INFO
        {
            public uint dwExitCode;
            public uint __pad;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct EXIT_PROCESS_DEBUG_INFO
        {
            public uint dwExitCode;
            public uint __pad;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct LOAD_DLL_DEBUG_INFO
        {
            public IntPtr hFile;
            public IntPtr lpBaseOfDll;
            public uint dwDebugInfoFileOffset;
            public uint nDebugInfoSize;
            public IntPtr lpImageName;
            public ushort fUnicode;
            public ushort __pad;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct UNLOAD_DLL_DEBUG_INFO
        {
            public IntPtr lpBaseOfDll;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct OUTPUT_DEBUG_STRING_INFO
        {
            public IntPtr lpDebugStringData;
            public ushort fUnicode;
            public ushort nDebugStringLength;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct RIP_INFO
        {
            public uint dwError;
            public uint dwType;
        }

        /// <summary>
        /// DEBUG_EVENT 的 union。
        /// x64 上 union 从偏移 16 开始、宽度 160 字节；x86 上从偏移 12 开始、宽度 84 字节。
        /// 这里按 8 字节对齐声明并显式 Size=176（x64 的 sizeof(DEBUG_EVENT)），
        /// 保证 Marshal.PtrToStructure 读到的各成员偏移与 Win32 完全一致。
        /// </summary>
        [StructLayout(LayoutKind.Explicit, Size = 176)]
        public struct DEBUG_EVENT
        {
            [FieldOffset(0)] public uint dwDebugEventCode;
            [FieldOffset(4)] public uint dwProcessId;
            [FieldOffset(8)] public uint dwThreadId;
            [FieldOffset(16)] public DebugEventUnion u;
        }

        [StructLayout(LayoutKind.Explicit, Size = 160)]
        public struct DebugEventUnion
        {
            [FieldOffset(0)] public EXCEPTION_DEBUG_INFO Exception;
            [FieldOffset(0)] public CREATE_THREAD_DEBUG_INFO CreateThread;
            [FieldOffset(0)] public CREATE_PROCESS_DEBUG_INFO CreateProcessInfo;
            [FieldOffset(0)] public EXIT_THREAD_DEBUG_INFO ExitThread;
            [FieldOffset(0)] public EXIT_PROCESS_DEBUG_INFO ExitProcess;
            [FieldOffset(0)] public LOAD_DLL_DEBUG_INFO LoadDll;
            [FieldOffset(0)] public UNLOAD_DLL_DEBUG_INFO UnloadDll;
            [FieldOffset(0)] public OUTPUT_DEBUG_STRING_INFO DebugString;
            [FieldOffset(0)] public RIP_INFO RipInfo;
        }

        // ---- x86 CONTEXT（只需要前 44 字节：ContextFlags + 整数寄存器，够写模块/RVA 和 cdb 看栈）----

        [StructLayout(LayoutKind.Explicit, Size = 716)]
        public struct CONTEXT_X86
        {
            [FieldOffset(0)] public uint ContextFlags;
            [FieldOffset(4)] public uint Dr0;
            [FieldOffset(8)] public uint Dr1;
            [FieldOffset(12)] public uint Dr2;
            [FieldOffset(16)] public uint Dr3;
            [FieldOffset(20)] public uint Dr6;
            [FieldOffset(24)] public uint Dr7;
            [FieldOffset(28)] public FLOATING_SAVE_AREA_X86 FloatSave;
            [FieldOffset(140)] public uint SegGs;
            [FieldOffset(144)] public uint SegFs;
            [FieldOffset(148)] public uint SegEs;
            [FieldOffset(152)] public uint SegDs;
            [FieldOffset(156)] public uint Edi;
            [FieldOffset(160)] public uint Esi;
            [FieldOffset(164)] public uint Ebx;
            [FieldOffset(168)] public uint Edx;
            [FieldOffset(172)] public uint Ecx;
            [FieldOffset(176)] public uint Eax;
            [FieldOffset(180)] public uint Ebp;
            [FieldOffset(184)] public uint Eip;
            [FieldOffset(188)] public uint SegCs;
            [FieldOffset(192)] public uint EFlags;
            [FieldOffset(196)] public uint Esp;
            [FieldOffset(200)] public uint SegSs;
        }

        [StructLayout(LayoutKind.Explicit, Size = 112)]
        public struct FLOATING_SAVE_AREA_X86
        {
            [FieldOffset(0)] public uint ControlWord;
            [FieldOffset(4)] public uint StatusWord;
            [FieldOffset(8)] public uint TagWord;
            [FieldOffset(12)] public uint ErrorOffset;
            [FieldOffset(16)] public uint ErrorSelector;
            [FieldOffset(20)] public uint DataOffset;
            [FieldOffset(24)] public uint DataSelector;
            [FieldOffset(28)] public byte RegisterArea0; // 80 字节，这里只占位
        }

        // ---- x64 CONTEXT（截断版：ContextFlags + 到 Rip 为止的全部寄存器字段）----
        //
        // 为什么要"截断"而不是声明完整的 1232 字节：
        //   完整 CONTEXT 尾部有 M128A[26] Header[2] 等嵌套联合，C# 里声明起来极易算错偏移。
        //   而 Marshal.PtrToStructure 允许结构体 Size 小于实际内存（只读我们声明的部分），
        //   所以只要"前 256 字节里每个字段的偏移"绝对正确，就能安全拿到全部通用寄存器——
        //   这已经足够写模块归属、RVA、以及让 cdb 的 .ecxr / k 正常工作（CONTEXT 本体是
        //   调试事件给的原生内存，我们只是读它）。
        [StructLayout(LayoutKind.Explicit, Size = 256)]
        public struct CONTEXT_X64
        {
            [FieldOffset(0)] public ulong P1Home;
            [FieldOffset(8)] public ulong P2Home;
            [FieldOffset(16)] public ulong P3Home;
            [FieldOffset(24)] public ulong P4Home;
            [FieldOffset(32)] public ulong P5Home;
            [FieldOffset(40)] public ulong P6Home;

            [FieldOffset(48)] public uint ContextFlags;
            [FieldOffset(52)] public uint MxCsr;

            [FieldOffset(56)] public ushort SegCs;
            [FieldOffset(58)] public ushort SegDs;
            [FieldOffset(60)] public ushort SegEs;
            [FieldOffset(62)] public ushort SegFs;
            [FieldOffset(64)] public ushort SegGs;
            [FieldOffset(66)] public ushort SegSs;
            [FieldOffset(68)] public uint EFlags;

            [FieldOffset(72)] public ulong Dr0;
            [FieldOffset(80)] public ulong Dr1;
            [FieldOffset(88)] public ulong Dr2;
            [FieldOffset(96)] public ulong Dr3;
            [FieldOffset(104)] public ulong Dr6;
            [FieldOffset(112)] public ulong Dr7;

            [FieldOffset(120)] public ulong Rax;
            [FieldOffset(128)] public ulong Rcx;
            [FieldOffset(136)] public ulong Rdx;
            [FieldOffset(144)] public ulong Rbx;
            [FieldOffset(152)] public ulong Rsp;
            [FieldOffset(160)] public ulong Rbp;
            [FieldOffset(168)] public ulong Rsi;
            [FieldOffset(176)] public ulong Rdi;
            [FieldOffset(184)] public ulong R8;
            [FieldOffset(192)] public ulong R9;
            [FieldOffset(200)] public ulong R10;
            [FieldOffset(208)] public ulong R11;
            [FieldOffset(216)] public ulong R12;
            [FieldOffset(224)] public ulong R13;
            [FieldOffset(232)] public ulong R14;
            [FieldOffset(240)] public ulong R15;
            [FieldOffset(248)] public ulong Rip;
        }

        // ---- 进程/模块信息辅助 ----

        [StructLayout(LayoutKind.Sequential)]
        public struct PROCESS_BASIC_INFORMATION
        {
            public IntPtr Reserved1;
            public IntPtr PebBaseAddress;
            public IntPtr Reserved2_0;
            public IntPtr Reserved2_1;
            public IntPtr UniqueProcessId;
            public IntPtr InheritedFromUniqueProcessId;
        }

        // ---------------------------------------------------------------- kernel32

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern IntPtr OpenProcess(int dwDesiredAccess, bool bInheritHandle, int dwProcessId);

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool DebugActiveProcess(int dwProcessId);

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool DebugActiveProcessStop(int dwProcessId);

        /// <summary>
        /// ★ 硬要求：必须传 FALSE。
        /// 默认（TRUE）语义是"调试器退出 → 内核把被调试进程一起干掉"。
        /// 看门狗自己是个会被 taskkill / 关机 / 崩溃带走的独立进程，一旦默认生效，
        /// 用户关掉的是看门狗，死的却是他正在写的 Word 文档 —— 对用户是灾难。
        /// </summary>
        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool DebugSetProcessKillOnExit([MarshalAs(UnmanagedType.Bool)] bool KillOnExit);

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool WaitForDebugEvent(out DEBUG_EVENT lpDebugEvent, uint dwMilliseconds);

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool ContinueDebugEvent(uint dwProcessId, uint dwThreadId, int dwContinueStatus);

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool CloseHandle(IntPtr hObject);

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool IsWow64Process2(IntPtr hProcess, out ushort pProcessMachine, out ushort pNativeMachine);

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool IsWow64Process(IntPtr hProcess, [MarshalAs(UnmanagedType.Bool)] out bool Wow64Process);

        /// <summary>句柄泄漏自检用：看门狗自己的句柄数在 attach/detach 前后应当基本不变。</summary>
        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool GetProcessHandleCount(IntPtr hProcess, out uint pdwHandleCount);

        [DllImport("kernel32.dll")]
        public static extern IntPtr GetCurrentProcess();

        [DllImport("kernel32.dll")]
        public static extern uint GetCurrentProcessId();

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern uint GetModuleFileNameW(IntPtr hModule, [Out] char[] lpFilename, uint nSize);

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern IntPtr GetModuleHandleW(string lpModuleName);

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern IntPtr OpenThread(int dwDesiredAccess, bool bInheritHandle, uint dwThreadId);

        /// <summary>ntdll!RtlMoveMemory。只在"本进程内存 → 本进程内存"上用。</summary>
        [DllImport("ntdll.dll")]
        public static extern void RtlMoveMemory(IntPtr destination, IntPtr source, UIntPtr length);

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool ReadProcessMemory(IntPtr hProcess, IntPtr lpBaseAddress,
            [Out] byte[] lpBuffer, IntPtr nSize, out IntPtr lpNumberOfBytesRead);

        // ---- 以调试方式启动进程（--exe 备用模式）----

        public const uint DEBUG_PROCESS = 0x00000001;
        public const uint DEBUG_ONLY_THIS_PROCESS = 0x00000002;
        public const uint CREATE_NEW_CONSOLE = 0x00000010;

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        public struct STARTUPINFO
        {
            public uint cb;
            public string lpReserved;
            public string lpDesktop;
            public string lpTitle;
            public uint dwX;
            public uint dwY;
            public uint dwXSize;
            public uint dwYSize;
            public uint dwXCountChars;
            public uint dwYCountChars;
            public uint dwFillAttribute;
            public uint dwFlags;
            public ushort wShowWindow;
            public ushort cbReserved2;
            public IntPtr lpReserved2;
            public IntPtr hStdInput;
            public IntPtr hStdOutput;
            public IntPtr hStdError;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct PROCESS_INFORMATION
        {
            public IntPtr hProcess;
            public IntPtr hThread;
            public uint dwProcessId;
            public uint dwThreadId;
        }

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool CreateProcessW(
            string lpApplicationName, string lpCommandLine, IntPtr lpProcessAttributes, IntPtr lpThreadAttributes,
            [MarshalAs(UnmanagedType.Bool)] bool bInheritHandles, uint dwCreationFlags,
            IntPtr lpEnvironment, string lpCurrentDirectory, ref STARTUPINFO lpStartupInfo,
            out PROCESS_INFORMATION lpProcessInformation);

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern uint SuspendThread(IntPtr hThread);

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern uint ResumeThread(IntPtr hThread);

        /// <summary>
        /// 用原始指针取 CONTEXT。
        /// 为什么不声明成 CONTEXT 结构体参数：x64 的 CONTEXT 尾部有嵌套联合，C# 里极易算错偏移；
        /// 而且 GetThreadContext 在 x64 上要求缓冲区 16 字节对齐。用 AllocHGlobal + 原始偏移读取
        /// 是最不容易出错的做法（偏移常量集中在上面，和 CONTEXT_X64 的 FieldOffset 对齐）。
        /// </summary>
        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool GetThreadContext(IntPtr hThread, IntPtr lpContext);

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool QueryFullProcessImageNameW(IntPtr hProcess, uint dwFlags, [Out] char[] lpExeName, ref uint lpdwSize);

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern IntPtr CreateToolhelp32Snapshot(uint dwFlags, uint th32ProcessID);

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool Thread32First(IntPtr hSnapshot, ref THREADENTRY32 lpte);

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool Thread32Next(IntPtr hSnapshot, ref THREADENTRY32 lpte);

        [StructLayout(LayoutKind.Sequential)]
        public struct THREADENTRY32
        {
            public uint dwSize;
            public uint cntUsage;
            public uint th32ThreadID;
            public uint th32OwnerProcessID;
            public int tpBasePri;
            public int tpDeltaPri;
            public uint dwFlags;
        }

        public const uint TH32CS_SNAPTHREAD = 0x00000002;

        // ---------------------------------------------------------------- psapi

        [DllImport("psapi.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool EnumProcessModulesEx(IntPtr hProcess, [Out] IntPtr[] lphModule, uint cb, out uint lpcbNeeded, uint dwFilterFlag);

        /// <summary>不带过滤标志的版本，作为 EnumProcessModulesEx 失败时的兜底。</summary>
        [DllImport("psapi.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool EnumProcessModules(IntPtr hProcess, [Out] IntPtr[] lphModule, uint cb, out uint lpcbNeeded);

        [DllImport("psapi.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern uint GetModuleFileNameExW(IntPtr hProcess, IntPtr hModule, [Out] System.Text.StringBuilder lpFilename, uint nSize);

        [DllImport("psapi.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool GetModuleInformation(IntPtr hProcess, IntPtr hModule, out MODULEINFO lpmodinfo, uint cb);

        /// <summary>
        /// 跨进程虚拟内存查询。
        /// 这是"模块归属"的第二条独立通道：即使 EnumProcessModules 挂了（场景 A 实测 win32=299
        /// ERROR_PARTIAL_COPY —— 被调试进程刚创建、loader 还没跑完时模块表读不出来），
        /// 也能从出错地址本身反查它落在哪个映像里（MEM_IMAGE 的 AllocationBase 就是模块基址）。
        /// </summary>
        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern UIntPtr VirtualQueryEx(IntPtr hProcess, IntPtr lpAddress,
            out MEMORY_BASIC_INFORMATION lpBuffer, UIntPtr dwLength);

        [DllImport("psapi.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern uint GetMappedFileNameW(IntPtr hProcess, IntPtr lpv, [Out] char[] lpFilename, uint nSize);

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern uint QueryDosDeviceW(string lpDeviceName, [Out] char[] lpTargetPath, uint ucchMax);

        [StructLayout(LayoutKind.Sequential)]
        public struct MEMORY_BASIC_INFORMATION
        {
            public IntPtr BaseAddress;       // 0
            public IntPtr AllocationBase;    // 8  ← MEM_IMAGE 时就是模块基址
            public uint AllocationProtect;   // 16
            public ushort PartitionId;       // 20
            public ushort __pad;             // 22
            public UIntPtr RegionSize;       // 24
            public uint State;               // 32
            public uint Protect;             // 36
            public uint Type;                // 40
            public uint __pad2;              // 44
        }

        public const uint MEM_IMAGE = 0x1000000;
        public const uint MEM_COMMIT = 0x1000;
        public const uint MEM_PRIVATE = 0x20000;
        public const uint MEM_MAPPED = 0x40000;
        public const uint PAGE_GUARD = 0x100;
        public const uint PAGE_NOACCESS = 0x01;
        public const uint PAGE_EXECUTE = 0x10;
        public const uint PAGE_EXECUTE_READ = 0x20;
        public const uint PAGE_EXECUTE_READWRITE = 0x40;
        public const uint PAGE_EXECUTE_WRITECOPY = 0x80;

        [StructLayout(LayoutKind.Sequential)]
        public struct MODULEINFO
        {
            public IntPtr lpBaseOfDll;
            public uint SizeOfImage;
            public IntPtr EntryPoint;
        }

        // ---------------------------------------------------------------- dbghelp

        [DllImport("dbghelp.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool MiniDumpWriteDump(
            IntPtr hProcess, uint ProcessId, IntPtr hFile, uint DumpType,
            ref MINIDUMP_EXCEPTION_INFORMATION ExceptionParam, IntPtr UserStreamParam, IntPtr CallbackParam);

        [DllImport("dbghelp.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool MiniDumpWriteDump(
            IntPtr hProcess, uint ProcessId, IntPtr hFile, uint DumpType,
            IntPtr ExceptionParam, IntPtr UserStreamParam, IntPtr CallbackParam);

        [DllImport("dbghelp.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SymInitialize(IntPtr hProcess, string UserSearchPath, [MarshalAs(UnmanagedType.Bool)] bool fInvadeProcess);

        [DllImport("dbghelp.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SymCleanup(IntPtr hProcess);

        [DllImport("dbghelp.dll", SetLastError = true)]
        public static extern uint SymSetOptions(uint SymOptions);

        [DllImport("dbghelp.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern ulong SymLoadModuleEx(
            IntPtr hProcess, IntPtr hFile, string ImageName, string ModuleName,
            ulong BaseOfDll, uint DllSize, IntPtr Data, uint Flags);

        [DllImport("dbghelp.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SymFromAddr(IntPtr hProcess, ulong Address, out ulong Displacement, ref SYMBOL_INFO Symbol);

        [DllImport("dbghelp.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SymGetLineFromAddr64(IntPtr hProcess, ulong dwAddr, out uint pdwDisplacement, out IMAGEHLP_LINE64 Line);

        /// <summary>
        /// SYMBOL_INFOW。
        /// ★ Name 必须是 2000 个 WCHAR（=4000 字节），让 Marshal.SizeOf 正好等于
        ///   Win32 的 sizeof(SYMBOL_INFOW) = 4088 —— DbgHelp 会校验 SizeOfStruct，
        ///   给 512 字节的小缓冲区会让 SymFromAddr 直接返回失败（且没有任何提示）。
        /// </summary>
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        public struct SYMBOL_INFO
        {
            public uint SizeOfStruct;
            public uint TypeIndex;
            public ulong Reserved0;
            public ulong Reserved1;
            public uint Index;
            public uint Size;
            public ulong ModBase;
            public uint Flags;
            public ulong Value;
            public ulong Address;
            public uint Register;
            public uint Scope;
            public uint Tag;
            public uint NameLen;
            public uint MaxNameLen;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst = 2000)]
            public char[] Name;
        }

        public const int SYMBOL_NAME_CHARS = 2000;

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        public struct IMAGEHLP_LINE64
        {
            public uint SizeOfStruct;
            public IntPtr Key;
            public uint LineNumber;
            public IntPtr FileName;
            public ulong Address;
        }

        // ---------------------------------------------------------------- 布局自检

        internal sealed class LayoutCheckResult
        {
            /// <summary>硬性结构错位：dump 会写出错误内容 → 禁用 dump，但进程继续跑并大声报错。</summary>
            public readonly System.Collections.Generic.List<string> Hard = new System.Collections.Generic.List<string>();
            /// <summary>软性结构错位：只影响日志细节（模块名/线程数），不影响 dump。</summary>
            public readonly System.Collections.Generic.List<string> Soft = new System.Collections.Generic.List<string>();
            public bool Ok { get { return Hard.Count == 0 && Soft.Count == 0; } }
            public bool CanDump { get { return Hard.Count == 0; } }
        }

        /// <summary>
        /// 结构体布局自检。期望值来自 Windows SDK 头文件（x64）：
        ///   sizeof(EXCEPTION_RECORD)             = 152
        ///   sizeof(EXCEPTION_DEBUG_INFO)         = 160
        ///   sizeof(DEBUG_EVENT)                  = 176
        ///   sizeof(CONTEXT)                      = 1232
        ///   sizeof(SYMBOL_INFOW)                 = 4088
        ///
        /// ★ 关键设计（Lead 明确要求）：
        ///   这段自检**绝不能成为启动路径上的硬依赖**。之前这里直接抛 TypeLoadException
        ///   把看门狗打死了 —— 一个跑不起来的看门狗等于没有，而"布局不对"最多让证据不完整。
        ///   所以：
        ///     · 每个 Marshal.SizeOf 单独 try/catch（TypeLoadException 就地捕获，不上抛）；
        ///     · 硬错位 → 只置 CanDump=false，进程继续运行并写 ERROR 日志；
        ///     · 软错位 → 仅 WARN。
        /// </summary>
        public static LayoutCheckResult CheckLayouts()
        {
            var r = new LayoutCheckResult();

            Action<string, Type, int, bool> check = (name, type, expected, hard) =>
            {
                try
                {
                    int actual = Marshal.SizeOf(type);
                    if (actual != expected)
                    {
                        string msg = string.Format("{0}: 实际 {1} != 期望 {2}", name, actual, expected);
                        if (hard) r.Hard.Add(msg); else r.Soft.Add(msg);
                    }
                }
                catch (Exception ex)
                {
                    string msg = string.Format("{0}: Marshal.SizeOf 抛 {1}: {2}", name, ex.GetType().Name, ex.Message);
                    if (hard) r.Hard.Add(msg); else r.Soft.Add(msg);
                }
            };

            // 硬性：这三个错了，写出来的 dump/异常上下文就是假证据
            check("EXCEPTION_RECORD", typeof(EXCEPTION_RECORD), EXCEPTION_RECORD.NativeSize, true);
            check("EXCEPTION_POINTERS", typeof(EXCEPTION_POINTERS), 16, true);
            check("MINIDUMP_EXCEPTION_INFORMATION", typeof(MINIDUMP_EXCEPTION_INFORMATION), 24, true);
            check("DEBUG_EVENT", typeof(DEBUG_EVENT), 176, true);
            check("DebugEventUnion", typeof(DebugEventUnion), 160, true);
            check("EXCEPTION_DEBUG_INFO", typeof(EXCEPTION_DEBUG_INFO), 160, true);

            // 软性：只影响日志字段
            check("CREATE_PROCESS_DEBUG_INFO", typeof(CREATE_PROCESS_DEBUG_INFO), 72, false);
            check("CREATE_THREAD_DEBUG_INFO", typeof(CREATE_THREAD_DEBUG_INFO), 24, false);
            check("LOAD_DLL_DEBUG_INFO", typeof(LOAD_DLL_DEBUG_INFO), 40, false);
            check("UNLOAD_DLL_DEBUG_INFO", typeof(UNLOAD_DLL_DEBUG_INFO), 8, false);
            check("OUTPUT_DEBUG_STRING_INFO", typeof(OUTPUT_DEBUG_STRING_INFO), 16, false);
            check("THREADENTRY32", typeof(THREADENTRY32), 28, false);
            check("MODULEINFO", typeof(MODULEINFO), 24, false);
            check("CONTEXT_X64(截断)", typeof(CONTEXT_X64), 256, false);
            check("CONTEXT_X86(截断)", typeof(CONTEXT_X86), 716, false);
            check("SYMBOL_INFO", typeof(SYMBOL_INFO), 4088, false);

            // 平台自检：x64 布局是硬编码的，跑在 32 位 CLR 上一切偏移都是错的
            try
            {
                if (IntPtr.Size != 8) r.Hard.Add("进程不是 64 位（IntPtr.Size=" + IntPtr.Size + "），x64 结构布局不适用");
            }
            catch (Exception ex)
            {
                r.Hard.Add("IntPtr.Size 读取失败: " + ex.GetType().Name);
            }

            return r;
        }

        /// <summary>旧签名保留（返回 null 表示全对），仅用于自检输出，不再作为启动门槛。</summary>
        public static string AssertLayouts()
        {
            LayoutCheckResult r = CheckLayouts();
            var all = new System.Collections.Generic.List<string>();
            all.AddRange(r.Hard);
            all.AddRange(r.Soft);
            return all.Count == 0 ? null : string.Join("; ", all.ToArray());
        }
    }
}
