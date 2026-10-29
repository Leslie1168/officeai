# OfficeAi.Watchdog —— 进程外崩溃看门狗

> T3 交付物。.NET Framework 4.8 console，x64，**零 NuGet 依赖**，单文件 exe 可整目录拷走。

## 为什么需要它

Word 闪退（`wwlib.dll` `0xc0000005` @ `0x183ba19`）一直拿不到任何 dump：

| 通道 | 为什么不行 |
|---|---|
| WER / LocalDumps | 本机 `HKLM\SOFTWARE\Microsoft\Windows\Windows Error Reporting\Disabled = 1`（全局禁用） |
| 进程内 `SetUnhandledExceptionFilter` | 进程级"后装者赢"，Office / CLR / WebView2 任一方后来装过滤器就顶掉它（报告 3.6） |
| 进程内 VEH | 仍活在目标进程里；崩溃发生在"过滤器跑不了"的状态时就没了 |

调试器（`DebugActiveProcess`）走的是**内核通道**：异常在分发到目标进程内任何处理器**之前**，
先以 `EXCEPTION_DEBUG_EVENT` 送到看门狗。这是唯一"先于一切处理器"的位置。

## 命令行契约（**addin 侧按这个调用，不要改**）

```
OfficeAi.Watchdog.exe --pid <PID> --out <dump目录>          # 主用（T2-F 就是拉起这个）
OfficeAi.Watchdog.exe --exe <path> [--args "..."] [--out]   # 备用：以调试方式启动
OfficeAi.Watchdog.exe --selftest [--out]                    # 自检：不需要 Word，自己制造崩溃并验证整条链
```

- 日志：`<out>\watchdog.log`（无缓冲逐行落盘；宿主是 `CreateNoWindow` 拉起的，stdout 无人接收，**日志是唯一通道**）
- 产物：`<out>\<proc>_<pid>_0x<码>_0x<地址>_<时间>.dmp` + `<同名>.exception.txt`（sidecar）
- 退出码：`0` 正常 / `2` 参数错误 / `3` 日志不可写 / `4` attach 失败 / `5` 启动目标失败 / `7` 顶层异常
- 可选：`--try-exception-param`、`--client-pointers`、`--symbol-budget`、`--symbol-wait`、`--symbol-cache`、`--probe-timeout`

## ★ 本机 `dbghelp.dll` 的缺陷（这份代码为什么长这样）

**实测结论（Lead 独立实验 + 本工具 `--selftest` / DumpProbe 双向复现）：**

1. `MiniDumpWriteDump(..., ExceptionParam = NULL, ...)` —— **稳定可用**。
2. `ExceptionParam != NULL` —— **本机必然失败**，返回 `0x800703E6`（`HRESULT_FROM_WIN32(998)` `ERROR_NOACCESS`），
   而且**会让调用进程直接崩掉**（probe 的 `T3.dmp` 是 0 字节、进程随后消失）。
   六种组合全部同样失败：`ClientPointers` 取反 / `ContextRecord=NULL` / 真实 `GetThreadContext` /
   只给 `ContextRecord` / 换 DumpType。**在自己的进程上自 dump 也一样失败** → 不是跨进程指针问题，是这条通道坏了。

**因此采用「dump + sidecar」双产物：**

- **dump**：`ExceptionParam=NULL` 的全内存 dump。全内存 dump 里**本来就包含**出事线程的 CONTEXT
  （在 `ThreadListStream` 的 `ThreadContext` 里），只是没有 `ExceptionStream` 把它串起来。
- **sidecar**：`<dump同名>.exception.txt` —— 异常码 / 异常地址 / **模块名+基址+RVA** / 访问类型 /
  出错线程 ID（含 cdb 需要的**十六进制**形式）/ 寄存器 / 栈候选 + 可直接复制的 cdb 命令。
  **定位 `wwlib+0x183BA19` 靠的是这一行，它完全不依赖 dbghelp。**
- 判据：**看文件是否真的 > 1MB，不看返回值**（本机出现过"返回值失败但文件其实是好的"与"文件 0 字节"两种情形）。
- `--try-exception-param`（默认关）会把③排在①**之后**：先保住成品，再赌那把 dbghelp 正常的机器。

## cdb 分析姿势

```
cdb -z <dump> -y "<空目录或本地符号缓存>" -c "~~[<十六进制tid>]s; r; k; q"
```

- ★ `~~[tid]` 用**十六进制**。sidecar 里两个进制都给了，直接用十六进制那个。
- ★ `!analyze -v` 会卡在微软符号服务器下载上（实测 180s 超时），**不要用**。
- ★ 本 dump 没有 `ExceptionStream`，`.ecxr` 会报 `Minidump doesn't have an exception context`——
  这是已知的本机 dbghelp 缺陷导致，不是 dump 坏了。用 `~~[tid]s` 达到同样目的。
- 已实测：`~~[tid]s; r; k` 能正常读出寄存器与完整调用栈（含 `ntdll!RtlMoveMemory+0xa7` 这样的符号）。

## 行为承诺（不可回退）

- **只取证，不改变目标行为**：所有真实异常一律 `DBG_EXCEPTION_NOT_HANDLED` 原样放行；
  只有 attach 后的初始断点用 `DBG_CONTINUE`。不吞异常、不改变 Word 的退出路径。
- **`DebugSetProcessKillOnExit(FALSE)`**：看门狗退出（含被 taskkill / 崩溃 / 关机）绝不会连带杀掉 Word。
  已用对照实验证明（见下）。
- 句柄纪律：`CREATE_PROCESS.hFile` 与各线程 `hThread` 一律 `CloseHandle`（Word 有几十个线程）；
  脚本执行前后打印自身句柄数。
- 符号解析：预检符号服务器可达性（2s 超时）→ 不可达则只走本地缓存 → 全程跑在**独立线程**上，
  主调试循环**绝不等待**。拿不到符号就退回「模块名+RVA」。
- 模块归属两条独立通道：`EnumProcessModulesEx` + `VirtualQueryEx`/`GetMappedFileNameW` 反查兜底。
  （实测：进程刚被调试启动时模块表会返回 `win32=299 ERROR_PARTIAL_COPY`，所以改为**延迟到异常发生时**再枚举。）

## 验收（可复现）

```
dotnet build office-ai\watchdog\OfficeAi.Watchdog.csproj -c Release      # 0 error 0 warning
OfficeAi.Watchdog.exe --selftest --out <dir>                            # 47 项检查全绿，退出码 0
_dev\watchdog-tests\run.ps1                                             # 27 项总验收全绿（含 cdb 复核与安全对照）
```

`--selftest` 覆盖 6 个场景：`--exe` 崩溃 / `--pid` attach 崩溃 / ExceptionParam 关闭与打开的降级顺序 /
不崩不写 dump / attach 失败必须留痕 / 结构体布局自检。

`run.ps1` 额外做 cdb 独立复核（逐位比对 RIP、模块基址、RVA）与 `DebugSetProcessKillOnExit` 对照实验。

## 已知限制

### ★ 第一号坑：本机 `dbghelp.dll` 的 `ExceptionParam` 通道是坏的（下一位接手者最容易踩）

> **这一条比代码值钱。踩进去的表现是「看门狗调用 dump 之后整个进程消失」，很容易被误判成"我们的代码有问题"。**

- **现象**：`MiniDumpWriteDump` 只要 `ExceptionParam != NULL`，就会**让调用进程直接崩溃**
  （实测 probe 产出 `T3.dmp` = **0 字节**，进程随后消失），而**不是**返回一个可以处理的错误码。
  因此**绝不能**把它放在任何"保命产物"之前执行 —— 这也是为什么
  `--try-exception-param` 默认关闭，且开启时也排在 NULL 版本 dump **落盘之后**。
- **错误码**：`0x800703E6` = `HRESULT_FROM_WIN32(998)` = **`ERROR_NOACCESS`（"Invalid access to memory location"）**。
  注意别把它读成 `0x80070057`(E_INVALIDARG) 或 `0x800700E6`：`998` 说明 dbghelp
  **尝试去读异常信息所在的内存并失败了**，而不是嫌参数格式错。
  ⇒ 所以"失败就降级重试一次"只改 `DumpType` 是**没有意义的**，必须整个去掉 `ExceptionParam`。
- **无关变量**（六种组合全部同样失败，别再逐个试了）：`ClientPointers` 取 `TRUE`/`FALSE`、
  `ContextRecord` 置 `NULL`、用真实 `GetThreadContext` 的 CONTEXT、只给 `ContextRecord`、
  换 `DumpType`。**在 `hProcess = GetCurrentProcess()` 上自 dump 也一样失败** →
  这不是"跨进程指针"问题，是本机这条通道本身坏了。
- **影响**：dump 里没有 `ExceptionStream`，于是 cdb 里
  `.ecxr` 会报 `Minidump doesn't have an exception context`
  （实测原文：`Unable to get exception context, HRESULT 0x80004002`）。
  **这是 dbghelp 缺陷的已知后果，不是 dump 损坏**，dump 本身完全可分析。
- **替代姿势**（已实测可读出寄存器与完整调用栈）：
  ```
  cdb -z <dump> -y "<空目录或本地符号缓存>" -c "~~[<十六进制tid>]s; r; kb; q"
  ```
  - **`~~[tid]` 用十六进制**。sidecar 里十进制与十六进制都给了，直接用十六进制那个；
    写成十进制会切到别的线程或报 `Illegal thread error`（这个坑已实测踩过）。
  - `kb` 会连没有符号的帧也带出原始地址，比只跑 `k` 更适合定位第三方 DLL 深处的问题。
  - 出错线程的 CONTEXT 仍然在全内存 dump 的 `ThreadListStream` 里，所以 `r`/`k` 都有数据。
  - **定位"是哪个 dll 的哪个偏移"根本不需要 cdb**：sidecar 第一节的
    `★ 模块归属 : <dll> + 0x<RVA>` 就是答案（`0x183BA19` 这类）。
- **不受影响的通道**：`ExceptionParam = NULL` 的写法（本工具默认路径）稳定可用；
  `--pid` attach、模块归属（`EnumProcessModulesEx` + `VirtualQueryEx` 双通道）、
  寄存器/栈候选、sidecar 全部与 dbghelp 无关。

### 其它

- **x64 only**：`DEBUG_EVENT`/`EXCEPTION_RECORD` 的偏移是按 x64 硬编码的，布局自检会拦下 32 位 CLR；
  面对 32 位目标时 `DebugActiveProcess` 通常直接失败（日志会说明）。
- `app.manifest` 是 `asInvoker` 而非 `requireAdministrator`：addin 用
  `Process.Start(UseShellExecute=false)` 拉起，`CreateProcess` **不会弹 UAC**，要求提权会直接
  `ERROR_ELEVATION_REQUIRED` 启动失败。`asInvoker` 会继承 Word 的完整性级别，普通与提权的 Word 都够用。
- 单个 dump 只写**第一次**符合条件的异常（防风暴）；后续异常只记日志。
- 全内存 dump 体积可能很大（Word 级进程 1~2GB），磁盘不足 3GB 时会提前告警。
- attach 模式下模块表会被枚举两次（`DoDebugActiveProcess` 一次、随后 `OnCreateProcess`
  置空并改为延迟枚举）。**功能正确**，代价只是 attach 瞬间多花约 24ms 与两行看起来矛盾的日志；
  已知但未修（冻结产物时不做非功能性改动）。
