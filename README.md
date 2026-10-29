# Office AI 助手（OfficeAI）

一个跑在本机的 Office 侧边栏 AI 插件，支持 **Word / Excel / PowerPoint**。
功能目标：替代海鹦 OfficeAI 助手那类商业插件，但**不依赖任何云端中转**——模型接口由你自己填。

当前版本：**0.5.2**

---

## 它能干什么

| 能力 | 说明 |
|---|---|
| **提问** | 读当前文档全文提问；可用「用选区」锁定一段只针对它回答。只回答，不改文档 |
| **写作** | 生成内容先在侧边栏预览，必须点「确认导入」才写进文档 |
| **精准定位** | 选区改写 / 查找替换；Word 表格可定位到具体单元格 |
| **生成 PPT** | 按大纲格式生成 .pptx，内置 tech-cyan / business-blue / forest-green / warm-amber 四套配色，本地 python-pptx 渲染 |
| **图片抓取** | 抓当前文档里所有内嵌图片，可预览、另存、批量导出 |
| **ima 资料库** | 接腾讯 ima 的 openapi，搜知识库并把结果插入对话（需自备 client_id / api_key） |

全部在本机运行，**只有对话内容会发到你配置的模型接口**。

---

## 系统要求

| 项 | 要求 |
|---|---|
| 操作系统 | Windows 10 / 11（x64） |
| Office | Microsoft Office 2013 及以上（Word / Excel / PowerPoint 桌面版） |
| 运行时 | .NET Framework 4.8（Win10 1903+ / Win11 自带） |
| Python | **不需要装**。Release 包内已自带便携版 Python 3.12.10 + 依赖（约 52 MB） |
| WebView2 | Microsoft Edge WebView2 Runtime（Win11 自带；Win10 若侧边栏白屏则装一次） |
| 权限 | 安装阶段需要**管理员**（要写 `HKLM` 注册 COM 加载项） |

> 注意：只支持 **Microsoft Office**。WPS 用的是另一套加载项机制，本插件挂不上。

---

## 安装（三步）

### 1. 下载

去 [Releases](../../releases/latest) 下载最新版 **`OfficeAI-0.5.2.zip`**（约 20 MB，已内置 PPT 引擎所需的独立 Python 环境）。

### 2. 解压

解压到一个**固定、不会被清理**的位置，例如：

```
D:\Apps\OfficeAI
```

> ⚠️ 不要放在下载目录后就删掉——安装脚本会把这里的文件复制进 `C:\Program Files\OfficeAi`，而这个解压目录你自己要留着以后升级/重装用。路径**不要带空格以外的特殊字符**。

### 3. 安装

双击解压目录里的 **`install.bat`**：

- 会弹一次 **UAC**，点「是」（要写注册表，必须提权）
- 看到「安装完成」后，**完全退出** Word / Excel / PowerPoint（不是关文档，是关进程；任务管理器里确认 `WINWORD.EXE` 已经没了）
- 重新打开 Word → 功能区出现 **「AI 助手」** 选项卡 → 打开任意文件时侧边栏自动弹出

### 4. 填 API Key

点侧边栏右上角 **⚙ 齿轮** → 填：

| 字段 | 填什么 |
|---|---|
| API Key | 你的 DeepSeek Key（`sk-` 开头） |
| Base URL | `https://api.deepseek.com`（默认）。任何 OpenAI 兼容接口都能填，比如自己的中转 |
| 模型 | `deepseek-chat`（默认） |

保存即生效，配置存在 `%LOCALAPPDATA%\OfficeAi\settings.json`。

---

## 使用要点

### 提问模式
勾选「带文档内容」后提问会自动把全文发给模型；点「用选区」可锁定一段只针对选中内容提问。**这个模式绝不改文档。**

### 写作模式 + 确认导入
生成的内容先出现在侧边栏，**必须点「确认导入」**才写进文档。写入方式在设置里二选一：

- `模拟输入`（推荐）——走宿主自己的"光标处输入"，等于你手打，格式与撤销栈都自然
- `VBA` —— 生成 VBA 代码注入执行；若宿主没开「信任对 VBA 工程对象模型的访问」，会自动退回对象模型并把代码显示出来

### 生成 PPT（仅 PowerPoint 宿主可见）
点底部「生成 PPT」，按这个格式写大纲：

```
@title 演示文稿标题
@subtitle 副标题
@theme  tech-cyan

# 第一页标题
- 要点一
  - 子要点（前面空两格）
```

「用当前文档内容起草大纲」可以一键生成骨架。生成的 pptx 默认存桌面并自动打开。

### 图片抓取
点顶部「图片」标签 → 「抓取文档图片」。右键缩略图可打开 / 复制路径 / 另存为 / 在资源管理器显示；「批量导出…」一次性复制整个文件夹。

提取结果落在：
```
%LOCALAPPDATA%\OfficeAi\imgcache\<文档名>_<hash>\orig\<序号>.<扩展名>
```

---

## 卸载

双击 **`uninstall.bat`**（同样会弹 UAC），会删掉：

- `HKLM\SOFTWARE\Classes\OfficeAi.Connect` 及其 CLSID
- `HKLM\SOFTWARE\Classes\OfficeAi.TaskPaneHost` 及其 CLSID
- `HKLM\SOFTWARE\Microsoft\Office\{Word,Excel,PowerPoint}\Addins\OfficeAi.Connect`
- 目录 `C:\Program Files\OfficeAi`

**不会删** `%LOCALAPPDATA%\OfficeAi`（里面有你的 Key 和日志）。要彻底清干净，卸载后手动删掉它。

---

## 故障排查

| 现象 | 处理 |
|---|---|
| 功能区没有「AI 助手」 | 确认 Office 是**完全退出后重开**的；文件 → 选项 → 加载项 → 管理选「COM 加载项」→ 进转到，勾选 `OfficeAi.Connect` |
| 侧边栏白屏 / 「WebView2 初始化失败」 | 装一次 [WebView2 Runtime](https://developer.microsoft.com/microsoft-edge/webview2/)；确认 `C:\Program Files\OfficeAi\WebView2Loader.dll` 存在 |
| 提示「还没有填 API Key」 | 点齿轮填 Key 并保存 |
| 图片抓取提示「请先保存」 | 文档还没存盘，Ctrl+S 之后重试 |
| ima 面板报错 | 设置里检查 `imaClientId` / `imaApiKey`；凭据从 <https://ima.qq.com/agent-interface> 领取 |
| 装完 Word 一开就崩 | 删 `HKCU\Software\Microsoft\Office\16.0\{Word,Excel,PowerPoint}\Resiliency\DisabledItems` 后重装；日志在 `%LOCALAPPDATA%\OfficeAi\addin.log` |

日志位置：`%LOCALAPPDATA%\OfficeAi\addin.log`

---

## 目录结构

```
office-ai/
├─ addin/          C# COM 加载项（Ribbon + TaskPane + WebView2 宿主）
│  ├─ ui/          侧边栏前端（原生 HTML/CSS/JS，无框架无构建）
│  └─ icons/       Ribbon 图标
├─ client/         独立窗口宿主（UI 模式 separate 时用）
├─ vba-worker/     VBA 注入执行器（独立进程，避免拖崩宿主）
├─ updater/        自更新器（把内嵌载荷写回 Program Files）
├─ watchdog/       进程外崩溃取证（仅排障时启用）
├─ engine/         PPT 引擎（make_pptx.py，python-pptx）
│  └─ python/      便携版 Python 3.12.10（自包含，不依赖系统 Python；仅随 Release 分发）
├─ shared/         前后端共用的数据结构
├─ install.bat     安装（提权写 HKLM）
├─ uninstall.bat   卸载
└─ deploy.bat      开发用：把新构建部署到本机 Program Files
```

---

## 从源码构建

需要 .NET SDK（能编 `net48` 的任意版本）与 NuGet 联网还原一次。

```powershell
dotnet build addin\OfficeAi.Addin.csproj        -c Release
dotnet build vba-worker\OfficeAi.VbaWorker.csproj -c Release
dotnet build client\OfficeAi.Client.csproj      -c Release
dotnet build updater\OfficeAi.Updater.csproj    -c Release
```

产物在各自 `bin\Release\`。交付目录 `build\` 的组装方式：把 `OfficeAi.Addin.dll`、`OfficeAi.Client.exe`、`OfficeAi.VbaWorker.exe`、三个 `Microsoft.Web.WebView2.*.dll` + `WebView2Loader.dll`、以及 `addin\ui\` 全量文件放进去。

> `build/`、`engine/python/`、`engine/venv/` 不进版本库，只随 Release 的 zip 分发。

### PPT 引擎（便携 Python）

`engine/python/` 是 python.org 的 **embeddable 包**（3.12.10 amd64）+ 手装依赖：

1. 解压 `python-3.12.10-embed-amd64.zip` 到 `engine/python`
2. 把 `python312._pth` 改成
   ```
   python312.zip
   .
   Lib\site-packages
   import site
   ```
3. `pip install --target engine/python/Lib/site-packages python-pptx lxml pillow xlsxwriter`

这样整包可搬走，目标机不需要装任何 Python。

---

## 已知限制

- 只支持 x64 Office，不支持 32 位 Office
- 只支持 Microsoft Office，不支持 WPS
- PPT 引擎用的是包内自带的便携 Python（`engine\python`），不走系统 Python，也不支持替换
- 崩溃取证（watchdog）默认关闭，日常不要开——它会附加原生调试器

---

## 许可

[MIT](LICENSE)
