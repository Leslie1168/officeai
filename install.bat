@echo off
chcp 936 >nul
setlocal
title Office AI 助手 - 安装

set "SRC=%~dp0"
set "ROOT=%ProgramFiles%\OfficeAi"
set "DLL=%ROOT%\OfficeAi.Addin.dll"
set "ASM=OfficeAi.Addin, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null"
set "CONNECT={7C3E9A21-4B6D-4F58-9E02-3A1B5C7D8E90}"
set "PANE={B4D8F1A2-9C3E-4A7B-8D15-6E2F0C9A3B47}"

echo ====================================================
echo   Office AI 助手 - 安装
echo ====================================================
echo.

net session >nul 2>&1
if errorlevel 1 (
    echo   需要管理员权限来注册 Office 加载项（写入 HKLM）。
    echo   即将弹出 UAC，请点「是」。
    echo.
    powershell -NoProfile -Command "Start-Process -FilePath '%~f0' -Verb RunAs" >nul 2>&1
    exit /b 0
)

echo [1/5] 检查交付文件...
if not exist "%SRC%build\OfficeAi.Addin.dll" (
    echo   错误：找不到 build\OfficeAi.Addin.dll
    goto :fail
)
if not exist "%SRC%engine\make_pptx.py" (
    echo   错误：找不到 engine\make_pptx.py
    goto :fail
)
if not exist "%SRC%engine\venv\Scripts\python.exe" (
    echo   错误：找不到 engine\venv，无法生成 PPT
    goto :fail
)

echo [2/5] 复制程序文件...
if not exist "%ROOT%" mkdir "%ROOT%"
copy /y "%SRC%build\OfficeAi.Addin.dll" "%ROOT%\" >nul || goto :fail
robocopy "%SRC%build" "%ROOT%" /e /njh /njs /ndl /np /nfl >nul
if errorlevel 8 goto :fail
rem .NET Framework 需要在程序根目录找到 WebView2Loader.dll
if exist "%ROOT%\runtimes\win-x64\native\WebView2Loader.dll" (
    copy /y "%ROOT%\runtimes\win-x64\native\WebView2Loader.dll" "%ROOT%\" >nul
)

echo [3/5] 复制 PPT 引擎（含独立 Python 环境，约 45 MB）...
robocopy "%SRC%engine" "%ROOT%\engine" /e /njh /njs /ndl /np /nfl >nul
if errorlevel 8 goto :fail

echo [4/5] 注册 COM 组件到 HKLM...
set "CB=%ROOT:\=/%"
set "CB=file:///%CB%/OfficeAi.Addin.dll"
set "K=HKLM\SOFTWARE\Classes"

rem ---- 加载项类 ----
reg add "%K%\OfficeAi.Connect" /ve /t REG_SZ /d "OfficeAi.Connect" /f >nul || goto :regfail
reg add "%K%\OfficeAi.Connect\CLSID" /ve /t REG_SZ /d "%CONNECT%" /f >nul || goto :regfail
reg add "%K%\CLSID\%CONNECT%" /ve /t REG_SZ /d "OfficeAi.Connect" /f >nul || goto :regfail
reg add "%K%\CLSID\%CONNECT%\ProgId" /ve /t REG_SZ /d "OfficeAi.Connect" /f >nul
reg add "%K%\CLSID\%CONNECT%\Implemented Categories\{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}" /f >nul
reg add "%K%\CLSID\%CONNECT%\InprocServer32" /ve /t REG_SZ /d "mscoree.dll" /f >nul
reg add "%K%\CLSID\%CONNECT%\InprocServer32" /v ThreadingModel /t REG_SZ /d "Both" /f >nul
reg add "%K%\CLSID\%CONNECT%\InprocServer32" /v Class /t REG_SZ /d "OfficeAi.Connect" /f >nul
reg add "%K%\CLSID\%CONNECT%\InprocServer32" /v Assembly /t REG_SZ /d "%ASM%" /f >nul
reg add "%K%\CLSID\%CONNECT%\InprocServer32" /v RuntimeVersion /t REG_SZ /d "v4.0.30319" /f >nul
reg add "%K%\CLSID\%CONNECT%\InprocServer32" /v CodeBase /t REG_SZ /d "%CB%" /f >nul

rem ---- 任务窗格宿主：必须注册成 ActiveX 控件，否则 CreateCTP 会失败 ----
reg add "%K%\OfficeAi.TaskPaneHost" /ve /t REG_SZ /d "OfficeAi.TaskPaneHost" /f >nul || goto :regfail
reg add "%K%\OfficeAi.TaskPaneHost\CLSID" /ve /t REG_SZ /d "%PANE%" /f >nul || goto :regfail
reg add "%K%\CLSID\%PANE%" /ve /t REG_SZ /d "OfficeAi.TaskPaneHost" /f >nul || goto :regfail
reg add "%K%\CLSID\%PANE%\ProgId" /ve /t REG_SZ /d "OfficeAi.TaskPaneHost" /f >nul
reg add "%K%\CLSID\%PANE%\Version" /ve /t REG_SZ /d "1.0" /f >nul
reg add "%K%\CLSID\%PANE%\InprocServer32" /ve /t REG_SZ /d "mscoree.dll" /f >nul
reg add "%K%\CLSID\%PANE%\InprocServer32" /v ThreadingModel /t REG_SZ /d "Both" /f >nul
reg add "%K%\CLSID\%PANE%\InprocServer32" /v Class /t REG_SZ /d "OfficeAi.TaskPaneHost" /f >nul
reg add "%K%\CLSID\%PANE%\InprocServer32" /v Assembly /t REG_SZ /d "%ASM%" /f >nul
reg add "%K%\CLSID\%PANE%\InprocServer32" /v RuntimeVersion /t REG_SZ /d "v4.0.30319" /f >nul
reg add "%K%\CLSID\%PANE%\InprocServer32" /v CodeBase /t REG_SZ /d "%CB%" /f >nul
reg add "%K%\CLSID\%PANE%\Control" /f >nul
reg add "%K%\CLSID\%PANE%\MiscStatus" /ve /t REG_SZ /d "131473" /f >nul
reg add "%K%\CLSID\%PANE%\ToolboxBitmap32" /ve /t REG_SZ /d "%DLL%,1" /f >nul

echo [5/5] 挂载到 Office 加载项列表...
for %%H in (Word Excel PowerPoint) do (
    reg add "HKLM\SOFTWARE\Microsoft\Office\%%H\Addins\OfficeAi.Connect" /v FriendlyName /t REG_SZ /d "Office AI 助手" /f >nul
    reg add "HKLM\SOFTWARE\Microsoft\Office\%%H\Addins\OfficeAi.Connect" /v Description /t REG_SZ /d "DeepSeek 驱动的文档助手" /f >nul
    reg add "HKLM\SOFTWARE\Microsoft\Office\%%H\Addins\OfficeAi.Connect" /v LoadBehavior /t REG_DWORD /d 3 /f >nul
    reg add "HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\%%H\Addins\OfficeAi.Connect" /v LoadBehavior /t REG_DWORD /d 3 /f >nul
    reg delete "HKCU\Software\Microsoft\Office\16.0\%%H\Resiliency\DisabledItems" /f >nul 2>nul
)

echo.
echo ====================================================
echo   安装完成
echo ====================================================
echo.
echo   下一步：
echo   1. 完全退出并重新打开 Word / Excel / PowerPoint
echo   2. 功能区会出现「AI 助手」选项卡，点「打开侧边栏」
echo   3. 在侧边栏右上角齿轮里填 DeepSeek 的 API Key
echo.
pause
exit /b 0

:regfail
echo.
echo   注册表写入失败（可能不是管理员权限）。
echo   请右键 install.bat -^> 以管理员身份运行。
echo.
pause
exit /b 1

:fail
echo.
echo   安装失败，请把上面的错误信息截图反馈。
echo.
pause
exit /b 1
