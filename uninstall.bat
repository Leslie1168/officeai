@echo off
chcp 936 >nul
setlocal
title Office AI 助手 - 卸载

set "ROOT=%ProgramFiles%\OfficeAi"
set "CONNECT={7C3E9A21-4B6D-4F58-9E02-3A1B5C7D8E90}"
set "PANE={B4D8F1A2-9C3E-4A7B-8D15-6E2F0C9A3B47}"

echo ====================================================
echo   Office AI 助手 - 卸载
echo ====================================================
echo.

net session >nul 2>&1
if errorlevel 1 (
    echo   需要管理员权限，即将弹出 UAC，请点「是」。
    powershell -NoProfile -Command "Start-Process -FilePath '%~f0' -Verb RunAs" >nul 2>&1
    exit /b 0
)

echo   将要删除：
echo     HKLM\SOFTWARE\Classes\OfficeAi.Connect 与 CLSID 注册
echo     HKLM\SOFTWARE\Classes\OfficeAi.TaskPaneHost 与 CLSID 注册
echo     HKLM\SOFTWARE\Microsoft\Office\{Word,Excel,PowerPoint}\Addins\OfficeAi.Connect
echo     目录 %ROOT%
echo.
echo   注意：不会删除 %LOCALAPPDATA%\OfficeAi（里面有日志和 API Key 设置）。
echo   如需彻底清除，卸载后手动删除该目录。
echo.
echo   请先完全退出 Word / Excel / PowerPoint。
echo.
set /p "YN=确认卸载？(Y/N) "
if /i not "%YN%"=="Y" (
    echo 已取消。
    pause
    exit /b 0
)

for %%H in (Word Excel PowerPoint) do (
    reg delete "HKLM\SOFTWARE\Microsoft\Office\%%H\Addins\OfficeAi.Connect" /f >nul 2>nul
    reg delete "HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\%%H\Addins\OfficeAi.Connect" /f >nul 2>nul
)
reg delete "HKLM\SOFTWARE\Classes\CLSID\%CONNECT%" /f >nul 2>nul
reg delete "HKLM\SOFTWARE\Classes\OfficeAi.Connect" /f >nul 2>nul
reg delete "HKLM\SOFTWARE\Classes\CLSID\%PANE%" /f >nul 2>nul
reg delete "HKLM\SOFTWARE\Classes\OfficeAi.TaskPaneHost" /f >nul 2>nul

if exist "%ROOT%" (
    rmdir /s /q "%ROOT%"
    if exist "%ROOT%" (
        echo   目录删除不完整，请关闭 Office 后手动删除：%ROOT%
    ) else (
        echo   目录已删除。
    )
)

echo.
echo 卸载完成。
pause
exit /b 0
