@echo off
setlocal

set "APP_NAME=LeagueSheetsApp"
set "SOURCE_DIR=%~dp0LeagueSheetsApp"
set "INSTALL_DIR=%LOCALAPPDATA%\LeagueSheetsApp"
set "EXE_PATH=%INSTALL_DIR%\LeagueSheetsApp.exe"
set "DESKTOP_LINK=%USERPROFILE%\Desktop\LeagueSheetsApp.lnk"
set "START_MENU_DIR=%APPDATA%\Microsoft\Windows\Start Menu\Programs\LeagueSheetsApp"
set "START_MENU_LINK=%START_MENU_DIR%\LeagueSheetsApp.lnk"

title LeagueSheetsApp installer
echo.
echo LeagueSheetsApp installer
echo =========================
echo.

if not exist "%SOURCE_DIR%\LeagueSheetsApp.exe" (
    echo Could not find the LeagueSheetsApp folder.
    echo.
    echo First unzip the package, then run Install.bat from the same folder
    echo where the LeagueSheetsApp folder is located.
    echo.
    pause
    exit /b 1
)

echo Installing to:
echo %INSTALL_DIR%
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference='Stop'; $source=$env:SOURCE_DIR; $target=$env:INSTALL_DIR; if (Test-Path -LiteralPath $target) { Remove-Item -LiteralPath $target -Recurse -Force }; Copy-Item -LiteralPath $source -Destination $target -Recurse -Force; New-Item -ItemType Directory -Force -Path $env:START_MENU_DIR | Out-Null; $shell=New-Object -ComObject WScript.Shell; $desktop=$shell.CreateShortcut($env:DESKTOP_LINK); $desktop.TargetPath=$env:EXE_PATH; $desktop.WorkingDirectory=$env:INSTALL_DIR; $desktop.IconLocation=$env:EXE_PATH; $desktop.Save(); $start=$shell.CreateShortcut($env:START_MENU_LINK); $start.TargetPath=$env:EXE_PATH; $start.WorkingDirectory=$env:INSTALL_DIR; $start.IconLocation=$env:EXE_PATH; $start.Save();"

if errorlevel 1 (
    echo.
    echo Installation failed.
    echo Try running Install.bat again.
    echo.
    pause
    exit /b 1
)

echo Done.
echo.
echo The app is installed and a desktop shortcut was created.
echo.
choice /C YN /N /M "Run the app now? [Y/N] "
if errorlevel 2 exit /b 0
start "" "%EXE_PATH%"

endlocal
