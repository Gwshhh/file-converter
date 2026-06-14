@echo off
chcp 65001 >nul
REM 一键更新发布脚本
REM 用途: 修改代码后，自动打包、提交、发布新版本

echo ==========================================
echo   File Converter - Auto Update Release
echo ==========================================
echo.

REM 检查是否有未提交的修改
git status -s >nul 2>&1
if %errorlevel% equ 0 (
    echo Checking for uncommitted changes...
    for /f %%i in ('git status -s') do (
        goto :has_changes
    )
    goto :no_changes

    :has_changes
    echo.
    echo [!] Uncommitted changes detected
    git status -s
    echo.
    set /p commit_msg="Enter commit message: "

    if "!commit_msg!"=="" (
        echo ERROR: Commit message cannot be empty
        pause
        exit /b 1
    )

    echo.
    echo [1/4] Committing changes...
    git add .
    git commit -m "!commit_msg!"

    if %errorlevel% neq 0 (
        echo ERROR: Commit failed
        pause
        exit /b 1
    )
    echo Done!
    goto :push

    :no_changes
    echo [OK] No uncommitted changes
)

:push
echo.
echo [2/4] Pushing to GitHub...
git push

if %errorlevel% neq 0 (
    echo ERROR: Push failed
    pause
    exit /b 1
)
echo Done!

echo.
echo [3/4] Building executable...
python -m PyInstaller 文件转换器.spec --clean

if %errorlevel% neq 0 (
    echo ERROR: Build failed
    pause
    exit /b 1
)

if not exist "dist\文件转换器.exe" (
    echo ERROR: Executable not found
    pause
    exit /b 1
)

echo Done!
for %%A in ("dist\文件转换器.exe") do echo    Size: %%~zA bytes

echo.
set /p new_version="Enter new version (e.g., 1.0.2): "

if "%new_version%"=="" (
    echo ERROR: Version cannot be empty
    pause
    exit /b 1
)

echo.
set /p release_notes="Enter release notes (Enter for default): "

if "%release_notes%"=="" (
    set "release_notes=Version %new_version% update"
)

echo.
echo [4/4] Publishing v%new_version% to GitHub...

gh release create "v%new_version%" "dist\文件转换器.exe" --title "v%new_version%" --notes "%release_notes%"

if %errorlevel% equ 0 (
    echo.
    echo ==========================================
    echo   SUCCESS!
    echo ==========================================
    echo.
    echo Version: v%new_version%
    echo Download: https://github.com/Gwshhh/file-converter/releases/tag/v%new_version%
    echo.
) else (
    echo ERROR: Release creation failed
    pause
    exit /b 1
)

pause
