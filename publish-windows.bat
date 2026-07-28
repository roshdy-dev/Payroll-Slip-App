@echo off
REM ============================================================================
REM  Payroll Slip Generator - Publish Script (Windows)
REM  Creates a self-contained, single-file executable that runs without .NET installed.
REM ============================================================================

echo.
echo ╔══════════════════════════════════════════════════╗
echo ║   Payroll Slip Generator - Publish (Windows)     ║
echo ╚══════════════════════════════════════════════════╝
echo.

set PROJECT_DIR=src\PayrollSlipApp
set OUTPUT_DIR=publish\win-x64

echo [1/2] Restoring packages...
dotnet restore "%PROJECT_DIR%\PayrollSlipApp.csproj"
if %ERRORLEVEL% NEQ 0 (
    echo ERROR: Restore failed.
    exit /b %ERRORLEVEL%
)

echo.
echo [2/2] Publishing self-contained single-file executable...
echo        This may take a few minutes on first run...
echo.

dotnet publish "%PROJECT_DIR%\PayrollSlipApp.csproj" ^
    -c Release ^
    -r win-x64 ^
    --self-contained true ^
    -p:PublishSingleFile=true ^
    -p:PublishTrimmed=false ^
    -p:IncludeNativeLibrariesForSelfExtract=true ^
    -o "%OUTPUT_DIR%"

if %ERRORLEVEL% NEQ 0 (
    echo ERROR: Publish failed.
    exit /b %ERRORLEVEL%
)

echo.
echo ╔══════════════════════════════════════════════════╗
echo ║   PUBLISH SUCCESSFUL                             ║
echo ╚══════════════════════════════════════════════════╝
echo.
echo    Output: %OUTPUT_DIR%\PayrollSlipGenerator.exe
echo.
echo    This file can be distributed and run on ANY Windows
echo    machine WITHOUT .NET installed!
echo.
