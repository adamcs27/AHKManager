@echo off
title AHK Script Manager - Build
echo ================================================
echo   AHK Script Manager - Build Tool
echo ================================================
echo.

REM ── Option 1: dotnet CLI (comes with .NET SDK) ──────────────────────────────
where dotnet >nul 2>&1
if %ERRORLEVEL% == 0 (
    echo Found: dotnet CLI
    goto :use_dotnet
)

REM ── Option 2: MSBuild from Visual Studio 2022 ────────────────────────────────
for %%p in (
    "%ProgramFiles%\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe"
    "%ProgramFiles%\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\MSBuild.exe"
    "%ProgramFiles%\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\MSBuild.exe"
    "%ProgramFiles%\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\MSBuild.exe"
    "%ProgramFiles(x86)%\Microsoft Visual Studio\2019\Community\MSBuild\Current\Bin\MSBuild.exe"
    "%ProgramFiles(x86)%\Microsoft Visual Studio\2019\Professional\MSBuild\Current\Bin\MSBuild.exe"
    "%ProgramFiles(x86)%\Microsoft Visual Studio\2019\BuildTools\MSBuild\Current\Bin\MSBuild.exe"
) do (
    if exist %%p (
        set MSBUILD=%%p
        echo Found: %%p
        goto :use_msbuild
    )
)

REM ── Nothing found ────────────────────────────────────────────────────────────
echo ERROR: Neither 'dotnet' nor MSBuild was found.
echo.
echo To fix this, install ONE of the following (both are free):
echo.
echo  A) .NET SDK  ^(recommended - small download^)
echo     https://dotnet.microsoft.com/download
echo     After installing, re-run this script.
echo.
echo  B) Visual Studio 2022 Community
echo     https://visualstudio.microsoft.com/downloads/
echo     Make sure to include the ".NET desktop development" workload.
echo.
pause
exit /b 1

REM ── Build with dotnet CLI ────────────────────────────────────────────────────
:use_dotnet
echo.
echo Building with dotnet CLI...
echo.
dotnet build AHKScriptManager.csproj -c Release --nologo
if %ERRORLEVEL% == 0 goto :success
goto :fail

REM ── Build with MSBuild ───────────────────────────────────────────────────────
:use_msbuild
echo.
echo Building with MSBuild...
echo.
%MSBUILD% AHKScriptManager.csproj /t:Build /p:Configuration=Release /nologo /verbosity:minimal
if %ERRORLEVEL% == 0 goto :success
goto :fail

REM ── Results ──────────────────────────────────────────────────────────────────
:success
echo.
REM Copy assets to output directory
if exist "iconAHKManager.png" (
    if exist "bin\Release\net472\" copy /y "iconAHKManager.png" "bin\Release\net472\" >nul
    if exist "bin\Release\" copy /y "iconAHKManager.png" "bin\Release\" >nul
)
echo.
echo ================================================
echo   Build SUCCESSFUL!
echo ================================================
echo.

REM Find the exe (dotnet puts it in net472 subfolder, msbuild in Release directly)
set EXE=bin\Release\AHKScriptManager.exe
if not exist "%EXE%" set EXE=bin\Release\net472\AHKScriptManager.exe

if exist "%EXE%" (
    echo Output: %EXE%
    echo.
    echo Launching...
    start "" "%EXE%"
) else (
    echo Could not locate the exe automatically.
    echo Check the bin\Release\ folder.
)
pause
exit /b 0

:fail
echo.
echo ================================================
echo   Build FAILED - see errors above
echo ================================================
pause
exit /b 1
