@echo off
rem ================================================================
rem  Build script for AppShortcutCreator.
rem  Uses the system built-in csc.exe (.NET Framework 4.x).
rem  NOTE: keep this file ASCII-only to avoid code page issues.
rem ================================================================

setlocal
set "CSC="

if exist "%SystemRoot%\Microsoft.NET\Framework64\v4.0.30319\csc.exe" set "CSC=%SystemRoot%\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
if not defined CSC if exist "%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\csc.exe" set "CSC=%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\csc.exe"

if not defined CSC echo [ERROR] csc.exe not found. Please install .NET Framework 4.x.
if not defined CSC exit /b 1

"%CSC%" /nologo /target:winexe /platform:anycpu /codepage:65001 /optimize+ /out:"%~dp0AppShortcutCreator.exe" /r:System.dll /r:System.Core.dll /r:System.Drawing.dll /r:System.Windows.Forms.dll /r:Microsoft.CSharp.dll "%~dp0AppShortcutCreator.cs"

if errorlevel 1 echo [ERROR] Build failed. See messages above.
if errorlevel 1 exit /b 1

echo [OK] Build succeeded: "%~dp0AppShortcutCreator.exe"
endlocal
