@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
set "SLIDEPACE_MSBUILD=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\MSBuild.exe"
if exist "%ProgramFiles(x86)%\Microsoft Visual Studio\Installer\vswhere.exe" (
  for /f "usebackq delims=" %%I in (`"%ProgramFiles(x86)%\Microsoft Visual Studio\Installer\vswhere.exe" -latest -products * -requires Microsoft.Component.MSBuild -find MSBuild\**\Bin\MSBuild.exe`) do set "SLIDEPACE_MSBUILD=%%I"
)
"%SLIDEPACE_MSBUILD%" SlidePace.sln /t:Build /p:Configuration=Release /nologo /v:minimal
if errorlevel 1 exit /b 1
if not exist dist mkdir dist
copy /y "src\SlidePace.Setup\bin\Release\SlidePace-Setup.exe" "dist\SlidePace-Setup.exe" >nul
copy /y "README.md" "dist\使用说明.md" >nul
echo Build completed: dist\SlidePace-Setup.exe
endlocal
