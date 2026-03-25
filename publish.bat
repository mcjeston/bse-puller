@echo off
setlocal
set "DOTNET_CLI_HOME=%~dp0"
set "DOTNET_SKIP_FIRST_TIME_EXPERIENCE=1"
set "DOTNET_NOLOGO=1"
set "APPDATA=%~dp0.appdata"
set "LOCALAPPDATA=%~dp0.localappdata"
if not exist "%APPDATA%" mkdir "%APPDATA%"
if not exist "%LOCALAPPDATA%" mkdir "%LOCALAPPDATA%"
if exist "%~dp0.dotnet\dotnet.exe" (
  "%~dp0.dotnet\dotnet.exe" publish BsePuller.csproj -c Release --no-restore
) else (
  dotnet publish BsePuller.csproj -c Release --no-restore
)
endlocal
