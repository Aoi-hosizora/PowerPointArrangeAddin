@echo off

cd %~dp0/ || exit /B 1

3rdparty\enigmavbconsole.exe PackLauncher.evb || exit /B 1
