@echo off

cd %~dp0/ || exit /B 1

"D:\Program Files (x86)\Enigma Virtual Box\enigmavbconsole.exe" PackLauncher.evb || exit /B 1
