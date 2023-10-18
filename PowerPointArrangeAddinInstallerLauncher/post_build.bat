@echo off

cd %~dp0/ || exit /B 1

@REM cp "D:\Program Files (x86)\Enigma Virtual Box\enigmavbconsole.exe" PowerPointArrangeAddinInstallerLauncher\3rdparty
3rdparty\enigmavbconsole.exe PackLauncher.evb || exit /B 1
