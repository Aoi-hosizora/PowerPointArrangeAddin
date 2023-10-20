@echo off

if "%VSCMD_VER%"=="" (
    call vsdevcmd.bat || goto :err
)

msbuild ./PowerPointArrangeAddin.sln /p:Configuration=Release /p:Platform=x64 /t:Clean || goto :err
rm -rf ./PowerPointArrangeAddin/bin/ ./PowerPointArrangeAddin/obj/
rm -rf ./PowerPointArrangeAddinSetup/PowerPointArrangeAddinInstaller/bin/ ./PowerPointArrangeAddinSetup/PowerPointArrangeAddinInstaller/obj/
rm -rf ./PowerPointArrangeAddinSetup/PowerPointArrangeAddinInstallerAction/bin/ ./PowerPointArrangeAddinSetup/PowerPointArrangeAddinInstallerAction/obj/
rm -rf ./PowerPointArrangeAddinSetup/PowerPointArrangeAddinInstallerLauncher/bin/ ./PowerPointArrangeAddinSetup/PowerPointArrangeAddinInstallerLauncher/obj/

echo.
echo Done!
goto :eof

:err
echo.
echo Failed to Clean!
exit /B 1
