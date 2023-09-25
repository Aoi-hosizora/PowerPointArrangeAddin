@echo off

if "%VSCMD_VER%"=="" (
    call vsdevcmd.bat || goto :err
)

cd ./PowerPointArrangeAddin/

msbuild ./PowerPointArrangeAddin.csproj /p:Configuration=Release /p:Platform=x64 /t:Clean,Build || goto :err

sed -i "s/asmv2:publisher=\"PowerPointArrangeAddin\"/asmv2:publisher=\"AoiHosizora\" asmv2:supportUrl=\"https:\/\/github.com\/Aoi-hosizora\/PowerPointArrangeAddin\"/" ^
    ./bin/x64/Release/PowerPointArrangeAddin.vsto || goto :err

for /f %%i in (./PowerPointArrangeAddin_TemporaryKey.pfx.pwd) do set PASSWORD=%%i
mage -Sign ./bin/x64/Release/PowerPointArrangeAddin.vsto -CertFile ./PowerPointArrangeAddin_TemporaryKey.pfx -Password %PASSWORD% || goto :err

cd ..
echo.
echo Done!
goto :eof

:err
cd ..
echo.
echo Failed to build!
exit /B 1
