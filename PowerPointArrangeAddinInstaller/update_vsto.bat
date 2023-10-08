@echo off

cd %~dp0/../PowerPointArrangeAddin || goto :err

sed -i "s/asmv2:publisher=\"PowerPointArrangeAddin\"/asmv2:publisher=\"AoiHosizora\" asmv2:supportUrl=\"https:\/\/github.com\/Aoi-hosizora\/PowerPointArrangeAddin\"/" ./bin/x64/Release/PowerPointArrangeAddin.vsto || goto :err

for /f %%i in (PowerPointArrangeAddin_TemporaryKey.pfx.pwd) do set PASSWORD=%%i
mage -Sign ./bin/x64/Release/PowerPointArrangeAddin.vsto -CertFile PowerPointArrangeAddin_TemporaryKey.pfx -Password %PASSWORD% || goto :err

goto :eof

:err
exit /B 1
