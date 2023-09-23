@echo off

call build.bat || goto :err

cd ./PowerPointArrangeAddin/

sed -i "s/asmv2:publisher=\"PowerPointArrangeAddin\"/asmv2:publisher=\"AoiHosizora\" asmv2:supportUrl=\"https:\/\/github.com\/Aoi-hosizora\/PowerPointArrangeAddin\"/" ./bin/x64/Release/PowerPointArrangeAddin.vsto || goto :err

for /f %%i in (PowerPointArrangeAddin_TemporaryKey.pfx.pwd) do set PASSWORD=%%i
mage -Sign ./bin/x64/Release/PowerPointArrangeAddin.vsto -CertFile PowerPointArrangeAddin_TemporaryKey.pfx -Password %PASSWORD% || goto :err

rm ../Release/ -rf
cp -r ./bin/x64/Release/ ../Release/

msbuild PowerPointArrangeAddin.csproj /p:Configuration=Release /p:Platform=x64 /t:Clean

cd ..
echo.
echo Done!
goto :eof

:err
cd ..
echo Failed to Publish!
exit /B 1
