@echo off

cd %~dp0/ || goto :err

set RELEASE_DIR=./bin/x64/Release

cp %RELEASE_DIR%/en-US/setup.msi %RELEASE_DIR%/setup.msi || goto :err
rm -rf %RELEASE_DIR%/zh-Hans %RELEASE_DIR%/zh-Hant || goto :err
mv %RELEASE_DIR%/zh-CN %RELEASE_DIR%/zh-Hans || goto :err
mv %RELEASE_DIR%/zh-TW %RELEASE_DIR%/zh-Hant || goto :err

torch -t language %RELEASE_DIR%/setup.msi %RELEASE_DIR%/ja-JP/setup.msi -out %RELEASE_DIR%/transforms/ja-JP.mst || goto :err
torch -t language %RELEASE_DIR%/setup.msi %RELEASE_DIR%/zh-Hans/setup.msi -out %RELEASE_DIR%/transforms/zh-Hans.mst || goto :err
torch -t language %RELEASE_DIR%/setup.msi %RELEASE_DIR%/zh-Hant/setup.msi -out %RELEASE_DIR%/transforms/zh-Hant.mst || goto :err

cscript 3rdparty\WiSubStg.vbs %RELEASE_DIR%/setup.msi %RELEASE_DIR%/transforms/ja-JP.mst 1041 || goto :err
cscript 3rdparty\WiSubStg.vbs %RELEASE_DIR%/setup.msi %RELEASE_DIR%/transforms/zh-Hans.mst 2052 || goto :err
cscript 3rdparty\WiSubStg.vbs %RELEASE_DIR%/setup.msi %RELEASE_DIR%/transforms/zh-Hant.mst 1028 || goto :err

cscript 3rdparty\WiLangId.vbs %RELEASE_DIR%/setup.msi Package 1033,1041,2052,1028 || goto :err

echo.
echo Done!
goto :eof

:err
echo.
echo Failed to transform!
exit /B 1
