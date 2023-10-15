@echo off

cd %~dp0/../PowerPointArrangeAddinInstaller || goto :err

set RELEASE_DIR=./bin/x64/Release

cp %RELEASE_DIR%/en-US/PowerPointArrangeAddinInstaller.msi %RELEASE_DIR%/PowerPointArrangeAddinInstaller.msi || goto :err

torch -t language %RELEASE_DIR%/PowerPointArrangeAddinInstaller.msi %RELEASE_DIR%/ja-JP/PowerPointArrangeAddinInstaller.msi -out %RELEASE_DIR%/transforms/ja-JP.mst || goto :err
torch -t language %RELEASE_DIR%/PowerPointArrangeAddinInstaller.msi %RELEASE_DIR%/zh-CN/PowerPointArrangeAddinInstaller.msi -out %RELEASE_DIR%/transforms/zh-CN.mst || goto :err
torch -t language %RELEASE_DIR%/PowerPointArrangeAddinInstaller.msi %RELEASE_DIR%/zh-TW/PowerPointArrangeAddinInstaller.msi -out %RELEASE_DIR%/transforms/zh-TW.mst || goto :err

cscript 3rdparty\WiSubStg.vbs %RELEASE_DIR%/PowerPointArrangeAddinInstaller.msi %RELEASE_DIR%/transforms/ja-JP.mst 1041 || goto :err
cscript 3rdparty\WiSubStg.vbs %RELEASE_DIR%/PowerPointArrangeAddinInstaller.msi %RELEASE_DIR%/transforms/zh-CN.mst 2052 || goto :err
cscript 3rdparty\WiSubStg.vbs %RELEASE_DIR%/PowerPointArrangeAddinInstaller.msi %RELEASE_DIR%/transforms/zh-TW.mst 1028 || goto :err

cscript 3rdparty\WiLangId.vbs %RELEASE_DIR%/PowerPointArrangeAddinInstaller.msi Package 1033,1041,2052,1028 || goto :err

echo.
echo Done!
goto :eof

:err
echo.
echo Failed to transform!
exit /B 1
