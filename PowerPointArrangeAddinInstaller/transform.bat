@echo off

set RELEASE_DIR=./bin/x64/Release

cp %RELEASE_DIR%/en-US/PowerPointArrangeAddinInstaller.msi %RELEASE_DIR%/PowerPointArrangeAddinInstaller.msi

torch -t language %RELEASE_DIR%/PowerPointArrangeAddinInstaller.msi %RELEASE_DIR%/ja-JP/PowerPointArrangeAddinInstaller.msi -out %RELEASE_DIR%/transforms/ja-JP.mst || goto :eof
torch -t language %RELEASE_DIR%/PowerPointArrangeAddinInstaller.msi %RELEASE_DIR%/zh-CN/PowerPointArrangeAddinInstaller.msi -out %RELEASE_DIR%/transforms/zh-CN.mst || goto :eof
torch -t language %RELEASE_DIR%/PowerPointArrangeAddinInstaller.msi %RELEASE_DIR%/zh-TW/PowerPointArrangeAddinInstaller.msi -out %RELEASE_DIR%/transforms/zh-TW.mst || goto :eof

WiSubStg.vbs %RELEASE_DIR%/PowerPointArrangeAddinInstaller.msi %RELEASE_DIR%/transforms/ja-JP.mst 1041 || goto :eof
WiSubStg.vbs %RELEASE_DIR%/PowerPointArrangeAddinInstaller.msi %RELEASE_DIR%/transforms/zh-CN.mst 2052 || goto :eof
WiSubStg.vbs %RELEASE_DIR%/PowerPointArrangeAddinInstaller.msi %RELEASE_DIR%/transforms/zh-TW.mst 1028 || goto :eof

WiLangId.vbs %RELEASE_DIR%/PowerPointArrangeAddinInstaller.msi Package 1033,1041,2052,1028 || goto :eof

@REM msiexec /i en-us\DIAViewSetup.msi TRANSFORMS=transforms\zh-tw.mst
@REM https://www.cnblogs.com/stoneniqiu/p/4725714.html
