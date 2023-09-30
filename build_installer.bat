@echo off

if "%VSCMD_VER%"=="" (
    call vsdevcmd.bat || goto :err
)

cd ./PowerPointArrangeAddinSetup/
mkdir Release_Temp
set VDPROJ_FILE=./PowerPointArrangeAddinSetup.vdproj
set MSI_FILE=./Release/PowerPointArrangeAddinSetup.msi
set SLN_FILE=../PowerPointArrangeAddin.sln

@REM English (1033)

devenv %SLN_FILE% /Project %VDPROJ_FILE% /Build Release || goto :err
cp %MSI_FILE% ./Release_Temp/PowerPointArrangeAddinSetup_en.msi || goto :err

@REM Simplified Chinese (2052)

sed -i 's/"LanguageId" = "3:1033"/"LanguageId" = "3:2052"/' %VDPROJ_FILE%
sed -i 's/"CodePage" = "3:1252"/"CodePage" = "3:936"/' %VDPROJ_FILE%
sed -i 's/"UILanguageId" = "3:1033"/"UILanguageId" = "3:2052"/' %VDPROJ_FILE%
sed -i 's/"LangId" = "3:1033"/"LangId" = "3:2052"/' %VDPROJ_FILE%

devenv %SLN_FILE% /Project %VDPROJ_FILE% /Build Release || goto :err
cp %MSI_FILE% ./Release_Temp/PowerPointArrangeAddinSetup_zh_hans.msi || goto :err

@REM Traditional Chinese (1028)

sed -i 's/"LanguageId" = "3:2052"/"LanguageId" = "3:1028"/' %VDPROJ_FILE%
sed -i 's/"CodePage" = "3:936"/"CodePage" = "3:950"/' %VDPROJ_FILE%
sed -i 's/"UILanguageId" = "3:2052"/"UILanguageId" = "3:1028"/' %VDPROJ_FILE%
sed -i 's/"LangId" = "3:2052"/"LangId" = "3:1028"/' %VDPROJ_FILE%

devenv %SLN_FILE% /Project %VDPROJ_FILE% /Build Release || goto :err
cp %MSI_FILE% ./Release_Temp/PowerPointArrangeAddinSetup_zh_hant.msi || goto :err

@REM Japanese (1041)

sed -i 's/"LanguageId" = "3:1028"/"LanguageId" = "3:1041"/' %VDPROJ_FILE%
sed -i 's/"CodePage" = "3:950"/"CodePage" = "3:932"/' %VDPROJ_FILE%
sed -i 's/"UILanguageId" = "3:1028"/"UILanguageId" = "3:1041"/' %VDPROJ_FILE%
sed -i 's/"LangId" = "3:1028"/"LangId" = "3:1041"/' %VDPROJ_FILE%

devenv %SLN_FILE% /Project %VDPROJ_FILE% /Build Release || goto :err
cp %MSI_FILE% ./Release_Temp/PowerPointArrangeAddinSetup_ja.msi || goto :err

@REM echo @echo off > ./Release_Temp/Uninstaller.bat
@REM echo for %%%%i in (*.msi) do set MSI_FILE=%%%%i >> ./Release_Temp/Uninstaller.bat
@REM echo msiexec /x %MSI_FILE% >> ./Release_Temp/Uninstaller.bat

@REM Collect files

rm ./Release/*
mv ./Release_Temp/* ./Release/
rm -rf ./Release_Temp/

@REM EOF and ERR

cd ..
call clean.bat
echo.
echo Done!
goto :eof

:err
cd ..
echo.
echo Failed to build!
exit /B 1
