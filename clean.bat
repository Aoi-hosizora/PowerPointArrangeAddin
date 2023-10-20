@echo off

if "%VSCMD_VER%"=="" (
    call vsdevcmd.bat || goto :err
)

@REM D:\Program Files (x86)\Microsoft Visual Studio\2019\Enterprise\MSBuild\Microsoft\VisualStudio\v16.0\OfficeTools\Microsoft.VisualStudio.Tools.Office.targets
msbuild ./PowerPointArrangeAddin/PowerPointArrangeAddin.csproj /p:Configuration=Release /p:Platform=x64 /t:VSTOClean || goto :err

echo.
echo Done!
goto :eof

:err
echo.
echo Failed to Clean!
exit /B 1
