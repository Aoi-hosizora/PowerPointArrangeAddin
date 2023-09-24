@echo off

if "%VSCMD_VER%"=="" (
    call vsdevcmd.bat || goto :err
)

msbuild ./PowerPointArrangeAddin/PowerPointArrangeAddin.csproj /p:Configuration=Release /p:Platform=x64 /t:Clean || goto :err

echo.
echo Done!
goto :eof

:err
echo Failed to Clean!
exit /B 1
