@echo off

if "%VSCMD_VER%"=="" (
    call vsdevcmd.bat || goto :err
)

msbuild ./PowerPointArrangeAddin/PowerPointArrangeAddin.csproj /p:Configuration=Release /p:Platform=x64 /t:Clean,Build || goto :err

echo.
echo Done!
goto :eof

:err
echo Failed to build!
exit /B 1
