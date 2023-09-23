@echo off

cd ./PowerPointArrangeAddin/

if "%VSCMD_VER%"=="" (
    call vsdevcmd.bat || goto :err
)

msbuild PowerPointArrangeAddin.csproj /p:Configuration=Release /p:Platform=x64 /t:Clean,Build || goto :err

cd ..
echo.
echo Done!
goto :eof

:err
cd ..
echo Failed to build!
exit /B 1
