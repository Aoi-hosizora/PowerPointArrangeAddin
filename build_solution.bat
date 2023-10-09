@echo off

if "%VSCMD_VER%"=="" (
    call vsdevcmd.bat || goto :err
)

msbuild PowerPointArrangeAddin.sln /p:Configuration=Release /p:Platform=x64 /t:Rebuild || goto :err
call clean.bat || goto :err

goto :eof

:err
exit /B 1
