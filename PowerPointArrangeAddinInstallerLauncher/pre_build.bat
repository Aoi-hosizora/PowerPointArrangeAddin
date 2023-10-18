@echo off

cd %~dp0/ || exit /B 1

@REM cp ..\PowerPointArrangeAddinInstaller\bin\x64\Release\PowerPointArrangeAddinInstaller.msi .\Resources\PowerPointArrangeAddinInstaller.msi || exit /B 1
cp ..\PowerPointArrangeAddinInstaller\bin\x64\Release\en-US\PowerPointArrangeAddinInstaller.msi .\Resources\PowerPointArrangeAddinInstaller.msi || exit /B 1
