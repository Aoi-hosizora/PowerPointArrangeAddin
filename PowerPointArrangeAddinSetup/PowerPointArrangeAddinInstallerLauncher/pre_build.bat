@echo off

cd %~dp0/ || exit /B 1

cp ..\PowerPointArrangeAddinInstaller\bin\x64\Release\setup.msi .\Resources\setup.msi || exit /B 1
