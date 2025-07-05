@echo off
cd %~dp0

if exist build\redmine-excel-sync.xlsm rmdir /s /q build\redmine-excel-sync.xlsm
mkdir build\redmine-excel-sync.xlsm

xcopy /y src\* build\redmine-excel-sync.xlsm\
xcopy /y lib\ariawase\src\Ariawase.xlsm\*.bas build\redmine-excel-sync.xlsm\
xcopy /y lib\ariawase\src\Ariawase.xlsm\*.cls build\redmine-excel-sync.xlsm\
xcopy /y lib\vba-json\*.bas build\redmine-excel-sync.xlsm\

cscript //nologo lib\ariawase\vbac.wsf combine /source ..\..\build /binary ..\..\dist
