if exist build\redmine-excel-sync.xlsm rmdir /s /q build\redmine-excel-sync.xlsm

cscript //nologo lib\ariawase\vbac.wsf decombine /source ..\..\build /binary ..\..\dist

xcopy /y build\redmine-excel-sync.xlsm\* src\
