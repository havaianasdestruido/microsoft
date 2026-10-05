rem Corporate Benefits Sample
rem ==========================
rem by Ken Bergmann

@echo off
setlocal
set dbName=%1

rem ***************  execute the source files  ***********************

echo _) BUILDING ------)----------)-----------)------------) DBA Procedures
PUSHD pc\sys
call build.bat > build.log
POPD

echo _) BUILDING ------)----------)-----------)------------) Code tables
PUSHD tb\code\
call build.bat %dbName% > build.log
POPD

echo _) BUILDING ------)----------)-----------)------------) Object tables
PUSHD tb\objs\
call build.bat %dbName% > build.log
POPD

echo _) BUILDING ------)----------)-----------)------------) Completed!

endlocal
