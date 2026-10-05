rem Corporate Benefits Sample
rem ==========================
rem by Ken Bergmann

@echo off
setlocal
set dbName=%1

rem ***************  execute the source files  ***********************

echo _) BUILDING ------)----------)-----------)------------) Delete Procedures
PUSHD pc\del\
call build.bat %dbName%
POPD

echo _) BUILDING ------)----------)-----------)------------) Get Procedures
PUSHD pc\get\
call build.bat %dbName%
POPD

echo _) BUILDING ------)----------)-----------)------------) Insert Procedures
PUSHD pc\ins\
call build.bat %dbName%
POPD

echo _) BUILDING ------)----------)-----------)------------) List View Procedures
PUSHD pc\lvw\
call build.bat %dbName%
POPD

echo _) BUILDING ------)----------)-----------)------------) Retrieve Procedures
PUSHD pc\ret\
call build.bat %dbName%
POPD

echo _) BUILDING ------)----------)-----------)------------) Update Procedures
PUSHD pc\upd\
call build.bat %dbName%
POPD

echo _) BUILDING ------)----------)-----------)------------) Completed!

endlocal
