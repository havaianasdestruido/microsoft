rem Corporate Benefits Sample
rem ==========================
rem by Ken Bergmann

@echo off
setlocal
set dbName=%1

rem ***************  loading the sample data  ***********************

echo _) LOADING ------)----------)-----------)------------) Sample Data

PUSHD dat
isql -E /P /d %dbName% /n /i loaddata.sql 
POPD

endlocal

