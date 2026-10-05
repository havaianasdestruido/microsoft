setlocal
set dbName=%1
isql -E /P /d %dbName% /n /i tbBank.sql
isql -E /P /d %dbName% /n /i tbBenefit.sql
isql -E /P /d %dbName% /n /i tbEmployee.sql
isql -E /P /d %dbName% /n /i tbDependant.sql
isql -E /P /d %dbName% /n /i tbEmplBnf.sql
endlocal

