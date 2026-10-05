setlocal
set dbName=%1
isql -E /P /d %dbName% /n /i pcuDepType.sql
isql -E /P /d %dbName% /n /i pcuBank.sql
isql -E /P /d %dbName% /n /i pcuBenefit.sql
isql -E /P /d %dbName% /n /i pcuBnfType.sql
isql -E /P /d %dbName% /n /i pcuEmplBnfCoverage.sql
isql -E /P /d %dbName% /n /i pcuEmplBnfTax.sql
isql -E /P /d %dbName% /n /i pcuBnfTax.sql
isql -E /P /d %dbName% /n /i pcuEmployee.sql
isql -E /P /d %dbName% /n /i pcuDependant.sql
isql -E /P /d %dbName% /n /i pcuEmplBnf.sql
endlocal

