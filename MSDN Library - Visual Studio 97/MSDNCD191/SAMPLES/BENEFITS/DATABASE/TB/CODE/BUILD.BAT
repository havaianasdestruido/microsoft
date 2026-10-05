setlocal
set dbName=%1
isql -E /P /d %dbName% /n /i tbDepType.sql
isql -E /P /d %dbName% /n /i tbBnfType.sql
isql -E /P /d %dbName% /n /i tbEmplBnfCoverage.sql
isql -E /P /d %dbName% /n /i tbEmplBnfTax.sql
isql -E /P /d %dbName% /n /i tbBnfTax.sql
endlocal

