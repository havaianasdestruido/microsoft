setlocal
set dbName=%1
isql -E /P /d %dbName% /n /i pcgDepType.sql
isql -E /P /d %dbName% /n /i pcgBank.sql
isql -E /P /d %dbName% /n /i pcgBenefit.sql
isql -E /P /d %dbName% /n /i pcgBnfType.sql
isql -E /P /d %dbName% /n /i pcgEmplBnfCoverage.sql
isql -E /P /d %dbName% /n /i pcgEmplBnfTax.sql
isql -E /P /d %dbName% /n /i pcgBnfTax.sql
isql -E /P /d %dbName% /n /i pcgEmployee.sql
isql -E /P /d %dbName% /n /i pcgDependant.sql
isql -E /P /d %dbName% /n /i pcgEmplBnf.sql
endlocal

