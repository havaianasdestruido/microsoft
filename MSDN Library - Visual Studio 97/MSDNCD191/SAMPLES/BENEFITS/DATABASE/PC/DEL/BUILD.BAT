setlocal
set dbName=%1
isql -E /P /d %dbName% /n /i pcdDepType.sql
isql -E /P /d %dbName% /n /i pcdBank.sql
isql -E /P /d %dbName% /n /i pcdBenefit.sql
isql -E /P /d %dbName% /n /i pcdBnfType.sql
isql -E /P /d %dbName% /n /i pcdEmplBnfCoverage.sql
isql -E /P /d %dbName% /n /i pcdEmplBnfTax.sql
isql -E /P /d %dbName% /n /i pcdBnfTax.sql
isql -E /P /d %dbName% /n /i pcdEmployee.sql
isql -E /P /d %dbName% /n /i pcdDependant.sql
isql -E /P /d %dbName% /n /i pcdEmplBnf.sql
endlocal

