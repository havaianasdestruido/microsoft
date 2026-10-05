setlocal
set dbName=%1
isql -E /P /d %dbName% /n /i pciDepType.sql
isql -E /P /d %dbName% /n /i pciBank.sql
isql -E /P /d %dbName% /n /i pciBenefit.sql
isql -E /P /d %dbName% /n /i pciBnfType.sql
isql -E /P /d %dbName% /n /i pciEmplBnfCoverage.sql
isql -E /P /d %dbName% /n /i pciEmplBnfTax.sql
isql -E /P /d %dbName% /n /i pciBnfTax.sql
isql -E /P /d %dbName% /n /i pciEmployee.sql
isql -E /P /d %dbName% /n /i pciDependant.sql
isql -E /P /d %dbName% /n /i pciEmplBnf.sql
endlocal

