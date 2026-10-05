setlocal
set dbName=%1
isql -E /P /d %dbName% /n /i pcrBenefitByBnfTaxId.sql
isql -E /P /d %dbName% /n /i pcrBenefitByBnfTypeId.sql
isql -E /P /d %dbName% /n /i pcrEmployeeByBankId.sql
isql -E /P /d %dbName% /n /i pcrDependantByEmployeeId.sql
isql -E /P /d %dbName% /n /i pcrDependantByDepTypeId.sql
isql -E /P /d %dbName% /n /i pcrEmplBnfByEmployeeId.sql
isql -E /P /d %dbName% /n /i pcrEmplBnfByBenefitId.sql
isql -E /P /d %dbName% /n /i pcrEmplBnfByTaxId.sql
isql -E /P /d %dbName% /n /i pcrEmplBnfByCoverageId.sql
endlocal
