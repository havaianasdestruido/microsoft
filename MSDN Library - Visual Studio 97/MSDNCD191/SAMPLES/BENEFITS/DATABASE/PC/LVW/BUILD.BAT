setlocal
set dbName=%1
isql -E /P /d %dbName% /n /i pcvDepType.sql
isql -E /P /d %dbName% /n /i pcvBank.sql
isql -E /P /d %dbName% /n /i pcvBenefit.sql
isql -E /P /d %dbName% /n /i pcvBenefitByBnfTaxId.sql
isql -E /P /d %dbName% /n /i pcvBenefitByBnfTypeId.sql
isql -E /P /d %dbName% /n /i pcvBnfType.sql
isql -E /P /d %dbName% /n /i pcvEmplBnfCoverage.sql
isql -E /P /d %dbName% /n /i pcvEmplBnfTax.sql
isql -E /P /d %dbName% /n /i pcvBnfTax.sql
isql -E /P /d %dbName% /n /i pcvEmployee.sql
isql -E /P /d %dbName% /n /i pcvEmployeeByBankId.sql
isql -E /P /d %dbName% /n /i pcvDependant.sql
isql -E /P /d %dbName% /n /i pcvDependantByDepTypeId.sql
isql -E /P /d %dbName% /n /i pcvDependantByEmployeeId.sql
isql -E /P /d %dbName% /n /i pcvEmplBnf.sql
isql -E /P /d %dbName% /n /i pcvEmplBnfByBenefitId.sql
isql -E /P /d %dbName% /n /i pcvEmplBnfByCoverageId.sql
isql -E /P /d %dbName% /n /i pcvEmplBnfByEmployeeId.sql
isql -E /P /d %dbName% /n /i pcvEmplBnfByTaxId.sql
endlocal

