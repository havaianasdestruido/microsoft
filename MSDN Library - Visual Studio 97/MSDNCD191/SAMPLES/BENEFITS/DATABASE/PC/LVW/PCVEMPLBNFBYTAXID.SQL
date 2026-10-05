GO
Set nocount on

IF EXISTS (SELECT * FROM sysobjects WHERE uid=user_id()
           AND      type='P'
           AND      name='pcLvw_EmplBnfByTaxId')
    BEGIN
        PRINT '   > DROPPING - pcLvw_EmplBnfByTaxId ...'
        DROP PROCEDURE pcLvw_EmplBnfByTaxId
    END

GO

PRINT 'CREATING      - pcLvw_EmplBnfByTaxId ...'
GO

CREATE PROCEDURE pcLvw_EmplBnfByTaxId
/* File - 09.25.96 - c:\work\projects\benefit\db\PC\LVW\pcvEmplBnfByTaxId.sql */

@TaxId  int

AS

SELECT EmployeeId 'Employee Id', PKId, BenefitId 'Benefit Id', TaxId 'Employee Benefit Tax', CoverageId 'Employee Benefit Coverage', Amount 'Employee Benefit Amount'
FROM   tbEmplBnf
WHERE  TaxId = @TaxId

GO

