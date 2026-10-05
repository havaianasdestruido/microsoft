GO
Set nocount on

IF EXISTS (SELECT * FROM sysobjects WHERE uid=user_id()
           AND      type='P'
           AND      name='pcLvw_EmplBnfByBenefitId')
    BEGIN
        PRINT '   > DROPPING - pcLvw_EmplBnfByBenefitId ...'
        DROP PROCEDURE pcLvw_EmplBnfByBenefitId
    END

GO

PRINT 'CREATING      - pcLvw_EmplBnfByBenefitId ...'
GO

CREATE PROCEDURE pcLvw_EmplBnfByBenefitId
/* File - 09.25.96 - c:\work\projects\benefit\db\PC\LVW\pcvEmplBnfByBenefitId.sql */

@BenefitId  int

AS

SELECT EmployeeId 'Employee Id', PKId, BenefitId 'Benefit Id', TaxId 'Employee Benefit Tax', CoverageId 'Employee Benefit Coverage', Amount 'Employee Benefit Amount'
FROM   tbEmplBnf
WHERE  BenefitId = @BenefitId

GO

