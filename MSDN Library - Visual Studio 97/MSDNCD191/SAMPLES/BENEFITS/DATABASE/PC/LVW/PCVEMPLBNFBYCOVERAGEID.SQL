GO
Set nocount on

IF EXISTS (SELECT * FROM sysobjects WHERE uid=user_id()
           AND      type='P'
           AND      name='pcLvw_EmplBnfByCoverageId')
    BEGIN
        PRINT '   > DROPPING - pcLvw_EmplBnfByCoverageId ...'
        DROP PROCEDURE pcLvw_EmplBnfByCoverageId
    END

GO

PRINT 'CREATING      - pcLvw_EmplBnfByCoverageId ...'
GO

CREATE PROCEDURE pcLvw_EmplBnfByCoverageId
/* File - 09.25.96 - c:\work\projects\benefit\db\PC\LVW\pcvEmplBnfByCoverageId.sql */

@CoverageId  int

AS

SELECT EmployeeId 'Employee Id', PKId, BenefitId 'Benefit Id', TaxId 'Employee Benefit Tax', CoverageId 'Employee Benefit Coverage', Amount 'Employee Benefit Amount'
FROM   tbEmplBnf
WHERE  CoverageId = @CoverageId

GO

