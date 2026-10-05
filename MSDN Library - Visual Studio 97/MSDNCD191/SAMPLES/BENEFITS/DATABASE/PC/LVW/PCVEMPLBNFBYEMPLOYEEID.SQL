GO
Set nocount on

IF EXISTS (SELECT * FROM sysobjects WHERE uid=user_id()
           AND      type='P'
           AND      name='pcLvw_EmplBnfByEmployeeId')
    BEGIN
        PRINT '   > DROPPING - pcLvw_EmplBnfByEmployeeId ...'
        DROP PROCEDURE pcLvw_EmplBnfByEmployeeId
    END

GO

PRINT 'CREATING      - pcLvw_EmplBnfByEmployeeId ...'
GO

CREATE PROCEDURE pcLvw_EmplBnfByEmployeeId
/* File - 09.25.96 - c:\work\projects\benefit\db\PC\LVW\pcvEmplBnfByEmployeeId.sql */

@EmployeeId  int

AS

SELECT EmployeeId 'Employee Id', PKId, BenefitId 'Benefit Id', TaxId 'Employee Benefit Tax', CoverageId 'Employee Benefit Coverage', Amount 'Employee Benefit Amount'
FROM   tbEmplBnf
WHERE  EmployeeId = @EmployeeId

GO

