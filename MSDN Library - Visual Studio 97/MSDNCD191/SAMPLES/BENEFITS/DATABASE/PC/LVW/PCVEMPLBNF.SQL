GO
Set nocount on

IF EXISTS (SELECT * FROM sysobjects WHERE uid=user_id()
           AND      type='P'
           AND      name='pcLvw_EmplBnf')
    BEGIN
        PRINT '   > DROPPING - pcLvw_EmplBnf ...'
        DROP PROCEDURE pcLvw_EmplBnf
    END

GO

PRINT 'CREATING      - pcLvw_EmplBnf ...'
GO

CREATE PROCEDURE pcLvw_EmplBnf
/* File - 09.25.96 - c:\work\projects\benefit\db\PC\LVW\pcvEmplBnf.sql */

AS

SELECT EmployeeId 'Employee Id', PKId, BenefitId 'Benefit Id', TaxId 'Employee Benefit Tax', CoverageId 'Employee Benefit Coverage', Amount 'Employee Benefit Amount'
FROM   tbEmplBnf

GO

