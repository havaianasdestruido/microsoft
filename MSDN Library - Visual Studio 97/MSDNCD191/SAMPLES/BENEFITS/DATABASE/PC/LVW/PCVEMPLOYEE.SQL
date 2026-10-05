GO
Set nocount on

IF EXISTS (SELECT * FROM sysobjects WHERE uid=user_id()
           AND      type='P'
           AND      name='pcLvw_Employee')
    BEGIN
        PRINT '   > DROPPING - pcLvw_Employee ...'
        DROP PROCEDURE pcLvw_Employee
    END

GO

PRINT 'CREATING      - pcLvw_Employee ...'
GO

CREATE PROCEDURE pcLvw_Employee
/* File - 09.25.96 - c:\work\projects\benefit\db\PC\LVW\pcvEmployee.sql */

AS

SELECT LastName 'Last Name', PKId, FirstName 'First Name', UserName 'Employee User Name', Position 'Employee Position', Address1 'Address 1', Address2 'Address 2', City 'City Name', DateOfBirth 'Date Of Birth', StartDate 'Employee Start Date', EndDate 'Employee End Date'
FROM   tbEmployee

GO

