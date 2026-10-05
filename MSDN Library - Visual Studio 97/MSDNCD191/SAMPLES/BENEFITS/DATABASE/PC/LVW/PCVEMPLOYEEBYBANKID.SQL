GO
Set nocount on

IF EXISTS (SELECT * FROM sysobjects WHERE uid=user_id()
           AND      type='P'
           AND      name='pcLvw_EmployeeByBankId')
    BEGIN
        PRINT '   > DROPPING - pcLvw_EmployeeByBankId ...'
        DROP PROCEDURE pcLvw_EmployeeByBankId
    END

GO

PRINT 'CREATING      - pcLvw_EmployeeByBankId ...'
GO

CREATE PROCEDURE pcLvw_EmployeeByBankId
/* File - 09.25.96 - c:\work\projects\benefit\db\PC\LVW\pcvEmployeeByBankId.sql */

@BankId  int

AS

SELECT LastName 'Last Name', PKId, FirstName 'First Name', UserName 'Employee User Name', Position 'Employee Position', Address1 'Address 1', Address2 'Address 2', City 'City Name', DateOfBirth 'Date Of Birth', StartDate 'Employee Start Date', EndDate 'Employee End Date'
FROM   tbEmployee
WHERE  BankId = @BankId

GO

