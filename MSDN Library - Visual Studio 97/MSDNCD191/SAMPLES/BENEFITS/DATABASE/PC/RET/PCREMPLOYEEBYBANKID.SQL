GO
Set nocount on

IF EXISTS (SELECT * FROM sysobjects WHERE uid=user_id()
           AND      type='P'
           AND      name='pcRet_EmployeeByBankId')
    BEGIN
        PRINT '   > DROPPING - pcRet_EmployeeByBankId ...'
        DROP PROCEDURE pcRet_EmployeeByBankId
    END

GO

PRINT 'CREATING      - pcRet_EmployeeByBankId ...'
GO

CREATE PROCEDURE pcRet_EmployeeByBankId
/* File - 09.25.96 - c:\work\projects\benefit\db\PC\RET\pcrEmployeeByBankId.sql */

@BankId  int

AS

SELECT *
FROM   tbEmployee
WHERE  BankId = @BankId

GO

