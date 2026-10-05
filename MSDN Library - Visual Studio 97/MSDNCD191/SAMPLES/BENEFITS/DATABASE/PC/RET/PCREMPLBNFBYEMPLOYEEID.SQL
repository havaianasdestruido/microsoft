GO
Set nocount on

IF EXISTS (SELECT * FROM sysobjects WHERE uid=user_id()
           AND      type='P'
           AND      name='pcRet_EmplBnfByEmployeeId')
    BEGIN
        PRINT '   > DROPPING - pcRet_EmplBnfByEmployeeId ...'
        DROP PROCEDURE pcRet_EmplBnfByEmployeeId
    END

GO

PRINT 'CREATING      - pcRet_EmplBnfByEmployeeId ...'
GO

CREATE PROCEDURE pcRet_EmplBnfByEmployeeId
/* File - 09.25.96 - c:\work\projects\benefit\db\PC\RET\pcrEmplBnfByEmployeeId.sql */

@EmployeeId  int

AS

SELECT *
FROM   tbEmplBnf
WHERE  EmployeeId = @EmployeeId

GO

