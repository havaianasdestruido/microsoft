GO
Set nocount on

IF EXISTS (SELECT * FROM sysobjects WHERE uid=user_id()
           AND      type='P'
           AND      name='pcRet_DependantByEmployeeId')
    BEGIN
        PRINT '   > DROPPING - pcRet_DependantByEmployeeId ...'
        DROP PROCEDURE pcRet_DependantByEmployeeId
    END

GO

PRINT 'CREATING      - pcRet_DependantByEmployeeId ...'
GO

CREATE PROCEDURE pcRet_DependantByEmployeeId
/* File - 09.25.96 - c:\work\projects\benefit\db\PC\RET\pcrDependantByEmployeeId.sql */

@EmployeeId  int

AS

SELECT *
FROM   tbDependant
WHERE  EmployeeId = @EmployeeId

GO

