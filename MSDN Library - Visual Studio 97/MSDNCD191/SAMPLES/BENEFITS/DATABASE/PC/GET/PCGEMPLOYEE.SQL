GO
Set nocount on

IF EXISTS (SELECT * FROM sysobjects WHERE uid=user_id()
           AND      type='P'
           AND      name='pcGet_Employee')
    BEGIN
        PRINT '   > DROPPING - pcGet_Employee ...'
        DROP PROCEDURE pcGet_Employee
    END

GO

PRINT 'CREATING      - pcGet_Employee ...'
GO

CREATE PROCEDURE pcGet_Employee
/* File - 09.25.96 - c:\work\projects\benefit\db\PC\GET\pcgEmployee.sql */

@PKId       int

AS

SELECT *
FROM   tbEmployee
WHERE  PKId = @PKId

GO

