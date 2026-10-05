GO
Set nocount on

IF EXISTS (SELECT * FROM sysobjects WHERE uid=user_id()
           AND      type='P'
           AND      name='pcGet_Bank')
    BEGIN
        PRINT '   > DROPPING - pcGet_Bank ...'
        DROP PROCEDURE pcGet_Bank
    END

GO

PRINT 'CREATING      - pcGet_Bank ...'
GO

CREATE PROCEDURE pcGet_Bank
/* File - 09.25.96 - c:\work\projects\benefit\db\PC\GET\pcgBank.sql */

@PKId       int

AS

SELECT *
FROM   tbBank
WHERE  PKId = @PKId

GO

