GO
Set nocount on

IF EXISTS (SELECT * FROM sysobjects WHERE uid=user_id()
           AND      type='P'
           AND      name='pcGet_EmplBnf')
    BEGIN
        PRINT '   > DROPPING - pcGet_EmplBnf ...'
        DROP PROCEDURE pcGet_EmplBnf
    END

GO

PRINT 'CREATING      - pcGet_EmplBnf ...'
GO

CREATE PROCEDURE pcGet_EmplBnf
/* File - 09.25.96 - c:\work\projects\benefit\db\PC\GET\pcgEmplBnf.sql */

@PKId       int

AS

SELECT *
FROM   tbEmplBnf
WHERE  PKId = @PKId

GO

