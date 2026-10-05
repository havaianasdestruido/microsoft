GO
Set nocount on

IF EXISTS (SELECT * FROM sysobjects WHERE uid=user_id()
           AND      type='P'
           AND      name='pcGet_EmplBnfCoverage')
    BEGIN
        PRINT '   > DROPPING - pcGet_EmplBnfCoverage ...'
        DROP PROCEDURE pcGet_EmplBnfCoverage
    END

GO

PRINT 'CREATING      - pcGet_EmplBnfCoverage ...'
GO

CREATE PROCEDURE pcGet_EmplBnfCoverage
/* File - 09.25.96 - c:\work\projects\benefit\db\PC\GET\pcgEmplBnfCoverage.sql */

@PKId       int

AS

SELECT *
FROM   tbEmplBnfCoverage
WHERE  PKId = @PKId

GO

