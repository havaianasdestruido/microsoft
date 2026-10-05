GO
Set nocount on

IF EXISTS (SELECT * FROM sysobjects WHERE uid=user_id()
           AND      type='P'
           AND      name='pcGet_EmplBnfTax')
    BEGIN
        PRINT '   > DROPPING - pcGet_EmplBnfTax ...'
        DROP PROCEDURE pcGet_EmplBnfTax
    END

GO

PRINT 'CREATING      - pcGet_EmplBnfTax ...'
GO

CREATE PROCEDURE pcGet_EmplBnfTax
/* File - 09.25.96 - c:\work\projects\benefit\db\PC\GET\pcgEmplBnfTax.sql */

@PKId       int

AS

SELECT *
FROM   tbEmplBnfTax
WHERE  PKId = @PKId

GO

