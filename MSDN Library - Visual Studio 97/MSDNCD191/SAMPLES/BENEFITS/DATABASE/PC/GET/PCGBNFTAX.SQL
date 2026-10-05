GO
Set nocount on

IF EXISTS (SELECT * FROM sysobjects WHERE uid=user_id()
           AND      type='P'
           AND      name='pcGet_BnfTax')
    BEGIN
        PRINT '   > DROPPING - pcGet_BnfTax ...'
        DROP PROCEDURE pcGet_BnfTax
    END

GO

PRINT 'CREATING      - pcGet_BnfTax ...'
GO

CREATE PROCEDURE pcGet_BnfTax
/* File - 09.25.96 - c:\work\projects\benefit\db\PC\GET\pcgBnfTax.sql */

@PKId       int

AS

SELECT *
FROM   tbBnfTax
WHERE  PKId = @PKId

GO

