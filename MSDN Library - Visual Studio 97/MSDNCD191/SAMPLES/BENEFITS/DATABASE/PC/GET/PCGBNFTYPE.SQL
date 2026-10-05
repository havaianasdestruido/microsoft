GO
Set nocount on

IF EXISTS (SELECT * FROM sysobjects WHERE uid=user_id()
           AND      type='P'
           AND      name='pcGet_BnfType')
    BEGIN
        PRINT '   > DROPPING - pcGet_BnfType ...'
        DROP PROCEDURE pcGet_BnfType
    END

GO

PRINT 'CREATING      - pcGet_BnfType ...'
GO

CREATE PROCEDURE pcGet_BnfType
/* File - 09.25.96 - c:\work\projects\benefit\db\PC\GET\pcgBnfType.sql */

@PKId       int

AS

SELECT *
FROM   tbBnfType
WHERE  PKId = @PKId

GO

