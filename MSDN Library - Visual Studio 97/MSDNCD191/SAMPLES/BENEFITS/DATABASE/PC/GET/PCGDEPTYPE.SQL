GO
Set nocount on

IF EXISTS (SELECT * FROM sysobjects WHERE uid=user_id()
           AND      type='P'
           AND      name='pcGet_DepType')
    BEGIN
        PRINT '   > DROPPING - pcGet_DepType ...'
        DROP PROCEDURE pcGet_DepType
    END

GO

PRINT 'CREATING      - pcGet_DepType ...'
GO

CREATE PROCEDURE pcGet_DepType
/* File - 09.25.96 - c:\work\projects\benefit\db\PC\GET\pcgDepType.sql */

@PKId       int

AS

SELECT *
FROM   tbDepType
WHERE  PKId = @PKId

GO

