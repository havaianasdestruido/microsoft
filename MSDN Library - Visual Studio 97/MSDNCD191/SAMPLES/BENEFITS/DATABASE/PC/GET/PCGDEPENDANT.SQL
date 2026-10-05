GO
Set nocount on

IF EXISTS (SELECT * FROM sysobjects WHERE uid=user_id()
           AND      type='P'
           AND      name='pcGet_Dependant')
    BEGIN
        PRINT '   > DROPPING - pcGet_Dependant ...'
        DROP PROCEDURE pcGet_Dependant
    END

GO

PRINT 'CREATING      - pcGet_Dependant ...'
GO

CREATE PROCEDURE pcGet_Dependant
/* File - 09.25.96 - c:\work\projects\benefit\db\PC\GET\pcgDependant.sql */

@PKId       int

AS

SELECT *
FROM   tbDependant
WHERE  PKId = @PKId

GO

