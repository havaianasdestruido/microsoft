GO
Set nocount on

IF EXISTS (SELECT * FROM sysobjects WHERE uid=user_id()
           AND      type='P'
           AND      name='pcGet_Benefit')
    BEGIN
        PRINT '   > DROPPING - pcGet_Benefit ...'
        DROP PROCEDURE pcGet_Benefit
    END

GO

PRINT 'CREATING      - pcGet_Benefit ...'
GO

CREATE PROCEDURE pcGet_Benefit
/* File - 09.25.96 - c:\work\projects\benefit\db\PC\GET\pcgBenefit.sql */

@PKId       int

AS

SELECT *
FROM   tbBenefit
WHERE  PKId = @PKId

GO

