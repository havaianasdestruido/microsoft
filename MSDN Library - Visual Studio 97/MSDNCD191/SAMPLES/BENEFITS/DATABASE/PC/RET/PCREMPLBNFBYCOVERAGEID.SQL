GO
Set nocount on

IF EXISTS (SELECT * FROM sysobjects WHERE uid=user_id()
           AND      type='P'
           AND      name='pcRet_EmplBnfByCoverageId')
    BEGIN
        PRINT '   > DROPPING - pcRet_EmplBnfByCoverageId ...'
        DROP PROCEDURE pcRet_EmplBnfByCoverageId
    END

GO

PRINT 'CREATING      - pcRet_EmplBnfByCoverageId ...'
GO

CREATE PROCEDURE pcRet_EmplBnfByCoverageId
/* File - 09.25.96 - c:\work\projects\benefit\db\PC\RET\pcrEmplBnfByCoverageId.sql */

@CoverageId  int

AS

SELECT *
FROM   tbEmplBnf
WHERE  CoverageId = @CoverageId

GO

