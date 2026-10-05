GO
Set nocount on

IF EXISTS (SELECT * FROM sysobjects WHERE uid=user_id()
           AND      type='P'
           AND      name='pcRet_DependantByDepTypeId')
    BEGIN
        PRINT '   > DROPPING - pcRet_DependantByDepTypeId ...'
        DROP PROCEDURE pcRet_DependantByDepTypeId
    END

GO

PRINT 'CREATING      - pcRet_DependantByDepTypeId ...'
GO

CREATE PROCEDURE pcRet_DependantByDepTypeId
/* File - 09.25.96 - c:\work\projects\benefit\db\PC\RET\pcrDependantByDepTypeId.sql */

@DepTypeId  int

AS

SELECT *
FROM   tbDependant
WHERE  DepTypeId = @DepTypeId

GO

