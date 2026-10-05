GO
Set nocount on

IF EXISTS (SELECT * FROM sysobjects WHERE uid=user_id()
           AND      type='P'
           AND      name='pcLvw_DependantByDepTypeId')
    BEGIN
        PRINT '   > DROPPING - pcLvw_DependantByDepTypeId ...'
        DROP PROCEDURE pcLvw_DependantByDepTypeId
    END

GO

PRINT 'CREATING      - pcLvw_DependantByDepTypeId ...'
GO

CREATE PROCEDURE pcLvw_DependantByDepTypeId
/* File - 09.25.96 - c:\work\projects\benefit\db\PC\LVW\pcvDependantByDepTypeId.sql */

@DepTypeId  int

AS

SELECT LastName 'Last Name', PKId, FirstName 'First Name', DateOfBirth 'Date Of Birth'
FROM   tbDependant
WHERE  DepTypeId = @DepTypeId

GO

