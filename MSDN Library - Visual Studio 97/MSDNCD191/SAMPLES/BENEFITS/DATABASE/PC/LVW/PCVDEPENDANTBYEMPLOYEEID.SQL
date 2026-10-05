GO
Set nocount on

IF EXISTS (SELECT * FROM sysobjects WHERE uid=user_id()
           AND      type='P'
           AND      name='pcLvw_DependantByEmployeeId')
    BEGIN
        PRINT '   > DROPPING - pcLvw_DependantByEmployeeId ...'
        DROP PROCEDURE pcLvw_DependantByEmployeeId
    END

GO

PRINT 'CREATING      - pcLvw_DependantByEmployeeId ...'
GO

CREATE PROCEDURE pcLvw_DependantByEmployeeId
/* File - 09.25.96 - c:\work\projects\benefit\db\PC\LVW\pcvDependantByEmployeeId.sql */

@EmployeeId  int

AS

SELECT LastName 'Last Name', PKId, FirstName 'First Name', DateOfBirth 'Date Of Birth'
FROM   tbDependant
WHERE  EmployeeId = @EmployeeId

GO

