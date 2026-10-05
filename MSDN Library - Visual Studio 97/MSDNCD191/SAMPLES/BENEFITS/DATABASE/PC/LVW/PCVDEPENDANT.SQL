GO
Set nocount on

IF EXISTS (SELECT * FROM sysobjects WHERE uid=user_id()
           AND      type='P'
           AND      name='pcLvw_Dependant')
    BEGIN
        PRINT '   > DROPPING - pcLvw_Dependant ...'
        DROP PROCEDURE pcLvw_Dependant
    END

GO

PRINT 'CREATING      - pcLvw_Dependant ...'
GO

CREATE PROCEDURE pcLvw_Dependant
/* File - 09.25.96 - c:\work\projects\benefit\db\PC\LVW\pcvDependant.sql */

AS

SELECT LastName 'Last Name', PKId, FirstName 'First Name', DateOfBirth 'Date Of Birth'
FROM   tbDependant

GO

