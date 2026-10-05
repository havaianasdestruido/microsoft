GO
Set nocount on

IF EXISTS (SELECT * FROM sysobjects WHERE uid=user_id()
           AND      type='P'
           AND      name='pcLvw_DepType')
    BEGIN
        PRINT '   > DROPPING - pcLvw_DepType ...'
        DROP PROCEDURE pcLvw_DepType
    END

GO

PRINT 'CREATING      - pcLvw_DepType ...'
GO

CREATE PROCEDURE pcLvw_DepType
/* File - 09.25.96 - c:\work\projects\benefit\db\PC\LVW\pcvDepType.sql */

AS

SELECT Descr Description, PKId, Code
FROM   tbDepType

GO

