GO
Set nocount on

IF EXISTS (SELECT * FROM sysobjects WHERE uid=user_id()
           AND      type='P'
           AND      name='pcLvw_BnfType')
    BEGIN
        PRINT '   > DROPPING - pcLvw_BnfType ...'
        DROP PROCEDURE pcLvw_BnfType
    END

GO

PRINT 'CREATING      - pcLvw_BnfType ...'
GO

CREATE PROCEDURE pcLvw_BnfType
/* File - 09.25.96 - c:\work\projects\benefit\db\PC\LVW\pcvBnfType.sql */

AS

SELECT Descr Description, PKId, Code
FROM   tbBnfType

GO

