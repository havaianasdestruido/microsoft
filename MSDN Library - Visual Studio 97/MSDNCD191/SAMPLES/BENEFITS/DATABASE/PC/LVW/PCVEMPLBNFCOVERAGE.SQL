GO
Set nocount on

IF EXISTS (SELECT * FROM sysobjects WHERE uid=user_id()
           AND      type='P'
           AND      name='pcLvw_EmplBnfCoverage')
    BEGIN
        PRINT '   > DROPPING - pcLvw_EmplBnfCoverage ...'
        DROP PROCEDURE pcLvw_EmplBnfCoverage
    END

GO

PRINT 'CREATING      - pcLvw_EmplBnfCoverage ...'
GO

CREATE PROCEDURE pcLvw_EmplBnfCoverage
/* File - 09.25.96 - c:\work\projects\benefit\db\PC\LVW\pcvEmplBnfCoverage.sql */

AS

SELECT Descr Description, PKId, Code
FROM   tbEmplBnfCoverage

GO

