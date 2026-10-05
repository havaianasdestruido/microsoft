GO
Set nocount on

IF EXISTS (SELECT * FROM sysobjects WHERE uid=user_id()
           AND      type='P'
           AND      name='pcLvw_EmplBnfTax')
    BEGIN
        PRINT '   > DROPPING - pcLvw_EmplBnfTax ...'
        DROP PROCEDURE pcLvw_EmplBnfTax
    END

GO

PRINT 'CREATING      - pcLvw_EmplBnfTax ...'
GO

CREATE PROCEDURE pcLvw_EmplBnfTax
/* File - 09.25.96 - c:\work\projects\benefit\db\PC\LVW\pcvEmplBnfTax.sql */

AS

SELECT Descr Description, PKId, Code
FROM   tbEmplBnfTax

GO

