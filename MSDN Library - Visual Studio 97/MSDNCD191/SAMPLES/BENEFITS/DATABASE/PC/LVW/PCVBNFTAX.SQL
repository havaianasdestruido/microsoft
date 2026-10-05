GO
Set nocount on

IF EXISTS (SELECT * FROM sysobjects WHERE uid=user_id()
           AND      type='P'
           AND      name='pcLvw_BnfTax')
    BEGIN
        PRINT '   > DROPPING - pcLvw_BnfTax ...'
        DROP PROCEDURE pcLvw_BnfTax
    END

GO

PRINT 'CREATING      - pcLvw_BnfTax ...'
GO

CREATE PROCEDURE pcLvw_BnfTax
/* File - 09.25.96 - c:\work\projects\benefit\db\PC\LVW\pcvBnfTax.sql */

AS

SELECT Descr Description, PKId, Code
FROM   tbBnfTax

GO

