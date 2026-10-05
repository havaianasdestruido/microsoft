GO
Set nocount on

IF EXISTS (SELECT * FROM sysobjects WHERE uid=user_id()
           AND      type='P'
           AND      name='pcLvw_Bank')
    BEGIN
        PRINT '   > DROPPING - pcLvw_Bank ...'
        DROP PROCEDURE pcLvw_Bank
    END

GO

PRINT 'CREATING      - pcLvw_Bank ...'
GO

CREATE PROCEDURE pcLvw_Bank
/* File - 09.25.96 - c:\work\projects\benefit\db\PC\LVW\pcvBank.sql */

AS

SELECT Name 'Bank Name', PKId, Code 'Bank Code', Address1 'Address Line 1', Address2 'Address Line 2', City 'City Name'
FROM   tbBank

GO

