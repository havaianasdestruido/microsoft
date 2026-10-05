GO
Set nocount on

IF EXISTS (SELECT * FROM sysobjects WHERE uid=user_id()
           AND      type='P'
           AND      name='pcLvw_Benefit')
    BEGIN
        PRINT '   > DROPPING - pcLvw_Benefit ...'
        DROP PROCEDURE pcLvw_Benefit
    END

GO

PRINT 'CREATING      - pcLvw_Benefit ...'
GO

CREATE PROCEDURE pcLvw_Benefit
/* File - 09.25.96 - c:\work\projects\benefit\db\PC\LVW\pcvBenefit.sql */

AS

SELECT Description 'Benefit Description', PKId, BnfTaxId 'Benefit Tax', Cost 'Benefit Cost', Minimum 'Benefit Minimum', Maximum 'Benefit Maximum', Increment 'Benefit Increment', Multiples 'Benefit Multiples'
FROM   tbBenefit

GO

