GO
Set nocount on

IF EXISTS (SELECT * FROM sysobjects WHERE uid=user_id()
           AND      type='P'
           AND      name='pcLvw_BenefitByBnfTaxId')
    BEGIN
        PRINT '   > DROPPING - pcLvw_BenefitByBnfTaxId ...'
        DROP PROCEDURE pcLvw_BenefitByBnfTaxId
    END

GO

PRINT 'CREATING      - pcLvw_BenefitByBnfTaxId ...'
GO

CREATE PROCEDURE pcLvw_BenefitByBnfTaxId
/* File - 09.25.96 - c:\work\projects\benefit\db\PC\LVW\pcvBenefitByBnfTaxId.sql */

@BnfTaxId  int

AS

SELECT Description 'Benefit Description', PKId, BnfTaxId 'Benefit Tax', Cost 'Benefit Cost', Minimum 'Benefit Minimum', Maximum 'Benefit Maximum', Increment 'Benefit Increment', Multiples 'Benefit Multiples'
FROM   tbBenefit
WHERE  BnfTaxId = @BnfTaxId

GO

