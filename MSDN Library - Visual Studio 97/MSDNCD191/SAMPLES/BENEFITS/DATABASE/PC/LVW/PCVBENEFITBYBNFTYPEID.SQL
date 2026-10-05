GO
Set nocount on

IF EXISTS (SELECT * FROM sysobjects WHERE uid=user_id()
           AND      type='P'
           AND      name='pcLvw_BenefitByBnfTypeId')
    BEGIN
        PRINT '   > DROPPING - pcLvw_BenefitByBnfTypeId ...'
        DROP PROCEDURE pcLvw_BenefitByBnfTypeId
    END

GO

PRINT 'CREATING      - pcLvw_BenefitByBnfTypeId ...'
GO

CREATE PROCEDURE pcLvw_BenefitByBnfTypeId
/* File - 09.25.96 - c:\work\projects\benefit\db\PC\LVW\pcvBenefitByBnfTypeId.sql */

@BnfTypeId  int

AS

SELECT Description 'Benefit Description', PKId, BnfTaxId 'Benefit Tax', Cost 'Benefit Cost', Minimum 'Benefit Minimum', Maximum 'Benefit Maximum', Increment 'Benefit Increment', Multiples 'Benefit Multiples'
FROM   tbBenefit
WHERE  BnfTypeId = @BnfTypeId

GO

