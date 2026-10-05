GO
Set nocount on

IF EXISTS (SELECT * FROM sysobjects WHERE uid=user_id()
           AND      type='P'
           AND      name='pcRet_BenefitByBnfTaxId')
    BEGIN
        PRINT '   > DROPPING - pcRet_BenefitByBnfTaxId ...'
        DROP PROCEDURE pcRet_BenefitByBnfTaxId
    END

GO

PRINT 'CREATING      - pcRet_BenefitByBnfTaxId ...'
GO

CREATE PROCEDURE pcRet_BenefitByBnfTaxId
/* File - 09.25.96 - c:\work\projects\benefit\db\PC\RET\pcrBenefitByBnfTaxId.sql */

@BnfTaxId  int

AS

SELECT *
FROM   tbBenefit
WHERE  BnfTaxId = @BnfTaxId

GO

