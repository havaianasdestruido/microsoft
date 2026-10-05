GO
Set nocount on

IF EXISTS (SELECT * FROM sysobjects WHERE uid=user_id()
           AND      type='P'
           AND      name='pcRet_BenefitByBnfTypeId')
    BEGIN
        PRINT '   > DROPPING - pcRet_BenefitByBnfTypeId ...'
        DROP PROCEDURE pcRet_BenefitByBnfTypeId
    END

GO

PRINT 'CREATING      - pcRet_BenefitByBnfTypeId ...'
GO

CREATE PROCEDURE pcRet_BenefitByBnfTypeId
/* File - 09.25.96 - c:\work\projects\benefit\db\PC\RET\pcrBenefitByBnfTypeId.sql */

@BnfTypeId  int

AS

SELECT *
FROM   tbBenefit
WHERE  BnfTypeId = @BnfTypeId

GO

