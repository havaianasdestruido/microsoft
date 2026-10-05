GO
Set nocount on

IF EXISTS (SELECT * FROM sysobjects WHERE uid=user_id()
           AND      type='P'
           AND      name='pcRet_EmplBnfByBenefitId')
    BEGIN
        PRINT '   > DROPPING - pcRet_EmplBnfByBenefitId ...'
        DROP PROCEDURE pcRet_EmplBnfByBenefitId
    END

GO

PRINT 'CREATING      - pcRet_EmplBnfByBenefitId ...'
GO

CREATE PROCEDURE pcRet_EmplBnfByBenefitId
/* File - 09.25.96 - c:\work\projects\benefit\db\PC\RET\pcrEmplBnfByBenefitId.sql */

@BenefitId  int

AS

SELECT *
FROM   tbEmplBnf
WHERE  BenefitId = @BenefitId

GO

