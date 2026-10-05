GO
Set nocount on

IF EXISTS (SELECT * FROM sysobjects WHERE uid=user_id()
           AND      type='P'
           AND      name='pcRet_EmplBnfByTaxId')
    BEGIN
        PRINT '   > DROPPING - pcRet_EmplBnfByTaxId ...'
        DROP PROCEDURE pcRet_EmplBnfByTaxId
    END

GO

PRINT 'CREATING      - pcRet_EmplBnfByTaxId ...'
GO

CREATE PROCEDURE pcRet_EmplBnfByTaxId
/* File - 09.25.96 - c:\work\projects\benefit\db\PC\RET\pcrEmplBnfByTaxId.sql */

@TaxId  int

AS

SELECT *
FROM   tbEmplBnf
WHERE  TaxId = @TaxId

GO

