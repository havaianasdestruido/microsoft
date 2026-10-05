GO
Set nocount on

IF EXISTS (SELECT * FROM sysobjects WHERE uid=user_id()
           AND      type='P'
           AND      name='pcUpd_Benefit')
    BEGIN
        PRINT '   > DROPPING - pcUpd_Benefit ...'
        DROP PROCEDURE pcUpd_Benefit
    END

GO

PRINT 'CREATING      - pcUpd_Benefit ...'
GO

CREATE PROCEDURE pcUpd_Benefit
/* File - 09.25.96 - c:\work\projects\benefit\db\PC\UPD\pcuBenefit.sql */

 @PKId         int
,@BnfTaxId     int          
,@BnfTypeId    int          
,@Cost         int          
,@Description  varchar(80)  
,@Increment    smallint     
,@Maximum      int          
,@Minimum      int          
,@Multiples    smallint     

AS

DECLARE @RetCode      int
DECLARE @_error       int
DECLARE @_rowcount    int

SELECT @RetCode = 0

BEGIN TRAN

UPDATE tbBenefit
    SET
        BnfTaxId      = @BnfTaxId    
       ,BnfTypeId     = @BnfTypeId   
       ,Cost          = @Cost        
       ,Description   = @Description 
       ,Increment     = @Increment   
       ,Maximum       = @Maximum     
       ,Minimum       = @Minimum     
       ,Multiples     = @Multiples   
WHERE   PKId          = @PKId

SELECT @_error = @@error ,@_rowcount = @@rowcount

IF (@_error <> 0) OR (@_rowcount <> 1)
  BEGIN
    ROLLBACK TRAN
    IF (@_error = 0) SELECT @_error = 333
    SELECT  @RetCode = -(Abs(@_error))
    RAISERROR('Could not update Benefit record [%d].', 16, 6, @PKId) 
    SELECT @RetCode
    RETURN @RetCode
  END
ELSE
  BEGIN
    COMMIT TRAN
    SELECT @RetCode
    RETURN @RetCode
  END
GO

