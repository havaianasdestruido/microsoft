GO
Set nocount on

IF EXISTS (SELECT * FROM sysobjects WHERE uid=user_id()
           AND      type='P'
           AND      name='pcUpd_EmplBnf')
    BEGIN
        PRINT '   > DROPPING - pcUpd_EmplBnf ...'
        DROP PROCEDURE pcUpd_EmplBnf
    END

GO

PRINT 'CREATING      - pcUpd_EmplBnf ...'
GO

CREATE PROCEDURE pcUpd_EmplBnf
/* File - 09.25.96 - c:\work\projects\benefit\db\PC\UPD\pcuEmplBnf.sql */

 @PKId         int
,@BenefitId    int          
,@CoverageId   int          
,@EmployeeId   int          
,@TaxId        int          
,@Amount       money        

AS

DECLARE @RetCode      int
DECLARE @_error       int
DECLARE @_rowcount    int

SELECT @RetCode = 0

BEGIN TRAN

UPDATE tbEmplBnf
    SET
        BenefitId     = @BenefitId   
       ,CoverageId    = @CoverageId  
       ,EmployeeId    = @EmployeeId  
       ,TaxId         = @TaxId       
       ,Amount        = @Amount      
WHERE   PKId          = @PKId

SELECT @_error = @@error ,@_rowcount = @@rowcount

IF (@_error <> 0) OR (@_rowcount <> 1)
  BEGIN
    ROLLBACK TRAN
    IF (@_error = 0) SELECT @_error = 333
    SELECT  @RetCode = -(Abs(@_error))
    RAISERROR('Could not update Employee Benefit record [%d].', 16, 6, @PKId) 
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

