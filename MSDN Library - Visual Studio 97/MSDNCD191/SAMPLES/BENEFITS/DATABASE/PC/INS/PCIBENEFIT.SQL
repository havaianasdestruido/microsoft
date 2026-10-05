GO
Set nocount on

IF EXISTS (SELECT * FROM sysobjects WHERE uid=user_id()
           AND      type='P'
           AND      name='pcIns_Benefit')
    BEGIN
        PRINT '   > DROPPING - pcIns_Benefit ...'
        DROP PROCEDURE pcIns_Benefit
    END

GO

PRINT 'CREATING      - pcIns_Benefit ...'
GO

CREATE PROCEDURE pcIns_Benefit
/* File - 09.25.96 - c:\work\projects\benefit\db\PC\INS\pciBenefit.sql */

 @BnfTaxId     int          
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

INSERT INTO tbBenefit
(
 BnfTaxId     
,BnfTypeId    
,Cost         
,Description  
,Increment    
,Maximum      
,Minimum      
,Multiples    
)
VALUES
(
 @BnfTaxId    
,@BnfTypeId   
,@Cost        
,@Description 
,@Increment   
,@Maximum     
,@Minimum     
,@Multiples   
)

SELECT @_error = @@error ,@_rowcount = @@rowcount

IF (@_error <> 0) OR (@_rowcount <> 1)
  BEGIN
    ROLLBACK TRAN
    IF (@_error = 0) SELECT @_error = 333
    SELECT  @RetCode = -(Abs(@_error))
    RAISERROR('Could not insert Benefit record.',16,4) 
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

