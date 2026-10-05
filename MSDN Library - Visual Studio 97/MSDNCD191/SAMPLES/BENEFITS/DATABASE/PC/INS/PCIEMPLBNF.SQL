GO
Set nocount on

IF EXISTS (SELECT * FROM sysobjects WHERE uid=user_id()
           AND      type='P'
           AND      name='pcIns_EmplBnf')
    BEGIN
        PRINT '   > DROPPING - pcIns_EmplBnf ...'
        DROP PROCEDURE pcIns_EmplBnf
    END

GO

PRINT 'CREATING      - pcIns_EmplBnf ...'
GO

CREATE PROCEDURE pcIns_EmplBnf
/* File - 09.25.96 - c:\work\projects\benefit\db\PC\INS\pciEmplBnf.sql */

 @BenefitId    int          
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

INSERT INTO tbEmplBnf
(
 BenefitId    
,CoverageId   
,EmployeeId   
,TaxId        
,Amount       
)
VALUES
(
 @BenefitId   
,@CoverageId  
,@EmployeeId  
,@TaxId       
,@Amount      
)

SELECT @_error = @@error ,@_rowcount = @@rowcount

IF (@_error <> 0) OR (@_rowcount <> 1)
  BEGIN
    ROLLBACK TRAN
    IF (@_error = 0) SELECT @_error = 333
    SELECT  @RetCode = -(Abs(@_error))
    RAISERROR('Could not insert Employee Benefit record.',16,4) 
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

