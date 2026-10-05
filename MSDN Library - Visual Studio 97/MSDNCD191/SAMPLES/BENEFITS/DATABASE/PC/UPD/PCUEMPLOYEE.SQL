GO
Set nocount on

IF EXISTS (SELECT * FROM sysobjects WHERE uid=user_id()
           AND      type='P'
           AND      name='pcUpd_Employee')
    BEGIN
        PRINT '   > DROPPING - pcUpd_Employee ...'
        DROP PROCEDURE pcUpd_Employee
    END

GO

PRINT 'CREATING      - pcUpd_Employee ...'
GO

CREATE PROCEDURE pcUpd_Employee
/* File - 09.25.96 - c:\work\projects\benefit\db\PC\UPD\pcuEmployee.sql */

 @PKId         int
,@BankId       int          
,@Address1     varchar(60)  
,@Address2     varchar(60)  
,@BankAccount  varchar(35)  
,@City         varchar(50)  
,@DateOfBirth  datetime     
,@EndDate      datetime     
,@FirstName    varchar(35)  
,@Gender       bit          
,@LastName     varchar(35)  
,@Position     varchar(50)  
,@Salary       money        
,@StartDate    datetime     
,@UserName     varchar(16)  

AS

DECLARE @RetCode      int
DECLARE @_error       int
DECLARE @_rowcount    int

SELECT @RetCode = 0

BEGIN TRAN

UPDATE tbEmployee
    SET
        BankId        = @BankId      
       ,Address1      = @Address1    
       ,Address2      = @Address2    
       ,BankAccount   = @BankAccount 
       ,City          = @City        
       ,DateOfBirth   = @DateOfBirth 
       ,EndDate       = @EndDate     
       ,FirstName     = @FirstName   
       ,Gender        = @Gender      
       ,LastName      = @LastName    
       ,Position      = @Position    
       ,Salary        = @Salary      
       ,StartDate     = @StartDate   
       ,UserName      = @UserName    
WHERE   PKId          = @PKId

SELECT @_error = @@error ,@_rowcount = @@rowcount

IF (@_error <> 0) OR (@_rowcount <> 1)
  BEGIN
    ROLLBACK TRAN
    IF (@_error = 0) SELECT @_error = 333
    SELECT  @RetCode = -(Abs(@_error))
    RAISERROR('Could not update Employee record [%d].', 16, 6, @PKId) 
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

