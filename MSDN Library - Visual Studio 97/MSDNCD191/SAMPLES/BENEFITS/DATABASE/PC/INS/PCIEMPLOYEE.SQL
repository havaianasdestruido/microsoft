GO
Set nocount on

IF EXISTS (SELECT * FROM sysobjects WHERE uid=user_id()
           AND      type='P'
           AND      name='pcIns_Employee')
    BEGIN
        PRINT '   > DROPPING - pcIns_Employee ...'
        DROP PROCEDURE pcIns_Employee
    END

GO

PRINT 'CREATING      - pcIns_Employee ...'
GO

CREATE PROCEDURE pcIns_Employee
/* File - 09.25.96 - c:\work\projects\benefit\db\PC\INS\pciEmployee.sql */

 @BankId       int          
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

INSERT INTO tbEmployee
(
 BankId       
,Address1     
,Address2     
,BankAccount  
,City         
,DateOfBirth  
,EndDate      
,FirstName    
,Gender       
,LastName     
,Position     
,Salary       
,StartDate    
,UserName     
)
VALUES
(
 @BankId      
,@Address1    
,@Address2    
,@BankAccount 
,@City        
,@DateOfBirth 
,@EndDate     
,@FirstName   
,@Gender      
,@LastName    
,@Position    
,@Salary      
,@StartDate   
,@UserName    
)

SELECT @_error = @@error ,@_rowcount = @@rowcount

IF (@_error <> 0) OR (@_rowcount <> 1)
  BEGIN
    ROLLBACK TRAN
    IF (@_error = 0) SELECT @_error = 333
    SELECT  @RetCode = -(Abs(@_error))
    RAISERROR('Could not insert Employee record.',16,4) 
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

