GO
Set nocount on

IF EXISTS (SELECT * FROM sysobjects WHERE uid=user_id()
           AND      type='P'
           AND      name='pcIns_Dependant')
    BEGIN
        PRINT '   > DROPPING - pcIns_Dependant ...'
        DROP PROCEDURE pcIns_Dependant
    END

GO

PRINT 'CREATING      - pcIns_Dependant ...'
GO

CREATE PROCEDURE pcIns_Dependant
/* File - 09.25.96 - c:\work\projects\benefit\db\PC\INS\pciDependant.sql */

 @DepTypeId    int          
,@EmployeeId   int          
,@DateOfBirth  datetime     
,@FirstName    varchar(35)  
,@Gender       bit          
,@LastName     varchar(35)  
,@Status       bit          

AS

DECLARE @RetCode      int
DECLARE @_error       int
DECLARE @_rowcount    int

SELECT @RetCode = 0

BEGIN TRAN

INSERT INTO tbDependant
(
 DepTypeId    
,EmployeeId   
,DateOfBirth  
,FirstName    
,Gender       
,LastName     
,Status       
)
VALUES
(
 @DepTypeId   
,@EmployeeId  
,@DateOfBirth 
,@FirstName   
,@Gender      
,@LastName    
,@Status      
)

SELECT @_error = @@error ,@_rowcount = @@rowcount

IF (@_error <> 0) OR (@_rowcount <> 1)
  BEGIN
    ROLLBACK TRAN
    IF (@_error = 0) SELECT @_error = 333
    SELECT  @RetCode = -(Abs(@_error))
    RAISERROR('Could not insert Dependant record.',16,4) 
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

