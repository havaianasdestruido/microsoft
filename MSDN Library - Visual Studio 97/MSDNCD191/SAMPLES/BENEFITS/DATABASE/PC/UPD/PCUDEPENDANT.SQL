GO
Set nocount on

IF EXISTS (SELECT * FROM sysobjects WHERE uid=user_id()
           AND      type='P'
           AND      name='pcUpd_Dependant')
    BEGIN
        PRINT '   > DROPPING - pcUpd_Dependant ...'
        DROP PROCEDURE pcUpd_Dependant
    END

GO

PRINT 'CREATING      - pcUpd_Dependant ...'
GO

CREATE PROCEDURE pcUpd_Dependant
/* File - 09.25.96 - c:\work\projects\benefit\db\PC\UPD\pcuDependant.sql */

 @PKId         int
,@DepTypeId    int          
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

UPDATE tbDependant
    SET
        DepTypeId     = @DepTypeId   
       ,EmployeeId    = @EmployeeId  
       ,DateOfBirth   = @DateOfBirth 
       ,FirstName     = @FirstName   
       ,Gender        = @Gender      
       ,LastName      = @LastName    
       ,Status        = @Status      
WHERE   PKId          = @PKId

SELECT @_error = @@error ,@_rowcount = @@rowcount

IF (@_error <> 0) OR (@_rowcount <> 1)
  BEGIN
    ROLLBACK TRAN
    IF (@_error = 0) SELECT @_error = 333
    SELECT  @RetCode = -(Abs(@_error))
    RAISERROR('Could not update Dependant record [%d].', 16, 6, @PKId) 
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

