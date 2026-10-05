GO
Set nocount on

IF EXISTS (SELECT * FROM sysobjects WHERE uid=user_id()
           AND      type='P'
           AND      name='pcIns_Bank')
    BEGIN
        PRINT '   > DROPPING - pcIns_Bank ...'
        DROP PROCEDURE pcIns_Bank
    END

GO

PRINT 'CREATING      - pcIns_Bank ...'
GO

CREATE PROCEDURE pcIns_Bank
/* File - 09.25.96 - c:\work\projects\benefit\db\PC\INS\pciBank.sql */

 @Address1     varchar(60)  
,@Address2     varchar(60)  
,@City         varchar(50)  
,@Code         varchar(16)  
,@Name         varchar(80)  

AS

DECLARE @RetCode      int
DECLARE @_error       int
DECLARE @_rowcount    int

SELECT @RetCode = 0

BEGIN TRAN

INSERT INTO tbBank
(
 Address1     
,Address2     
,City         
,Code         
,Name         
)
VALUES
(
 @Address1    
,@Address2    
,@City        
,@Code        
,@Name        
)

SELECT @_error = @@error ,@_rowcount = @@rowcount

IF (@_error <> 0) OR (@_rowcount <> 1)
  BEGIN
    ROLLBACK TRAN
    IF (@_error = 0) SELECT @_error = 333
    SELECT  @RetCode = -(Abs(@_error))
    RAISERROR('Could not insert Bank record.',16,4) 
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

