GO
Set nocount on

IF EXISTS (SELECT * FROM sysobjects WHERE uid=user_id()
           AND      type='P'
           AND      name='pcUpd_Bank')
    BEGIN
        PRINT '   > DROPPING - pcUpd_Bank ...'
        DROP PROCEDURE pcUpd_Bank
    END

GO

PRINT 'CREATING      - pcUpd_Bank ...'
GO

CREATE PROCEDURE pcUpd_Bank
/* File - 09.25.96 - c:\work\projects\benefit\db\PC\UPD\pcuBank.sql */

 @PKId         int
,@Address1     varchar(60)  
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

UPDATE tbBank
    SET
        Address1      = @Address1    
       ,Address2      = @Address2    
       ,City          = @City        
       ,Code          = @Code        
       ,Name          = @Name        
WHERE   PKId          = @PKId

SELECT @_error = @@error ,@_rowcount = @@rowcount

IF (@_error <> 0) OR (@_rowcount <> 1)
  BEGIN
    ROLLBACK TRAN
    IF (@_error = 0) SELECT @_error = 333
    SELECT  @RetCode = -(Abs(@_error))
    RAISERROR('Could not update Bank record [%d].', 16, 6, @PKId) 
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

