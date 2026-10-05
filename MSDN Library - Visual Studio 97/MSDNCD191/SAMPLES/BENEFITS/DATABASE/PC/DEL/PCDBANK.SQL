GO
Set nocount on

IF EXISTS (SELECT * FROM sysobjects WHERE uid=user_id()
           AND      type='P'
           AND      name='pcDel_Bank')
    BEGIN
        PRINT '   > DROPPING - pcDel_Bank ...'
        DROP PROCEDURE pcDel_Bank
    END

GO

PRINT 'CREATING      - pcDel_Bank ...'
GO

CREATE PROCEDURE pcDel_Bank
/* File - 09.23.96 - c:\work\projects\benefit\db\PC\DEL\pcdBank.sql */

@PKId       int

AS

DECLARE @RetCode      int
DECLARE @_error       int
DECLARE @_rowcount    int

IF EXISTS ( SELECT PKId
            FROM   tbEmployee
            WHERE  BankId = @PKId)
  BEGIN
    SELECT  @RetCode = -555
    RAISERROR('Could not delete Bank record [%d]. Object is currently being referenced by one or more Employees.', 16, 5, @PKId)
    SELECT @RetCode
    RETURN @RetCode
  END

SELECT @RetCode = 0

BEGIN TRAN

DELETE FROM tbBank
WHERE  PKId = @PKId

SELECT @_error = @@error ,@_rowcount = @@rowcount

IF (@_error <> 0) OR (@_rowcount <> 1)
  BEGIN
    ROLLBACK TRAN
    IF (@_error = 0) SELECT @_error = 333
    SELECT  @RetCode = -(Abs(@_error))
    RAISERROR('Could not delete Bank record [%d].', 16, 5, @PKId)
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

