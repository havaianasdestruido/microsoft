GO
Set nocount on

IF EXISTS (SELECT * FROM sysobjects WHERE uid=user_id()
           AND      type='P'
           AND      name='pcDel_EmplBnfTax')
    BEGIN
        PRINT '   > DROPPING - pcDel_EmplBnfTax ...'
        DROP PROCEDURE pcDel_EmplBnfTax
    END

GO

PRINT 'CREATING      - pcDel_EmplBnfTax ...'
GO

CREATE PROCEDURE pcDel_EmplBnfTax
/* File - 09.25.96 - c:\work\projects\benefit\db\PC\DEL\pcdEmplBnfTax.sql */

@PKId       int

AS

DECLARE @RetCode      int
DECLARE @_error       int
DECLARE @_rowcount    int

SELECT @RetCode = 0

BEGIN TRAN

DELETE FROM tbEmplBnfTax
WHERE  PKId = @PKId

SELECT @_error = @@error ,@_rowcount = @@rowcount

IF (@_error <> 0) OR (@_rowcount <> 1)
  BEGIN
    ROLLBACK TRAN
    IF (@_error = 0) SELECT @_error = 333
    SELECT  @RetCode = -(Abs(@_error))
    RAISERROR('Could not delete Employee Benifit Tax Domain record [%d].', 16, 5, @PKId)
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

