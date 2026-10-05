GO
Set nocount on

IF EXISTS (SELECT * FROM sysobjects WHERE uid=user_id()
           AND      type='P'
           AND      name='pcDel_EmplBnf')
    BEGIN
        PRINT '   > DROPPING - pcDel_EmplBnf ...'
        DROP PROCEDURE pcDel_EmplBnf
    END

GO

PRINT 'CREATING      - pcDel_EmplBnf ...'
GO

CREATE PROCEDURE pcDel_EmplBnf
/* File - 09.25.96 - c:\work\projects\benefit\db\PC\DEL\pcdEmplBnf.sql */

@PKId       int

AS

DECLARE @RetCode      int
DECLARE @_error       int
DECLARE @_rowcount    int

SELECT @RetCode = 0

BEGIN TRAN

DELETE FROM tbEmplBnf
WHERE  PKId = @PKId

SELECT @_error = @@error ,@_rowcount = @@rowcount

IF (@_error <> 0) OR (@_rowcount <> 1)
  BEGIN
    ROLLBACK TRAN
    IF (@_error = 0) SELECT @_error = 333
    SELECT  @RetCode = -(Abs(@_error))
    RAISERROR('Could not delete Employee Benefit record [%d].', 16, 5, @PKId)
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

