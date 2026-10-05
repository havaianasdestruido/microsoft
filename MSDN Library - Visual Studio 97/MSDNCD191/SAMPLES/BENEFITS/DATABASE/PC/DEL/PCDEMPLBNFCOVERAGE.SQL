GO
Set nocount on

IF EXISTS (SELECT * FROM sysobjects WHERE uid=user_id()
           AND      type='P'
           AND      name='pcDel_EmplBnfCoverage')
    BEGIN
        PRINT '   > DROPPING - pcDel_EmplBnfCoverage ...'
        DROP PROCEDURE pcDel_EmplBnfCoverage
    END

GO

PRINT 'CREATING      - pcDel_EmplBnfCoverage ...'
GO

CREATE PROCEDURE pcDel_EmplBnfCoverage
/* File - 09.25.96 - c:\work\projects\benefit\db\PC\DEL\pcdEmplBnfCoverage.sql */

@PKId       int

AS

DECLARE @RetCode      int
DECLARE @_error       int
DECLARE @_rowcount    int

SELECT @RetCode = 0

BEGIN TRAN

DELETE FROM tbEmplBnfCoverage
WHERE  PKId = @PKId

SELECT @_error = @@error ,@_rowcount = @@rowcount

IF (@_error <> 0) OR (@_rowcount <> 1)
  BEGIN
    ROLLBACK TRAN
    IF (@_error = 0) SELECT @_error = 333
    SELECT  @RetCode = -(Abs(@_error))
    RAISERROR('Could not delete Employee Benefit Coverage Domain record [%d].', 16, 5, @PKId)
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

