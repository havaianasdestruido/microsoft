GO
Set nocount on

IF EXISTS (SELECT * FROM sysobjects WHERE uid=user_id()
           AND      type='P'
           AND      name='pcDel_DepType')
    BEGIN
        PRINT '   > DROPPING - pcDel_DepType ...'
        DROP PROCEDURE pcDel_DepType
    END

GO

PRINT 'CREATING      - pcDel_DepType ...'
GO

CREATE PROCEDURE pcDel_DepType
/* File - 09.25.96 - c:\work\projects\benefit\db\PC\DEL\pcdDepType.sql */

@PKId       int

AS

DECLARE @RetCode      int
DECLARE @_error       int
DECLARE @_rowcount    int

SELECT @RetCode = 0

BEGIN TRAN

DELETE FROM tbDepType
WHERE  PKId = @PKId

SELECT @_error = @@error ,@_rowcount = @@rowcount

IF (@_error <> 0) OR (@_rowcount <> 1)
  BEGIN
    ROLLBACK TRAN
    IF (@_error = 0) SELECT @_error = 333
    SELECT  @RetCode = -(Abs(@_error))
    RAISERROR('Could not delete Dependant Type Domain record [%d].', 16, 5, @PKId)
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

