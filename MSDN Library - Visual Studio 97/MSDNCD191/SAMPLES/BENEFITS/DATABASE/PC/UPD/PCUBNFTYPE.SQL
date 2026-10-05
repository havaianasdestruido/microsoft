GO
Set nocount on

IF EXISTS (SELECT * FROM sysobjects WHERE uid=user_id()
           AND      type='P'
           AND      name='pcUpd_BnfType')
    BEGIN
        PRINT '   > DROPPING - pcUpd_BnfType ...'
        DROP PROCEDURE pcUpd_BnfType
    END

GO

PRINT 'CREATING      - pcUpd_BnfType ...'
GO

CREATE PROCEDURE pcUpd_BnfType
/* File - 09.25.96 - c:\work\projects\benefit\db\PC\UPD\pcuBnfType.sql */

 @PKId        int
,@Descr       varchar(50)
,@Code        char(7)

AS

DECLARE @RetCode      int
DECLARE @_error       int
DECLARE @_rowcount    int

SELECT @RetCode = 0

BEGIN TRAN

UPDATE tbBnfType
    SET  Descr       = @Descr
        ,Code        = @Code
WHERE    PKId        = @PKId

SELECT @_error = @@error ,@_rowcount = @@rowcount

IF (@_error <> 0) OR (@_rowcount <> 1)
  BEGIN
    ROLLBACK TRAN
    IF (@_error = 0) SELECT @_error = 333
    SELECT  @RetCode = -(Abs(@_error))
    RAISERROR('Could not update Benefit Type Domain record [%d].', 16, 6, @PKId) 
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

