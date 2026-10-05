GO
Set nocount on

IF EXISTS (SELECT * FROM sysobjects WHERE uid=user_id()
           AND      type='P'
           AND      name='pcUpd_EmplBnfTax')
    BEGIN
        PRINT '   > DROPPING - pcUpd_EmplBnfTax ...'
        DROP PROCEDURE pcUpd_EmplBnfTax
    END

GO

PRINT 'CREATING      - pcUpd_EmplBnfTax ...'
GO

CREATE PROCEDURE pcUpd_EmplBnfTax
/* File - 09.25.96 - c:\work\projects\benefit\db\PC\UPD\pcuEmplBnfTax.sql */

 @PKId        int
,@Descr       varchar(50)
,@Code        char(7)

AS

DECLARE @RetCode      int
DECLARE @_error       int
DECLARE @_rowcount    int

SELECT @RetCode = 0

BEGIN TRAN

UPDATE tbEmplBnfTax
    SET  Descr       = @Descr
        ,Code        = @Code
WHERE    PKId        = @PKId

SELECT @_error = @@error ,@_rowcount = @@rowcount

IF (@_error <> 0) OR (@_rowcount <> 1)
  BEGIN
    ROLLBACK TRAN
    IF (@_error = 0) SELECT @_error = 333
    SELECT  @RetCode = -(Abs(@_error))
    RAISERROR('Could not update Employee Benifit Tax Domain record [%d].', 16, 6, @PKId) 
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

