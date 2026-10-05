GO
Set nocount on

IF EXISTS (SELECT * FROM sysobjects WHERE uid=user_id()
           AND      type='P'
           AND      name='pcIns_EmplBnfTax')
    BEGIN
        PRINT '   > DROPPING - pcIns_EmplBnfTax ...'
        DROP PROCEDURE pcIns_EmplBnfTax
    END

GO

PRINT 'CREATING      - pcIns_EmplBnfTax ...'
GO

CREATE PROCEDURE pcIns_EmplBnfTax
/* File - 09.25.96 - c:\work\projects\benefit\db\PC\INS\pciEmplBnfTax.sql */

 @Descr       varchar(50)
,@Code        char(7)

AS

DECLARE @RetCode      int
DECLARE @_error       int
DECLARE @_rowcount    int

SELECT @RetCode = 0

BEGIN TRAN

INSERT INTO tbEmplBnfTax
    (Descr, Code)
VALUES
    (@Descr, @Code)

SELECT @_error = @@error ,@_rowcount = @@rowcount

IF (@_error <> 0) OR (@_rowcount <> 1)
  BEGIN
    ROLLBACK TRAN
    IF (@_error = 0) SELECT @_error = 333
    SELECT  @RetCode = -(Abs(@_error))
    RAISERROR('Could not insert Employee Benifit Tax Domain record.',16,4) 
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

