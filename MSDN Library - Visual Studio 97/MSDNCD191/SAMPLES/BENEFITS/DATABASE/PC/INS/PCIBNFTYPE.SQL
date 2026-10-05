GO
Set nocount on

IF EXISTS (SELECT * FROM sysobjects WHERE uid=user_id()
           AND      type='P'
           AND      name='pcIns_BnfType')
    BEGIN
        PRINT '   > DROPPING - pcIns_BnfType ...'
        DROP PROCEDURE pcIns_BnfType
    END

GO

PRINT 'CREATING      - pcIns_BnfType ...'
GO

CREATE PROCEDURE pcIns_BnfType
/* File - 09.25.96 - c:\work\projects\benefit\db\PC\INS\pciBnfType.sql */

 @Descr       varchar(50)
,@Code        char(7)

AS

DECLARE @RetCode      int
DECLARE @_error       int
DECLARE @_rowcount    int

SELECT @RetCode = 0

BEGIN TRAN

INSERT INTO tbBnfType
    (Descr, Code)
VALUES
    (@Descr, @Code)

SELECT @_error = @@error ,@_rowcount = @@rowcount

IF (@_error <> 0) OR (@_rowcount <> 1)
  BEGIN
    ROLLBACK TRAN
    IF (@_error = 0) SELECT @_error = 333
    SELECT  @RetCode = -(Abs(@_error))
    RAISERROR('Could not insert Benefit Type Domain record.',16,4) 
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

