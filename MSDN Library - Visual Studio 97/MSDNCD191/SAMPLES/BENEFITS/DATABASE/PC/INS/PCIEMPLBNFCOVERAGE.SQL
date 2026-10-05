GO
Set nocount on

IF EXISTS (SELECT * FROM sysobjects WHERE uid=user_id()
           AND      type='P'
           AND      name='pcIns_EmplBnfCoverage')
    BEGIN
        PRINT '   > DROPPING - pcIns_EmplBnfCoverage ...'
        DROP PROCEDURE pcIns_EmplBnfCoverage
    END

GO

PRINT 'CREATING      - pcIns_EmplBnfCoverage ...'
GO

CREATE PROCEDURE pcIns_EmplBnfCoverage
/* File - 09.25.96 - c:\work\projects\benefit\db\PC\INS\pciEmplBnfCoverage.sql */

 @Descr       varchar(50)
,@Code        char(7)

AS

DECLARE @RetCode      int
DECLARE @_error       int
DECLARE @_rowcount    int

SELECT @RetCode = 0

BEGIN TRAN

INSERT INTO tbEmplBnfCoverage
    (Descr, Code)
VALUES
    (@Descr, @Code)

SELECT @_error = @@error ,@_rowcount = @@rowcount

IF (@_error <> 0) OR (@_rowcount <> 1)
  BEGIN
    ROLLBACK TRAN
    IF (@_error = 0) SELECT @_error = 333
    SELECT  @RetCode = -(Abs(@_error))
    RAISERROR('Could not insert Employee Benefit Coverage Domain record.',16,4) 
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

