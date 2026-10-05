GO
Set nocount on

IF EXISTS (SELECT * FROM sysobjects WHERE uid=user_id()
           AND      type='P'
           AND      name='pcDel_Benefit')
    BEGIN
        PRINT '   > DROPPING - pcDel_Benefit ...'
        DROP PROCEDURE pcDel_Benefit
    END

GO

PRINT 'CREATING      - pcDel_Benefit ...'
GO

CREATE PROCEDURE pcDel_Benefit
/* File - 09.23.96 - c:\work\projects\benefit\db\PC\DEL\pcdBenefit.sql */

@PKId       int

AS

DECLARE @RetCode      int
DECLARE @_error       int
DECLARE @_rowcount    int

IF EXISTS ( SELECT PKId
            FROM   tbEmplBnf
            WHERE  BenefitId = @PKId)
  BEGIN
    SELECT  @RetCode = -555
    RAISERROR('Could not delete Benefit record [%d]. Object is currently being referenced by one or more Employees.', 16, 5, @PKId)
    SELECT @RetCode
    RETURN @RetCode
  END

SELECT @RetCode = 0

BEGIN TRAN

DELETE FROM tbBenefit
WHERE  PKId = @PKId

SELECT @_error = @@error ,@_rowcount = @@rowcount

IF (@_error <> 0) OR (@_rowcount <> 1)
  BEGIN
    ROLLBACK TRAN
    IF (@_error = 0) SELECT @_error = 333
    SELECT  @RetCode = -(Abs(@_error))
    RAISERROR('Could not delete Benefit record [%d].', 16, 5, @PKId)
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

