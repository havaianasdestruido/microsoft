GO
Set nocount on

IF EXISTS (SELECT * FROM sysobjects WHERE uid=user_id()
           AND      type='P'
           AND      name='pcDel_Employee')
    BEGIN
        PRINT '   > DROPPING - pcDel_Employee ...'
        DROP PROCEDURE pcDel_Employee
    END

GO

PRINT 'CREATING      - pcDel_Employee ...'
GO

CREATE PROCEDURE pcDel_Employee
/* File - 09.23.96 - c:\work\projects\benefit\db\PC\DEL\pcdEmployee.sql */

@PKId       int

AS

DECLARE @RetCode      int
DECLARE @_error       int
DECLARE @_rowcount    int

SELECT @RetCode = 0

BEGIN TRAN

--First remove the dependants
DELETE FROM tbDependant
WHERE  EmployeeId = @PKId

SELECT @_error = @@error ,@_rowcount = @@rowcount

IF (@_error <> 0)
  BEGIN
    ROLLBACK TRAN
    SELECT  @RetCode = -(Abs(@_error))
    RAISERROR('Could not delete Employees Dependant records [%d].', 16, 5, @PKId)
    SELECT @RetCode
    RETURN @RetCode
  END
ELSE
  BEGIN
    --Next remove the benefits
    DELETE FROM tbEmplBnf
    WHERE  EmployeeId = @PKId

    SELECT @_error = @@error ,@_rowcount = @@rowcount
    IF (@_error <> 0)
      BEGIN
        ROLLBACK TRAN
        SELECT  @RetCode = -(Abs(@_error))
        RAISERROR('Could not delete Employees Benefit records [%d].', 16, 5, @PKId)
        SELECT @RetCode
        RETURN @RetCode
      END
    ELSE
      BEGIN
        --Then remove the actual employee
        DELETE FROM tbEmployee
        WHERE  PKId = @PKId
        
        SELECT @_error = @@error ,@_rowcount = @@rowcount
        IF (@_error <> 0) OR (@_rowcount <> 1)
          BEGIN
            ROLLBACK TRAN
            IF (@_error = 0) SELECT @_error = 333
            SELECT  @RetCode = -(Abs(@_error))
            RAISERROR('Could not delete Employee record [%d].', 16, 5, @PKId)
            SELECT @RetCode
            RETURN @RetCode
          END
        ELSE
          BEGIN
            COMMIT TRAN
            SELECT @RetCode
            RETURN @RetCode
         END
      END
  END

GO

