--Corporate Benefits Sample
--==========================
--by Ken Bergmann

/**

declare @int1 int
exec @int1 = sp_dba_DropPK  'tbERAT_EventRating' ,'Y'
select 'returned value = ' ,@int1

**/

----+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Go
-- USE master
Go

Set nocount on
Declare @vc255z varchar (80)
Print  ' '
Select @vc255z = '(Is  master  the current DB?:  ' + Db_name() + ')'
Print  @vc255z


IF exists (select * from sysobjects where uid=user_id() and type='P ' and
            name =           'sp_dba_DropPK')
   begin
   Print  'D-ropping SP       sp_dba_DropPK ...'
   DROP Procedure             sp_dba_DropPK
   end
GO

/***************************************************************
 This SP, like all SP (not PC) is most useful if created in
 the master DB.

 This SP assumes both parent & child TBs are in the current DB.
***************************************************************/

Print  'C-reating SP          sp_dba_DropPK ...'
GO

CREATE Procedure              sp_dba_DropPK

-- InfoTag: SPDropPK.SQL ,1995/04/05 19:26b ,sp_dba_DropPK

       @pParentTBName         varchar (30)
      ,@pDiagnosticsYN        varchar (1)    = 'N'   -- All nonY = N
As

Set nocount on

Declare
       @vc255a             varchar (255)
      ,@RetCode            int

Declare
       @CSName             varchar (30)  -- fk, assume in home db


Select
       @CSName             = Null
      ,@RetCode            = 0  -- 0=zero PKs dropped



IF (@pDiagnosticsYN Is Null)
   begin
   Select @pDiagnosticsYN  = 'N'
   end
ELSE
   begin
   Select @pDiagnosticsYN  = Upper(@pDiagnosticsYN)
   end


--------------------  get PK constraint name  -------------


SELECT
             @CSName          = Object_name(con1.constid)
      From
             sysconstraints   con1
            ,sysobjects       obj2
            ,sysobjects       obj3
      Where
             con1.id          = Object_id(@pParentTBName)
      And    con1.id          = obj2.id
      And    con1.constid     = obj3.id
      And    obj2.type        = 'U '
      And    obj3.type        = 'K '



IF (@CSName Is Null)
   begin
   GOTO LABEL_86DONE
   end



Select @vc255a = 'ALTER Table '
               + @pParentTBName
               + ' Drop Constraint '
               + @CSName



IF (@pDiagnosticsYN = 'Y')
   begin
   Select 'Will now EXEC: ' ,@vc255a
   end


EXECUTE (@vc255a)


Select @RetCode = @RetCode + 1



LABEL_86DONE:


Return @RetCode

Go




Print  'G-ranting sp_dba_DropPK to p-ublic ...'

GRANT Execute on sp_dba_DropPK to Public

Print  ' '
Go
-------


