--Corporate Benefits Sample
--==========================
--by Ken Bergmann

Go
-- USE master
Go

Set nocount on
Declare @vc80z varchar (80)
Print   ' '
Select  @vc80z = '(Is  master  the current DB? ==> ' + Db_name() + ')'
Print   @vc80z


IF exists (select * from sysobjects where type='P ' and
            name =           'sp_dba_DropFKs')
   begin
   Print  'D-ropping SP       sp_dba_DropFKs ...'
   DROP Procedure             sp_dba_DropFKs
   end
GO

/***************************************************************
 This system sproc will drop 1 -to- all
 of the foreign key constraints within the current DB.

 This SP, like all SPs (not PCs) is most useful if
 created in the master DB.

 This SP assumes both parent & child TBs are in the
 current DB.


WHEN        WHO            WHAT
----------  -------------  -----------------------------------
1995/04/14  Ken Bergmann   Original code.
***************************************************************/

Print  'C-reating SP          sp_dba_DropFKs ...'
GO


CREATE Procedure              sp_dba_DropFKs

-- InfoTag: SPDropFK.SQL ,1995/04/14 18:21c ,sp_dba_DropFKs

-- The varchar (30) parms may be any LIKE patterns or exact names.
-- DOS wildcards *? can be used instead of %_, due to .BAT limits re %.

       @pParentTBName         varchar (30)         -- Too dangerous to default
      ,@pChildTBName          varchar (30)  = '%'
      ,@pFKConstraintName     varchar (30)  = '%'
      ,@pDiagnosticsYN        varchar ( 1)  = 'N'  -- All nonY = N = No
As

Set nocount on


Declare
       @vc80a              varchar ( 80)
      ,@vc255a             varchar (255)
      ,@LoopCount11        int
      ,@LoopCount21        int
      ,@TotalDrops         int
      ,@NewLine            char (2)
      ,@Pos                int
      ,@Datalength         int

Declare
       @ChildTBName        varchar (30)
      ,@FKConstraintName   varchar (30)
      ,@ParentTBName       varchar (30)
      ,@ParentTBId         int
      ,@charParentTBId     varchar (12)


Select
       @TotalDrops         = 0
      ,@NewLine            = Char(13) + Char(10)


--------------------------  edit input parms  ---------------------

---  For .BAT limits, swap all * for % (and all _ for ? to be DOS-ish)


Select @Datalength = Datalength(@pParentTBName)
IF (Substring(@pParentTBName,1,1) = '*')
   Select @pParentTBName = '%' + Substring(@pParentTBName ,2
                                          ,@Datalength-1)
WHILE (@Datalength > 0)
   Begin
   Select @Pos = CharIndex('*',@pParentTBName)
   IF (@Pos = 0) BREAK
   Select @pParentTBName = Substring(@pParentTBName,1,@Pos-1) + '%'
                         + Substring(@pParentTBName,@Pos+1
                                    ,@Datalength-@Pos)
   End

--------

Select @Datalength = Datalength(@pChildTBName)
IF (Substring(@pChildTBName,1,1) = '*')
   Select @pChildTBName = '%' + Substring(@pChildTBName ,2
                                          ,@Datalength-1)
WHILE (@Datalength > 0)
   Begin
   Select @Pos = CharIndex('*',@pChildTBName)
   IF (@Pos = 0) BREAK
   Select @pChildTBName = Substring(@pChildTBName,1,@Pos-1) + '%'
                        + Substring(@pChildTBName,@Pos+1
                                   ,@Datalength-@Pos)
   End

--------

Select @Datalength = Datalength(@pFKConstraintName)
IF (Substring(@pFKConstraintName,1,1) = '*')
   Select @pFKConstraintName = '%' + Substring(@pFKConstraintName ,2
                                          ,@Datalength-1)
WHILE (@Datalength > 0)
   Begin
   Select @Pos = CharIndex('*',@pFKConstraintName)
   IF (@Pos = 0) BREAK
   Select @pFKConstraintName = Substring(@pFKConstraintName,1,@Pos-1) + '%'
                             + Substring(@pFKConstraintName,@Pos+1
                                        ,@Datalength-@Pos)
   End


-------------------------  Above was */%, below is ?/_


Select @Datalength = Datalength(@pParentTBName)
IF (Substring(@pParentTBName,1,1) = '?')
   Select @pParentTBName = '_' + Substring(@pParentTBName ,2
                                          ,@Datalength-1)
WHILE (@Datalength > 0)
   Begin
   Select @Pos = CharIndex('?',@pParentTBName)
   IF (@Pos = 0) BREAK
   Select @pParentTBName = Substring(@pParentTBName,1,@Pos-1) + '_'
                         + Substring(@pParentTBName,@Pos+1
                                    ,@Datalength-@Pos)
   End

--------

Select @Datalength = Datalength(@pChildTBName)
IF (Substring(@pChildTBName,1,1) = '?')
   Select @pChildTBName = '_' + Substring(@pChildTBName ,2
                                          ,@Datalength-1)
WHILE (@Datalength > 0)
   Begin
   Select @Pos = CharIndex('?',@pChildTBName)
   IF (@Pos = 0) BREAK
   Select @pChildTBName = Substring(@pChildTBName,1,@Pos-1) + '_'
                        + Substring(@pChildTBName,@Pos+1
                                   ,@Datalength-@Pos)
   End

--------

Select @Datalength = Datalength(@pFKConstraintName)
IF (Substring(@pFKConstraintName,1,1) = '?')
   Select @pFKConstraintName = '_' + Substring(@pFKConstraintName ,2
                                          ,@Datalength-1)
WHILE (@Datalength > 0)
   Begin
   Select @Pos = CharIndex('?',@pFKConstraintName)
   IF (@Pos = 0) BREAK
   Select @pFKConstraintName = Substring(@pFKConstraintName,1,@Pos-1) + '_'
                             + Substring(@pFKConstraintName,@Pos+1
                                        ,@Datalength-@Pos)
   End


----------


IF (@pDiagnosticsYN Is Null)
   begin
   Select @pDiagnosticsYN  = 'N'
   end
ELSE
   begin
   Select @pDiagnosticsYN  = Upper(@pDiagnosticsYN)
   end


/* Print  ' '
Print  'sp_dba_DropFKs:  Here are the edited input parms...'

Select @vc255a = @pParentTBName     + '  = ParentTBName'      + @NewLine
               + @pChildTBName      + '  = ChildTBName'       + @NewLine
               + @pFKConstraintName + '  = FKConstraintName'  + @NewLine
               + @pDiagnosticsYN    + '  = DiagnosticsYN'
Print  @vc255a

*/
---------------  temp# TB for tracking each drop, for report  --------

SELECT
             "ParentTBName"      = @pParentTBName
            ,"ChildTBName"       = @pChildTBName
            ,"FKConstraintName"  = @pFKConstraintName
      Into
             #tb99Trail
      Where
             1=2



-------------------  MainLine  --------------------------

/*****
   The logic structure involves these loops nested:

      1 = Parent TBs
         2 = Child TB / Constraint FK pairs
*****/




EXECUTE (
"
DECLARE Cursor_11ParentTBs
   Insensitive Scroll
   Cursor For
"+"
SELECT
          ob1.name
         ,ob1.id
   From
          sysobjects    ob1
"+"
   Where
          ob1.name LIKE '" + @pParentTBName + "'
   And    ob1.type = 'U '
   Order By
          ob1.name
"
)


--------------------  Loop 11:  Thru ParentTBs  ------------------


OPEN Cursor_11ParentTBs


Select
       @LoopCount11  = 0


WHILE (11=11)
   Begin


   FETCH
      Next From Cursor_11ParentTBs Into
          @ParentTBName
         ,@ParentTBId


   IF (@@fetch_status <> 0)
      begin
      CLOSE      Cursor_11ParentTBs
      DEALLOCATE Cursor_11ParentTBs

      BREAK
      end


   Select @LoopCount11 = @LoopCount11 + 1



   Select @charParentTBId = Convert(varchar,@ParentTBId)


   EXECUTE (
   "
   DECLARE Cursor_21ChildTBs
      Insensitive Scroll
      Cursor For
   "+"
   SELECT
             ob1.name
            ,ob2.name
   "+"
      From
             sysobjects       ob1
            ,sysobjects       ob2
            ,sysreferences    rf1
   "+"
      Where
             rf1.fkeyid  = ob1.id
      And    rf1.constid = ob2.id
      And
             rf1.rkeyid = "     + @charParentTBId     + "
      And    ob1.name   LIKE '" + @pChildTBName       + "'
      And    ob2.name   LIKE '" + @pFKConstraintName  + "'
      Order By
             ob1.name
            ,ob2.name
   "
   )



-------------  Loop 21:  Thru ChildTB/Constraint pairs, for ParentTB. ---


   OPEN Cursor_21ChildTBs

   Select @LoopCount21  = 0


   WHILE (21=21)
      Begin


      FETCH
         Next From Cursor_21ChildTBs Into
             @ChildTBName
            ,@FKConstraintName


      IF (@@fetch_status <> 0)
         begin
         CLOSE      Cursor_21ChildTBs
         DEALLOCATE Cursor_21ChildTBs

         BREAK
         end


      Select @LoopCount21 = @LoopCount21 + 1


      Select @vc255a = 'ALTER Table '
                     + @ChildTBName
                     + @NewLine + '   '
                     + ' Drop Constraint '
                     + @FKConstraintName


      IF (@pDiagnosticsYN = 'Y')
         begin
         Print  ' '

         Select @vc80a = 'Re ParentTB=' + @ParentTBName
                       + ' will now dynamically execute this:'
         Print  @vc80a

         Print  @vc255a
         end


      EXECUTE (@vc255a)



      INSERT #tb99Trail
            ( ParentTBName , ChildTBName , FKConstraintName)
         Values
            (@ParentTBName ,@ChildTBName ,@FKConstraintName)


      End
-- loop 21



   Select @TotalDrops = @TotalDrops + @LoopCount21


   End
-- Loop 11



----------------------  Finalization  ----------------------

/*
Print  ' '

Select @vc255a = Convert(varchar,@TotalDrops)
               + '  = Total number of FKs dropped.'
Print  @vc255a


IF (@TotalDrops > 0)
   begin
   Print  ' '
   Print  'Here is a list of all the FK Constraints dropped:'

   SELECT       ParentTBName ,ChildTBName ,FKConstraintName
      From      #tb99Trail
      Order By  ParentTBName ,ChildTBName ,FKConstraintName
   end

*/
DROP Table      #tb99Trail


-- Print  ' '

Return @TotalDrops

Go



Print  'G-ranting sp_dba_DropFKs to p-ublic ...'

GRANT Execute on sp_dba_DropFKs to Public

Print  ' '
Go
-------



