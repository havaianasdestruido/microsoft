GO
Set nocount on

PRINT '   > DROP FKs - tbBank ...'
EXECUTE sp_dba_DropFKs 'tbBank'
GO

IF EXISTS (SELECT * FROM sysobjects WHERE uid=user_id()
           AND      type='U'
           AND      name='tbBank')
    BEGIN
        PRINT '   > DROPPING - tbBank ...'
        DROP TABLE tbBank
    END

GO

/*****************************************************************
Table Create for - c:\work\projects\benefit\db\TB\OBJS\tbBank.sql

WHEN        WHO         WHAT
--------  ----------  --------------------------------------------
mm.dd.yy  You?        (Next change?)
09.25.96  kenbe       Original.
*****************************************************************/

PRINT 'CREATING      - tbBank ...'
GO

CREATE TABLE tbBank
(

     /* File - 09.25.96 - c:\work\projects\benefit\db\TB\OBJS\tbBank.sql */

     PKId         int          Identity(1,1)  PRIMARY KEY  CLUSTERED
    ,Address1     varchar(60)  NULL    
    ,Address2     varchar(60)  NULL    
    ,City         varchar(50)  NULL    
    ,Code         varchar(16)  NOT NULL
    ,Name         varchar(80)  NOT NULL
)
GO

