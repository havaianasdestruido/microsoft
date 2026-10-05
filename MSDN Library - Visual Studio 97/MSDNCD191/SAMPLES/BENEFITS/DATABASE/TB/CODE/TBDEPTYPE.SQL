GO
Set nocount on

PRINT '   > DROP FKs - tbDepType ...'
EXECUTE sp_dba_DropFKs 'tbDepType'
GO

IF EXISTS (SELECT * FROM sysobjects WHERE uid=user_id()
           AND      type='U'
           AND      name='tbDepType')
    BEGIN
        PRINT '   > DROPPING - tbDepType ...'
        DROP TABLE tbDepType
    END

GO

/*****************************************************************
Table Create for - c:\work\projects\benefit\db\TB\CODE\tbDepType.sql

WHEN        WHO         WHAT
--------  ----------  --------------------------------------------
mm.dd.yy  You?        (Next change?)
09.25.96  kenbe       Original.
*****************************************************************/

PRINT 'CREATING      - tbDepType ...'
GO

CREATE TABLE tbDepType
(

     /* File - 09.25.96 - c:\work\projects\benefit\db\TB\CODE\tbDepType.sql */

     PKId       int       Identity(1,1)  PRIMARY KEY  CLUSTERED
    ,Descr      varchar(50)  NOT NULL
    ,Code       char(7)   NOT NULL
)
GO

