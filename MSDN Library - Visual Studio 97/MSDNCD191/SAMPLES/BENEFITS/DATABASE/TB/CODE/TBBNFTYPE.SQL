GO
Set nocount on

PRINT '   > DROP FKs - tbBnfType ...'
EXECUTE sp_dba_DropFKs 'tbBnfType'
GO

IF EXISTS (SELECT * FROM sysobjects WHERE uid=user_id()
           AND      type='U'
           AND      name='tbBnfType')
    BEGIN
        PRINT '   > DROPPING - tbBnfType ...'
        DROP TABLE tbBnfType
    END

GO

/*****************************************************************
Table Create for - c:\work\projects\benefit\db\TB\CODE\tbBnfType.sql

WHEN        WHO         WHAT
--------  ----------  --------------------------------------------
mm.dd.yy  You?        (Next change?)
09.25.96  kenbe       Original.
*****************************************************************/

PRINT 'CREATING      - tbBnfType ...'
GO

CREATE TABLE tbBnfType
(

     /* File - 09.25.96 - c:\work\projects\benefit\db\TB\CODE\tbBnfType.sql */

     PKId       int       Identity(1,1)  PRIMARY KEY  CLUSTERED
    ,Descr      varchar(50)  NOT NULL
    ,Code       char(7)   NOT NULL
)
GO

