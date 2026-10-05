GO
Set nocount on

PRINT '   > DROP FKs - tbEmplBnfTax ...'
EXECUTE sp_dba_DropFKs 'tbEmplBnfTax'
GO

IF EXISTS (SELECT * FROM sysobjects WHERE uid=user_id()
           AND      type='U'
           AND      name='tbEmplBnfTax')
    BEGIN
        PRINT '   > DROPPING - tbEmplBnfTax ...'
        DROP TABLE tbEmplBnfTax
    END

GO

/*****************************************************************
Table Create for - c:\work\projects\benefit\db\TB\CODE\tbEmplBnfTax.sql

WHEN        WHO         WHAT
--------  ----------  --------------------------------------------
mm.dd.yy  You?        (Next change?)
09.25.96  kenbe       Original.
*****************************************************************/

PRINT 'CREATING      - tbEmplBnfTax ...'
GO

CREATE TABLE tbEmplBnfTax
(

     /* File - 09.25.96 - c:\work\projects\benefit\db\TB\CODE\tbEmplBnfTax.sql */

     PKId       int       Identity(1,1)  PRIMARY KEY  CLUSTERED
    ,Descr      varchar(50)  NOT NULL
    ,Code       char(7)   NOT NULL
)
GO

