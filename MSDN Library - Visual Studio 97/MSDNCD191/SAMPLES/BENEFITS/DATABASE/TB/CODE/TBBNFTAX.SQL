GO
Set nocount on

PRINT '   > DROP FKs - tbBnfTax ...'
EXECUTE sp_dba_DropFKs 'tbBnfTax'
GO

IF EXISTS (SELECT * FROM sysobjects WHERE uid=user_id()
           AND      type='U'
           AND      name='tbBnfTax')
    BEGIN
        PRINT '   > DROPPING - tbBnfTax ...'
        DROP TABLE tbBnfTax
    END

GO

/*****************************************************************
Table Create for - c:\work\projects\benefit\db\TB\CODE\tbBnfTax.sql

WHEN        WHO         WHAT
--------  ----------  --------------------------------------------
mm.dd.yy  You?        (Next change?)
09.25.96  kenbe       Original.
*****************************************************************/

PRINT 'CREATING      - tbBnfTax ...'
GO

CREATE TABLE tbBnfTax
(

     /* File - 09.25.96 - c:\work\projects\benefit\db\TB\CODE\tbBnfTax.sql */

     PKId       int       Identity(1,1)  PRIMARY KEY  CLUSTERED
    ,Descr      varchar(50)  NOT NULL
    ,Code       char(7)   NOT NULL
)
GO

