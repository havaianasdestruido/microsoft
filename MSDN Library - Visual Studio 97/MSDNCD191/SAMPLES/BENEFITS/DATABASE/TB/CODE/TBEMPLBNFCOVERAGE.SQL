GO
Set nocount on

PRINT '   > DROP FKs - tbEmplBnfCoverage ...'
EXECUTE sp_dba_DropFKs 'tbEmplBnfCoverage'
GO

IF EXISTS (SELECT * FROM sysobjects WHERE uid=user_id()
           AND      type='U'
           AND      name='tbEmplBnfCoverage')
    BEGIN
        PRINT '   > DROPPING - tbEmplBnfCoverage ...'
        DROP TABLE tbEmplBnfCoverage
    END

GO

/*****************************************************************
Table Create for - c:\work\projects\benefit\db\TB\CODE\tbEmplBnfCoverage.sql

WHEN        WHO         WHAT
--------  ----------  --------------------------------------------
mm.dd.yy  You?        (Next change?)
09.25.96  kenbe       Original.
*****************************************************************/

PRINT 'CREATING      - tbEmplBnfCoverage ...'
GO

CREATE TABLE tbEmplBnfCoverage
(

     /* File - 09.25.96 - c:\work\projects\benefit\db\TB\CODE\tbEmplBnfCoverage.sql */

     PKId       int       Identity(1,1)  PRIMARY KEY  CLUSTERED
    ,Descr      varchar(50)  NOT NULL
    ,Code       char(7)   NOT NULL
)
GO

