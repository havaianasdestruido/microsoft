GO
Set nocount on

PRINT '   > DROP FKs - tbEmplBnf ...'
EXECUTE sp_dba_DropFKs 'tbEmplBnf'
GO

IF EXISTS (SELECT * FROM sysobjects WHERE uid=user_id()
           AND      type='U'
           AND      name='tbEmplBnf')
    BEGIN
        PRINT '   > DROPPING - tbEmplBnf ...'
        DROP TABLE tbEmplBnf
    END

GO

/*****************************************************************
Table Create for - c:\work\projects\benefit\db\TB\OBJS\tbEmplBnf.sql

WHEN        WHO         WHAT
--------  ----------  --------------------------------------------
mm.dd.yy  You?        (Next change?)
09.25.96  kenbe       Original.
*****************************************************************/

PRINT 'CREATING      - tbEmplBnf ...'
GO

CREATE TABLE tbEmplBnf
(

     /* File - 09.25.96 - c:\work\projects\benefit\db\TB\OBJS\tbEmplBnf.sql */

     PKId         int          Identity(1,1)  PRIMARY KEY  CLUSTERED
    ,BenefitId    int          NOT NULL
                  REFERENCES  tbBenefit(PKId)
    ,CoverageId   int          NOT NULL
                  REFERENCES  tbEmplBnfCoverage(PKId)
    ,EmployeeId   int          NOT NULL
                  REFERENCES  tbEmployee(PKId)
    ,TaxId        int          NOT NULL
                  REFERENCES  tbEmplBnfTax(PKId)
    ,Amount       money        NULL    
)
GO

