GO
Set nocount on

PRINT '   > DROP FKs - tbBenefit ...'
EXECUTE sp_dba_DropFKs 'tbBenefit'
GO

IF EXISTS (SELECT * FROM sysobjects WHERE uid=user_id()
           AND      type='U'
           AND      name='tbBenefit')
    BEGIN
        PRINT '   > DROPPING - tbBenefit ...'
        DROP TABLE tbBenefit
    END

GO

/*****************************************************************
Table Create for - c:\work\projects\benefit\db\TB\OBJS\tbBenefit.sql

WHEN        WHO         WHAT
--------  ----------  --------------------------------------------
mm.dd.yy  You?        (Next change?)
09.25.96  kenbe       Original.
*****************************************************************/

PRINT 'CREATING      - tbBenefit ...'
GO

CREATE TABLE tbBenefit
(

     /* File - 09.25.96 - c:\work\projects\benefit\db\TB\OBJS\tbBenefit.sql */

     PKId         int          Identity(1,1)  PRIMARY KEY  CLUSTERED
    ,BnfTaxId     int          NOT NULL
                  REFERENCES  tbBnfTax(PKId)
    ,BnfTypeId    int          NOT NULL
                  REFERENCES  tbBnfType(PKId)
    ,Cost         int          NOT NULL
    ,Description  varchar(80)  NOT NULL
    ,Increment    smallint     NOT NULL
    ,Maximum      int          NOT NULL
    ,Minimum      int          NOT NULL
    ,Multiples    smallint     NOT NULL
)
GO

