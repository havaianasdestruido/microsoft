GO
Set nocount on

PRINT '   > DROP FKs - tbEmployee ...'
EXECUTE sp_dba_DropFKs 'tbEmployee'
GO

IF EXISTS (SELECT * FROM sysobjects WHERE uid=user_id()
           AND      type='U'
           AND      name='tbEmployee')
    BEGIN
        PRINT '   > DROPPING - tbEmployee ...'
        DROP TABLE tbEmployee
    END

GO

/*****************************************************************
Table Create for - c:\work\projects\benefit\db\TB\OBJS\tbEmployee.sql

WHEN        WHO         WHAT
--------  ----------  --------------------------------------------
mm.dd.yy  You?        (Next change?)
09.25.96  kenbe       Original.
*****************************************************************/

PRINT 'CREATING      - tbEmployee ...'
GO

CREATE TABLE tbEmployee
(

     /* File - 09.25.96 - c:\work\projects\benefit\db\TB\OBJS\tbEmployee.sql */

     PKId         int          Identity(1,1)  PRIMARY KEY  CLUSTERED
    ,BankId       int          NOT NULL
                  REFERENCES  tbBank(PKId)
    ,Address1     varchar(60)  NULL    
    ,Address2     varchar(60)  NULL    
    ,BankAccount  varchar(35)  NULL    
    ,City         varchar(50)  NULL    
    ,DateOfBirth  datetime     NULL    
    ,EndDate      datetime     NULL    
    ,FirstName    varchar(35)  NOT NULL
    ,Gender       bit          NOT NULL
    ,LastName     varchar(35)  NOT NULL
    ,Position     varchar(50)  NOT NULL
    ,Salary       money        NULL    
    ,StartDate    datetime     NOT NULL
    ,UserName     varchar(16)  NOT NULL
)
GO

