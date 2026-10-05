GO
Set nocount on

PRINT '   > DROP FKs - tbDependant ...'
EXECUTE sp_dba_DropFKs 'tbDependant'
GO

IF EXISTS (SELECT * FROM sysobjects WHERE uid=user_id()
           AND      type='U'
           AND      name='tbDependant')
    BEGIN
        PRINT '   > DROPPING - tbDependant ...'
        DROP TABLE tbDependant
    END

GO

/*****************************************************************
Table Create for - c:\work\projects\benefit\db\TB\OBJS\tbDependant.sql

WHEN        WHO         WHAT
--------  ----------  --------------------------------------------
mm.dd.yy  You?        (Next change?)
09.25.96  kenbe       Original.
*****************************************************************/

PRINT 'CREATING      - tbDependant ...'
GO

CREATE TABLE tbDependant
(

     /* File - 09.25.96 - c:\work\projects\benefit\db\TB\OBJS\tbDependant.sql */

     PKId         int          Identity(1,1)  PRIMARY KEY  CLUSTERED
    ,DepTypeId    int          NOT NULL
                  REFERENCES  tbDepType(PKId)
    ,EmployeeId   int          NOT NULL
                  REFERENCES  tbEmployee(PKId)
    ,DateOfBirth  datetime     NULL    
    ,FirstName    varchar(35)  NOT NULL
    ,Gender       bit          NOT NULL
    ,LastName     varchar(35)  NOT NULL
    ,Status       bit          NOT NULL
)
GO

