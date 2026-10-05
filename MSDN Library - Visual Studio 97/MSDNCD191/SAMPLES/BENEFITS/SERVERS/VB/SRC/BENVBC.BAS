Attribute VB_Name = "basBenefitsVBConstants"
'BenVBC.bas  - Copyright (C) 1996 Microsoft Corporation
'            - All rights reserved.
'Date        - 12/20/96
'Authors     - Ken Bergmann, James Braum, Steve Kirk
'Purpose     - Constants for Benefits API
Option Explicit

Public Const icZero = 0

Public Const scBlank = ""
Public Const scDateFormat = "dd mmm yyyy"
Public Const scNull = "NULL"
Public Const scOneStr = " 1 "
Public Const scSpace = " "
Public Const scSpaceCommaSpace = " , "
Public Const scSQ = "'"
Public Const scZeroStr = " 0 "

Public Enum DataCaption
    scBnfName = 9
    scBnfTypeName = 10
    scEmplCnt = 11
    scLvwDesc = 12
    scLvwCode = 13
    scLvwPKId = 14
End Enum

Public Enum Error
    icErrSrcType = 1
    icErrOpenDB = 2
    icErrReqFld = 3
    icErrFormat = 4
    icErrExpand = 5
    icErrPKId = 6
    icErrObjectOp = 7
    icErrBadIndex = 8
End Enum

Public Enum FunctionName
    scDel = 15
    scDelEmpl = 16
    scDelBank = 17
    scDelDep = 18
    scDelBnf = 19
    scDelBnfLnk = 20
    scExecFillArray = 21
    scExecFillList = 22
    scFillLstBank = 23
    scFillLstBnf = 24
    scFillLstEmpl = 25
    scGetLstBank = 26
    scGetLstBnf = 27
    scGetLstEmpl = 28
    scGetDomain = 29
    scInit = 30
    scIns = 31
    scInsEmpl = 32
    scInsBank = 33
    scInsDep = 34
    scInsBnf = 35
    scInsBnfLnk = 36
    scUpd = 37
    scUpdEmpl = 38
    scUpdBank = 39
    scUpdDep = 40
    scUpdBnf = 41
    scUpdBnfLnk = 42
    scGetBnfTaxTypes = 43
    scGetBnfTypes = 44
    scGetDepTypes = 45
    scGetEmplBnfCvrgTypes = 46
    scGetEmplBnfTaxTypes = 47
    scIsUserValid = 48
    scGetNumEmplsPerBnf = 49
    scGetNumEmplsPerBnfType = 50
    scExecBoolMDB = 51
    scExecBoolSQL = 52
    scBankExpand = 53
    scDepExpand = 54
    scEmplExpand = 55
    scBnfExpand = 56
    scBnfLnkExpand = 57
End Enum

Public Enum ListViewIndex
     icLvwDesc = 0
     icLvwPKId = 1
     icLvwCode = 2
End Enum

Public Enum SQL_Bank
    scIns_Bank = 68
    scUpd_Bank = 69
    scDel_Bank = 70
    scGet_Bank_MDB = 71
    scGet_Bank_SQL = 72
    scLvw_Bank_MDB = 73
    scLvw_Bank_SQL = 74
End Enum

Public Enum SQL_Benefit
    scIns_Benefit = 75
    scUpd_Benefit = 76
    scDel_Benefit = 77
    scGet_Benefit_MDB = 78
    scGet_Benefit_SQL = 79
    scLvw_Benefit_MDB = 80
    scLvw_Benefit_SQL = 81
End Enum

Public Enum SQL_BenefitTaxDomain
    scLvw_BnfTax_MDB = 82
    scLvw_BnfTax_SQL = 83
End Enum

Public Enum SQL_BenefitTypeDomain
    scLvw_BnfType_MDB = 84
    scLvw_BnfType_SQL = 85
End Enum

Public Enum SQL_Dependant
    scIns_Dependant = 86
    scUpd_Dependant = 87
    scDel_Dependant = 88
    scGet_Dependant_MDB = 89
    scGet_Dependant_SQL = 90
    scLvw_DependantByEmployeeId_MDB = 91
    scLvw_DependantByEmployeeId_SQL = 92
End Enum

Public Enum SQL_DependantTypeDomain
    scLvw_DepType_MDB = 93
    scLvw_DepType_SQL = 94
End Enum

Public Enum SQL_Dynamic
    scExecute = 58
    scGNEPBSQL1 = 59
    scGNEPBSQL2 = 60
    scGNEPBSQL3 = 61
    scGNEPBSQL4 = 62
    scGNEPBSQL5 = 63
    scGNEPBSQL6 = 64
    scGNEPBTSQL1 = 65
    scGNEPBTSQL2 = 66
    scCHKUSER = 67
End Enum

Public Enum SQL_Employee
    scIns_Employee = 106
    scUpd_Employee = 107
    scDel_Employee = 108
    scGet_Employee_MDB = 109
    scGet_Employee_SQL = 110
    scLvw_Employee_MDB = 111
    scLvw_Employee_SQL = 112
    scLvw_EmployeeByBankId_MDB = 113
    scLvw_EmployeeByBankId_SQL = 114
End Enum

Public Enum SQL_EmployeeBenefit
    scIns_EmplBnf = 95
    scUpd_EmplBnf = 96
    scDel_EmplBnf = 97
    scGet_EmplBnf_MDB = 98
    scGet_EmplBnf_SQL = 99
    scLvw_EmplBnfByEmployeeId_MDB = 100
    scLvw_EmplBnfByEmployeeId_SQL = 101
End Enum

Public Enum SQL_EmployeeBenefitCoverageDomain
    scLvw_EmplBnfCoverage_MDB = 102
    scLvw_EmplBnfCoverage_SQL = 103
End Enum

Public Enum SQL_EmployeeBenifitTaxDomain
    scLvw_EmplBnfTax_MDB = 104
    scLvw_EmplBnfTax_SQL = 105
End Enum

