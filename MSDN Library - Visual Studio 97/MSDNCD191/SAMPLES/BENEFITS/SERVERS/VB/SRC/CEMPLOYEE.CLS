VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CEmployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'CEmployee.cls - Copyright (C) 1996 Microsoft Corporation
'              - All rights reserved.
'Date          - 12/20/96
'Authors       - Ken Bergmann, James Braum, Steve Kirk
'Purpose       - Employee data class implementation for benefits servers

Option Explicit
Private m_oBank             As CBank
Private m_lstBenefitLinks   As CListBenefitLink
Private m_lstDependant      As CListDependant

Private m_bExpanded         As Boolean
Private m_bGender           As Boolean
Private m_cSalary           As Currency
Private m_dtDateOfBirth     As Date
Private m_dtEndDate         As Date
Private m_dtStartDate       As Date
Private m_iIdNo             As Integer
Private m_lPKId             As Long
Private m_sAddress1         As String
Private m_sAddress2         As String
Private m_sBankAccount      As String
Private m_sCity             As String
Private m_sFirstName        As String
Private m_sLastName         As String
Private m_sPosition         As String
Private m_sUserName         As String

#If UseEvents Then
    Private WithEvents m_oAdmin As CAdmin
Attribute m_oAdmin.VB_VarHelpID = -1
#Else
    Private m_oAdmin            As CAdmin
#End If

'Clear
'Purpose     - Restore object to a pre-expanded state
Private Sub Clear()
Dim oBnf    As CBenefitLink
Dim oDep    As CDependant
Dim cZero As Currency
Dim dtEmpty As Date

On Error Resume Next
m_bExpanded = False
'Release BenefitLinks
Set m_lstBenefitLinks = Nothing
'Release Dependants
Set m_lstDependant = Nothing
'Clear remaining member variables
Set m_oBank = Nothing
m_bGender = False
m_cSalary = cZero
m_dtDateOfBirth = dtEmpty
m_dtEndDate = dtEmpty
m_dtStartDate = dtEmpty
m_sAddress1 = scBlank
m_sAddress2 = scBlank
m_sBankAccount = scBlank
m_sCity = scBlank
m_sFirstName = scBlank
m_sLastName = scBlank
m_sPosition = scBlank
m_sUserName = scBlank
End Sub

'Expand
'Purpose     - do look ups and fill in all object properties
'Outputs     - (Boolean) Success
Private Function Expand() As Boolean
Dim oBenefitLink    As CBenefitLink
Dim oDependant      As CDependant
Dim lstData         As CList
Dim iIndex          As Integer
Dim lIndexPos       As Long
Dim sQry            As String
Dim sTemp()         As String

On Error Resume Next

'Nothing to do without m_lPKId
If ((CBool(m_lPKId)) And (Not m_oAdmin Is Nothing)) Then
    'Set up SQL for Employee query
    If m_oAdmin.UseODBC Then
        sQry = LoadResString(scGet_Employee_SQL) & CStr(m_lPKId)
    Else
        sQry = LoadResString(scGet_Employee_MDB) & CStr(m_lPKId)
    End If
    'Restore object to a pre-expanded state
    Clear
    'Initialize list for Employee query data
    Set lstData = New CList
    'If Employee query fills list, populate member variables
    If m_oAdmin.ExecFillList(sQry, lstData) Then
        'Mark Employee object as expanded
        m_bExpanded = True
        'Assign member variables
        Set m_oBank = Nothing
        If Not IsNull(lstData("FirstName")) Then m_sFirstName = lstData("FirstName")
        If Not IsNull(lstData("LastName")) Then m_sLastName = lstData("LastName")
        If Not IsNull(lstData("Gender")) Then m_bGender = lstData("Gender")
        If Not IsNull(lstData("DateOfBirth")) Then m_dtDateOfBirth = lstData("DateOfBirth")
        If Not IsNull(lstData("Address1")) Then m_sAddress1 = lstData("Address1")
        If Not IsNull(lstData("Address2")) Then m_sAddress2 = lstData("Address2")
        If Not IsNull(lstData("City")) Then m_sCity = lstData("City")
        If Not IsNull(lstData("BankId")) Then Set m_oBank = m_oAdmin.Banks(lstData("BankId"))
        If Not IsNull(lstData("BankAccount")) Then m_sBankAccount = lstData("BankAccount")
        If Not IsNull(lstData("Position")) Then m_sPosition = lstData("Position")
        If Not IsNull(lstData("Salary")) Then m_cSalary = lstData("Salary")
        If Not IsNull(lstData("StartDate")) Then m_dtStartDate = lstData("StartDate")
        If Not IsNull(lstData("EndDate")) Then m_dtEndDate = lstData("EndDate")
        If Not IsNull(lstData("UserName")) Then m_sUserName = lstData("UserName")
        'Initialize Dependants list
        Set m_lstDependant = New CListDependant
        'Set up SQL for Dependants query
        If m_oAdmin.UseODBC Then
            sQry = LoadResString(scLvw_DependantByEmployeeId_SQL) & CStr(m_lPKId)
        Else
            sQry = LoadResString(scLvw_DependantByEmployeeId_MDB) & CStr(m_lPKId)
        End If
        'If Dependants query fills list, populate Dependants list
        If m_oAdmin.ExecFillArray(sQry, sTemp()) Then
            For iIndex = LBound(sTemp, 2) To UBound(sTemp, 2)
                Set oDependant = New CDependant
                Set oDependant.Admin = m_oAdmin
                oDependant.PKId = CLng(sTemp(icLvwPKId, iIndex))
                m_lstDependant.Add oDependant
                Set oDependant = Nothing
            Next iIndex
        End If
        'Initialize BenefitLinks list
        Set m_lstBenefitLinks = New CListBenefitLink
        'Set up SQL for BenefitLinks query
        If m_oAdmin.UseODBC Then
            sQry = LoadResString(scLvw_EmplBnfByEmployeeId_SQL) & CStr(m_lPKId)
        Else
            sQry = LoadResString(scLvw_EmplBnfByEmployeeId_MDB) & CStr(m_lPKId)
        End If
        'If BenefitLinks query fills list, populate BenefitLinks list
        If m_oAdmin.ExecFillArray(sQry, sTemp()) Then
            For iIndex = LBound(sTemp, 2) To UBound(sTemp, 2)
                Set oBenefitLink = New CBenefitLink
                Set oBenefitLink.Admin = m_oAdmin
                oBenefitLink.PKId = CLng(sTemp(icLvwPKId, iIndex))
                m_lstBenefitLinks.Add oBenefitLink
                Set oBenefitLink = Nothing
            Next iIndex
        End If
    End If
Else  'no pkid
    'Mark as expanded to prevent repeated attempts to expand
    m_bExpanded = True
End If
Expand = m_bExpanded
End Function

'Term
'Purpose     - Terminate an object, release all object references
Public Sub Term()
On Error Resume Next
'Restore object to a pre-expanded state
Clear
End Sub

Public Property Set Admin(oAdmin As CAdmin)
Set m_oAdmin = oAdmin
End Property

Public Property Get Address1() As String
On Error Resume Next
If m_bExpanded Then
    Address1 = m_sAddress1
Else
    If Expand() Then
        Address1 = m_sAddress1
    Else
        On Error GoTo 0
        Err.Raise icErrExpand + vbObjectError, Trim$(App.Title) & LoadResString(scEmplExpand), icErrExpand
        Exit Property
    End If
End If
End Property

Public Property Let Address1(ByVal sAddress1 As String)
m_sAddress1 = sAddress1
End Property

Public Property Get Address2() As String
On Error Resume Next
If m_bExpanded Then
    Address2 = m_sAddress2
Else
    If Expand() Then
        Address2 = m_sAddress2
    Else
        On Error GoTo 0
        Err.Raise icErrExpand + vbObjectError, Trim$(App.Title) & LoadResString(scEmplExpand), icErrExpand
        Exit Property
    End If
End If
End Property

Public Property Let Address2(ByVal sAddress2 As String)
m_sAddress2 = sAddress2
End Property

Public Property Get Bank() As CBank
On Error Resume Next
If m_bExpanded Then
    Set Bank = m_oBank
Else
    If Expand() Then
        Set Bank = m_oBank
    Else
        On Error GoTo 0
        Err.Raise icErrExpand + vbObjectError, Trim$(App.Title) & LoadResString(scEmplExpand), icErrExpand
        Exit Property
    End If
End If
End Property

Public Property Set Bank(ByVal oBank As CBank)
Set m_oBank = oBank
End Property

Public Property Get BankAccount() As String
On Error Resume Next
If m_bExpanded Then
    BankAccount = m_sBankAccount
Else
    If Expand() Then
        BankAccount = m_sBankAccount
    Else
        On Error GoTo 0
        Err.Raise icErrExpand + vbObjectError, Trim$(App.Title) & LoadResString(scEmplExpand), icErrExpand
        Exit Property
    End If
End If
End Property

Public Property Let BankAccount(ByVal sBankAccount As String)
m_sBankAccount = sBankAccount
End Property

Public Property Get BenefitLinks() As CListBenefitLink
On Error Resume Next
If m_bExpanded Then
    Set BenefitLinks = m_lstBenefitLinks
Else
    If Expand() Then
        Set BenefitLinks = m_lstBenefitLinks
    Else
        On Error GoTo 0
        Err.Raise icErrExpand + vbObjectError, Trim$(App.Title) & LoadResString(scEmplExpand), icErrExpand
        Exit Property
    End If
End If
End Property

Public Property Get City() As String
On Error Resume Next
If m_bExpanded Then
    City = m_sCity
Else
    If Expand() Then
        City = m_sCity
    Else
        On Error GoTo 0
        Err.Raise icErrExpand + vbObjectError, Trim$(App.Title) & LoadResString(scEmplExpand), icErrExpand
        Exit Property
    End If
End If
End Property

Public Property Let City(ByVal sCity As String)
m_sCity = sCity
End Property

Public Property Get DateOfBirth() As Date
On Error Resume Next
If m_bExpanded Then
    DateOfBirth = m_dtDateOfBirth
Else
    If Expand() Then
        DateOfBirth = m_dtDateOfBirth
    Else
        On Error GoTo 0
        Err.Raise icErrExpand + vbObjectError, Trim$(App.Title) & LoadResString(scEmplExpand), icErrExpand
        Exit Property
    End If
End If
End Property

Public Property Let DateOfBirth(ByVal dtDateOfBirth As Date)
m_dtDateOfBirth = dtDateOfBirth
End Property

Public Property Get Dependants() As CListDependant
On Error Resume Next
If m_bExpanded Then
    Set Dependants = m_lstDependant
Else
    If Expand() Then
        Set Dependants = m_lstDependant
    Else
        On Error GoTo 0
        Err.Raise icErrExpand + vbObjectError, Trim$(App.Title) & LoadResString(scEmplExpand), icErrExpand
        Exit Property
    End If
End If
End Property

Public Property Get EndDate() As Date
On Error Resume Next
If m_bExpanded Then
    EndDate = m_dtEndDate
Else
    If Expand() Then
        EndDate = m_dtEndDate
    Else
        On Error GoTo 0
        Err.Raise icErrExpand + vbObjectError, Trim$(App.Title) & LoadResString(scEmplExpand), icErrExpand
        Exit Property
    End If
End If
End Property

Public Property Let EndDate(ByVal dtEndDate As Date)
m_dtEndDate = dtEndDate
End Property

Public Property Get FirstName() As String
On Error Resume Next
If m_bExpanded Then
    FirstName = m_sFirstName
Else
    If Expand() Then
        FirstName = m_sFirstName
    Else
        On Error GoTo 0
        Err.Raise icErrExpand + vbObjectError, Trim$(App.Title) & LoadResString(scEmplExpand), icErrExpand
        Exit Property
    End If
End If
End Property

Public Property Let FirstName(ByVal sFirstName As String)
m_sFirstName = sFirstName
End Property

Public Property Get Gender() As Boolean
On Error Resume Next
If m_bExpanded Then
    Gender = m_bGender
Else
    If Expand() Then
        Gender = m_bGender
    Else
        On Error GoTo 0
        Err.Raise icErrExpand + vbObjectError, Trim$(App.Title) & LoadResString(scEmplExpand), icErrExpand
        Exit Property
    End If
End If
End Property

Public Property Let Gender(ByVal bGender As Boolean)
m_bGender = bGender
End Property

Public Property Get LastName() As String
On Error Resume Next
If m_bExpanded Then
    LastName = m_sLastName
Else
    If Expand() Then
        LastName = m_sLastName
    Else
        On Error GoTo 0
        Err.Raise icErrExpand + vbObjectError, Trim$(App.Title) & LoadResString(scEmplExpand), icErrExpand
        Exit Property
    End If
End If
End Property

Public Property Let LastName(ByVal sLastName As String)
m_sLastName = sLastName
End Property

Public Property Get PKId() As Long
PKId = m_lPKId
End Property

Public Property Let PKId(ByVal lPKId As Long)
m_lPKId = lPKId
End Property

Public Property Get Position() As String
On Error Resume Next
If m_bExpanded Then
    Position = m_sPosition
Else
    If Expand() Then
        Position = m_sPosition
    Else
        On Error GoTo 0
        Err.Raise icErrExpand + vbObjectError, Trim$(App.Title) & LoadResString(scEmplExpand), icErrExpand
        Exit Property
    End If
End If
End Property

Public Property Let Position(ByVal sPosition As String)
m_sPosition = sPosition
End Property

Public Property Get Salary() As Currency
On Error Resume Next
If m_bExpanded Then
    Salary = m_cSalary
Else
    If Expand() Then
        Salary = m_cSalary
    Else
        On Error GoTo 0
        Err.Raise icErrExpand + vbObjectError, Trim$(App.Title) & LoadResString(scEmplExpand), icErrExpand
        Exit Property
    End If
End If
End Property

Public Property Let Salary(ByVal cSalary As Currency)
m_cSalary = cSalary
End Property

Public Property Get StartDate() As Date
On Error Resume Next
If m_bExpanded Then
    StartDate = m_dtStartDate
Else
    If Expand() Then
        StartDate = m_dtStartDate
    Else
        On Error GoTo 0
        Err.Raise icErrExpand + vbObjectError, Trim$(App.Title) & LoadResString(scEmplExpand), icErrExpand
        Exit Property
    End If
End If
End Property

Public Property Let StartDate(ByVal dtStartDate As Date)
m_dtStartDate = dtStartDate
End Property

Public Property Get UserName() As String
On Error Resume Next
If m_bExpanded Then
    UserName = m_sUserName
Else
    If Expand() Then
        UserName = m_sUserName
    Else
        On Error GoTo 0
        Err.Raise icErrExpand + vbObjectError, Trim$(App.Title) & LoadResString(scEmplExpand), icErrExpand
        Exit Property
    End If
End If
End Property

Public Property Let UserName(ByVal sUserName As String)
m_sUserName = sUserName
End Property

Private Sub Class_Terminate()
On Error Resume Next
Term
'Release Admin object reference
Set m_oAdmin = Nothing
End Sub

#If UseEvents Then
Private Sub m_oAdmin_BanksChanged()
If Not m_oBank Is Nothing Then Set m_oBank = m_oAdmin.Banks(m_oBank.PKId)
End Sub
#End If

#If UseEvents Then
Private Sub m_oAdmin_DependantChanged(ByVal lEmployeePKId As Long)
On Error Resume Next
If m_lPKId = lEmployeePKId Then Expand
End Sub
#End If
