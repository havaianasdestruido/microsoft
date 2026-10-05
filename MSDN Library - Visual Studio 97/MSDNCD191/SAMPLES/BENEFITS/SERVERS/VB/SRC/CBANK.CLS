VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CBank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'CBank.cls   - Copyright (C) 1996 Microsoft Corporation
'            - All rights reserved.
'Date        - 12/20/96
'Authors     - Ken Bergmann, James Braum, Steve Kirk
'Purpose     - Bank data class implementation for benefits servers

Option Explicit

Private m_lstEmployees      As CListEmployee
Private m_bExpanded         As Boolean
Private m_iIdNo             As Integer
Private m_lPKId             As Long
Private m_sAddress1         As String
Private m_sAddress2         As String
Private m_sCity             As String
Private m_sCode             As String
Private m_sName             As String
#If UseEvents Then
    Private WithEvents m_oAdmin As CAdmin
Attribute m_oAdmin.VB_VarHelpID = -1
#Else
    Private m_oAdmin            As CAdmin
#End If

'Clear
'Purpose     - Restore object to a pre-expanded state
Private Sub Clear()
Dim oEmp    As CEmployee
On Error Resume Next

'Return all remaining member variables to initial values
m_bExpanded = False
m_sAddress1 = scBlank
m_sAddress2 = scBlank
m_sCity = scBlank
m_sCode = scBlank
m_sName = scBlank
'Remove employees list
Set m_lstEmployees = Nothing
End Sub

'Expand
'Purpose     - do look ups and fill in all object properties
'Outputs     - (Boolean) Success
Private Function Expand() As Boolean
Dim oEmp        As CEmployee
Dim lstData     As CList
Dim iIndex      As Integer
Dim lIndexPos   As Long
Dim sQry        As String
Dim sTemp()     As String

On Error Resume Next
'Nothing to do without m_lPKId
If ((CBool(m_lPKId)) And (Not m_oAdmin Is Nothing)) Then
    'Set up SQL for Bank query
    If m_oAdmin.UseODBC Then
        sQry = LoadResString(scGet_Bank_SQL) & CStr(m_lPKId)
    Else
        sQry = LoadResString(scGet_Bank_MDB) & CStr(m_lPKId)
    End If
    'Restore object to a pre-expanded state
    Clear
    'Initialize list for Bank query data
    Set lstData = New CList
    'If Bank query fills list, populate member variables
    If m_oAdmin.ExecFillList(sQry, lstData) Then
        'Mark Bank object as expanded
        m_bExpanded = True
        'Assign member variables
        If Not IsNull(lstData("Address1")) Then m_sAddress1 = lstData("Address1")
        If Not IsNull(lstData("Address2")) Then m_sAddress2 = lstData("Address2")
        If Not IsNull(lstData("City")) Then m_sCity = lstData("City")
        If Not IsNull(lstData("Code")) Then m_sCode = lstData("Code")
        If Not IsNull(lstData("Name")) Then m_sName = lstData("Name")
        'Initialize Employees list
        Set m_lstEmployees = New CListEmployee
        'Set up SQL for Employees query
        If m_oAdmin.UseODBC Then
            sQry = LoadResString(scLvw_EmployeeByBankId_SQL) & CStr(m_lPKId)
        Else
            sQry = LoadResString(scLvw_EmployeeByBankId_MDB) & CStr(m_lPKId)
        End If
        'If Employees query fills list, populate Employees list
        If m_oAdmin.ExecFillArray(sQry, sTemp()) Then
            For iIndex = LBound(sTemp, 2) To UBound(sTemp, 2)
                Set oEmp = m_oAdmin.Employees(CLng(sTemp(icLvwPKId, iIndex)))
                If Not oEmp Is Nothing Then m_lstEmployees.Add oEmp
                Set oEmp = Nothing
            Next iIndex
        End If
    End If
Else 'no pkid
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

Public Property Set Admin(oNewAdmin As CAdmin)
On Error Resume Next
Set m_oAdmin = oNewAdmin
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
        Err.Raise icErrExpand + vbObjectError, Trim$(App.Title) & LoadResString(scBankExpand), icErrExpand
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
        Err.Raise icErrExpand + vbObjectError, Trim$(App.Title) & LoadResString(scBankExpand), icErrExpand
        Exit Property
    End If
End If
End Property

Public Property Let Address2(ByVal sAddress2 As String)
m_sAddress2 = sAddress2
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
        Err.Raise icErrExpand + vbObjectError, Trim$(App.Title) & LoadResString(scBankExpand), icErrExpand
        Exit Property
    End If
End If
End Property

Public Property Let City(ByVal sCity As String)
m_sCity = sCity
End Property

Public Property Get Code() As String
On Error Resume Next
If m_bExpanded Then
    Code = m_sCode
Else
    If Expand() Then
        Code = m_sCode
    Else
        On Error GoTo 0
        Err.Raise icErrExpand + vbObjectError, Trim$(App.Title) & LoadResString(scBankExpand), icErrExpand
        Exit Property
    End If
End If
End Property

Public Property Let Code(ByVal sCode As String)
m_sCode = sCode
End Property

Property Get Employees() As CListEmployee
On Error Resume Next
If m_bExpanded Then
    Set Employees = m_lstEmployees
Else
    If Expand() Then
        Set Employees = m_lstEmployees
    Else
        On Error GoTo 0
        Err.Raise icErrExpand + vbObjectError, Trim$(App.Title) & LoadResString(scBankExpand), icErrExpand
        Exit Property
    End If
End If
End Property

Public Property Get Name() As String
On Error Resume Next
If m_bExpanded Then
    Name = m_sName
Else
    If Expand() Then
        Name = m_sName
    Else
        On Error GoTo 0
        Err.Raise icErrExpand + vbObjectError, Trim$(App.Title) & LoadResString(scBankExpand), icErrExpand
        Exit Property
    End If
End If
End Property

Public Property Let Name(ByVal sName As String)
m_sName = sName
End Property

Public Property Get PKId() As Long
On Error Resume Next
PKId = m_lPKId
End Property

Public Property Let PKId(ByVal lPKId As Long)
On Error Resume Next
m_lPKId = lPKId
End Property

Private Sub Class_Terminate()
On Error Resume Next
'preserve m_oAdmin object reference return other object properties to initial state
Term
'Release Admin object reference
Set m_oAdmin = Nothing
End Sub

#If UseEvents Then
Private Sub m_oAdmin_EmployeesChanged()
On Error Resume Next
Expand
End Sub
#End If
