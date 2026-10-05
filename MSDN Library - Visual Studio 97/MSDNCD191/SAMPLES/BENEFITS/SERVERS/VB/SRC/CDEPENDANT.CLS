VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDependant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'CDependant.cls - Copyright (C) 1996 Microsoft Corporation
'               - All rights reserved.
'Date           - 12/20/96
'Authors        - Ken Bergmann, James Braum, Steve Kirk
'Purpose        - Dependant data class implementation for benefits servers

Option Explicit
Private m_oEmployee         As CEmployee
Private m_bExpanded         As Boolean
Private m_bGender           As Boolean
Private m_bStatus           As Boolean
Private m_dtDateOfBirth     As Date
Private m_iIdNo             As Integer
Private m_lDepTypeId        As Long
Private m_lPKId             As Long
Private m_sFirstName        As String
Private m_sLastName         As String

#If UseEvents Then
    Private WithEvents m_oAdmin As CAdmin
Attribute m_oAdmin.VB_VarHelpID = -1
#Else
    Private m_oAdmin            As CAdmin
#End If

'Clear
'Purpose     - Restore object to a pre-expanded state
Private Sub Clear()
Dim dtEmpty As Date

On Error Resume Next
m_bGender = False
m_bExpanded = False
m_bStatus = False
m_dtDateOfBirth = dtEmpty
m_lDepTypeId = icZero
Set m_oEmployee = Nothing
m_sFirstName = scBlank
m_sLastName = scBlank
End Sub

'Expand
'Purpose     - do look ups and fill in all object properties
'Outputs     - (Boolean) Success
Private Function Expand() As Boolean
Dim lstData   As CList
Dim lIndexPos As Long
Dim sQry      As String

On Error Resume Next
'Nothing to do without m_lPKId
If ((CBool(m_lPKId)) And (Not m_oAdmin Is Nothing)) Then
    'Set up SQL for Dependant query
    If m_oAdmin.UseODBC Then
        sQry = LoadResString(scGet_Dependant_SQL) & CStr(m_lPKId)
    Else
        sQry = LoadResString(scGet_Dependant_MDB) & CStr(m_lPKId)
    End If
    'Restore object to a pre-expanded state
    Clear
    'Initialize list for Dependant query data
    Set lstData = New CList
    'If Dependant query fills list, populate member variables
    If m_oAdmin.ExecFillList(sQry, lstData) Then
        'Mark Benefit object as expanded
        m_bExpanded = True
        'Assign member variables
        If Not IsNull(lstData("DepTypeId")) Then m_lDepTypeId = lstData("DepTypeId")
        If Not IsNull(lstData("DateOfBirth")) Then m_dtDateOfBirth = lstData("DateOfBirth")
        If Not IsNull(lstData("FirstName")) Then m_sFirstName = lstData("FirstName")
        If Not IsNull(lstData("Gender")) Then m_bGender = lstData("Gender")
        If Not IsNull(lstData("LastName")) Then m_sLastName = lstData("LastName")
        If Not IsNull(lstData("Status")) Then m_bStatus = lstData("Status")
        Set m_oEmployee = Nothing
        If Not IsNull(lstData("EmployeeId")) Then Set m_oEmployee = m_oAdmin.Employees(lstData("EmployeeId"))
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

Public Property Get DateOfBirth() As Date
On Error Resume Next
If m_bExpanded Then
    DateOfBirth = m_dtDateOfBirth
Else
    If Expand() Then
        DateOfBirth = m_dtDateOfBirth
    Else
        On Error GoTo 0
        Err.Raise icErrExpand + vbObjectError, Trim$(App.Title) & LoadResString(scDepExpand), icErrExpand
        Exit Property
    End If
End If
End Property

Public Property Let DateOfBirth(ByVal dtDateOfBirth As Date)
m_dtDateOfBirth = dtDateOfBirth
End Property

Public Property Get DepTypeId() As Long
On Error Resume Next
If m_bExpanded Then
    DepTypeId = m_lDepTypeId
Else
    If Expand() Then
        DepTypeId = m_lDepTypeId
    Else
        On Error GoTo 0
        Err.Raise icErrExpand + vbObjectError, Trim$(App.Title) & LoadResString(scDepExpand), icErrExpand
        Exit Property
    End If
End If
End Property

Public Property Let DepTypeId(ByVal lDepTypeId As Long)
m_lDepTypeId = lDepTypeId
End Property

Property Get Employee() As CEmployee
On Error Resume Next
If m_bExpanded Then
    Set Employee = m_oEmployee
Else
    If Expand() Then
        Set Employee = m_oEmployee
    Else
        On Error GoTo 0
        Err.Raise icErrExpand + vbObjectError, Trim$(App.Title) & LoadResString(scDepExpand), icErrExpand
        Exit Property
    End If
End If
End Property

Public Property Set Employee(oEmployee As CEmployee)
Set m_oEmployee = oEmployee
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
        Err.Raise icErrExpand + vbObjectError, Trim$(App.Title) & LoadResString(scDepExpand), icErrExpand
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
        Err.Raise icErrExpand + vbObjectError, Trim$(App.Title) & LoadResString(scDepExpand), icErrExpand
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
        Err.Raise icErrExpand + vbObjectError, Trim$(App.Title) & LoadResString(scDepExpand), icErrExpand
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

Public Property Get Status() As Boolean
On Error Resume Next
If m_bExpanded Then
    Status = m_bStatus
Else
    If Expand() Then
        Status = m_bStatus
    Else
        On Error GoTo 0
        Err.Raise icErrExpand + vbObjectError, Trim$(App.Title) & LoadResString(scDepExpand), icErrExpand
        Exit Property
    End If
End If
End Property

Public Property Let Status(bStatus As Boolean)
m_bStatus = bStatus
End Property

Private Sub Class_Terminate()
On Error Resume Next
Term
'Release Admin object reference
Set m_oAdmin = Nothing
End Sub


