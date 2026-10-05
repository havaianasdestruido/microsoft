VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CBenefitLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'CBenefitLink.cls - Copyright (C) 1996 Microsoft Corporation
'                 - All rights reserved.
'Date             - 12/20/96
'Authors          - Ken Bergmann, James Braum, Steve Kirk
'Purpose          - BenefitLink data class implementation for benefits servers

Option Explicit
Private m_oBenefit          As CBenefit
Private m_oEmployee         As CEmployee
Private m_bExpanded         As Boolean
Private m_curAmount         As Currency
Private m_iIdNo             As Integer
Private m_lCoverageId       As Long
Private m_lPKId             As Long
Private m_lTaxId            As Long

#If UseEvents Then
    Private WithEvents m_oAdmin As CAdmin
Attribute m_oAdmin.VB_VarHelpID = -1
#Else
    Private m_oAdmin            As CAdmin
#End If

'Clear
'Purpose     - preserve m_oAdmin object reference return other object properties to initial state
Private Sub Clear()
On Error Resume Next
m_bExpanded = False
m_lCoverageId = icZero
m_lTaxId = icZero
m_curAmount = CCur(icZero)
Set m_oEmployee = Nothing
Set m_oBenefit = Nothing
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
    'Set up SQL for BenefitLink query
    If m_oAdmin.UseODBC Then
        sQry = LoadResString(scGet_EmplBnf_SQL) & CStr(m_lPKId)
    Else
        sQry = LoadResString(scGet_EmplBnf_MDB) & CStr(m_lPKId)
    End If
    'Restore object to a pre-expanded state
    Clear
    'Initialize list for BenefitLink query data
    Set lstData = New CList
    'If BenefitLink query fills list, populate member variables
    If m_oAdmin.ExecFillList(sQry, lstData) Then
        'Mark BenefitLink object as expanded
        m_bExpanded = True
        'Assign member variables
        If Not IsNull(lstData("CoverageId")) Then m_lCoverageId = lstData("CoverageId")
        If Not IsNull(lstData("TaxId")) Then m_lTaxId = lstData("TaxId")
        If Not IsNull(lstData("Amount")) Then m_curAmount = lstData("Amount")
        Set m_oEmployee = Nothing
        If Not IsNull(lstData("EmployeeId")) Then Set m_oEmployee = m_oAdmin.Employees(lstData("EmployeeId"))
        Set m_oBenefit = Nothing
        If Not IsNull(lstData("BenefitId")) Then Set m_oBenefit = m_oAdmin.Benefits(lstData("BenefitId"))
    End If
Else  'No PKId
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

Public Property Get Amount() As Currency
On Error Resume Next
If m_bExpanded Then
    Amount = m_curAmount
Else
    If Expand() Then
        Amount = m_curAmount
    Else
        On Error GoTo 0
        Err.Raise icErrExpand + vbObjectError, Trim$(App.Title) & LoadResString(scBnfLnkExpand), LoadResString(icErrExpand)
        Exit Property
    End If
End If
End Property

Public Property Let Amount(ByVal cAmount As Currency)
m_curAmount = cAmount
End Property

Property Get Benefit() As CBenefit
On Error Resume Next
If m_bExpanded Then
    Set Benefit = m_oBenefit
Else
    If Expand() Then
        Set Benefit = m_oBenefit
    Else
        On Error GoTo 0
        Err.Raise icErrExpand + vbObjectError, Trim$(App.Title) & LoadResString(scBnfLnkExpand), LoadResString(icErrExpand)
        Exit Property
    End If
End If
End Property

Public Property Set Benefit(oBenefit As CBenefit)
Set m_oBenefit = oBenefit
End Property

Public Property Get CoverageId() As Long
On Error Resume Next
If m_bExpanded Then
    CoverageId = m_lCoverageId
Else
    If Expand() Then
        CoverageId = m_lCoverageId
    Else
        On Error GoTo 0
        Err.Raise icErrExpand + vbObjectError, Trim$(App.Title) & LoadResString(scBnfLnkExpand), LoadResString(icErrExpand)
        Exit Property
    End If
End If
End Property

Public Property Let CoverageId(ByVal lCoverageId As Long)
m_lCoverageId = lCoverageId
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
        Err.Raise icErrExpand + vbObjectError, Trim$(App.Title) & LoadResString(scBnfLnkExpand), LoadResString(icErrExpand)
        Exit Property
    End If
End If
End Property

Public Property Set Employee(oEmployee As CEmployee)
Set m_oEmployee = oEmployee
End Property

Public Property Get PKId() As Long
PKId = m_lPKId
End Property

Public Property Let PKId(ByVal lPKId As Long)
m_lPKId = lPKId
End Property

Public Property Get TaxId() As Long
On Error Resume Next
If m_bExpanded Then
    TaxId = m_lTaxId
Else
    If Expand() Then
        TaxId = m_lTaxId
    Else
        On Error GoTo 0
        Err.Raise icErrExpand + vbObjectError, Trim$(App.Title) & LoadResString(scBnfLnkExpand), LoadResString(icErrExpand)
        Exit Property
    End If
End If
End Property

Public Property Let TaxId(ByVal lTaxId As Long)
m_lTaxId = lTaxId
End Property

Private Sub Class_Terminate()
On Error Resume Next
Term
'Release Admin object reference
Set m_oAdmin = Nothing
End Sub

#If UseEvents Then
Private Sub m_oAdmin_BenefitsChanged()
On Error Resume Next
If Not m_oBenefit Is Nothing Then Set m_oBenefit = m_oAdmin.Benefits(m_oBenefit.PKId)
End Sub
#End If

#If UseEvents Then
Private Sub m_oAdmin_EmployeesChanged()
On Error Resume Next
If Not m_oEmployee Is Nothing Then Set m_oEmployee = m_oAdmin.Employees(m_oEmployee.PKId)
End Sub
#End If
