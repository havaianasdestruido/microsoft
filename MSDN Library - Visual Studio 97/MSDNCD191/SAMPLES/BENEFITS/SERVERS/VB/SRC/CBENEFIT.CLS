VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CBenefit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'CBenefit.cls - Copyright (C) 1996 Microsoft Corporation
'             - All rights reserved.
'Date         - 12/20/96
'Authors      - Ken Bergmann, James Braum, Steve Kirk
'Purpose      - Benefit data class implementation for benefits servers

Option Explicit

Private m_bExpanded         As Boolean
Private m_iIdNo             As Integer
Private m_iIncrement        As Integer
Private m_iMultiples        As Integer
Private m_lBnfTaxId         As Long
Private m_lBnfTypeId        As Long
Private m_lCost             As Long
Private m_lMaximum          As Long
Private m_lMinimum          As Long
Private m_lPKId             As Long
Private m_sDescription      As String

#If UseEvents Then
    Private WithEvents m_oAdmin As CAdmin
Attribute m_oAdmin.VB_VarHelpID = -1
#Else
    Private m_oAdmin            As CAdmin
#End If

'Clear
'Purpose     - Restore object to a pre-expanded state
Private Sub Clear()
On Error Resume Next
m_bExpanded = False
m_iIncrement = icZero
m_iMultiples = icZero
m_lBnfTaxId = icZero
m_lBnfTypeId = icZero
m_lCost = icZero
m_lMaximum = icZero
m_lMinimum = icZero
m_sDescription = scBlank
End Sub

'Expand
'Purpose     - do look ups and fill in all object properties
'Outputs     - (Boolean) Success
Private Function Expand() As Boolean
Dim lstData     As CList
Dim sQry        As String

On Error Resume Next

'Nothing to do without m_lPKId
If ((CBool(m_lPKId)) And (Not m_oAdmin Is Nothing)) Then
    'Set up SQL for Benefit query
    If m_oAdmin.UseODBC Then
        sQry = LoadResString(scGet_Benefit_SQL) & CStr(m_lPKId)
    Else
        sQry = LoadResString(scGet_Benefit_MDB) & CStr(m_lPKId)
    End If
    'Restore object to a pre-expanded state
    Clear
    'Initialize list for Benefit query data
    Set lstData = New CList
    'If Benefit query fills list, populate member variables
    If m_oAdmin.ExecFillList(sQry, lstData) Then
        'Mark Benefit object as expanded
        m_bExpanded = True
        'Assign member variables
        If Not IsNull(lstData("BnfTaxId")) Then m_lBnfTaxId = lstData("BnfTaxId")
        If Not IsNull(lstData("BnfTypeId")) Then m_lBnfTypeId = lstData("BnfTypeId")
        If Not IsNull(lstData("Cost")) Then m_lCost = lstData("Cost")
        If Not IsNull(lstData("Description")) Then m_sDescription = lstData("Description")
        If Not IsNull(lstData("Increment")) Then m_iIncrement = lstData("Increment")
        If Not IsNull(lstData("Maximum")) Then m_lMaximum = lstData("Maximum")
        If Not IsNull(lstData("Minimum")) Then m_lMinimum = lstData("Minimum")
        If Not IsNull(lstData("Multiples")) Then m_iMultiples = lstData("Multiples")
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
Set m_oAdmin = oNewAdmin
End Property

Public Property Get BnfTaxId() As Long
On Error Resume Next
If m_bExpanded Then
    BnfTaxId = m_lBnfTaxId
Else
    If Expand() Then
        BnfTaxId = m_lBnfTaxId
    Else
        On Error GoTo 0
        Err.Raise icErrExpand + vbObjectError, Trim$(App.Title) & LoadResString(scBnfExpand), icErrExpand
        Exit Property
    End If
End If
End Property

Public Property Let BnfTaxId(ByVal lNewID As Long)
m_lBnfTaxId = lNewID
End Property

Public Property Get BnfTypeId() As Long
On Error Resume Next
If m_bExpanded Then
    BnfTypeId = m_lBnfTypeId
Else
    If Expand() Then
        BnfTypeId = m_lBnfTypeId
    Else
        On Error GoTo 0
        Err.Raise icErrExpand + vbObjectError, Trim$(App.Title) & LoadResString(scBnfExpand), icErrExpand
        Exit Property
    End If
End If
End Property

Public Property Let BnfTypeId(ByVal lNewID As Long)
m_lBnfTypeId = lNewID
End Property

Public Property Get Cost() As Long
On Error Resume Next
If m_bExpanded Then
    Cost = m_lCost
Else
    If Expand() Then
        Cost = m_lCost
    Else
        On Error GoTo 0
        Err.Raise icErrExpand + vbObjectError, Trim$(App.Title) & LoadResString(scBnfExpand), icErrExpand
        Exit Property
    End If
End If
End Property

Public Property Let Cost(ByVal lNewCost As Long)
m_lCost = lNewCost
End Property

Public Property Get Description() As String
On Error Resume Next
If m_bExpanded Then
    Description = m_sDescription
Else
    If Expand() Then
        Description = m_sDescription
    Else
        On Error GoTo 0
        Err.Raise icErrExpand + vbObjectError, Trim$(App.Title) & LoadResString(scBnfExpand), icErrExpand
        Exit Property
    End If
End If
End Property

Public Property Let Description(ByVal sNewDesc As String)
m_sDescription = sNewDesc
End Property

Public Property Get Increment() As Integer
On Error Resume Next
If m_bExpanded Then
    Increment = m_iIncrement
Else
    If Expand() Then
        Increment = m_iIncrement
    Else
        On Error GoTo 0
        Err.Raise icErrExpand + vbObjectError, Trim$(App.Title) & LoadResString(scBnfExpand), icErrExpand
        Exit Property
    End If
End If
End Property

Public Property Let Increment(ByVal iNewInc As Integer)
m_iIncrement = iNewInc
End Property

Public Property Get Maximum() As Long
On Error Resume Next
If m_bExpanded Then
    Maximum = m_lMaximum
Else
    If Expand() Then
        Maximum = m_lMaximum
    Else
        On Error GoTo 0
        Err.Raise icErrExpand + vbObjectError, Trim$(App.Title) & LoadResString(scBnfExpand), icErrExpand
        Exit Property
    End If
End If
End Property

Public Property Let Maximum(ByVal lNewMax As Long)
m_lMaximum = lNewMax
End Property

Public Property Get Minimum() As Long
On Error Resume Next
If m_bExpanded Then
    Minimum = m_lMinimum
Else
    If Expand() Then
        Minimum = m_lMinimum
    Else
        On Error GoTo 0
        Err.Raise icErrExpand + vbObjectError, Trim$(App.Title) & LoadResString(scBnfExpand), icErrExpand
        Exit Property
    End If
End If
End Property

Public Property Let Minimum(ByVal lNewMin As Long)
m_lMinimum = lNewMin
End Property

Public Property Get Multiples() As Integer
On Error Resume Next
If m_bExpanded Then
    Multiples = m_iMultiples
Else
    If Expand() Then
        Multiples = m_iMultiples
    Else
        On Error GoTo 0
        Err.Raise icErrExpand + vbObjectError, Trim$(App.Title) & LoadResString(scBnfExpand), icErrExpand
        Exit Property
    End If
End If
End Property

Public Property Let Multiples(ByVal iNewMult As Integer)
m_iMultiples = iNewMult
End Property

Public Property Get PKId() As Integer
On Error Resume Next
PKId = m_lPKId
End Property

Public Property Let PKId(ByVal iNewId As Integer)
m_lPKId = iNewId
End Property

Private Sub Class_Terminate()
On Error Resume Next
Term
'Release Admin object reference
Set m_oAdmin = Nothing
End Sub

