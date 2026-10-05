VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'CList.cls - Copyright (C) 1996 Microsoft Corporation
'          - All rights reserved.
'Date      - 12/20/96
'Authors   - Ken Bergmann, James Braum, Steve Kirk
'Purpose   - Clist list implementation for benefits servers

Option Explicit
Private m_cList As Collection
Private m_iIdNo As Integer

'Add
'Purpose     - Add to list
'Inputs      - (String)  sKey
'            - (Variant) vItem
Public Sub Add(ByVal sKey As String, vItem As Variant)
m_cList.Add vItem, sKey
End Sub

'Remove
'Purpose     - Remove from list
'Inputs      - (String)  sKey
Public Sub Remove(sKey As String)
m_cList.Remove sKey
End Sub

Public Property Get Count() As Long
Count = m_cList.Count
End Property

Public Property Get Item(sKey As String) As Variant
Attribute Item.VB_UserMemId = 0
On Error Resume Next
If m_cList.Count > 0 Then
    If IsObject(m_cList(sKey)) Then
        Set Item = m_cList(sKey)
    Else
        Item = m_cList(sKey)
    End If
End If
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
On Error Resume Next
Set NewEnum = m_cList.[_NewEnum]
End Property

Private Sub Class_Initialize()
On Error Resume Next
Set m_cList = New Collection
End Sub

Private Sub Class_Terminate()
On Error Resume Next
Set m_cList = Nothing
End Sub

