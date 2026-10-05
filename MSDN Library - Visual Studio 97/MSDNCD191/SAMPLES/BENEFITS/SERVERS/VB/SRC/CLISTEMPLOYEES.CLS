VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CListEmployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'CListEmployees.cls - Copyright (C) 1996 Microsoft Corporation
'                   - All rights reserved.
'Date               - 12/20/96
'Authors            - Ken Bergmann, James Braum, Steve Kirk
'Purpose            - CListEmployee list implementation for benefits servers

Option Explicit
Private m_lstEmployees  As CList
Private m_iIdNo         As Integer

'Add
'Purpose     - Add to list
'Inputs      - (CEmployee)    oEmployee
'Outputs     - (CEmployee)
Public Function Add(oEmployee As CEmployee) As CEmployee
m_lstEmployees.Add "ID=" & Trim$(CStr(oEmployee.PKId)), oEmployee
Set Add = oEmployee
End Function

'Remove
'Purpose     - Remove from the list
'Inputs      - (Long)    lIndexKey
Public Sub Remove(lIndexKey As Long)
m_lstEmployees.Remove "ID=" & Trim$(CStr(lIndexKey))
End Sub

Public Property Get Count() As Long
Count = m_lstEmployees.Count
End Property

Public Property Get Item(lIndexKey As Long) As CEmployee
Attribute Item.VB_UserMemId = 0
Set Item = m_lstEmployees("ID=" & Trim$(CStr(lIndexKey)))
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set NewEnum = m_lstEmployees.NewEnum
End Property

Private Sub Class_Initialize()
On Error Resume Next
Set m_lstEmployees = New CList
End Sub

Private Sub Class_Terminate()
On Error Resume Next
Set m_lstEmployees = Nothing
End Sub

