VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CListDependant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'CListDependants.cls - Copyright (C) 1996 Microsoft Corporation
'                    - All rights reserved.
'Date                - 12/20/96
'Authors             - Ken Bergmann, James Braum, Steve Kirk
'Purpose             - CListDependant list implementation for benefits servers

Option Explicit
Private m_lstDependants As CList
Private m_iIdNo         As Integer

'Add
'Purpose     - Add to list
'Inputs      - (CDependant)    oDependant
'Outputs     - (CDependant)
Public Function Add(oDependant As CDependant) As CDependant
m_lstDependants.Add "ID=" & Trim$(CStr(oDependant.PKId)), oDependant
Set Add = oDependant
End Function

'Remove
'Purpose     - Remove from the list
'Inputs      - (Long)   lIndexKey
Public Sub Remove(lIndexKey As Long)
m_lstDependants.Remove "ID=" & Trim$(CStr(lIndexKey))
End Sub

Public Property Get Count() As Long
Count = m_lstDependants.Count
End Property

Public Property Get Item(lIndexKey As Long) As CDependant
Attribute Item.VB_UserMemId = 0
Set Item = m_lstDependants("ID=" & Trim$(CStr(lIndexKey)))
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set NewEnum = m_lstDependants.NewEnum
End Property

Private Sub Class_Initialize()
On Error Resume Next
Set m_lstDependants = New CList
End Sub

Private Sub Class_Terminate()
On Error Resume Next
Set m_lstDependants = Nothing
End Sub

