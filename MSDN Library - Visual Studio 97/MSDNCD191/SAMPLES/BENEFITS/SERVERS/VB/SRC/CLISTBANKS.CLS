VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CListBank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'CListBanks.cls - Copyright (C) 1996 Microsoft Corporation
'               - All rights reserved.
'Date           - 12/20/96
'Authors        - Ken Bergmann, James Braum, Steve Kirk
'Purpose        - CListBank list implementation for benefits servers

Option Explicit
Private m_lstBanks  As CList
Private m_iIdNo     As Integer

'Add
'Purpose     - Add to list
'Inputs      - (CBank)  oBank
'Outputs     - (CBank)
Public Function Add(oBank As CBank) As CBank
m_lstBanks.Add "ID=" & Trim$(CStr(oBank.PKId)), oBank
Set Add = oBank
End Function

'Remove
'Purpose     - Remove from list
'Inputs      - (Long)   lIndexKey
Public Sub Remove(lIndexKey As Long)
m_lstBanks.Remove "ID=" & Trim$(CStr(lIndexKey))
End Sub

Public Property Get Count() As Long
Count = m_lstBanks.Count
End Property

Public Property Get Item(lIndexKey As Long) As CBank
Attribute Item.VB_UserMemId = 0
Set Item = m_lstBanks("ID=" & Trim$(CStr(lIndexKey)))
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set NewEnum = m_lstBanks.NewEnum
End Property

Private Sub Class_Initialize()
On Error Resume Next
Set m_lstBanks = New CList
End Sub

Private Sub Class_Terminate()
On Error Resume Next
Set m_lstBanks = Nothing
End Sub

