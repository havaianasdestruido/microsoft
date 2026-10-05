VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CListBenefitLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'CListBenefitLinks.cls - Copyright (C) 1996 Microsoft Corporation
'                      - All rights reserved.
'Date                  - 12/20/96
'Authors               - Ken Bergmann, James Braum, Steve Kirk
'Purpose               - CListBenefitLink list implementation for benefits servers

Option Explicit
Private m_lstBenefitLinks   As CList
Private m_iIdNo             As Integer

'Add
'Purpose     - Add to list
'Inputs      - (CBenefitLink)    oBenefitLink
'Outputs     - (CBenefitLink)
Public Function Add(oBenefitLink As CBenefitLink) As CBenefitLink
m_lstBenefitLinks.Add "ID=" & Trim$(CStr(oBenefitLink.PKId)), oBenefitLink
Set Add = oBenefitLink
End Function

'Add
'Purpose     - Remove from the list
'Inputs      - (Long)    lIndexKey
Public Sub Remove(lIndexKey As Long)
m_lstBenefitLinks.Remove "ID=" & Trim$(CStr(lIndexKey))
End Sub

Public Property Get Count() As Long
Count = m_lstBenefitLinks.Count
End Property

Public Property Get Item(lIndexKey As Long) As CBenefitLink
Attribute Item.VB_UserMemId = 0
Set Item = m_lstBenefitLinks("ID=" & Trim$(CStr(lIndexKey)))
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set NewEnum = m_lstBenefitLinks.NewEnum
End Property

Private Sub Class_Initialize()
On Error Resume Next
Set m_lstBenefitLinks = New CList
End Sub

Private Sub Class_Terminate()
On Error Resume Next
Set m_lstBenefitLinks = Nothing
End Sub

