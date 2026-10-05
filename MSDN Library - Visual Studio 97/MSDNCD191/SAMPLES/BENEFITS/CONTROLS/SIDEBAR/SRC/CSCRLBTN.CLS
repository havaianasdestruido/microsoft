VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ScrollButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'CScrollButton.cls
'Date        - 10/24/96
'Author      - Phil Johnson, MG Components
'Purpose     - ScrollButton class used by VB SideBar Control code

Public Canvas As PictureBox     ' Reference to Controls picture box
Public Name As String           ' Scrollbuttons key
Attribute Name.VB_VarMemberFlags = "40"

Public Left As Long             ' Scrollbutton Location Member Fields
Attribute Left.VB_VarMemberFlags = "40"
Public Top As Long
Attribute Top.VB_VarMemberFlags = "40"
Public Right As Long
Attribute Right.VB_VarMemberFlags = "40"
Public Bottom As Long

Private m_Visible As Boolean
Public Sub Init(Name As String, Canvas As PictureBox)
Attribute Init.VB_MemberFlags = "40"
    'Set Class initial values
    Me.Name = Name
    Set Me.Canvas = Canvas
End Sub
Public Sub SetLocation(Left As Long, Top As Long, Right As Long, Bottom As Long)
Attribute SetLocation.VB_MemberFlags = "40"
    'Locates scrollbutton
    Me.Left = Left
    Me.Top = Top
    Me.Right = Right
    Me.Bottom = Bottom
    Me.Visible = False
End Sub
Public Sub Draw()
Attribute Draw.VB_MemberFlags = "40"

    If Visible Then
        Dim rRect As RECT, _
            lRet As Long
            
        'Draws the Icon...
        With rRect
            .Left = Left
            .Top = Top
            .Right = Right
            .Bottom = Bottom
            
            If Name = "Up" Then
                lRet = DrawFrameControl(Canvas.hDC, rRect, DFC_SCROLL, DFCS_SCROLLUP)
            Else
                lRet = DrawFrameControl(Canvas.hDC, rRect, DFC_SCROLL, DFCS_SCROLLDOWN)
            End If
        End With
    End If
End Sub
' Visible Property
Public Property Get Visible()
Attribute Visible.VB_MemberFlags = "40"
    Visible = m_Visible
End Property
Public Property Let Visible(vNewValue)
    m_Visible = vNewValue
End Property
Private Sub Class_Initialize()
    'Set Class initial values
    m_Visible = False
End Sub
