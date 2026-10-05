VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FolderItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'CFolderItem.cls
'Date        - 10/24/96
'Author      - Phil Johnson, MG Components
'Purpose     - FolderItem class used by VB SideBar Control code

Public Left As Long                         'FolderItem Location Member Fields
Attribute Left.VB_VarMemberFlags = "40"
Public Top As Long
Attribute Top.VB_VarMemberFlags = "40"
Public Right As Long
Attribute Right.VB_VarMemberFlags = "40"
Public Bottom As Long
Attribute Bottom.VB_VarMemberFlags = "40"

Private IconTop As Long                     'Location of Icon Portion
Private IconLeft As Long

Private BorderLeft As Long                  'Location of Border Portion
Private BorderTop As Long
Private BorderRight As Long
Private BorderBottom As Long

Public Canvas As PictureBox                 'Reference to Controls picture box
Attribute Canvas.VB_VarMemberFlags = "40"
Public Hidden As Boolean                    'FolderItems visible state
Public Parent As Folder                     'The parent Folder object
Attribute Parent.VB_VarMemberFlags = "40"
Public HiLite As Boolean                    'FolderItems hilighted state
Attribute HiLite.VB_VarMemberFlags = "40"

Private m_Key As Variant                    'Key
Private m_Caption As Variant                'Caption
Private m_Visible As Boolean                'Visibility
Private m_IconKey As Variant                'Icon Key
Public Sub SetLocation(Left As Long, Top As Long, Right As Long, Bottom As Long)
Attribute SetLocation.VB_MemberFlags = "40"

    'Calculations on where to draw each item
    Me.Left = Left
    Me.Top = Top
    Me.Right = Right
    Me.Bottom = Bottom
    
    'Border Size, the border is three pixels high
    BorderTop = Top + 2
    BorderLeft = ((Right - GlbUserControl.ImageWidth) \ 2) - 3
    BorderRight = (BorderLeft + GlbUserControl.ImageWidth) + 6
    BorderBottom = BorderTop + GlbUserControl.ImageHeight + 6
    
    'Check for ImageList - If the imagelist is not set we have no Icons
    'therefore we do not have to calculate a Top and Left Offset
    If Not GlbUserControl.CTLImageList Is Nothing Then
        IconTop = Canvas.ScaleY(Top + 5, vbPixels, GlbUserControl.ContainerScaleMode)
        IconLeft = Canvas.ScaleX(BorderLeft + 3, vbPixels, GlbUserControl.ContainerScaleMode)
    End If
      
End Sub
Public Sub Init(Caption As Variant, Key As Variant, Icon As Variant, Parent As Folder, Canvas As Object)
Attribute Init.VB_MemberFlags = "40"
    'Initialize this FolderItem
    m_Caption = Caption
    m_Key = Key
    m_IconKey = Icon
    Set Me.Canvas = Canvas
    Set Me.Parent = Parent
End Sub
Public Sub Draw(Optional Outline As Boolean)
Attribute Draw.VB_MemberFlags = "40"
    'Draw this folderItem
    DrawIcon Outline
    DrawCaption
End Sub
Public Sub DrawIcon(Optional DrawBox As Boolean)
Attribute DrawIcon.VB_MemberFlags = "40"

On Error GoTo DrawError

If Visible And Not Hidden Then
    'Wipe the existing graphic
    Canvas.Line (BorderLeft, BorderTop)-(BorderRight, BorderBottom), GlbUserControl.BackColor, BF
    'Draw the Icon
    GlbUserControl.CTLImageList.ListImages(m_IconKey).Draw Canvas.hDC, IconLeft, IconTop, 1
    'Draw the box?
    If DrawBox Then
        HiLite = True
        Canvas.Line (BorderLeft, BorderBottom)-(BorderLeft, BorderTop), vb3DLight
        Canvas.Line -(BorderRight, BorderTop), vb3DLight
        Canvas.Line -(BorderRight, BorderBottom), vb3DDKShadow
        Canvas.Line -(BorderLeft - 1, BorderBottom), vb3DDKShadow
    Else
        HiLite = False
    End If

    'We may have to redraw the top Scroll Button
    GlbUserControl.ScrollButtons("Up").Draw
    
End If

DrawError:
    Select Case Err.Number
    Case 380
        'Not enough space for Icon, will be truncated on right hand side
        IconLeft = 0
        Resume
    Case 35600, 35601
        Err.Raise Number:=40000 + 2918, Description:="No Image '" & m_IconKey & "' in " & GlbUserControl.CTLImageList.Name
        Exit Sub
    End Select

End Sub
Public Sub DrawCaption()
Attribute DrawCaption.VB_MemberFlags = "40"

If Visible And Not Hidden Then
    
    Dim rRect As RECT, Ret As Long
    
    With rRect
        .Top = BorderBottom + 2
        .Bottom = Me.Bottom
        .Left = Me.Left
        .Right = Me.Right
        'Wipe the previous text
        Canvas.Line (.Left, .Top)-(.Right, .Bottom), GlbUserControl.BackColor, BF
    End With
    
    'Draw the text
    Canvas.ForeColor = GlbUserControl.ForeColor
    Ret = DrawText(Canvas.hDC, Caption, Len(Caption), rRect, DTCCenter)
    Canvas.ForeColor = vbBlack
    
    'Bottom Scroll Button Overwritten?
    With GlbUserControl.ScrollButtons("Down")
        If PtInRect(rRect, .Left, .Top) Then
            .Draw
        End If
    End With

End If

End Sub
'Caption Property
Public Property Get Caption() As Variant
    Caption = m_Caption
End Property
Public Property Let Caption(ByVal vNewValue As Variant)
    m_Caption = vNewValue
    If Parent Is GlbUserControl.ActiveFolder Then DrawCaption
End Property
'Key Property
Public Property Get Key() As Variant
    Key = m_Key
End Property
Public Property Let Key(ByVal vNewValue As Variant)
    m_Key = vNewValue
End Property
'Visible Property
Public Property Get Visible() As Variant
    Visible = m_Visible
End Property
Public Property Let Visible(ByVal New_Visible As Variant)
    
    Dim PrevVisible As Boolean
    PrevVisible = m_Visible
    m_Visible = New_Visible
    
    If Not Parent Is Nothing And Parent Is GlbUserControl.ActiveFolder Then
        If New_Visible <> PrevVisible Then Parent.DrawFolderItems
    End If
    
End Property
'Icon Property
Public Property Get Icon() As Variant
    Icon = m_IconKey
End Property
Public Property Let Icon(ByVal vNewValue As Variant)
    m_IconKey = vNewValue
    If Parent Is GlbUserControl.ActiveFolder Then DrawIcon
End Property
Private Sub Class_Initialize()
    m_Visible = True
End Sub
Private Sub Class_Terminate()
    'Class is terminating so, free contained objects
    Set Canvas = Nothing
    Set Parent = Nothing
End Sub
