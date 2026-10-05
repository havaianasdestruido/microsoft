VERSION 5.00
Begin VB.UserControl SideBar 
   Alignable       =   -1  'True
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   1230
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2745
   ClipControls    =   0   'False
   PaletteMode     =   0  'Halftone
   PropertyPages   =   "SBarVB.ctx":0000
   ScaleHeight     =   1230
   ScaleWidth      =   2745
   ToolboxBitmap   =   "SBarVB.ctx":0014
   Begin VB.PictureBox picBorder 
      Align           =   1  'Align Top
      Height          =   735
      Left            =   0
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   179
      TabIndex        =   0
      Top             =   0
      Width           =   2745
   End
End
Attribute VB_Name = "SideBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Visual Basic SideBar Control"
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'VBSideBar.ctl
'Date        - 10/25/96
'Author      - Phil Johnson, MG Components
'Purpose     - VB SideBar Control code

'The Collections
Public ScrollButtons As New Collection  'Hidden
Attribute ScrollButtons.VB_VarMemberFlags = "440"
Public Folders As New Collection        'Public
Attribute Folders.VB_VarMemberFlags = "400"

Enum m_BorderStyle
    [None] = 0
    [Fixed Single] = 1
End Enum

Enum m_ImageStyle
    [Small Icons] = 0
    [Large Icons] = 1
End Enum

Enum m_HiliteMode
    [No Highlight] = 0
    [On Click] = 1
    [On MouseMove] = 2
End Enum

'Property Variables
Dim m_BackColor As Long
Dim m_ForeColor As Long
Dim m_BorderStyle As Integer
Dim m_Enabled As Boolean
Dim m_MousePointer As Integer
Dim m_ActiveFolder As Folder
Dim m_ImageListName As String
Dim m_FoldersVisible As Integer
Dim m_FolderHeight As Long
Dim m_Style As Integer
Dim m_SmallImageList As String
Dim m_LargeImageList As String
Dim m_Hilight As Integer
Dim m_CTLImageList As Control
Dim m_ImageHeight As Long
Dim m_ImageWidth As Long
Dim m_ContainerScaleMode As Long

'Internal Variables
Dim BorderHeight As Long
Dim BorderWidth As Long
Dim LastHiLite As FolderItem

'Event Declarations
Event FolderClick(Folder As Folder)
Event FolderItemClick(FolderItem As FolderItem)
'ParentControls
Public Property Get ParentControls() As ParentControls
  Set ParentControls = UserControl.ParentControls
End Property
Private Sub picBorder_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Only process these events in run mode
If Ambient.UserMode Then

    'Scroll button?
    Dim ScrollButton As ScrollButton
    For Each ScrollButton In ScrollButtons
        If Y >= ScrollButton.Top And Y <= ScrollButton.Bottom And ScrollButton.Visible Then
            'Up
            If ScrollButton.Name = "Up" Then
                ActiveFolder.TopIcon = ActiveFolder.TopIcon - 1
            'Down
            Else
                ActiveFolder.TopIcon = ActiveFolder.TopIcon + 1
            End If
            'Draw the items
            Draw
            'Leave
            Exit Sub
        End If
    Next
    
    'Folder?
    Dim Folder As Folder
    For Each Folder In Folders
        If Folder.Top <= Y And Folder.Bottom >= Y And Folder.Visible Then
            'Wipe the previous hilite on folder change
            If Not LastHiLite Is Nothing Then
                LastHiLite.DrawIcon
                Set LastHiLite = Nothing
            End If
            'Why are you click on me more than once? - Don't Redraw
            If Not ActiveFolder Is Folder Then
                Set ActiveFolder = Folder
                Draw
            End If
            'Raise the event
            RaiseEvent FolderClick(Folder)
            'Was there a default folder item?
            If Not Folder.DefaultItem Is Nothing Then
                RaiseEvent FolderItemClick(Folder.DefaultItem)
            End If
            'Leave
            Exit Sub
        End If
    Next
                
    'Folder Item?
    If Not ActiveFolder Is Nothing Then
        Dim FolderItem As FolderItem
        For Each FolderItem In ActiveFolder.FolderItems
            If FolderItem.Top <= Y And FolderItem.Bottom >= Y And FolderItem.Visible And Not FolderItem.Hidden Then
                'Wipe the previous Hilite
                If Highlight = 1 And Not LastHiLite Is Nothing Then
                    LastHiLite.DrawIcon
                End If
                'Draw the new Hilite
                If Highlight = 1 Then
                    Set LastHiLite = FolderItem
                    LastHiLite.DrawIcon DrawBox:=True
                End If
                'Raise the event
                Set ActiveFolder.ActiveFolderItem = FolderItem
                RaiseEvent FolderItemClick(FolderItem)
                'Leave
                Exit Sub
            End If
        Next
    End If
End If
    
End Sub
Private Sub picBorder_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Only process these events in run mode
If Ambient.UserMode And Highlight = 2 Then
    'Folder Item?
    Dim FolderItem As FolderItem
    If Not ActiveFolder Is Nothing Then
        For Each FolderItem In ActiveFolder.FolderItems
        'Should the control Hilite what's under the cursor?
        If FolderItem.Top <= Y And FolderItem.Bottom >= Y And Not FolderItem.Hidden And FolderItem.Visible Then
            'Check if it already is...
            If Not FolderItem.HiLite Then
                If Not LastHiLite Is Nothing Then
                    LastHiLite.DrawIcon
                End If
                Set LastHiLite = FolderItem
                FolderItem.DrawIcon DrawBox:=True
            End If
            'Leave
            Exit Sub
        End If
        Next
    End If
End If

End Sub
Private Sub picBorder_Paint()

    Draw

End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)

    If PropertyName = "DisplayName" Then
        Draw
    End If
       
End Sub

Private Sub UserControl_Initialize()

    'Creates two instances of the CSCrollButton Class
    'Assigns the names Up and Down to them
    
    Dim NewScroll As ScrollButton
    Set NewScroll = New ScrollButton
    NewScroll.Init "Up", picBorder
    ScrollButtons.Add Item:=NewScroll, Key:="Up"

    Set NewScroll = New ScrollButton
    NewScroll.Init "Down", picBorder
    ScrollButtons.Add Item:=NewScroll, Key:="Down"
    
End Sub

Private Sub UserControl_Resize()
    
    'Size and paint the control
    With picBorder
        .Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
        BorderHeight = .ScaleHeight
        BorderWidth = .ScaleWidth
        Draw
    End With
    
End Sub
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute BackColor.VB_UserMemId = -501
    BackColor = m_BackColor
End Property
Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
    
    'Redraw the items
    If Ambient.UserMode Then
        If Not ActiveFolder Is Nothing Then ActiveFolder.DrawFolderItems
    Else
        DrawDesTime
    End If
    
End Property
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute ForeColor.VB_UserMemId = -513
    ForeColor = m_ForeColor
End Property
Public Property Let ForeColor(ByVal New_BackColor As OLE_COLOR)
    m_ForeColor = New_BackColor
        
    'Redraw the items
    If Ambient.UserMode Then
        If Not ActiveFolder Is Nothing Then ActiveFolder.DrawFolderItems
    End If
    
    PropertyChanged "ForeColor"
    
End Property
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_ProcData.VB_Invoke_Property = ";Behavior"
    Enabled = m_Enabled
End Property
Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    picBorder.Enabled = m_Enabled
    PropertyChanged "Enabled"
End Property
Public Property Get hDC() As Long
Attribute hDC.VB_MemberFlags = "400"
    hDC = picBorder.hDC
End Property
Public Property Get hWnd() As Long
Attribute hWnd.VB_UserMemId = -515
Attribute hWnd.VB_MemberFlags = "400"
    hWnd = picBorder.hWnd
End Property
Public Property Get MousePointer() As Integer
Attribute MousePointer.VB_ProcData.VB_Invoke_Property = ";Appearance"
    MousePointer = m_MousePointer
    picBorder.MousePointer = MousePointer
End Property
Public Property Let MousePointer(ByVal New_MousePointer As Integer)
    m_MousePointer = New_MousePointer
    PropertyChanged "MousePointer"
End Property
Public Sub Refresh()
    'Draw the control
    Draw
End Sub
Public Property Get ActiveFolder() As Folder
Attribute ActiveFolder.VB_MemberFlags = "400"
    Set ActiveFolder = m_ActiveFolder
End Property
Public Property Set ActiveFolder(ByVal New_ActiveFolder As Folder)
    Set m_ActiveFolder = New_ActiveFolder
End Property
Public Sub AddFolder(FolderCaption As Variant, FolderKey As String)

On Error GoTo AddFolderHandler

    'Create and initialise a new folder item
    Set GlbUserControl = Me
    Dim NewFolder As New Folder
    NewFolder.Init Caption:=FolderCaption, Key:=FolderKey, Parent:=Me, Canvas:=picBorder
    
    'Add to the Folders collection
    Folders.Add Item:=NewFolder, Key:=FolderKey
    If Folders.Count = 1 Then Set ActiveFolder = NewFolder
    FoldersVisible = FoldersVisible + 1
    
    'Redraw
    Draw
    
AddFolderHandler:
    If Err.Number = 457 Then
        Err.Raise Number:=40000 + 2914, Description:="There is already a Folder with '" & FolderKey & "' as its key"
        Exit Sub
    End If

End Sub
'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    Set Font = Ambient.Font
    Align = 3
    Highlight = 2
    FolderHeight = 18
    BorderStyle = 1
    MousePointer = 0
    BackColor = vbHighlight
    ForeColor = vbHighlightText
    Enabled = True
    SmallImageList = ""
    LargeImageList = ""
    ImageList = ""
    Style = 1
End Sub
'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Set Font = PropBag.ReadProperty("Font", Ambient.Font)
    Align = PropBag.ReadProperty("Align", 3)
    BackColor = PropBag.ReadProperty("BackColor", vbHighlight)
    ForeColor = PropBag.ReadProperty("ForeColor", vbHighlightText)
    BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    FolderHeight = PropBag.ReadProperty("FolderHeight", 18)
    Enabled = PropBag.ReadProperty("Enabled", True)
    MousePointer = PropBag.ReadProperty("MousePointer", 0)
    SmallImageList = PropBag.ReadProperty("SmallImageList", "")
    LargeImageList = PropBag.ReadProperty("LargeImageList", "")
    Style = PropBag.ReadProperty("Style", 1)
    Highlight = PropBag.ReadProperty("Highlight", 2)
End Sub
Private Sub UserControl_Terminate()
    Set ScrollButtons = Nothing
    Set Folders = Nothing
End Sub
'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    
    On Error Resume Next
    
    Call PropBag.WriteProperty("Font", Font, Ambient.Font)
    Call PropBag.WriteProperty("Align", Align, 3)
    Call PropBag.WriteProperty("BackColor", BackColor, vbHighlight)
    Call PropBag.WriteProperty("ForeColor", ForeColor, vbHighlightText)
    Call PropBag.WriteProperty("BorderStyle", BorderStyle, 1)
    Call PropBag.WriteProperty("Enabled", Enabled, True)
    Call PropBag.WriteProperty("MousePointer", MousePointer, 0)
    Call PropBag.WriteProperty("SmallImageList", SmallImageList, "")
    Call PropBag.WriteProperty("LargeImageList", LargeImageList, "")
    Call PropBag.WriteProperty("FolderHeight", FolderHeight, 18)
    Call PropBag.WriteProperty("Style", Style, 1)
    Call PropBag.WriteProperty("Highlight", Highlight, 2)

End Sub
Friend Sub Draw()
    
On Error GoTo DrawError
    
If Ambient.UserMode Then
    
    'This is where we tell the things to draw
    Dim Flder As Folder, _
        rRect As RECT, _
        DrawCount As Integer, _
        WorkArea As Long
    
    'Have I an ImageList?
    If CTLImageList Is Nothing Then CheckImageListState
    
    'Calculate how much space there is to work with
    WorkArea = BorderHeight - (FoldersVisible * m_FolderHeight)
    rRect.Right = BorderWidth
    
    For Each Flder In Folders
        DrawCount = DrawCount + 1
        'Maths
        If Flder.Visible Then
            rRect.Top = rRect.Bottom
            rRect.Bottom = rRect.Top + m_FolderHeight
            'Position it
            Flder.SetLocation Left:=(rRect.Left), Top:=(rRect.Top), Right:=(rRect.Right), Bottom:=(rRect.Bottom), WorkingHeight:=(WorkArea)
            Flder.Draw
            'Was it the current item?
            If Flder Is ActiveFolder Then
                ScrollButtons("Up").SetLocation Left:=(rRect.Right - 14), Top:=(rRect.Bottom + 2), Right:=(rRect.Right - 2), Bottom:=(rRect.Bottom + 14)
                rRect.Bottom = rRect.Bottom + WorkArea
                ScrollButtons("Down").SetLocation Left:=(rRect.Right - 14), Top:=(rRect.Bottom - 14), Right:=(rRect.Right - 2), Bottom:=(rRect.Bottom - 2)
                Flder.DrawFolderItems
            End If
        End If
    Next
Else
    DrawDesTime
End If

DrawError:
    Select Case Err.Number
        Case 398
            'Client Site Not Available
            Exit Sub
    End Select

End Sub
Private Property Get ImageList() As Variant
Attribute ImageList.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute ImageList.VB_MemberFlags = "440"
    ImageList = m_ImageListName
End Property
Private Property Let ImageList(ByVal New_ImageListName As Variant)
    m_ImageListName = New_ImageListName
    
    If Ambient.UserMode Then
        If Not CTLImageList Is Nothing Then
            Set m_CTLImageList = Nothing
            Draw
        End If
    End If
End Property
Public Property Get FoldersVisible() As Variant
Attribute FoldersVisible.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute FoldersVisible.VB_MemberFlags = "400"
    FoldersVisible = m_FoldersVisible
End Property
Public Property Let FoldersVisible(ByVal vNewValue As Variant)
    m_FoldersVisible = vNewValue
End Property
Public Property Get BorderStyle() As m_BorderStyle
Attribute BorderStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute BorderStyle.VB_UserMemId = -504
    BorderStyle = picBorder.BorderStyle
End Property
Public Property Let BorderStyle(ByVal vNewValue As m_BorderStyle)
    picBorder.BorderStyle = vNewValue
End Property
Friend Sub CheckImageListState()

    If CTLImageList Is Nothing Then
        If Len(ImageList) Then
            Dim iLoopCounter As Long
            For iLoopCounter = 0 To ParentControls.Count - 1
                If ParentControls(iLoopCounter).Name = ImageList Then
                    Set m_CTLImageList = ParentControls(iLoopCounter)
                    m_ImageHeight = CTLImageList.ImageHeight
                    m_ImageWidth = CTLImageList.ImageWidth
                    m_ContainerScaleMode = UserControl.Extender.Container.ScaleMode
                    Exit For
                End If
            Next
        End If
    End If

End Sub
Public Property Get Font() As Font
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute Font.VB_UserMemId = -512
    Set Font = picBorder.Font
End Property
Public Property Set Font(ByVal New_Font As Font)
    If Not New_Font Is Nothing Then
        Set picBorder.Font = New_Font
        PropertyChanged "Font"
        Draw
    End If
End Property
Public Property Get FolderHeight() As Variant
Attribute FolderHeight.VB_ProcData.VB_Invoke_Property = ";Appearance"
    FolderHeight = m_FolderHeight
End Property
Public Property Let FolderHeight(ByVal vNewValue As Variant)
    If vNewValue > -1 Then
        m_FolderHeight = vNewValue
        Draw
    End If
    PropertyChanged "FolderHeight"
End Property
Friend Function ImageHeight() As Long
    ImageHeight = m_ImageHeight
End Function
Friend Function ImageWidth() As Long
    ImageWidth = m_ImageWidth
End Function
Friend Function CTLImageList() As Control
    If Not m_CTLImageList Is Nothing Then
        Set CTLImageList = m_CTLImageList
    End If
End Function
'Is the control displaying Small or Large Icons
Public Property Get Style() As m_ImageStyle
Attribute Style.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Style = m_Style
End Property
Public Property Let Style(ByVal vNewValue As m_ImageStyle)
    m_Style = vNewValue
    'Change the ImageList
    If m_Style = 1 Then
        ImageList = LargeImageList
    Else
        ImageList = SmallImageList
    End If
    'Update the properties window
    PropertyChanged "Style"
End Property
'The imagelist for small icons
Public Property Get SmallImageList() As Variant
Attribute SmallImageList.VB_ProcData.VB_Invoke_Property = "ImageListPage;Appearance"
    SmallImageList = m_SmallImageList
End Property
Public Property Let SmallImageList(ByVal vNewValue As Variant)
    m_SmallImageList = vNewValue
    PropertyChanged "SmallImageList"
End Property
'The image list for large icons
Public Property Get LargeImageList() As Variant
Attribute LargeImageList.VB_ProcData.VB_Invoke_Property = "ImageListPage;Appearance"
    LargeImageList = m_LargeImageList
End Property
Public Property Let LargeImageList(ByVal vNewValue As Variant)
    m_LargeImageList = vNewValue
    PropertyChanged "LargeImageList"
End Property
Friend Function FolderItemHeight() As Long
    FolderItemHeight = ImageHeight + 38
End Function
Public Sub Aboutbox()
Attribute Aboutbox.VB_UserMemId = -552
Attribute Aboutbox.VB_MemberFlags = "40"

    MsgBox "VB SideBar Control  v" & App.Major & "." & App.Minor & "." & _
           App.Revision & vbCrLf & "By Phil Johnson 1996-97", vbOKOnly Or vbInformation, _
           "MG Components Present"

End Sub
Friend Sub DrawDesTime()
    
    'Draws the Folder object onto the Canvas
    Dim Ret As Long, rRect As RECT
    'Initialise the Rectangle
    With rRect
        .Right = picBorder.ScaleWidth
        .Bottom = FolderHeight
        'Draw it (square first then text)
        picBorder.Line (.Left, .Top)-(.Right, .Bottom), vbButtonFace, BF
        Ret = DrawEdge(picBorder.hDC, rRect, EDGE_RAISED, BF_RECT)
        Ret = DrawText(picBorder.hDC, Ambient.DisplayName, Len(Ambient.DisplayName), rRect, DTCCenter)
        picBorder.Line (.Left, .Bottom)-(.Right, picBorder.ScaleHeight), BackColor, BF
    End With
    
End Sub
Public Property Get Highlight() As m_HiliteMode
Attribute Highlight.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Highlight = m_Hilight
End Property
Public Property Let Highlight(ByVal vNewValue As m_HiliteMode)
    m_Hilight = vNewValue
    PropertyChanged "Highlight"
End Property
Friend Function ContainerScaleMode() As Long
    ContainerScaleMode = m_ContainerScaleMode
End Function

