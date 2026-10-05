VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Folder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'CFolder.cls
'Date        - 10/24/96
'Author      - Phil Johnson, MG Components
'Purpose     - Folder class used by VB SideBar Control code

Option Explicit

Public FolderItems As New Collection     ' A collection (of FolderItem)
Public Parent As SideBar                 ' The User Control
Attribute Parent.VB_VarMemberFlags = "40"
Public Canvas As PictureBox              ' The drawing region
Attribute Canvas.VB_VarMemberFlags = "40"

Public Left As Long                      ' Location of Folder
Attribute Left.VB_VarMemberFlags = "40"
Public Top As Long                       ' Not a property as saves doing
Attribute Top.VB_VarMemberFlags = "40"
Public Right As Long                     ' Hidden Member etc
Attribute Right.VB_VarMemberFlags = "40"
Public Bottom As Long
Attribute Bottom.VB_VarMemberFlags = "40"
 
Public WorkingHeight As Long             ' The 'drawable' area of the control
Attribute WorkingHeight.VB_VarMemberFlags = "40"
Public BottomIcon As Long                ' How many icons can be displayed
Attribute BottomIcon.VB_VarMemberFlags = "40"
Public TopIcon As Long                   ' # of the topmost icon
Attribute TopIcon.VB_VarMemberFlags = "40"

Private m_Key As Variant                 ' Key for this Folder
Private m_Caption As Variant             ' Caption for this Folder
Private m_Visible As Boolean             ' Visible flag
Private m_ActiveFolderItem As FolderItem ' ActiveFolderItem
Private m_DefaultItem As Variant         ' Default FolderItem

Public Sub SetLocation(Left As Long, Top As Long, Right As Long, Bottom As Long, WorkingHeight As Long)
Attribute SetLocation.VB_MemberFlags = "40"

    'Stores the location of the Folder Object
    Me.Left = Left
    Me.Top = Top
    Me.Right = Right
    Me.Bottom = Bottom
    Me.WorkingHeight = WorkingHeight

End Sub

Public Sub Init(Caption As Variant, Key As Variant, Parent As SideBar, Canvas As Object)
Attribute Init.VB_MemberFlags = "40"
    
    'Initialise the properties
    m_Caption = Caption
    Me.Key = Key
    Set Me.Parent = Parent
    Set Me.Canvas = Canvas
    
End Sub

Public Sub Draw()
Attribute Draw.VB_MemberFlags = "40"
    
    'Draws the Folder object onto the Canvas
    Dim Ret As Long, rRect As RECT
    'Initialise the Rectangle
    With rRect
        .Left = Left
        .Top = Top
        .Right = Right
        .Bottom = Bottom
    End With
    
    'Draw it (square first then text)
    Canvas.Line (Left, Top)-(Right, Bottom), vbButtonFace, BF
    Ret = DrawEdge(Canvas.hDC, rRect, EDGE_RAISED, BF_RECT)
    Ret = DrawText(Canvas.hDC, Caption, Len(Caption), rRect, DTCCenter)

End Sub
Public Sub DrawFolderItems()
Attribute DrawFolderItems.VB_MemberFlags = "40"
    
    'Enumerates the FolderItems and decides whether
    'to draw them onto the canvas
    
    Dim Ret As Long, rRect As RECT, _
        FolderItem As FolderItem, iLoopCounter As Long, _
        iVisCount As Long, IconItemHeight As Long
    
    'Decide upon the topmost and bottommost FolderItems
    IconItemHeight = Parent.FolderItemHeight
    BottomIcon = (WorkingHeight \ IconItemHeight) + (TopIcon - 1)
    If BottomIcon > FolderItems.Count Then TopIcon = 1
    
    'Initialise the Rectangle
    With rRect
        .Left = Left
        .Top = Bottom
        .Right = Right
    End With
    
    'Draw the background
    Canvas.Line (Left, Bottom)-(Right, Bottom + WorkingHeight), Parent.BackColor, BF
    
    'Change the text colour
    For iLoopCounter = 1 To FolderItems.Count
        Set FolderItem = FolderItems(iLoopCounter)
        If FolderItem.Visible Then
            iVisCount = iVisCount + 1
            'Can this FolderItem be displayed?
            If iVisCount >= TopIcon And iVisCount <= BottomIcon Then
                'If yes position and the tell it to draw itself
                rRect.Bottom = rRect.Top + IconItemHeight
                FolderItem.Hidden = False
                FolderItem.SetLocation Left:=(rRect.Left), Top:=(rRect.Top), Right:=(rRect.Right), Bottom:=(rRect.Bottom)
                'Do I need to draw a border?
                If Parent.Highlight = 1 And FolderItem.HiLite Then
                    FolderItem.Draw Outline:=True
                Else
                    FolderItem.Draw
                End If
                rRect.Top = rRect.Bottom
            Else
                'No - set the hidden property
                FolderItem.Hidden = True
            End If
        End If
    Next iLoopCounter

    'Draw the Scroll Buttons (as appropriate)
    With Parent.ScrollButtons
        'The up arrow
        If TopIcon > 1 And BottomIcon <> 0 Then
            .Item("Up").Visible = True
            .Item("Up").Draw
        Else
            .Item("Up").Visible = False
        End If
        'The down arrow
        If BottomIcon < iVisCount And BottomIcon <> 0 Then
            .Item("Down").Visible = True
            .Item("Down").Draw
        Else
            .Item("Down").Visible = False
        End If
    End With 'With Parent.ScrollButtons

End Sub
Public Sub AddItem(ItemCaption As Variant, ItemKey As Variant, ItemIconKey As Variant)

On Error GoTo AddFolderHander

    'Creates a new instance of FolderItem (a Folder) and adds it to the collection
    Dim NewFolderItem As New FolderItem
    NewFolderItem.Init Caption:=ItemCaption, Key:=ItemKey, Icon:=ItemIconKey, Canvas:=Me.Canvas, Parent:=Me
    FolderItems.Add Item:=NewFolderItem, Key:=ItemKey
    If FolderItems.Count = 1 Then Set ActiveFolderItem = NewFolderItem
    
    'Draw Item
    If Me Is GlbUserControl.ActiveFolder Then
        If Not GlbUserControl.CTLImageList Is Nothing _
               Then GlbUserControl.CheckImageListState
        DrawFolderItems
    End If
    
AddFolderHander:
    If Err.Number = 457 Then
        Err.Raise Number:=40000 + 2912, Description:="Folder '" & Key & "' already has a FolderItem with '" & ItemKey & "' as its key"
        Exit Sub
    End If

End Sub
Public Sub RemoveItem(ItemKey As Variant)

On Error GoTo RemoveHandler

    'Removal by numeric
    If IsNumeric(ItemKey) Then
        FolderItems.Remove ItemKey
        If Me Is GlbUserControl.ActiveFolder Then DrawFolderItems
    'Removal by key
    Else
        Dim iLoopCounter As Long
        For iLoopCounter = FolderItems.Count To 1 Step -1
            If FolderItems(iLoopCounter).Key = ItemKey Then
                FolderItems.Remove iLoopCounter
                If Me Is GlbUserControl.ActiveFolder Then DrawFolderItems
                Exit For
            End If
        Next iLoopCounter
    End If
        
    
RemoveHandler:
    If Err.Number = 5 Then
        Err.Raise 40000 + 2916, Description:="FolderItem '" & ItemKey & "' doesn't exist"
        Exit Sub
    End If

End Sub
'Caption Property
Public Property Get Caption() As Variant
    Caption = m_Caption
End Property
Public Property Let Caption(ByVal vNewValue As Variant)
    m_Caption = vNewValue
    If Visible Then Draw
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
    
    'Update the Control's FoldersVisible property
    If Not Parent Is Nothing Then
        If m_Visible = False And New_Visible Then
            Parent.FoldersVisible = Parent.FoldersVisible + 1
        ElseIf m_Visible And New_Visible = False Then
            Parent.FoldersVisible = Parent.FoldersVisible - 1
        End If
    End If
    
    'Store value
    m_Visible = New_Visible
    
    'If the now invisible item was the current folder
    'then we must choose another
    If m_Visible = False Then
        If Not Parent Is Nothing And Parent.ActiveFolder Is Me Then
            Dim Folder As Folder
            For Each Folder In Parent.Folders
                If Not Folder Is Me Then
                    Set Parent.ActiveFolder = Folder
                    Exit For
                End If
            Next
        End If
    End If
    
    'Redraw Control
    If Not Parent Is Nothing Then
        Parent.Draw
    End If
    
End Property
Private Sub Class_Initialize()
    'Set Class initial values
    m_Visible = True
    TopIcon = 1
End Sub
'The currently active item in this folder
Public Property Get ActiveFolderItem() As FolderItem
    Set ActiveFolderItem = m_ActiveFolderItem
End Property
Public Property Set ActiveFolderItem(ByVal New_ActiveFolderItem As FolderItem)
    Set m_ActiveFolderItem = New_ActiveFolderItem
End Property
Private Sub Class_Terminate()
    'Class is terminating so, free contained objects
    Set Canvas = Nothing
    Set Parent = Nothing
    Set FolderItems = Nothing
End Sub
Public Property Get DefaultItem() As Variant
    'A default item for the FolderItem can be optionally set
    
    'This will fire the FolderItemClick event for this
    'item when the parent Folder is selected
    If TypeOf m_DefaultItem Is FolderItem Then
        Set DefaultItem = m_DefaultItem
    Else
        If m_DefaultItem = Empty Then
            Set DefaultItem = Nothing
        Else
            Set DefaultItem = FolderItems(m_DefaultItem)
        End If
    End If
End Property
Public Property Let DefaultItem(ByVal vNewValue As Variant)
    If VarType(vNewValue) = vbObject Then
        If TypeOf vNewValue Is FolderItem Then
            Set m_DefaultItem = vNewValue
        End If
    Else
        m_DefaultItem = vNewValue
    End If
End Property
