Attribute VB_Name = "SideBarDecls"
'SideBarDecls.bas
'Date        - 10/24/96
'Author      - Phil Johnson, MG Components
'Purpose     - General Module for utility functions and
'            - declarations used by VB SideBar Control code

' Reference to the User Control
Public GlbUserControl As SideBar
Attribute GlbUserControl.VB_VarMemberFlags = "40"

' PtInRect() API Function

Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long

' Windows Structure

Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

' DrawEdge() API Function

Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, _
    qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Boolean

' Constants for the DrawEdge() API Function

Public Const BDR_RAISEDOUTER = &H1
Public Const BDR_SUNKENOUTER = &H2
Public Const BDR_RAISEDINNER = &H4
Public Const BDR_SUNKENINNER = &H8

Public Const BDR_OUTER = &H3
Public Const BDR_INNER = &HC
Public Const BDR_RAISED = &H5
Public Const BDR_SUNKEN = &HA

Public Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Public Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
Public Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
Public Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)

Public Const BF_LEFT = &H1
Public Const BF_TOP = &H2
Public Const BF_RIGHT = &H4
Public Const BF_BOTTOM = &H8

Public Const BF_TOPLEFT = (BF_TOP Or BF_LEFT)
Public Const BF_TOPRIGHT = (BF_TOP Or BF_RIGHT)
Public Const BF_BOTTOMLEFT = (BF_BOTTOM Or BF_LEFT)
Public Const BF_BOTTOMRIGHT = (BF_BOTTOM Or BF_RIGHT)
Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

Public Const BF_DIAGONAL = &H10
Public Const BF_DIAGONAL_ENDTOPRIGHT = (BF_DIAGONAL Or BF_TOP Or BF_RIGHT)
Public Const BF_DIAGONAL_ENDTOPLEFT = (BF_DIAGONAL Or BF_TOP Or BF_LEFT)
Public Const BF_DIAGONAL_ENDBOTTOMLEFT = (BF_DIAGONAL Or BF_BOTTOM Or BF_LEFT)
Public Const BF_DIAGONAL_ENDBOTTOMRIGHT = (BF_DIAGONAL Or BF_BOTTOM Or BF_RIGHT)

Public Const BF_MIDDLE = &H800        ' Fill in the middle
Public Const BF_SOFT = &H1000         ' For softer buttons
Public Const BF_ADJUST = &H2000       ' Calculate the space left over
Public Const BF_FLAT = &H4000         ' For flat rather than 3D borders
Public Const BF_MONO = &H8000         ' For monochrome borders

' DrawText() API Function

Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, _
    ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long

' Constants for the DrawText() API Function

Public Const DT_NOCLIP = &H100
Public Const DT_CENTER = &H1
Public Const DT_VCENTER = &H4
Public Const DT_SINGLELINE = &H20

Public Const DTCenter As Long = DT_CENTER Or DT_VCENTER Or DT_SINGLELINE Or DT_NOCLIP
Public Const DTCCenter As Long = DT_CENTER Or DT_VCENTER Or DT_SINGLELINE

' DrawFrameControl() API Function

Public Declare Function DrawFrameControl Lib "user32" (ByVal _
    hDC&, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) _
    As Boolean

' Constants for the DrawFrameControl() API Function

Public Const DFC_CAPTION = 1            'Title bar
Public Const DFC_MENU = 2               'Menu
Public Const DFC_SCROLL = 3             'Scroll bar
Public Const DFC_BUTTON = 4             'Standard button
Public Const DFCS_SCROLLUP = &H0        'Up arrow of scroll bar
Public Const DFCS_SCROLLDOWN = &H1      'Down arrow of scroll bar
 
Public Const DFCS_FLAT = &H4000         'Button has a flat border
Public Const DFCS_MONO = &H8000         'Button has a monochrome border

