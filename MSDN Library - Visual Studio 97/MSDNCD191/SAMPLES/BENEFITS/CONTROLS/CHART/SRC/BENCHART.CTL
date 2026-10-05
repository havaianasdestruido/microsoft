VERSION 5.00
Begin VB.UserControl Chart 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000C&
   CanGetFocus     =   0   'False
   ClientHeight    =   2535
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3615
   EditAtDesignTime=   -1  'True
   FillColor       =   &H80000005&
   FillStyle       =   0  'Solid
   ForeColor       =   &H80000001&
   ForwardFocus    =   -1  'True
   ScaleHeight     =   2535
   ScaleWidth      =   3615
   Begin VB.Shape shpLine 
      BackColor       =   &H80000008&
      BackStyle       =   1  'Opaque
      Height          =   40
      Left            =   330
      Top             =   2025
      Width           =   3150
   End
   Begin VB.Label lblPost 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Pre"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   570
      Left            =   1995
      TabIndex        =   1
      Top             =   1200
      Width           =   555
   End
   Begin VB.Label lblPre 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Pre"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   570
      Left            =   585
      TabIndex        =   0
      Top             =   1440
      Width           =   555
      WordWrap        =   -1  'True
   End
   Begin VB.Shape shpPost 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H80000002&
      FillStyle       =   0  'Solid
      Height          =   1830
      Left            =   1860
      Top             =   360
      Width           =   840
   End
   Begin VB.Shape shpPre 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H80000001&
      FillStyle       =   0  'Solid
      Height          =   1830
      Left            =   405
      Top             =   315
      Width           =   840
   End
End
Attribute VB_Name = "Chart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'BenChart.ctl - Copyright (C) 1996 Microsoft Corporation
'             - All rights reserved.
'Date         - 12/20/96
'Authors      - Ken Bergmann
'Purpose      - Sample chart control
Option Explicit

Private dMaxBen         As Double
Private dValueOfPre     As Double
Private dValueOfPost    As Double
Private sPreLabel       As String
Private sPostLabel      As String

Const m_def_BackColor = &H80000005

'Draw
'Purpose     - Draw the control
Private Sub Draw()
Dim iChunk  As Integer
Dim iHDiv   As Integer
Dim iHMax   As Integer

UserControl.Cls

lblPre.Visible = False
lblPost.Visible = False
shpPre.Visible = False
shpPost.Visible = False
shpLine.Visible = False

iHMax = (UserControl.Height - 40) \ 2
iHDiv = (UserControl.Height / 4)

shpLine.Left = 0
shpLine.Width = UserControl.Width
shpLine.Top = iHMax

If dValueOfPre >= 0 Then
    lblPre.Top = iHDiv
Else
    lblPre.Top = iHDiv * 3
End If
lblPre.Width = UserControl.Width \ 2
lblPre.Left = 0

If dValueOfPost >= 0 Then
    lblPost.Top = iHDiv
Else
    lblPost.Top = iHDiv * 3
End If
lblPost.Width = lblPre.Width
lblPost.Left = lblPre.Width

iChunk = UserControl.Width \ 5
shpPre.Width = iChunk
shpPost.Width = iChunk

shpPre.Left = iChunk
shpPost.Left = iChunk * 3

lblPre = sPreLabel
lblPost = sPostLabel
lblPre = lblPre & " = "
lblPost = lblPost & " = "
lblPre = lblPre & Format$(CCur(dValueOfPre), "Currency")
lblPost = lblPost & Format$(CCur(dValueOfPost), "Currency")

If dMaxBen > 0 Then
    'Set the height of the pre column
    If dValueOfPre > dMaxBen Then
        shpPre.Height = iHMax
    Else
        shpPre.Height = iHMax * (Abs(dValueOfPre) / dMaxBen)
    End If
    'Set the height of the post column
    If dValueOfPost > dMaxBen Then
        shpPost.Height = iHMax
    Else
        shpPost.Height = iHMax * (Abs(dValueOfPost) / dMaxBen)
    End If
Else
    shpPre.Height = 0
    shpPost.Height = 0
End If
If dValueOfPre >= 0 Then
    shpPre.FillColor = &H0&
    shpPre.Top = (iHMax + 20) - shpPre.Height
Else
    shpPre.FillColor = &HFF&
    shpPre.Top = iHMax + 40
End If
If dValueOfPost >= 0 Then
    shpPost.FillColor = &H0&
    shpPost.Top = (iHMax + 20) - shpPost.Height
Else
    shpPost.FillColor = &HFF&
    shpPost.Top = iHMax + 40
End If

lblPre.Visible = True
lblPost.Visible = True
shpPre.Visible = True
shpPost.Visible = True
shpLine.Visible = True

Refresh
End Sub

'ReCalc
'Purpose     - Recalc chart
'Inputs      - (dPre)   Double
'            - (dPost)  Double
'Outputs     - (Integer)Success
Public Function ReCalc(ByVal dPre As Double, ByVal dPost As Double) As Integer
'Add the value
dValueOfPre = dValueOfPre + dPre
'Add the value
dValueOfPost = dValueOfPost + dPost
Draw
ReCalc = True
End Function

'Reset
'Purpose     - Reset chart
'Outputs     - Integer indicates success
Public Function Reset() As Integer
dValueOfPre = dMaxBen
dValueOfPost = 0
Draw
Reset = True
End Function

Public Property Get BGColor() As OLE_COLOR
BGColor = UserControl.BackColor
End Property

Public Property Let BGColor(ByVal NewBackColor As OLE_COLOR)
UserControl.BackColor = NewBackColor
PropertyChanged "BGColor"
End Property

Public Property Get FGColor() As OLE_COLOR
End Property

Public Property Let FGColor(ByVal NewForeColor As OLE_COLOR)
End Property

Public Property Get MaxBen() As Double
MaxBen = dMaxBen
End Property

Public Property Let MaxBen(ByVal dNewMaxBen As Double)
dMaxBen = dNewMaxBen
PropertyChanged "MaxBen"
End Property

Public Property Get PostLabel() As String
PostLabel = sPostLabel
End Property

Public Property Let PostLabel(ByVal sNewPostLabel As String)
sPostLabel = sNewPostLabel
PropertyChanged "PostLabel"
End Property

Public Property Get PreLabel() As String
PreLabel = sPreLabel
End Property

Public Property Let PreLabel(ByVal sNewPreLabel As String)
sPreLabel = sNewPreLabel
PropertyChanged "PreLabel"
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next
BGColor = PropBag.ReadProperty("BGColor", m_def_BackColor)
MaxBen = PropBag.ReadProperty("MaxBen", 1)
PostLabel = PropBag.ReadProperty("PostLabel", "")
PreLabel = PropBag.ReadProperty("PreLabel", "")
End Sub

Private Sub UserControl_Show()
On Error Resume Next
Draw
Refresh
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
On Error Resume Next
PropBag.WriteProperty "BGColor", UserControl.BackColor, m_def_BackColor
PropBag.WriteProperty "MaxBen", MaxBen, 1
PropBag.WriteProperty "PostLabel", PostLabel, ""
PropBag.WriteProperty "PreLabel", PreLabel, ""
End Sub

