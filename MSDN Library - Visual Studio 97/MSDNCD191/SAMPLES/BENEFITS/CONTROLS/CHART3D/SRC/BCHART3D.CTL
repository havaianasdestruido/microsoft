VERSION 5.00
Object = "{02B5E320-7292-11CF-93D5-0020AF99504A}#1.0#0"; "MSCHART.OCX"
Begin VB.UserControl BenChart 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   4500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   4500
   ScaleWidth      =   4800
   Begin MSChartLib.MSChart MSChart1 
      Height          =   2490
      Left            =   315
      OleObjectBlob   =   "bchart3d.ctx":0000
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1740
      Width           =   4080
   End
   Begin VB.Label lblMain 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pre Label"
      Height          =   1560
      Left            =   360
      TabIndex        =   1
      Top             =   -120
      Width           =   675
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "BenChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'bchart3d.ctl - Copyright (C) 1996 Microsoft Corporation
'             - All rights reserved.
'Date         - 12/20/96
'Authors      - Ken Bergmann
'Purpose      - Benefits Chart3D sample control
Option Explicit

Private m_dMaxBen         As Double
Private m_dValueOfPre     As Double
Private m_dValueOfPost    As Double
Private m_sPreLabel       As String
Private m_sPostLabel      As String

Dim lVal(1 To 2)          As Long

'Draw
'Purpose     - Draw control
Private Sub Draw()
If m_dMaxBen > 0 Then
    'Set the height of the pre column
    If m_dValueOfPre > m_dMaxBen Then
        lVal(1) = m_dMaxBen
    Else
        lVal(1) = m_dValueOfPre
    End If
    'Set the height of the post column
    If m_dValueOfPost > m_dMaxBen Then
        lVal(2) = m_dMaxBen
    Else
        lVal(2) = m_dValueOfPost
    End If
Else
    lVal(1) = 0
    lVal(2) = 0
End If

lblMain = m_sPreLabel & " = " & Format$(CCur(m_dValueOfPre), "Currency")
lblMain = lblMain & " ; " & m_sPostLabel & " = " & Format$(CCur(m_dValueOfPost), "Currency")
MSChart1.ChartData = lVal
MSChart1.Column = 1
MSChart1.ColumnLabel = m_sPreLabel & " = " & Format$(CCur(m_dValueOfPre), "Currency")
MSChart1.Column = 2
MSChart1.ColumnLabel = m_sPostLabel & " = " & Format$(CCur(m_dValueOfPost), "Currency")
MSChart1.RowLabel = ""
MSChart1.Plot.View3d.Elevation = 0
MSChart1.Plot.View3d.Rotation = 72
MSChart1.Plot.PlotBase.BaseHeight = 1
MSChart1.Plot.PlotBase.Brush.Style = VtBrushStyleNull
MSChart1.Plot.PlotBase.Pen.Style = VtPenStyleNull
MSChart1.Plot.LocationRect.Min.Set 0, 0
MSChart1.Plot.LocationRect.Max.Set ScaleWidth, ScaleHeight
MSChart1.Refresh
End Sub

'ReCalc
'Purpose     - Recalc chart
'Inputs      - (dPre)   Double
'            - (dPost)  Double
'Outputs     - (Integer)Success
Public Function ReCalc(ByVal dPre As Double, ByVal dPost As Double) As Integer
m_dValueOfPre = m_dValueOfPre + dPre
m_dValueOfPost = m_dValueOfPost + dPost
Draw
ReCalc = True
End Function

'Reset
'Purpose     - Reset chart
'Outputs     - (Integer) Success
Public Function Reset() As Integer
m_dValueOfPre = m_dMaxBen
m_dValueOfPost = 0
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

Public Property Get Chart() As Object
Set Chart = MSChart1
End Property

Public Property Get MaxBen() As Double
MaxBen = m_dMaxBen
End Property

Public Property Let MaxBen(ByVal dNewMaxBen As Double)
m_dMaxBen = dNewMaxBen
PropertyChanged "MaxBen"
End Property

Public Property Get PostLabel() As String
PostLabel = m_sPostLabel
End Property

Public Property Let PostLabel(ByVal sNewPostLabel As String)
m_sPostLabel = sNewPostLabel
PropertyChanged "PostLabel"
End Property

Public Property Get PreLabel() As String
PreLabel = m_sPreLabel
End Property

Public Property Let PreLabel(ByVal sNewPreLabel As String)
m_sPreLabel = sNewPreLabel
PropertyChanged "PreLabel"
End Property


Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next
BGColor = PropBag.ReadProperty("BGColor", &HFFFFFF)
MaxBen = PropBag.ReadProperty("MaxBen", 1)
PostLabel = PropBag.ReadProperty("PostLabel", "")
PreLabel = PropBag.ReadProperty("PreLabel", "")
End Sub

Private Sub UserControl_Resize()
lblMain.Move 0, 0, ScaleWidth
MSChart1.Move 0, lblMain.Height, ScaleWidth, ScaleHeight - lblMain.Height
End Sub

Private Sub UserControl_Show()
Draw
Refresh
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
On Error Resume Next
PropBag.WriteProperty "BGColor", UserControl.BackColor, &HFFFFFF
PropBag.WriteProperty "MaxBen", MaxBen, 1
PropBag.WriteProperty "PostLabel", PostLabel, ""
PropBag.WriteProperty "PreLabel", PreLabel, ""
End Sub

