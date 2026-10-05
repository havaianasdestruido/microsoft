VERSION 5.00
Object = "{AEC70996-47A5-11D0-BADA-000000000000}#4.0#0"; "bchart3d.ocx"
Begin VB.Form Form1 
   BackColor       =   &H8000000C&
   Caption         =   "Form1"
   ClientHeight    =   3645
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   8250
   LinkTopic       =   "Form1"
   ScaleHeight     =   3645
   ScaleWidth      =   8250
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox R 
      Height          =   330
      Left            =   7380
      TabIndex        =   1
      Text            =   "85"
      Top             =   465
      Width           =   675
   End
   Begin VB.TextBox E 
      Height          =   300
      Left            =   7380
      TabIndex        =   0
      Text            =   "0"
      Top             =   105
      Width           =   705
   End
   Begin BenChart3D.BenChart BenChart1 
      Height          =   3555
      Left            =   15
      TabIndex        =   2
      Top             =   60
      Width           =   4770
      _ExtentX        =   8414
      _ExtentY        =   6271
      MaxBen          =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub E_Change()
On Error Resume Next
BenChart1.Chart.Plot.View3d.Elevation = CInt(E)
End Sub

Private Sub Form_Load()
BenChart1.PostLabel = "Foo"
BenChart1.PreLabel = "Foo"
'BenChart1.BGColor = vbWindowBackground
BenChart1.MaxBen = 100
BenChart1.ReCalc -40, 60
End Sub

Private Sub Form_Resize()
BenChart1.Move 0, 0, ScaleWidth, ScaleHeight
End Sub

Private Sub R_Change()
On Error Resume Next
BenChart1.Chart.Plot.View3d.Rotation = CInt(R)
End Sub
