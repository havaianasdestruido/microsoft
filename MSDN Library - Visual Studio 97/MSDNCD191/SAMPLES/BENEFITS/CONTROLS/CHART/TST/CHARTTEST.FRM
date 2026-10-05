VERSION 5.00
Object = "{E8D74294-3386-11D0-BADA-000000000000}#1.0#0"; "BENCHART.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H80000005&
   Caption         =   "Chart Control Test"
   ClientHeight    =   5475
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   5250
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   ScaleHeight     =   5475
   ScaleWidth      =   5250
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btSmallDU 
      Caption         =   "Small - D/U"
      Height          =   465
      Left            =   2085
      TabIndex        =   4
      Top             =   4905
      Width           =   1125
   End
   Begin VB.CommandButton btSmallUD 
      Caption         =   "Small - U/D"
      Height          =   465
      Left            =   915
      TabIndex        =   3
      Top             =   4905
      Width           =   1125
   End
   Begin VB.CommandButton btReset 
      Caption         =   "Reset"
      Height          =   975
      Left            =   3270
      TabIndex        =   2
      Top             =   4395
      Width           =   1125
   End
   Begin VB.CommandButton btBigDU 
      Caption         =   "Big - D/U"
      Height          =   465
      Left            =   2085
      TabIndex        =   1
      Top             =   4380
      Width           =   1125
   End
   Begin VB.CommandButton btBigUD 
      Caption         =   "Big - U/D"
      Height          =   465
      Left            =   915
      TabIndex        =   0
      Top             =   4380
      Width           =   1125
   End
   Begin BenefitChart.Chart chartMain 
      Height          =   2535
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   4471
      BGColor         =   -2147483636
      MaxBen          =   0
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'ChartTest.frm  - Copyright (C) 1996 Microsoft Corporation
'               - All rights reserved.
'Date           - 12/18/96
'Authors        - Ken Bergmann, James Braum, Steve Kirk
'Purpose        - Test for Chart control

Option Explicit

Private Sub btBigDU_Click()
chartMain.Width = 5040
chartMain.Height = 4035
chartMain.ReCalc -20, 20
End Sub

Private Sub btBigUD_Click()
chartMain.Width = 5040
chartMain.Height = 4035
chartMain.ReCalc 20, -20
End Sub

Private Sub btReset_Click()
chartMain.Reset
End Sub

Private Sub btSmallDU_Click()
chartMain.Width = 3630
chartMain.Height = 2655
chartMain.ReCalc -20, 20
End Sub

Private Sub btSmallUD_Click()
chartMain.Width = 3630
chartMain.Height = 2655
chartMain.ReCalc 20, -20
End Sub

Private Sub Form_Load()
chartMain.Width = 5040
chartMain.Height = 4035
chartMain.MaxBen = 300
chartMain.prelabel = "Pre-Tax"
chartMain.postlabel = "Post-Tax"
chartMain.Reset
End Sub
