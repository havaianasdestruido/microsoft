VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "User Login"
   ClientHeight    =   1125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6360
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1125
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "2001"
   Begin VB.PictureBox pctAppIcon 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   120
      Picture         =   "frmLogin.frx":0000
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   5
      Top             =   165
      Width           =   480
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5010
      TabIndex        =   3
      Tag             =   "2005"
      Top             =   615
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   5010
      TabIndex        =   2
      Tag             =   "2004"
      Top             =   135
      Width           =   1215
   End
   Begin VB.TextBox txtUserName 
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Top             =   705
      Width           =   3225
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Welcome to the Corporate Benefits Application. To log in, please enter your User name."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   675
      TabIndex        =   4
      Tag             =   "2002"
      Top             =   120
      Width           =   4350
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblUserName 
      Caption         =   "&User Name:"
      Height          =   255
      Left            =   690
      TabIndex        =   0
      Tag             =   "2003"
      Top             =   735
      Width           =   870
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Visual Studio '97  - Corporate Benefits Sample
'frmLogin.frm       - Copyright (C) 1997 Microsoft Corporation
'                   - All rights reserved.
'Date               - 10/01/97
'Authors            - Simon J. Poile
'Purpose            - Login form for the Corporate Benefit Client Application


Option Explicit

'LoadScreenResStrings
'Purpose     - Load all/any localized string resources
'              for each control on the form.
Sub LoadScreenResStrings()
On Error Resume Next

Dim ctl As Control
Dim obj As Object

'Set the form's caption
Me.Caption = LoadResString(CInt(Me.Tag))

'Set the controls'captions using the Tag property
'for all controls
For Each ctl In Me.Controls
    Err.Clear
    If IsNumeric(ctl.Tag) Then
        If Err = 0 Then
            ctl.Caption = LoadResString(CInt(ctl.Tag))
        End If
    End If
Next
End Sub

Private Sub cmdCancel_Click()
'User cancelled so close the application
Unload Me
End Sub

Private Sub cmdOK_Click()
Dim lValid      As Integer
Dim sErrorText  As String
Dim sErrorTitle As String
      
'Disable the OK & Cancel keys to avoid double-click
cmdOK.Enabled = False
cmdCancel.Enabled = False
  
On Error GoTo errOkay
  
lValid = oAdmin.IsUserValid(txtUserName)

If lValid = icInvalidUser Then
    'An Invalid User name was entered. Warn the user
    sErrorTitle = LoadResString(es.LoginErrorTitle)
    sErrorText = LoadResString(es.LoginErrorText)
    MsgBox sErrorText, vbExclamation, sErrorTitle
    
    'Re-enable OK & Cancel for re-use
    cmdOK.Enabled = True
    cmdCancel.Enabled = True
Else
    'A Valid User name was entered - save the details
    sUser = txtUserName
    iRights = lValid
    
    'Initialize and show the main screen
    frmMain.Init
    frmMain.Show
    
    'Fix for bad redraw in StatusBar
    frmMain.statMain.Panels.Item(1).AutoSize = sbrContents
    frmMain.statMain.Panels.Item(1).AutoSize = sbrSpring
    
    'Now the Main Screen is shown, unload
    Unload Me
End If
Exit Sub

errOkay:
    If Err > vbObjectError Then
        MsgBox Err.Description & vbCrLf & LoadResString(es.ErrorCode) & "[" & CStr(Err - vbObjectError) & "]", vbCritical, LoadResString(es.BenefitsErrorTitle)
    Else
        MsgBox Err.Description & vbCrLf & LoadResString(es.ErrorCode) & "[" & CStr(Err) & "]", vbCritical, LoadResString(es.BenefitsErrorTitle)
    End If
    lValid = icInvalidUser
    Resume Next
End Sub

Private Sub Form_Load()
LoadScreenResStrings
End Sub
