VERSION 5.00
Begin VB.UserControl Spinner 
   ClientHeight    =   870
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1890
   PropertyPages   =   "spinvb.ctx":0000
   ScaleHeight     =   870
   ScaleWidth      =   1890
   Begin VB.TextBox txtValue 
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdDecrement 
      Height          =   300
      Left            =   1320
      Picture         =   "spinvb.ctx":0018
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   420
      Width           =   400
   End
   Begin VB.CommandButton cmdIncrement 
      Height          =   300
      Left            =   1320
      Picture         =   "spinvb.ctx":055A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   400
   End
End
Attribute VB_Name = "Spinner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'SpinVB.ctl  - Copyright (C) 1996 Microsoft Corporation
'            - All rights reserved.
'Date        - 12/20/96
'Authors     - Ken Bergmann, James Braum, Steve Kirk
'Purpose     - SpinVB sample control
Option Explicit

Const m_def_BackColor = &H80000005
Const m_def_ForeColor = &H80000008
Const m_def_Enabled = True
Const m_def_Increment = 0
Const m_def_Maximum = 0
Const m_def_Minimum = 0
Const m_def_Value = 0

Private m_cIncrement As Currency
Private m_cMaximum   As Currency
Private m_cMinimum   As Currency
Private m_cValue     As Currency

Event Click()
Event DblClick()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Public Event OnChange()

Public Property Get BackColor() As OLE_COLOR
BackColor = txtValue.BackColor
End Property

Public Property Let BackColor(ByVal NewBackColor As OLE_COLOR)
txtValue.BackColor = NewBackColor
PropertyChanged "BackColor"
End Property

Public Property Get Enabled() As Boolean
Enabled = txtValue.Enabled
End Property

Public Property Let Enabled(ByVal bNewEnabled As Boolean)
txtValue.Enabled = bNewEnabled
cmdIncrement.Enabled = bNewEnabled
cmdDecrement.Enabled = bNewEnabled
PropertyChanged "Enabled"
End Property

Public Property Get Font() As Font
Set Font = txtValue.Font
End Property

Public Property Set Font(ByVal NewFont As Font)
Set txtValue.Font = NewFont
PropertyChanged "Font"
End Property

Public Property Get ForeColor() As OLE_COLOR
ForeColor = txtValue.ForeColor
End Property

Public Property Let ForeColor(ByVal NewForeColor As OLE_COLOR)
txtValue.ForeColor = NewForeColor
PropertyChanged "ForeColor"
End Property

Public Property Get Increment() As Currency
Increment = m_cIncrement
End Property

Public Property Let Increment(ByVal cNewIncrement As Currency)
m_cIncrement = cNewIncrement
PropertyChanged "Increment"
End Property

Public Property Get Maximum() As Currency
Maximum = m_cMaximum
End Property

Public Property Let Maximum(ByVal cNewMaximum As Currency)
If cNewMaximum >= Minimum Then
    m_cMaximum = cNewMaximum
    PropertyChanged "Maximum"
End If
End Property

Public Property Get Minimum() As Currency
Minimum = m_cMinimum
End Property

Public Property Let Minimum(ByVal cNewMinimum As Currency)
If cNewMinimum <= Maximum Then
    m_cMinimum = cNewMinimum
    PropertyChanged "Minimum"
End If
End Property

Public Property Get Value() As Currency
Value = m_cValue
End Property

Public Property Let Value(ByVal cNewValue As Currency)
If cNewValue <> m_cValue Then
    If ((cNewValue >= m_cMinimum) And (cNewValue <= m_cMaximum)) Or (cNewValue = 0) Then
        m_cValue = cNewValue
    Else
        If cNewValue < m_cMinimum Then
            If m_cMinimum < 0 Then
                m_cValue = 0
            Else
                m_cValue = m_cMinimum
            End If
        Else
            m_cValue = m_cMaximum
        End If
    End If
    txtValue.Text = Format(m_cValue, "Currency")
    PropertyChanged "Value"
    RaiseEvent OnChange
End If
End Property

Private Sub cmdDecrement_Click()
Value = m_cValue - m_cIncrement
End Sub

Private Sub cmdIncrement_Click()
Value = m_cValue + m_cIncrement
End Sub

Private Sub txtValue_Click()
RaiseEvent Click
End Sub

Private Sub txtValue_DblClick()
RaiseEvent DblClick
End Sub

Private Sub txtValue_KeyDown(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub txtValue_KeyPress(KeyAscii As Integer)
RaiseEvent KeyPress(KeyAscii)
On Error Resume Next
Select Case KeyAscii
    Case vbKey0 To vbKey9, vbKeyBack, vbKeyDecimal, Asc("."), Asc("-"):
        'valid
    Case vbKeyUp
        KeyAscii = 0
        cmdIncrement_Click
    Case vbKeyDown
        KeyAscii = 0
        cmdDecrement_Click
    Case vbKeyReturn:
        KeyAscii = 0
        Value = CCur(txtValue.Text)
    Case Else
        KeyAscii = 0
        Beep
End Select
End Sub

Private Sub txtValue_KeyUp(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub txtValue_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
RaiseEvent MouseDown(Button, Shift, x, Y)
End Sub

Private Sub txtValue_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
RaiseEvent MouseMove(Button, Shift, x, Y)
End Sub

Private Sub txtValue_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
RaiseEvent MouseUp(Button, Shift, x, Y)
End Sub

Private Sub UserControl_Initialize()
txtValue.Text = Format(m_cValue, "Currency")
End Sub

Private Sub UserControl_InitProperties()
Maximum = m_def_Maximum
Minimum = m_def_Minimum
Increment = m_def_Increment
Value = m_def_Value
BackColor = m_def_BackColor
ForeColor = m_def_ForeColor
Enabled = m_def_Enabled
Set Font = Ambient.Font
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
Enabled = PropBag.ReadProperty("Enabled", True)
Set Font = PropBag.ReadProperty("Font", Ambient.Font)
Maximum = PropBag.ReadProperty("Maximum", m_def_Maximum)
Minimum = PropBag.ReadProperty("Minimum", m_def_Minimum)
Increment = PropBag.ReadProperty("Increment", m_def_Increment)
Value = PropBag.ReadProperty("Value", m_def_Value)
End Sub

Private Sub UserControl_Resize()
txtValue.Move 0, 0, ScaleWidth * 0.7, ScaleHeight
cmdIncrement.Move txtValue.Width, 0, ScaleWidth * 0.3, ScaleHeight / 2
cmdDecrement.Move txtValue.Width, cmdIncrement.Height, ScaleWidth * 0.3, ScaleHeight / 2
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
PropBag.WriteProperty "BackColor", txtValue.BackColor, m_def_BackColor
PropBag.WriteProperty "ForeColor", txtValue.ForeColor, m_def_ForeColor
PropBag.WriteProperty "Enabled", txtValue.Enabled, True
Call PropBag.WriteProperty("Font", Font)
Call PropBag.WriteProperty("Maximum", m_cMaximum, m_def_Maximum)
Call PropBag.WriteProperty("Minimum", m_cMinimum, m_def_Minimum)
Call PropBag.WriteProperty("Increment", m_cIncrement, m_def_Increment)
Call PropBag.WriteProperty("Value", m_cValue, m_def_Value)
End Sub

