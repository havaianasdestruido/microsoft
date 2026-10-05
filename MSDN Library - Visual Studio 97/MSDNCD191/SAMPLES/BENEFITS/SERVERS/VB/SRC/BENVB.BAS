Attribute VB_Name = "basBenefitsVB"
'benvb.bas   - Copyright (C) 1996 Microsoft Corporation
'            - All rights reserved.
'Date        - 12/20/96
'Authors     - Ken Bergmann, James Braum, Steve Kirk
'Purpose     - Utility procedures for Benefits servers

Option Explicit
Global g_iFileNum   As Integer
Global g_iIdCounter As Integer

'Required for ActiveX DLLs
Sub Main()
End Sub

'DoQuotes
'Purpose     - Double quotes where needed for stored procedures
'Inputs      - (String) sData
'Outputs     - (String)
Public Function DoQuotes(ByVal sData As String) As String
If Len(sData) = 0 Then Exit Function
Dim iLast As Integer
Dim sPart As String

'Find first single quote
iLast = InStr(sData, Chr$(39))
While iLast
    'Build the string from the left, include two single quotes
    sPart = sPart & Left$(sData, iLast - 1) & Chr$(39) & Chr$(39)
    'Truncate the working string from the right
    sData = Right$(sData, Len(sData) - iLast)
    'Find next single quote in the remainder
    iLast = InStr(sData, Chr$(39))
Wend
'Put any remaining string data on the end
sData = sPart & sData
'Return valid string
DoQuotes = Trim$(sData)
End Function
