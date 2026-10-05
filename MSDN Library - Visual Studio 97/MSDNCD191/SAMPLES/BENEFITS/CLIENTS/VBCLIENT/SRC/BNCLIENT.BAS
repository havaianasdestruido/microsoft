Attribute VB_Name = "basBenefits"
'Visual Studio '97  - Corporate Benefits Sample
'bnclient.bas       - Copyright (C) 1997 Microsoft Corporation
'                   - All rights reserved.
'Date               - 10/01/97
'Authors            - Simon J. Poile
'Purpose            - Constant, functions, etc...


'Used to hold information about the
'currently logged on user
Public iRights   As UserType
Public sUser     As String

Public cRemTypes As Collection
Public cLnkTypes As Collection

'Server Type Flags
Public Const stTotalServers = 8
Public Const stPassthrough = 0
Public Const stCPPDLL = 1
Public Const stCPPEXE = 2
Public Const stVBDLL = 3
Public Const stVBEXE = 4
Public Const stATLDLL = 5
Public Const stATLEXE = 6
Public Const stMFCDLL = 7
Public Const stMFCEXE = 8

'Local Server Flag
'Based on Compilation flag
'See Project | Benefits Properties
#If defLOCALSERVER = 1 Then
Public Const stLocal = stCPPDLL
#ElseIf defLOCALSERVER = 2 Then
Public Const stLocal = stCPPEXE
#ElseIf defLOCALSERVER = 4 Then
Public Const stLocal = stVBEXE
#ElseIf defLOCALSERVER = 5 Then
Public Const stLocal = stATLDLL
#ElseIf defLOCALSERVER = 6 Then
Public Const stLocal = stATLEXE
#ElseIf defLOCALSERVER = 7 Then
Public Const stLocal = stMFCDLL
#ElseIf defLOCALSERVER = 8 Then
Public Const stLocal = stMFCEXE
#Else
Public Const stLocal = stVBDLL
#End If

'Public settings flags

Public bUseMSGraph          As Boolean
Public sMSGraphRegistryKey  As String

Type Settings
    bChanged As Boolean
    sType    As String
    sDB      As String
    lSrc     As SourceServer
    lRem     As RemoteServer
    lLnk     As LinkServer
End Type
Public tSet As Settings

'Public Reference to the Benefits Admin Server
Public oAdmin As CAdmin

'Consts for usage of Server
Public sGraphPID

'Declarations of used Win32 API Types and Functions

Type RECT
    Left    As Long
    Top     As Long
    Right   As Long
    Bottom  As Long
End Type

Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long

Type OPENFILENAME
    lStructSize         As Long
    hwndOwner           As Long
    hInstance           As Long
    lpstrFilter         As String
    lpstrCustomFilter   As String
    nMaxCustFilter      As Long
    nFilterIndex        As Long
    lpstrFile           As String
    nMaxFile            As Long
    lpstrFileTitle      As String
    nMaxFileTitle       As Long
    lpstrInitialDir     As String
    lpstrTitle          As String
    flags               As Long
    nFileOffset         As Integer
    nFileExtension      As Integer
    lpstrDefExt         As String
    lCustData           As Long
    lpfnHook            As Long
    lpTemplateName      As String
End Type

Declare Function GetOpenFileName Lib "comdlg32.dll" Alias _
      "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Declare Function ReleaseCapture Lib "user32" () As Long

Public Const SW_SHOWNORMAL = 1

Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
  (ByVal hwnd As Long, ByVal lpOperation As Long, ByVal lpFile As String, _
   ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Const SQL_ERROR              As Long = -1
Public Const SQL_INVALID_HANDLE     As Long = -2
Public Const SQL_NO_DATA_FOUND      As Long = 100
Public Const SQL_SUCCESS            As Long = 0
Public Const SQL_SUCCESS_WITH_INFO  As Long = 1

Public Const SQL_DRIVER_NOPROMPT = 0
Public Const SQL_DRIVER_COMPLETE = 1
Public Const SQL_DRIVER_PROMPT = 2
Public Const SQL_DRIVER_COMPLETE_REQUIRED = 3

Declare Function SQLDriverConnect Lib "odbc32.dll" (ByVal hdbc&, ByVal hwnd As Long, ByVal szCSIn$, ByVal cbCSIn%, ByVal szCSOut$, ByVal cbCSMax%, cbCSOut%, ByVal fDrvrComp%) As Integer
Declare Function SQLAllocConnect Lib "odbc32.dll" (ByVal henv&, phdbc&) As Integer
Declare Function SQLAllocEnv Lib "odbc32.dll" (phenv&) As Integer
Declare Function SQLFreeConnect Lib "odbc32.dll" (ByVal hdbc&) As Integer
Declare Function SQLFreeEnv Lib "odbc32.dll" (ByVal henv&) As Integer

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const KEY_QUERY_VALUE = &H1

Public Const ERROR_NONE = 0

Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias _
  "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, _
ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As _
  Long) As Long
Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

'Enum for String Localization
Public Enum eUpdate
    emClearFields = 1
    emFieldsToData = 2
    emDataToFields = 3
End Enum

'Enum for String Localization
Public Enum es
    AppName = 1000
    MSGraphProgID = 1001
    BenGlossaryAddress = 1002
    
    LoginErrorTitle = 1003
    LoginErrorText = 1004
    DBInitFailed = 1005
    
    Statistics = 1006
    Admin = 1007
    DataAdmin = 1008
    
    Queries = 1009
    Results = 1010
    Benefits = 1011
    Employees = 1012
    Dependants = 1013
    Banks = 1014
    DataServices = 1015
    
    BenGlossaryNotFound = 1017
    
    ChartErrorTitle = 1018
    ChartErrorText = 1019
    
    BenefitsErrorTitle = 1020
    DBBenTypesFailed = 1021
    DBCovTypesFailed = 1022
    DBQueryTaxTypesFailed = 1023
    DBDepRelFailed = 1024
    DBBenTaxTypesFailed = 1025
    
    AboutTitle = 1026
    AboutText = 1027
    
    NoBenefitErrorText = 1028
    DbFailAddBenefit = 1029
    DbFailEditBenefit = 1030
    DeleteBenefit = 1031
    DbFailedDeleteBenefit = 1032
    
    NoBankErrorText = 1033
    DbFailAddBank = 1034
    DbFailEditBank = 1035
    DeleteBank = 1036
    DbFailedDeleteBank = 1037
    
    NoEmplErrorText = 1038
    DbFailAddEmployee = 1039
    DbFailEditEmployee = 1040
    DeleteEmployee = 1041
    DbFailedDeleteEmployee = 1042
    
    NoDepErrorText = 1043
    DbFailAddDependant = 1044
    DbFailEditDependant = 1045
    DeleteDependant = 1046
    DbFailedDeleteDependant = 1047
    
    DbNoEmplsFoundTitle = 1048
    DbNoEmplsFoundText = 1049
    
    AcceptDOChangesTitle = 1050
    AcceptDOChangesText = 1051
    
    InvalidMDBText = 1052
    
    MDBSelectText = 1053
    MDBFileDesc = 1054
    MDBFileExt = 1055
    
    RACMgr = 1056
    
    CFTEXT = 1057
    
    BankInUse = 1058
    
    ErrorCode = 1059
    
    srvPassthrough = 1060
    srvCPPDLL = 1061
    srvCPPEXE = 1062
    srvVBDLL = 1063
    srvVBEXE = 1064
    srvATLDLL = 1065
    srvATLEXE = 1066
    srvMFCDLL = 1067
    srvMFCEXE = 1068
    
    DirDelimiter = 2176
    
    DSVBApp = 2180
    DSLocal = 2181
    DSNetComm = 2182
    DSRemote = 2183
    DSRemoteLink = 2184
    DSDatabase = 2185
    
    DOPASSTHROUGH = 2186
    DOCPPDLL = 2187
    DOCPPEXE = 2188
    DOVBDLL = 2189
    DOVBEXE = 2190
    DOATLDLL = 2191
    DOATLEXE = 2192
    DOMFCDLL = 2193
    DOMFCEXE = 2194
    
    NCRemoteAutomation = 2200
    NCDistributedCOM = 2201
        
    NetComm = 2212
    RemFuncNA = 2215
    RemFuncNoSrvAvail = 2216
        
    qryAllBenefits = 2220
    qryAllBenefitTypes = 2221
    qryIgnoreCoverage = 2222
    qryPaymentOfTax = 2223
    
    emplBankUnknown = 2224
        
    InitSetup = 2226
        
    CancelSetupTitle = 2230
    CancelSetupText = 2231

    qryEmplsPerCat = 2240
    qryEmplsPerBen = 2241
    qryEmplsOneCat = 2242
    qryEmplsOneBen = 2243
    qryResults = 2244
End Enum
Dim iRemSvr As Integer
Dim iLnkSvr As Integer

'Main
'Purpose     - Application startup routine
Sub Main()
On Error Resume Next
Set cRemTypes = New Collection
cRemTypes.Add icRemCPP, CStr(stCPPEXE)
cRemTypes.Add icRemVB, CStr(stVBEXE)
cRemTypes.Add icRemATL, CStr(stATLEXE)
cRemTypes.Add icRemMFC, CStr(stMFCEXE)
Set cLnkTypes = New Collection
cLnkTypes.Add icLnkCPP, CStr(stCPPEXE)
cLnkTypes.Add icLnkVB, CStr(stVBEXE)
cLnkTypes.Add icLnkATL, CStr(stATLEXE)
cLnkTypes.Add icLnkMFC, CStr(stMFCEXE)

'Load the registry settings
If Init() Then
    If StartAdmin() Then
        frmLogin.Show
    Else
        Unload frmLogin
        Unload frmMain
    End If
Else
    Unload frmLogin
    Unload frmMain
End If
End Sub

'GetDSN
'Purpose     - Get Valid ODBC DSN
'Inputs      - (String)     DSN
'            - (String)     UID
'            - (String)     PWD
'            - (String)     DATABASE
'Outputs     - (Boolean)    Success
Public Function GetDSN(ByRef DSN As String, ByRef UID As String, ByRef PWD As String, ByRef DATABASE As String) As Boolean
Dim iLoop           As Integer
Dim rc              As Integer
Dim henv            As Long
Dim hdbc            As Long
Dim Connect         As String
Dim ConnectOut      As String * 1024
Dim ConnectOutLen   As Integer

GetDSN = False

ConnectOutLen = 0
ConnectOut = String(1024, 0)  'Initialize string.
    
'Allocate the environment & connection handles.
rc = SQLAllocEnv(henv)
rc = SQLAllocConnect(ByVal henv, hdbc)
    
'Connect using the "Connect string" technique.
rc = SQLDriverConnect(hdbc, Screen.ActiveForm.hwnd, Connect, Len(Connect), _
ConnectOut, 1024, ConnectOutLen, SQL_DRIVER_COMPLETE)
  
If rc = SQL_SUCCESS_WITH_INFO Then
    'User selected a DSN
    'Now parse it
     tsSettingsType = "ODBC"
     For iLoop = 1 To Len(ConnectOut)
        If Mid$(ConnectOut, iLoop, 4) = "DSN=" Then
            DSN = ""
            iLoop = iLoop + 4
            While ((Mid$(ConnectOut, iLoop, 1) <> ";") And (iLoop < Len(ConnectOut)) And (Asc(Mid$(ConnectOut, iLoop, 1)) > 31))
                DSN = DSN & Mid$(ConnectOut, iLoop, 1)
                iLoop = iLoop + 1
            Wend
        End If
        If Mid$(ConnectOut, iLoop, 4) = "UID=" Then
            UID = ""
            iLoop = iLoop + 4
            While ((Mid$(ConnectOut, iLoop, 1) <> ";") And (iLoop < Len(ConnectOut)) And (Asc(Mid$(ConnectOut, iLoop, 1)) > 31))
                UID = UID & Mid$(ConnectOut, iLoop, 1)
                iLoop = iLoop + 1
            Wend
        End If
        If Mid$(ConnectOut, iLoop, 4) = "PWD=" Then
            PWD = ""
            iLoop = iLoop + 4
            While ((Mid$(ConnectOut, iLoop, 1) <> ";") And (iLoop < Len(ConnectOut)) And (Asc(Mid$(ConnectOut, iLoop, 1)) > 31))
                PWD = PWD & Mid$(ConnectOut, iLoop, 1)
                iLoop = iLoop + 1
            Wend
        End If
        If Mid$(ConnectOut, iLoop, 9) = "DATABASE=" Then
            DATABASE = ""
            iLoop = iLoop + 9
            While ((Mid$(ConnectOut, iLoop, 1) <> ";") And (iLoop < Len(ConnectOut)) And (Asc(Mid$(ConnectOut, iLoop, 1)) > 31))
                DATABASE = DATABASE & Mid$(ConnectOut, iLoop, 1)
                iLoop = iLoop + 1
            Wend
        End If
    Next iLoop
    GetDSN = True
End If
    
'Free the environment & connection handles.
rc = SQLFreeConnect(hdbc)
rc = SQLFreeEnv(henv)
End Function

'Init
'Purpose     - Get registry settings
'Outputs     - (Boolean)    Success
Function Init() As Boolean
Dim sTemp As String

'Load Registry Settings
tSet.sType = Trim$(GetSetting("Corporate Benefits Sample", "Configuration", "Type"))
tSet.sDB = Trim$(GetSetting("Corporate Benefits Sample", "Configuration", "DataSource"))
sTemp = GetSetting("Corporate Benefits Sample", "Configuration", "SrvLocal")
If CBool(Len(sTemp)) Then tSet.lSrc = CLng(sTemp)
sTemp = GetSetting("Corporate Benefits Sample", "Configuration", "SrvRemote")
If CBool(Len(sTemp)) Then tSet.lRem = CLng(sTemp)
sTemp = GetSetting("Corporate Benefits Sample", "Configuration", "SrvRemoteLink")
If CBool(Len(sTemp)) Then tSet.lLnk = CLng(sTemp)

If CBool(Len(tSet.sType)) And CBool(Len(tSet.sDB)) Then
    Init = True
Else
    'Registry not set yet - create some defaults
    'Get correct database information form the user
    Load frmMain
    If frmMain.DoConfigure Then
        'Load the new Registry Settings
        Init = Init()
    Else
        'User cancelled request for settings so exit now.
        Init = False
        Unload frmMain
    End If
End If
End Function

'MsgError
'Purpose     - General Message Handler
'Inputs      - (Integer)    iErrorTitle
'            - (Integer)    iErrorText
Public Sub MsgError(iErrorTitle As Integer, iErrorText As Integer)
Dim sErrorTitle As String
Dim sErrorText As String

On Error Resume Next
sErrorTitle = Trim$(LoadResString(iErrorTitle))
If Len(sErrorTitle) = 0 Then sErrorTitle = "Benefits Sample Error"

sErrorText = Trim$(LoadResString(iErrorText))
If Len(sErrorText) = 0 Then sErrorText = "Unable to load error text. Please check installation."

MsgBox sErrorText, vbExclamation, sErrorTitle
End Sub

'SaveSettings
'Purpose     - Update the registry
Sub SaveSettings()
SaveSetting "Corporate Benefits Sample", "Configuration", "Type", tSet.sType
SaveSetting "Corporate Benefits Sample", "Configuration", "DataSource", tSet.sDB
SaveSetting "Corporate Benefits Sample", "Configuration", "SrvLocal", CStr(tSet.lSrc)
SaveSetting "Corporate Benefits Sample", "Configuration", "SrvRemote", CStr(tSet.lRem)
SaveSetting "Corporate Benefits Sample", "Configuration", "SrvRemoteLink", CStr(tSet.lLnk)
End Sub

'StartAdmin
'Purpose     - Start local ActiveX server
'Outputs     - (Boolean)    Success
Function StartAdmin() As Boolean
Dim bODBC   As Boolean
Dim iSource As Integer
Dim sDS     As String

On Error GoTo errStartAdmin

If tSet.sType = "Access" Then
    'Database is to be Access
    sDS = tSet.sDB
    bODBC = False
Else
    'Database access must be via ODBC
    sDS = "ODBC;DSN=" & tSet.sDB
    bODBC = True
End If

On Error Resume Next
iRemSvr = 0
iLnkSvr = 0
iSource = tSet.lSrc
iRemSvr = cRemTypes(CStr(tSet.lRem))
iLnkSvr = cLnkTypes(CStr(tSet.lLnk))

On Error GoTo errStartAdmin
'Initialize Admin Object
Set oAdmin = New CAdmin
If oAdmin.Init(bODBC, sDS, tSet.lSrc, iRemSvr, iLnkSvr) Then
    StartAdmin = True
Else
    StartAdmin = False
    MsgError es.BenefitsErrorTitle, es.DBInitFailed
    tSet.lLnk = 0
    tSet.lRem = 0
    tSet.lSrc = 0
    tSet.sDB = ""
    tSet.sType = ""
    SaveSettings
End If
Exit Function

errStartAdmin:
    StartAdmin = False
    If Err > vbObjectError Then
        MsgBox Err.Description & vbCrLf & LoadResString(es.ErrorCode) & "[" & CStr(Err - vbObjectError) & "]", vbCritical, LoadResString(es.BenefitsErrorTitle)
    Else
        MsgBox Err.Description & vbCrLf & LoadResString(es.ErrorCode) & "[" & CStr(Err) & "]", vbCritical, LoadResString(es.BenefitsErrorTitle)
    End If
    tSet.lLnk = 0
    tSet.lRem = 0
    tSet.lSrc = 0
    tSet.sDB = ""
    tSet.sType = ""
    SaveSettings
    Exit Function
End Function

