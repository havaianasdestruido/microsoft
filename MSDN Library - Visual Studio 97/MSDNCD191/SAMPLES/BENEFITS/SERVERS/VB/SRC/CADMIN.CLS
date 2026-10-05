VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'CAdmin.cls  - Copyright (C) 1996 Microsoft Corporation
'            - All rights reserved.
'Date        - 12/20/96
'Authors     - Ken Bergmann, James Braum, Steve Kirk
'Purpose     - Admin operator class implementation for benefits servers

Option Explicit

' Internal object containment lists
Private m_lstBanks     As CListBank
Private m_lstBenefits  As CListBenefit
Private m_lstEmployees As CListEmployee

' Database objects
Private m_dbBenefit    As Database
Private m_wsBenefit    As Workspace

' Remote server reference
Private m_oRemote      As Object

' Object id number for logging information
Private m_iIdNo        As Integer

'Used for temporary storage of error information
Private m_lErrNo       As Long
Private m_sErrDesc     As String

' Database instantiation parameters
Private m_bUseODBC     As Boolean
Private m_iSrcType     As Integer

' Public events raised
#If UseEvents Then
    Public Event BanksChanged()
    Public Event BenefitsChanged()
    Public Event DependantChanged(ByVal lEmployeePKId As Long)
    Public Event EmployeesChanged()
#End If

Public Enum LinkServer
     icLnkVB = 512
     icLnkATL = 1024
     icLnkCPP = 2048
     icLnkMFC = 4096
End Enum

Public Enum RemoteServer
     icRemVB = 16
     icRemATL = 32
     icRemCPP = 64
     icRemMFC = 128
End Enum

Public Enum SourceServer
     icDoDataDirect = 1
     icDoDataRemote = 2
     icDoDataLinked = 4
End Enum

Public Enum UserType
     icInvalidUser = 0
     icValidUser = 1
     icValidAdmin = 2
End Enum

'Delete
'Purpose     - Delete data object from system
'Inputs      - (CBank)        Optional oBank
'              (CBenefit)     Optional oBenefit
'              (CBenefitLink) Optional oBenefitLink
'              (CDependant)   Optional oDependant
'              (CEmployee)    Optional oEmployee
'Outputs     - (Boolean) Success
Public Function Delete(Optional oBank As CBank, Optional oBenefit As CBenefit, Optional oBenefitLink As CBenefitLink, Optional oDependant As CDependant, Optional oEmployee As CEmployee) As Boolean
Dim lstParms     As CList
Dim iParameters  As Integer
Dim sSQL         As String

On Error GoTo errDelete

' Count number of parameters
If Not (oBank Is Nothing) Then iParameters = iParameters + 1
If Not (oBenefit Is Nothing) Then iParameters = iParameters + 1
If Not (oBenefitLink Is Nothing) Then iParameters = iParameters + 1
If Not (oDependant Is Nothing) Then iParameters = iParameters + 1
If Not (oEmployee Is Nothing) Then iParameters = iParameters + 1

' Check for invalid number of parameters
If iParameters <> 1 Then
    Delete = False
    On Error GoTo 0
    Err.Raise vbObjectError + icErrObjectOp, Trim$(App.Title), LoadResString(icErrObjectOp)
    Exit Function
End If

'Deleting a Bank?
If Not (oBank Is Nothing) Then
    ' Only delete if you actually have an object
    If CBool(oBank.PKId) Then
        If m_bUseODBC Then
            sSQL = LoadResString(scExecute) & LoadResString(scDel_Bank) & scSpace & CStr(oBank.PKId)
            If Not ExecBooleanSQL(sSQL) Then
                Delete = False
                Exit Function
            End If
        Else
            ' Create CList with parameter information
            Set lstParms = New CList
            lstParms.Add "@PKId", oBank.PKId
            If Not ExecBooleanMDB(LoadResString(scDel_Bank), lstParms) Then
                Delete = False
                Exit Function
            End If
            Set lstParms = Nothing
        End If
    Else   'No PKId
        On Error GoTo 0
        Err.Raise icErrPKId, Trim$(App.Title) & LoadResString(scDelBank), LoadResString(icErrPKId)
        Exit Function
    End If
    oBank.Term
    Set oBank = Nothing
    FillListBanks
End If

'Deleting a Benefit?
If Not (oBenefit Is Nothing) Then
    ' Only delete if you actually have an object
    If CBool(oBenefit.PKId) Then
        If m_bUseODBC Then
            sSQL = LoadResString(scExecute) & LoadResString(scDel_Benefit) & scSpace & CStr(oBenefit.PKId)
            If Not ExecBooleanSQL(sSQL) Then
                Delete = False
                Exit Function
            End If
         Else
            ' Create CList with parameter information
            Set lstParms = New CList
            lstParms.Add "@PKId", oBenefit.PKId
            If Not ExecBooleanMDB(LoadResString(scDel_Benefit), lstParms) Then
                Delete = False
                Exit Function
            End If
            Set lstParms = Nothing
        End If
    Else   'No PKId
        On Error GoTo 0
        Err.Raise icErrPKId, Trim$(App.Title) & LoadResString(scDelBnf), LoadResString(icErrPKId)
        Exit Function
    End If
    oBenefit.Term
    Set oBenefit = Nothing
    FillListBenefits
End If

'Deleting a BenefitLink?
If Not (oBenefitLink Is Nothing) Then
    ' Only delete if you actually have an object
    If CBool(oBenefitLink.PKId) Then
        If m_bUseODBC Then
            sSQL = LoadResString(scExecute) & LoadResString(scDel_EmplBnf) & scSpace & CStr(oBenefitLink.PKId)
            If Not ExecBooleanSQL(sSQL) Then
                Delete = False
                Exit Function
            End If
        Else
            ' Create CList with parameter information
            Set lstParms = New CList
            lstParms.Add "@PKId", oBenefitLink.PKId
            If Not ExecBooleanMDB(LoadResString(scDel_EmplBnf), lstParms) Then
                Delete = False
                Exit Function
            End If
            Set lstParms = Nothing
        End If
    Else   'No PKId
        On Error GoTo 0
        Err.Raise icErrPKId, Trim$(App.Title) & LoadResString(scDelBnfLnk), LoadResString(icErrPKId)
        Exit Function
    End If
    oBenefitLink.Term
    Set oBenefitLink = Nothing
    FillListEmployees
End If

'Deleting a Dependant?
If Not (oDependant Is Nothing) Then
    ' Only delete if you actually have an object
    If CBool(oDependant.PKId) Then
        If m_bUseODBC Then
            sSQL = LoadResString(scExecute) & LoadResString(scDel_Dependant) & scSpace & CStr(oDependant.PKId)
            If Not ExecBooleanSQL(sSQL) Then
                Delete = False
                Exit Function
            End If
        Else
            'Create CList with parameter information
            Set lstParms = New CList
            lstParms.Add "@PKId", oDependant.PKId
            If Not ExecBooleanMDB(LoadResString(scDel_Dependant), lstParms) Then
                Delete = False
                Exit Function
            End If
            Set lstParms = Nothing
#If UseEvents Then
            RaiseEvent DependantChanged(oDependant.Employee.PKId)
#End If
        End If
    Else   'No PKId
        On Error GoTo 0
        Err.Raise icErrPKId, Trim$(App.Title) & LoadResString(scDelDep), LoadResString(icErrPKId)
        Exit Function
    End If
    oDependant.Term
    Set oDependant = Nothing
End If

'Deleting an Employee?
If Not (oEmployee Is Nothing) Then
    ' Only delete if you actually have an object
    If CBool(oEmployee.PKId) Then
        If m_bUseODBC Then
            sSQL = LoadResString(scExecute) & LoadResString(scDel_Employee) & scSpace & CStr(oEmployee.PKId)
            If Not ExecBooleanSQL(sSQL) Then
                Delete = False
                Exit Function
            End If
        Else
            ' Create CList with parameter information
            Set lstParms = New CList
            lstParms.Add "@PKId", oEmployee.PKId
            If Not ExecBooleanMDB(LoadResString(scDel_Employee), lstParms) Then
                Delete = False
                Exit Function
            End If
            Set lstParms = Nothing
        End If
    Else   'No PKId
        On Error GoTo 0
        Err.Raise icErrPKId, Trim$(App.Title) & LoadResString(scDelEmpl), LoadResString(icErrPKId)
        Exit Function
    End If
    oEmployee.Term
    Set oEmployee = Nothing
    FillListEmployees
End If
Delete = True
Exit Function

errDelete:
    Delete = False
    If Err > vbObjectError Then
        Err.Raise Err, Trim$(App.Title) & LoadResString(scDel), Err.Description
    Else
        Err.Raise Err + vbObjectError, Trim$(App.Title) & LoadResString(scDel), Err.Description
    End If
    Exit Function
    
End Function
'ExecBooleanMDB
'Purpose     - Execute a JET QueryDef
'Assumtions - The lstParameters object is a CList
'Inputs      - (Sting)   sQry
'              (Object)  lstParameters
'Outputs     - (Boolean) Success
Public Function ExecBooleanMDB(ByVal sQry As String, ByVal lstParameters As Object) As Boolean
Dim pParm   As Parameter
Dim qryDef  As QueryDef

On Error Resume Next

If m_iSrcType <> icDoDataDirect Then
    'If acting as remote or link pass request along
    ExecBooleanMDB = m_oRemote.ExecBooleanMDB(sQry, lstParameters)
    If CBool(Err) Then
        m_sErrDesc = Err.Description
        If Err > vbObjectError Then
            m_lErrNo = Err
        Else
            m_lErrNo = Err + vbObjectError
        End If
        On Error GoTo 0
        Err.Raise m_lErrNo, Trim$(App.Title) & LoadResString(scExecBoolMDB), m_sErrDesc
        ExecBooleanMDB = False
        Exit Function
    End If
Else
    m_wsBenefit.BeginTrans
    Set qryDef = m_dbBenefit.QueryDefs(sQry)
    With qryDef
        For Each pParm In .Parameters
            pParm.Value = lstParameters(pParm.Name)
        Next pParm
        Err = False
        .Execute dbFailOnError + dbSeeChanges
        If CBool(Err) Then
            m_wsBenefit.Rollback
            If Errors.Count > 1 Then
                m_lErrNo = Errors(0).Number
                m_sErrDesc = Errors(0).Description
            Else
                m_lErrNo = Err
                m_sErrDesc = Err.Description
            End If
            If Err < vbObjectError Then m_lErrNo = Err + vbObjectError
            On Error GoTo 0
            Err.Raise m_lErrNo, Trim$(App.Title) & LoadResString(scExecBoolMDB), m_sErrDesc
            ExecBooleanMDB = False
            Exit Function
        End If
    End With
    m_wsBenefit.CommitTrans dbForceOSFlush
    ExecBooleanMDB = True
    Set qryDef = Nothing
End If

End Function
'ExecBooleanSQL
'Purpose     - Execute a stored SQL server procedure
'Inputs      - (String)  sSQL
'Outputs     - (Boolean) Success
Public Function ExecBooleanSQL(ByVal sSQL As String) As Boolean

On Error Resume Next

If m_iSrcType <> icDoDataDirect Then
    'If acting as remote or link pass request along
    ExecBooleanSQL = m_oRemote.ExecBooleanSQL(sSQL)
    If CBool(Err) Then
        m_sErrDesc = Err.Description
        If Err > vbObjectError Then
            m_lErrNo = Err
        Else
            m_lErrNo = Err + vbObjectError
        End If
        On Error GoTo 0
        Err.Raise m_lErrNo, Trim$(App.Title) & LoadResString(scExecBoolSQL), m_sErrDesc
        ExecBooleanSQL = False
        Exit Function
    End If
Else
    m_wsBenefit.BeginTrans
    m_dbBenefit.Execute sSQL, dbFailOnError + dbSQLPassThrough
    If CBool(Err) Then
        m_wsBenefit.Rollback
        If Errors.Count > 1 Then
            m_lErrNo = Errors(0).Number
            m_sErrDesc = Errors(0).Description
        Else
            m_lErrNo = Err
            m_sErrDesc = Err.Description
        End If
        If Err < vbObjectError Then m_lErrNo = Err + vbObjectError
        On Error GoTo 0
        Err.Raise m_lErrNo, Trim$(App.Title) & LoadResString(scExecBoolSQL), m_sErrDesc
        ExecBooleanSQL = False
        Exit Function
    End If
    m_wsBenefit.CommitTrans dbForceOSFlush
    ExecBooleanSQL = True
End If

End Function

'ExecFillArray
'Purpose     - Fill an array based on query result
'Assumptions - The optional parameter is only used for internal calls
'Inputs      - (String)     sQry
'            - (String)     sArray()
'            - (Boolean)    Optional bAddColumnName
'Outputs     - (Boolean)    Success
Public Function ExecFillArray(ByVal sQry As String, sArray() As String, Optional ByVal bAddColumnName As Boolean) As Boolean
Dim iCols       As Integer
Dim iCurCol     As Integer
Dim lCurRow     As Long
Dim rsTemp      As Recordset

On Error GoTo errFillArray

If m_iSrcType <> icDoDataDirect Then
    'If acting as remote or link pass request along
    ExecFillArray = m_oRemote.ExecFillArray(sQry, sArray, bAddColumnName)
Else
    If m_bUseODBC Then
        Set rsTemp = m_dbBenefit.OpenRecordset(sQry, dbOpenSnapshot, dbSQLPassThrough)
    Else
        Set rsTemp = m_dbBenefit.OpenRecordset(sQry, dbOpenSnapshot)
    End If
    If Not (rsTemp.BOF And rsTemp.EOF) Then
        rsTemp.MoveLast 0
        iCols = rsTemp.Fields.Count - 1
        lCurRow = rsTemp.RecordCount - 1
        If bAddColumnName Then
            ReDim sArray(iCols, lCurRow + 1) As String
        Else
            ReDim sArray(iCols, lCurRow) As String
        End If
        lCurRow = 0
        rsTemp.MoveFirst
        Do Until rsTemp.EOF
            For iCurCol = 0 To iCols
                If IsNull(rsTemp(iCurCol)) Then
                    sArray(iCurCol, lCurRow) = scBlank
                Else
                    sArray(iCurCol, lCurRow) = Trim$(CStr(rsTemp(iCurCol)))
                End If
                If bAddColumnName Then sArray(iCurCol, 1) = rsTemp(iCurCol).Name
            Next iCurCol
            lCurRow = lCurRow + 1
            rsTemp.MoveNext
        Loop
        ExecFillArray = True
    End If
    rsTemp.Close
End If
Exit Function

errFillArray:
    ExecFillArray = False
    m_sErrDesc = Err.Description
    If Errors.Count > 1 Then
        m_lErrNo = Errors(0).Number
        m_sErrDesc = Errors(0).Description
    Else
        m_lErrNo = Err
        m_sErrDesc = Err.Description
    End If
    If Err < vbObjectError Then m_lErrNo = m_lErrNo + vbObjectError
    Err.Raise m_lErrNo, Trim$(App.Title) & LoadResString(scExecFillArray), m_sErrDesc
    Exit Function

End Function

'ExecFillList
'Purpose     - Fill a list based on query result
'Inputs      - (String)     sQry
'            - (CList)      lstTmp
'Outputs     - (Boolean)    Success
Public Function ExecFillList(ByVal sQry As String, lstTmp As CList) As Boolean
Dim iCols       As Integer
Dim iCurCol     As Integer
Dim iCurRow     As Integer
Dim rsTemp      As Recordset
Dim sTmp()      As String

On Error GoTo errFillList

Set lstTmp = Nothing
Set lstTmp = New CList

If ExecFillArray(sQry, sTmp(), True) Then
    For iCurRow = 0 To UBound(sTmp)
        lstTmp.Add sTmp(iCurRow, 1), sTmp(iCurRow, 0)
    Next iCurRow
    ExecFillList = True
Else
    ExecFillList = False
End If
Exit Function

errFillList:
    ExecFillList = False
    If Err > vbObjectError Then
        m_lErrNo = Err
    Else
        m_lErrNo = Err + vbObjectError
    End If
    Err.Raise m_lErrNo, Trim$(App.Title) & LoadResString(scExecFillList), Err.Description
    Exit Function

End Function

'FillListBanks
'Purpose     - Fill Banks list
'Assumptions - The CAdmin object has been initialized
Public Sub FillListBanks()
Dim oBank       As CBank
Dim iIndex      As Integer
Dim sSQL        As String
Dim sTemp()     As String
Dim vData       As Variant

On Error Resume Next

If m_iSrcType <> icDoDataDirect Then
    'If acting as remote or link pass request along
    Err = False
    vData = m_oRemote.GetListBanks()
    If CBool(Err) Then
        m_lErrNo = Err
        m_sErrDesc = Err.Description
        If Err < vbObjectError Then m_lErrNo = m_lErrNo + vbObjectError
        On Error GoTo 0
        Err.Raise m_lErrNo, Trim$(App.Title) & LoadResString(scFillLstBank), m_sErrDesc
        Exit Sub
    End If
Else
    'If acting as direct data server, get data
    If m_bUseODBC Then
        sSQL = LoadResString(scLvw_Bank_SQL)
    Else
        sSQL = LoadResString(scLvw_Bank_MDB)
    End If
    Err = False
    If ExecFillArray(sSQL, sTemp()) Then vData = CVar(sTemp)
    If CBool(Err) Then
        ' Save error number, source, and description and raise error.
        m_lErrNo = Err
        m_sErrDesc = Err.Description
        If Err < vbObjectError Then m_lErrNo = m_lErrNo + vbObjectError
        On Error GoTo 0
        Err.Raise m_lErrNo, Trim$(App.Title) & LoadResString(scFillLstBank), m_sErrDesc
        Exit Sub
    End If
End If

'Reset the Bank List
If Not (Banks Is Nothing) Then
    For Each oBank In Banks
        oBank.Term
    Next oBank
    Set oBank = Nothing
    Set Banks = Nothing
End If
Set Banks = New CListBank

For iIndex = LBound(vData, 2) To UBound(vData, 2)
    ' add new CBank to CBankList
    Set oBank = New CBank
    Set oBank.Admin = Me
    oBank.PKId = vData(icLvwPKId, iIndex)
    Banks.Add oBank
    Set oBank = Nothing
Next iIndex
#If UseEvents Then
    RaiseEvent BanksChanged
#End If
End Sub

'FillListBenefits
'Purpose     - Fill Benefits list
'Assumptions - The CAdmin object has been initialized
Public Sub FillListBenefits()
Dim oBenefit    As CBenefit
Dim iIndex      As Integer
Dim sSQL        As String
Dim sTemp()     As String
Dim vData       As Variant

On Error Resume Next

If m_iSrcType <> icDoDataDirect Then
    'If acting as remote or link pass request along
    Err = False
    vData = m_oRemote.GetListBenefits()
    If CBool(Err) Then
        m_lErrNo = Err
        m_sErrDesc = Err.Description
        If Err < vbObjectError Then m_lErrNo = m_lErrNo + vbObjectError
        On Error GoTo 0
        Err.Raise m_lErrNo, Trim$(App.Title) & LoadResString(scFillLstBnf), m_sErrDesc
        Exit Sub
    End If
Else
    'If acting as direct data server, get data
    If m_bUseODBC Then
        sSQL = LoadResString(scLvw_Benefit_SQL)
    Else
        sSQL = LoadResString(scLvw_Benefit_MDB)
    End If
    Err = False
    If ExecFillArray(sSQL, sTemp()) Then
        vData = CVar(sTemp)
    End If
    If CBool(Err) Then
        ' Save error number, source, and description and raise error.
        m_lErrNo = Err
        m_sErrDesc = Err.Description
        If Err < vbObjectError Then m_lErrNo = m_lErrNo + vbObjectError
        On Error GoTo 0
        Err.Raise m_lErrNo, Trim$(App.Title) & LoadResString(scFillLstBnf), m_sErrDesc
        Exit Sub
    End If
End If

'Reset the Benefit list
If Not (Benefits Is Nothing) Then
    For Each oBenefit In Benefits
        oBenefit.Term
    Next oBenefit
    Set oBenefit = Nothing
    Set Benefits = Nothing
End If
Set Benefits = New CListBenefit

For iIndex = LBound(vData, 2) To UBound(vData, 2)
    ' add new CBenefit to list
    Set oBenefit = New CBenefit
    Set oBenefit.Admin = Me
    oBenefit.PKId = vData(icLvwPKId, iIndex)
    Benefits.Add oBenefit
    Set oBenefit = Nothing
Next iIndex
#If UseEvents Then
    RaiseEvent BenefitsChanged
#End If
End Sub

'FillListEmployees
'Purpose     - Fill Employees List
'Assumptions - The CAdmin object has been initialized
Public Sub FillListEmployees()
Dim oEmployee   As CEmployee
Dim iIndex      As Integer
Dim sSQL        As String
Dim sTemp()     As String
Dim vData       As Variant

On Error Resume Next

If m_iSrcType <> icDoDataDirect Then
    'If acting as remote or link pass request along
    Err = False
    vData = m_oRemote.GetListEmployees()
    If CBool(Err) Then
        m_lErrNo = Err
        m_sErrDesc = Err.Description
        If Err < vbObjectError Then m_lErrNo = m_lErrNo + vbObjectError
        On Error GoTo 0
        Err.Raise m_lErrNo, Trim$(App.Title) & LoadResString(scFillLstEmpl), m_sErrDesc
        Exit Sub
    End If
Else
    'If acting as direct data server, get data
    If m_bUseODBC Then
        sSQL = LoadResString(scLvw_Employee_SQL)
    Else
        sSQL = LoadResString(scLvw_Employee_MDB)
    End If
    Err = False
    If ExecFillArray(sSQL, sTemp()) Then
        vData = CVar(sTemp)
    End If
    If CBool(Err) Then
        ' Save error number, source, and description and raise error.
        m_lErrNo = Err
        m_sErrDesc = Err.Description
        If Err < vbObjectError Then m_lErrNo = m_lErrNo + vbObjectError
        On Error GoTo 0
        Err.Raise m_lErrNo, Trim$(App.Title) & LoadResString(scFillLstEmpl), m_sErrDesc
        Exit Sub
    End If
End If

'Reset the Employees list
If Not (Employees Is Nothing) Then
    For Each oEmployee In Employees
        oEmployee.Term
    Next oEmployee
    Set oEmployee = Nothing
    Set Employees = Nothing
End If
Set Employees = New CListEmployee

For iIndex = LBound(vData, 2) To UBound(vData, 2)
    ' add new CEmployee to list
    Set oEmployee = New CEmployee
    Set oEmployee.Admin = Me
    oEmployee.PKId = vData(icLvwPKId, iIndex)
    Employees.Add oEmployee
    Set oEmployee = Nothing
Next iIndex
#If UseEvents Then
    RaiseEvent EmployeesChanged
#End If
End Sub

'GetBenefitTaxTypes
'Purpose     - Fill list of tax types
'Inputs      - (CList)      lstOut
'Outputs     - (Boolean)    Success
Public Function GetBenefitTaxTypes(lstOut As CList) As Boolean
Dim sQry    As String

On Error Resume Next

If m_bUseODBC Then
    sQry = LoadResString(scLvw_BnfTax_SQL)
Else
    sQry = LoadResString(scLvw_BnfTax_MDB)
End If
GetBenefitTaxTypes = GetDomain(sQry, lstOut)
If CBool(Err) Then
    m_sErrDesc = Err.Description
    If Err > vbObjectError Then
        m_lErrNo = Err
    Else
        m_lErrNo = Err + vbObjectError
    End If
    On Error GoTo 0
    Err.Raise m_lErrNo, Trim$(App.Title) & LoadResString(scGetBnfTaxTypes), m_sErrDesc
    GetBenefitTaxTypes = False
    Exit Function
End If
End Function

'GetBenefitTypes
'Purpose     - Fill list of benefit types
'Inputs      - (CList)      lstOut
'Outputs     - (Boolean)    Success
Public Function GetBenefitTypes(lstOut As CList) As Boolean
Dim sQry    As String

On Error Resume Next

If m_bUseODBC Then
    sQry = LoadResString(scLvw_BnfType_SQL)
Else
    sQry = LoadResString(scLvw_BnfType_MDB)
End If
GetBenefitTypes = GetDomain(sQry, lstOut)
If CBool(Err) Then
    m_sErrDesc = Err.Description
    If Err > vbObjectError Then
        m_lErrNo = Err
    Else
        m_lErrNo = Err + vbObjectError
    End If
    On Error GoTo 0
    Err.Raise m_lErrNo, Trim$(App.Title) & LoadResString(scGetBnfTypes), m_sErrDesc
    GetBenefitTypes = False
    Exit Function
End If
End Function

'GetDependantTypes
'Purpose     - Fill list of dependant types
'Inputs      - (CList)      lstOut
'Outputs     - (Boolean)    Success
Public Function GetDependantTypes(lstOut As CList) As Boolean
Dim sQry    As String

On Error Resume Next
If m_bUseODBC Then
    sQry = LoadResString(scLvw_DepType_SQL)
Else
    sQry = LoadResString(scLvw_DepType_MDB)
End If
GetDependantTypes = GetDomain(sQry, lstOut)
If CBool(Err) Then
    m_sErrDesc = Err.Description
    If Err > vbObjectError Then
        m_lErrNo = Err
    Else
        m_lErrNo = Err + vbObjectError
    End If
    On Error GoTo 0
    Err.Raise m_lErrNo, Trim$(App.Title) & LoadResString(scGetBnfTypes), m_sErrDesc
    GetDependantTypes = False
    Exit Function
End If
End Function

'GetDomain
'Purpose     - Fill list with domain values
'Inputs      - (String)     sSQL
'Inputs      - (CList)      lstout
'Outputs     - (Boolean)    Success
Private Function GetDomain(sQry As String, lstOut As CList) As Boolean
Dim lstItem     As CList
Dim iIndex      As Integer
Dim sCapCode    As String
Dim sCapDesc    As String
Dim sCapPKId    As String
Dim sData()     As String

On Error GoTo errGetDomain

sCapDesc = LoadResString(scLvwDesc)
sCapPKId = LoadResString(scLvwPKId)
sCapCode = LoadResString(scLvwCode)

If ExecFillArray(sQry, sData) Then
    Set lstOut = New CList
    For iIndex = LBound(sData, 2) To UBound(sData, 2)
        Set lstItem = New CList
        lstItem.Add sCapDesc, sData(icLvwDesc, iIndex)
        lstItem.Add sCapPKId, sData(icLvwPKId, iIndex)
        lstItem.Add sCapCode, sData(icLvwCode, iIndex)
        lstOut.Add "ID=" & sData(icLvwPKId, iIndex), lstItem
        Set lstItem = Nothing
    Next iIndex
    GetDomain = True
End If
Exit Function

errGetDomain:
    ' Save error number, source, and description and raise error.
    m_lErrNo = Err
    m_sErrDesc = Err.Description
    If Err < vbObjectError Then m_lErrNo = m_lErrNo + vbObjectError
    Err.Raise m_lErrNo, Trim$(App.Title) & LoadResString(scGetDomain), m_sErrDesc
    GetDomain = False
    Exit Function

End Function

'GetEmployeeBenefitCoverageTypes
'Purpose     - Fill list with coverage types
'Inputs      - (CList)      lstOut
'Outputs     - (Boolean)    Success
Public Function GetEmployeeBenefitCoverageTypes(lstOut As CList) As Boolean
Dim sQry    As String

On Error Resume Next

If m_bUseODBC Then
    sQry = LoadResString(scLvw_EmplBnfCoverage_SQL)
Else
    sQry = LoadResString(scLvw_EmplBnfCoverage_MDB)
End If
GetEmployeeBenefitCoverageTypes = GetDomain(sQry, lstOut)
If CBool(Err) Then
    m_sErrDesc = Err.Description
    If Err > vbObjectError Then
        m_lErrNo = Err
    Else
        m_lErrNo = Err + vbObjectError
    End If
    On Error GoTo 0
    Err.Raise m_lErrNo, Trim$(App.Title) & LoadResString(scGetEmplBnfCvrgTypes), m_sErrDesc
    GetEmployeeBenefitCoverageTypes = False
    Exit Function
End If
End Function

'GetEmployeeBenefitTaxTypes
'Purpose     - Fill list with benefit tax types
'Inputs      - (CList)      lstOut
'Outputs     - (Boolean)    Success
Public Function GetEmployeeBenefitTaxTypes(lstOut As CList) As Boolean
Dim sQry    As String

On Error Resume Next

If m_bUseODBC Then
    sQry = LoadResString(scLvw_EmplBnfTax_SQL)
Else
    sQry = LoadResString(scLvw_EmplBnfTax_MDB)
End If
GetEmployeeBenefitTaxTypes = GetDomain(sQry, lstOut)
If CBool(Err) Then
    m_sErrDesc = Err.Description
    If Err > vbObjectError Then
        m_lErrNo = Err
    Else
        m_lErrNo = Err + vbObjectError
    End If
    On Error GoTo 0
    Err.Raise m_lErrNo, Trim$(App.Title) & LoadResString(scGetEmplBnfTaxTypes), m_sErrDesc
    GetEmployeeBenefitTaxTypes = False
    Exit Function
End If
End Function

'GetListBanks
'Purpose     - Fill Banks list
'Outputs     - (Variant)
Public Function GetListBanks() As Variant
Dim sData() As String
Dim sSQL    As String

On Error Resume Next

If m_iSrcType <> icDoDataDirect Then
    GetListBanks = m_oRemote.GetListBanks()
    If CBool(Err) Then
        ' Save error number, source, and description and raise error.
        m_lErrNo = Err
        m_sErrDesc = Err.Description
        If Err < vbObjectError Then m_lErrNo = m_lErrNo + vbObjectError
        On Error GoTo 0
        Err.Raise m_lErrNo, Trim$(App.Title) & LoadResString(scGetLstBank), m_sErrDesc
        Exit Function
    End If
Else
    If m_bUseODBC Then
        sSQL = LoadResString(scLvw_Bank_SQL)
    Else
        sSQL = LoadResString(scLvw_Bank_MDB)
    End If
    If ExecFillArray(sSQL, sData) Then
        GetListBanks = CVar(sData)
    End If
    If CBool(Err) Then
        ' Save error number, source, and description and raise error.
        m_lErrNo = Err
        m_sErrDesc = Err.Description
        If Err < vbObjectError Then m_lErrNo = m_lErrNo + vbObjectError
        On Error GoTo 0
        Err.Raise m_lErrNo, Trim$(App.Title) & LoadResString(scGetLstBank), m_sErrDesc
        Exit Function
    End If
End If
End Function

'GetListBenefits
'Purpose     - Fill Benefits List
'Outputs     - (Variant)
Public Function GetListBenefits() As Variant
Dim sData() As String
Dim sSQL    As String

On Error Resume Next

If m_iSrcType <> icDoDataDirect Then
    GetListBenefits = m_oRemote.GetListBenefits()
    If CBool(Err) Then
        ' Save error number, source, and description and raise error.
        m_lErrNo = Err
        m_sErrDesc = Err.Description
        If Err < vbObjectError Then m_lErrNo = m_lErrNo + vbObjectError
        On Error GoTo 0
        Err.Raise m_lErrNo, Trim$(App.Title) & LoadResString(scGetLstBnf), m_sErrDesc
        Exit Function
    End If
Else
    If m_bUseODBC Then
        sSQL = LoadResString(scLvw_Benefit_SQL)
    Else
        sSQL = LoadResString(scLvw_Benefit_MDB)
    End If
    If ExecFillArray(sSQL, sData) Then
        GetListBenefits = CVar(sData)
    End If
    If CBool(Err) Then
        ' Save error number, source, and description and raise error.
        m_lErrNo = Err
        m_sErrDesc = Err.Description
        If Err < vbObjectError Then m_lErrNo = m_lErrNo + vbObjectError
        On Error GoTo 0
        Err.Raise m_lErrNo, Trim$(App.Title) & LoadResString(scGetLstBnf), m_sErrDesc
        Exit Function
    End If
End If
End Function

'GetListEmployees
'Purpose     - Fill Employees list
'Outputs     - Variant
Public Function GetListEmployees() As Variant
Dim bRetVal     As Boolean
Dim sData()     As String
Dim sQry        As String

On Error Resume Next

If m_iSrcType <> icDoDataDirect Then
    GetListEmployees = m_oRemote.GetListEmployees()
    If CBool(Err) Then
        ' Save error number, source, and description and raise error.
        m_lErrNo = Err
        m_sErrDesc = Err.Description
        If Err < vbObjectError Then m_lErrNo = m_lErrNo + vbObjectError
        On Error GoTo 0
        Err.Raise m_lErrNo, Trim$(App.Title) & LoadResString(scGetLstEmpl), m_sErrDesc
        Exit Function
    End If
Else
    If m_bUseODBC Then
        sQry = LoadResString(scLvw_Employee_SQL)
    Else
        sQry = LoadResString(scLvw_Employee_MDB)
    End If
    If ExecFillArray(sQry, sData) Then GetListEmployees = CVar(sData)
    If CBool(Err) Then
        ' Save error number, source, and description and raise error.
        m_lErrNo = Err
        m_sErrDesc = Err.Description
        If Err < vbObjectError Then m_lErrNo = m_lErrNo + vbObjectError
        On Error GoTo 0
        Err.Raise m_lErrNo, Trim$(App.Title) & LoadResString(scGetLstEmpl), m_sErrDesc
        Exit Function
    End If
End If
End Function

'GetNumEmplsPerBenefit
'Purpose     - Fill list of employee benefits matching coverage, TaxID and optional gender template
'Inputs      - (CList)      lstOut
'            - (Long)       lCoverageId
'            - (Long)       lTaxId
'            - (Boolean)    Optional bGender
'Outputs     - (Boolean)    Success
Public Function GetNumEmplsPerBenefit(lstOut As CList, ByVal lCoverageId As Long, ByVal lTaxId As Long, Optional bGender As Boolean) As Boolean
Dim cItem       As CList
Dim bRetVal     As Boolean
Dim iIndex      As Integer
Dim sData()     As String
Dim sQry        As String

On Error GoTo errGetNumEmplsPerBenefit

sQry = LoadResString(scGNEPBSQL1)
If Not IsMissing(bGender) Then
    If CBool(bGender) Then
        sQry = sQry & LoadResString(scGNEPBSQL2)
    Else
        sQry = sQry & LoadResString(scGNEPBSQL3)
    End If
End If
If lTaxId <> icZero Then sQry = sQry & LoadResString(scGNEPBSQL4) & Trim$(CStr(lTaxId))
If lCoverageId <> icZero Then sQry = sQry & LoadResString(scGNEPBSQL5) & Trim$(CStr(lCoverageId))
'Add the group by
sQry = sQry & LoadResString(scGNEPBSQL6)
If ExecFillArray(sQry, sData) Then
    Set lstOut = New CList
    For iIndex = LBound(sData, 2) To UBound(sData, 2)
        Set cItem = New CList
        cItem.Add LoadResString(scBnfName), sData(icLvwDesc, iIndex)
        cItem.Add LoadResString(scEmplCnt), CLng(sData(icLvwPKId, iIndex))
        lstOut.Add "ID=" & CStr(iIndex + 1), cItem
        Set cItem = Nothing
    Next iIndex
    GetNumEmplsPerBenefit = True
End If
Exit Function

errGetNumEmplsPerBenefit:
    GetNumEmplsPerBenefit = False
    m_lErrNo = Err
    If Err < vbObjectError Then m_lErrNo = m_lErrNo + vbObjectError
    Err.Raise m_lErrNo, Trim$(App.Title) & LoadResString(scGetNumEmplsPerBnf), Err.Description
    Exit Function

End Function

'GetNumEmplsPerBenefitType
'Purpose     - Fill list of employee benefit types matching coverage, TaxID and optional gender template
'Inputs      - (CList)      lstOut
'            - (Long)       lCoverageId
'            - (Long)       lTaxId
'            - (Boolean)    Optional bGender
'Outputs     - (Boolean)    Success
Public Function GetNumEmplsPerBenefitType(lstOut As CList, ByVal lCoverageId As Long, ByVal lTaxId As Long, Optional ByVal bGender As Boolean) As Boolean
Dim lstItem     As CList
Dim bRetVal     As Boolean
Dim iIndex      As Integer
Dim sData()     As String
Dim sQry        As String

On Error GoTo errGetNumEmplsPerBenefitType

sQry = LoadResString(scGNEPBTSQL1)
If Not IsMissing(bGender) Then
    If CBool(bGender) Then
        sQry = sQry & LoadResString(scGNEPBSQL2)
    Else
        sQry = sQry & LoadResString(scGNEPBSQL3)
    End If
End If
If lTaxId <> icZero Then sQry = sQry & LoadResString(scGNEPBSQL4) & Trim$(CStr(lTaxId))
If lCoverageId <> icZero Then sQry = sQry & LoadResString(scGNEPBSQL5) & Trim$(CStr(lCoverageId))
'Add the group by
sQry = sQry & LoadResString(scGNEPBTSQL2)
If ExecFillArray(sQry, sData) Then
    Set lstOut = New CList
    For iIndex = LBound(sData, 2) To UBound(sData, 2)
        Set lstItem = New CList
        lstItem.Add LoadResString(scBnfTypeName), sData(icLvwDesc, iIndex)
        lstItem.Add LoadResString(scEmplCnt), CLng(sData(icLvwPKId, iIndex))
        lstOut.Add "ID=" & CStr(iIndex + 1), lstItem
        Set lstItem = Nothing
    Next iIndex
    bRetVal = True
End If
GetNumEmplsPerBenefitType = bRetVal
Exit Function

errGetNumEmplsPerBenefitType:
    GetNumEmplsPerBenefitType = False
    m_lErrNo = Err
    If Err < vbObjectError Then m_lErrNo = m_lErrNo + vbObjectError
    Err.Raise m_lErrNo, Trim$(App.Title) & LoadResString(scGetNumEmplsPerBnfType), Err.Description
    Exit Function
    
End Function

'Init
'Purpose     - Instantiate remote Benefits server if required, open database if SourceType is DataDirect
'Inputs      - (Boolean)    bUseODBC
'            - (String)     sDataSource
'            - (Integer)    iSourceType
'            - (Integer)    iRemoteSvr
'            - (Integer)    iLinkSvr
'Outputs     - (Boolean)    Success
Public Function Init(ByVal bUseODBC As Boolean, ByVal sDataSource As String, ByVal iSourceType As Integer, ByVal iRemoteSvr As Integer, ByVal iLinkSvr As Integer) As Boolean
On Error GoTo errInit
Term
m_bUseODBC = bUseODBC
m_iSrcType = iSourceType

Err = False
Init = True
Select Case iSourceType
    Case icDoDataDirect:
        Set m_wsBenefit = DBEngine.Workspaces(0)
        If m_bUseODBC Then
            Set m_dbBenefit = m_wsBenefit.OpenDatabase(scBlank, False, False, sDataSource)
            If CBool(Err) Then
                Err.Raise Err.Number + vbObjectError, Trim$(App.Title) & LoadResString(icErrOpenDB), _
                    Err.Description
                Init = False
            End If
        Else
            Set m_dbBenefit = m_wsBenefit.OpenDatabase(sDataSource, False, False)
            If CBool(Err) Then
                Err.Raise Err.Number + vbObjectError, Trim$(App.Title) & LoadResString(icErrOpenDB), _
                    Err.Description
                Init = False
            End If
        End If
    Case icDoDataRemote:
        Select Case iRemoteSvr
            Case icRemVB:
                Set m_oRemote = GetObject("", "BenefitsVBRemote.CAdmin")
                If CBool(Err) Then
                    Err.Raise Err.Number + vbObjectError, Trim$(App.Title) & LoadResString(scInit), _
                        Err.Description
                Init = False
                End If
                Init = m_oRemote.Init(m_bUseODBC, sDataSource, icDoDataDirect, 0, 0)
            Case icRemCPP:
                Set m_oRemote = GetObject("", "BenefitsCPPRemote.CAdmin")
                If CBool(Err) Then
                    Err.Raise Err.Number + vbObjectError, Trim$(App.Title) & LoadResString(scInit), Err.Description
                    Init = False
                End If
                Init = m_oRemote.Init(m_bUseODBC, sDataSource, icDoDataDirect, 0, 0)
            Case Else
                Err.Raise icErrSrcType + vbObjectError, Trim$(App.Title) & LoadResString(scInit), _
                    "Invalid Remote Server Type On Initialization"
                Init = False
        End Select
    Case icDoDataLinked:
        Select Case iLinkSvr
            Case icLnkVB:
                Set m_oRemote = GetObject("", "BenefitsVBRemote.CAdmin")
                If CBool(Err) Then
                    Err.Raise Err.Number + vbObjectError, Trim$(App.Title) & LoadResString(scInit), _
                        Err.Description
                    Init = False
                End If
                Init = m_oRemote.Init(m_bUseODBC, sDataSource, icDoDataRemote, iRemoteSvr, 0)
            Case icLnkCPP:
                Set m_oRemote = GetObject("", "BenefitsCPPRemote.CAdmin")
                If CBool(Err) Then
                    Err.Raise Err.Number + vbObjectError, Trim$(App.Title) & LoadResString(scInit), _
                        Err.Description
                    Init = False
                End If
                Init = m_oRemote.Init(m_bUseODBC, sDataSource, icDoDataRemote, iRemoteSvr, 0)
            Case Else
                Err.Raise icErrSrcType + vbObjectError, Trim$(App.Title) & LoadResString(scInit), _
                    "Invalid Remote Link Server Type On Initialization"
                Init = False
        End Select
    Case Else
        Err.Raise icErrSrcType + vbObjectError, Trim$(App.Title) & LoadResString(scInit), _
            LoadResString(icErrSrcType)
        Init = False
End Select

Exit Function
errInit:
    Init = False
    If Err > vbObjectError Then
        Err.Raise Err.Number, Trim$(App.Title) & LoadResString(scInit), Err.Description
    Else
        Err.Raise Err.Number + vbObjectError, Trim$(App.Title) & LoadResString(scInit), Err.Description
    End If
    Exit Function
End Function

'Insert
'Purpose     - Insert data object into system
'Inputs      - (CBank)        Optional oBank
'              (CBenefit)     Optional oBenefit
'              (CBenefitLink) Optional oBenefitLink
'              (CDependant)   Optional oDependant
'              (CEmployee)    Optional oEmployee
'Outputs     - (Boolean) Success
Public Function Insert(Optional oBank As CBank, Optional oBenefit As CBenefit, Optional oBenefitLink As CBenefitLink, Optional oDependant As CDependant, Optional oEmployee As CEmployee) As Boolean
Dim lstParms        As CList
Dim iParameters     As Integer
Dim sQry            As String

On Error GoTo errInsert

' count number of parameters
If Not (oBank Is Nothing) Then iParameters = iParameters + 1
If Not (oBenefit Is Nothing) Then iParameters = iParameters + 1
If Not (oBenefitLink Is Nothing) Then iParameters = iParameters + 1
If Not (oDependant Is Nothing) Then iParameters = iParameters + 1
If Not (oEmployee Is Nothing) Then iParameters = iParameters + 1

If iParameters <> 1 Then
    Insert = False
    On Error GoTo 0
    Err.Raise vbObjectError + icErrObjectOp, Trim$(App.Title), LoadResString(icErrObjectOp)
    Exit Function
End If

'Inserting a Bank?
If Not (oBank Is Nothing) Then
    ' validate oBank object
    With oBank
        If Not CBool(Len(.Code)) Then
            Insert = False
            On Error GoTo 0
            Err.Raise icErrReqFld + vbObjectError, Trim$(App.Title) & LoadResString(scInsBank), LoadResString(icErrReqFld) & "Code"
            Exit Function
        End If
        
        If Not CBool(Len(.Name)) Then
            Insert = False
            On Error GoTo 0
            Err.Raise icErrReqFld + vbObjectError, Trim$(App.Title) & LoadResString(scInsBank), LoadResString(icErrReqFld) & "Name"
            Exit Function
        End If
    End With
    If m_bUseODBC Then
        sQry = LoadResString(scExecute) & LoadResString(scIns_Bank) & scSpace
        With oBank
            sQry = sQry & scSQ & Trim$(Left$(DoQuotes(.Address1), 60)) & scSQ & scSpaceCommaSpace
            sQry = sQry & scSQ & Trim$(Left$(DoQuotes(.Address2), 60)) & scSQ & scSpaceCommaSpace
            sQry = sQry & scSQ & Trim$(Left$(DoQuotes(.City), 50)) & scSQ & scSpaceCommaSpace
            sQry = sQry & scSQ & Trim$(Left$(DoQuotes(.Code), 16)) & scSQ & scSpaceCommaSpace
            sQry = sQry & scSQ & Trim$(Left$(DoQuotes(.Name), 80)) & scSQ
        End With
        If Not ExecBooleanSQL(sQry) Then
            Insert = False
            Exit Function
        End If
    Else
        ' Create CList with parameter information
        Set lstParms = New CList
        lstParms.Add "@Address1", Trim$(Left$(oBank.Address1, 60))
        lstParms.Add "@Address2", Trim$(Left$(oBank.Address2, 60))
        lstParms.Add "@City", Trim$(Left$(oBank.City, 15))
        lstParms.Add "@Code", Trim$(Left$(oBank.Code, 16))
        lstParms.Add "@Name", Trim$(Left$(oBank.Name, 80))
        If Not ExecBooleanMDB(LoadResString(scIns_Bank), lstParms) Then
            Insert = False
            Exit Function
        End If
        Set lstParms = Nothing
    End If
    oBank.Term
    Set oBank = Nothing
    FillListBanks
End If

' Inserting a benefit?
If Not (oBenefit Is Nothing) Then
    ' validate oBenefit object
    With oBenefit
        If Not CBool(.BnfTaxId) Then
            Insert = False
            On Error GoTo 0
            Err.Raise icErrReqFld + vbObjectError, Trim$(App.Title) & LoadResString(scInsBnf), LoadResString(icErrReqFld) & "BnfTaxId"
            Exit Function
        End If
        If Not CBool(.BnfTypeId) Then
            Insert = False
            On Error GoTo 0
            Err.Raise icErrReqFld + vbObjectError, Trim$(App.Title) & LoadResString(scInsBnf), LoadResString(icErrReqFld) & "BnfTypeId"
            Exit Function
        End If
        If Not CBool(Len(.Description)) Then
            Insert = False
            On Error GoTo 0
            Err.Raise icErrReqFld + vbObjectError, Trim$(App.Title) & LoadResString(scInsBnf), LoadResString(icErrReqFld) & "Description"
            Exit Function
        End If
    End With
    If m_bUseODBC Then
        sQry = LoadResString(scExecute) & LoadResString(scIns_Benefit) & scSpace
        With oBenefit
            sQry = sQry & CStr(.BnfTaxId) & scSpaceCommaSpace
            sQry = sQry & CStr(.BnfTypeId) & scSpaceCommaSpace
            sQry = sQry & CStr(.Cost) & scSpaceCommaSpace
            sQry = sQry & scSQ & Trim$(Left$(DoQuotes(.Description), 80)) & scSQ & scSpaceCommaSpace
            sQry = sQry & CStr(.Increment) & scSpaceCommaSpace
            sQry = sQry & CStr(.Maximum) & scSpaceCommaSpace
            sQry = sQry & CStr(.Minimum) & scSpaceCommaSpace
            sQry = sQry & CStr(.Multiples)
        End With
        If Not ExecBooleanSQL(sQry) Then
            Insert = False
            Exit Function
        End If
     Else
        ' Create CList with parameter information
        Set lstParms = New CList
        lstParms.Add "@BnfTaxId", oBenefit.BnfTaxId
        lstParms.Add "@BnfTypeId", oBenefit.BnfTypeId
        lstParms.Add "@Cost", oBenefit.Cost
        lstParms.Add "@Description", Trim$(Left$(oBenefit.Description, 80))
        lstParms.Add "@Increment", oBenefit.Increment
        lstParms.Add "@Maximum", oBenefit.Maximum
        lstParms.Add "@Minimum", oBenefit.Minimum
        lstParms.Add "@Multiples", oBenefit.Multiples
        If Not ExecBooleanMDB(LoadResString(scIns_Benefit), lstParms) Then
            Insert = False
            Exit Function
        End If
        Set lstParms = Nothing
    End If
    oBenefit.Term
    Set oBenefit = Nothing
    FillListBenefits
End If

'Inserting a BenefitLink?
If Not (oBenefitLink Is Nothing) Then
    ' validate oEmployee object
    With oBenefitLink
        If .Employee Is Nothing Then
                Insert = False
                On Error GoTo 0
                Err.Raise icErrReqFld + vbObjectError, Trim$(App.Title) & LoadResString(scInsBnfLnk), LoadResString(icErrReqFld) & "Employee Object"
                Exit Function
        Else
            If Not CBool(.Employee.PKId) Then
                Insert = False
                On Error GoTo 0
                Err.Raise icErrReqFld + vbObjectError, Trim$(App.Title) & LoadResString(scInsBnfLnk), LoadResString(icErrReqFld) & "Employee.PKId"
                Exit Function
            End If
        End If
        If .Benefit Is Nothing Then
                Insert = False
                On Error GoTo 0
                Err.Raise icErrReqFld + vbObjectError, Trim$(App.Title) & LoadResString(scInsBnfLnk), LoadResString(icErrReqFld) & "Benefit"
                Exit Function
        Else
            If Not CBool(.Benefit.PKId) Then
                Insert = False
                On Error GoTo 0
                Err.Raise icErrReqFld + vbObjectError, Trim$(App.Title) & LoadResString(scInsBnfLnk), LoadResString(icErrReqFld) & "Benefit.PKId"
                Exit Function
            End If
        End If
        If Not CBool(.CoverageId) Then
            Insert = False
            On Error GoTo 0
            Err.Raise icErrReqFld + vbObjectError, Trim$(App.Title) & LoadResString(scInsBnfLnk), LoadResString(icErrReqFld) & "Benefit.CoverageID"
            Exit Function
        End If
        If Not CBool(.TaxId) Then
            Insert = False
            On Error GoTo 0
            Err.Raise icErrReqFld + vbObjectError, Trim$(App.Title) & LoadResString(scInsBnfLnk), LoadResString(icErrReqFld) & "Benefit.TaxID"
            Exit Function
        End If
    End With
    
    If m_bUseODBC Then
        With oBenefitLink
            sQry = LoadResString(scExecute) & LoadResString(scIns_EmplBnf) & scSpace
            sQry = sQry & CStr(.Benefit.PKId) & scSpaceCommaSpace
            sQry = sQry & CStr(.CoverageId) & scSpaceCommaSpace
            sQry = sQry & CStr(.Employee.PKId) & scSpaceCommaSpace
            sQry = sQry & CStr(.TaxId) & scSpaceCommaSpace
            sQry = sQry & CStr(.Amount)
        End With
        If Not ExecBooleanSQL(sQry) Then
            Insert = False
            Exit Function
        End If
    Else
        ' Create CList with parameter information
        Set lstParms = New CList
        lstParms.Add "@BenefitId", oBenefitLink.Benefit.PKId
        lstParms.Add "@CoverageId", oBenefitLink.CoverageId
        lstParms.Add "@EmployeeId", oBenefitLink.Employee.PKId
        lstParms.Add "@TaxId", oBenefitLink.TaxId
        lstParms.Add "@Amount", oBenefitLink.Amount
        If Not ExecBooleanMDB(LoadResString(scIns_EmplBnf), lstParms) Then
            Insert = False
            Exit Function
        End If
        Set lstParms = Nothing
    End If
    oBenefitLink.Term
    Set oBenefitLink = Nothing
End If

'Inserting a Dependant?
If Not (oDependant Is Nothing) Then
    ' validate oDependant object
    With oDependant
        If Not CBool(.DepTypeId) Then
            Insert = False
            On Error GoTo 0
            Err.Raise icErrReqFld + vbObjectError, Trim$(App.Title) & LoadResString(scInsDep), LoadResString(icErrReqFld) & "DepTypeId"
            Exit Function
        End If
        If .Employee Is Nothing Then
                Insert = False
                On Error GoTo 0
                Err.Raise icErrReqFld + vbObjectError, Trim$(App.Title) & LoadResString(scInsBnfLnk), LoadResString(icErrReqFld) & "Employee Object"
                Exit Function
        Else
            If Not CBool(.Employee.PKId) Then
                Insert = False
                On Error GoTo 0
                Err.Raise icErrReqFld + vbObjectError, Trim$(App.Title) & LoadResString(scInsDep), LoadResString(icErrReqFld) & "Employee.PKId"
                Exit Function
            End If
        End If
        If Not CBool(Len(.FirstName)) Then
            Insert = False
            On Error GoTo 0
            Err.Raise icErrReqFld + vbObjectError, Trim$(App.Title) & LoadResString(scInsDep), LoadResString(icErrReqFld) & "FirstName"
            Exit Function
        End If
        If Not CBool(Len(.LastName)) Then
            Insert = False
            On Error GoTo 0
            Err.Raise icErrReqFld + vbObjectError, Trim$(App.Title) & LoadResString(scInsDep), LoadResString(icErrReqFld) & "LastName"
            Exit Function
        End If
    End With
    If m_bUseODBC Then
        sQry = LoadResString(scExecute) & LoadResString(scIns_Dependant) & scSpace
        With oDependant
            sQry = sQry & CStr(.DepTypeId) & scSpaceCommaSpace
            sQry = sQry & CStr(.Employee.PKId) & scSpaceCommaSpace
            If IsDate(.DateOfBirth) Then
                sQry = sQry & scSQ & Format$(.DateOfBirth, scDateFormat) & scSQ & scSpaceCommaSpace
            Else
                sQry = sQry & scNull & scSpaceCommaSpace
            End If
            sQry = sQry & scSQ & Trim$(Left$(DoQuotes(.FirstName), 35)) & scSQ & scSpaceCommaSpace
            If .Gender Then
                sQry = sQry & scZeroStr & scSpaceCommaSpace
            Else
                sQry = sQry & scOneStr & scSpaceCommaSpace
            End If
            sQry = sQry & scSQ & Trim$(Left$(DoQuotes(.LastName), 35)) & scSQ & scSpaceCommaSpace
            If .Status Then
                sQry = sQry & scOneStr
            Else
                sQry = sQry & scZeroStr
            End If
        End With
        If Not ExecBooleanSQL(sQry) Then
            Insert = False
            Exit Function
        End If
    Else
      ' Create CList with parameter information
        Set lstParms = New CList
        lstParms.Add "@DepTypeId", oDependant.DepTypeId
        lstParms.Add "@EmployeeId", oDependant.Employee.PKId
        lstParms.Add "@DateOfBirth", oDependant.DateOfBirth
        lstParms.Add "@FirstName", Trim$(Left$(oDependant.FirstName, 35))
        lstParms.Add "@Gender", oDependant.Gender
        lstParms.Add "@LastName", Trim$(Left$(oDependant.LastName, 35))
        lstParms.Add "@Status", oDependant.Status
        If Not ExecBooleanMDB(LoadResString(scIns_Dependant), lstParms) Then
            Insert = False
            Exit Function
        End If
        Set lstParms = Nothing
    End If
#If UseEvents Then
    RaiseEvent DependantChanged(oDependant.Employee.PKId)
#End If
    oDependant.Term
    Set oDependant = Nothing
End If

'Inserting an Employee?
If Not (oEmployee Is Nothing) Then
    ' validate oEmployee object
    With oEmployee
        If .Bank Is Nothing Then
            Insert = False
            On Error GoTo 0
            Err.Raise icErrReqFld + vbObjectError, Trim$(App.Title) & LoadResString(scInsEmpl), LoadResString(icErrReqFld) & "Bank Object"
            Exit Function
        Else
            If Not CBool(.Bank.PKId) Then
                Insert = False
                On Error GoTo 0
                Err.Raise icErrReqFld + vbObjectError, Trim$(App.Title) & LoadResString(scInsEmpl), LoadResString(icErrReqFld) & "Bank.PKId"
                Exit Function
            End If
        End If
        If Not CBool(Len(.FirstName)) Then
            Insert = False
            On Error GoTo 0
            Err.Raise icErrReqFld + vbObjectError, Trim$(App.Title) & LoadResString(scInsEmpl), LoadResString(icErrReqFld) & "FirstName"
            Exit Function
        End If
        If Not CBool(Len(.LastName)) Then
            Insert = False
            On Error GoTo 0
            Err.Raise icErrReqFld + vbObjectError, Trim$(App.Title) & LoadResString(scInsEmpl), LoadResString(icErrReqFld) & "LastName"
            Exit Function
        End If
        If Not CBool(Len(.Position)) Then
            Insert = False
            On Error GoTo 0
            Err.Raise icErrReqFld + vbObjectError, Trim$(App.Title) & LoadResString(scInsEmpl), LoadResString(icErrReqFld) & "Position"
            Exit Function
        End If
        If Not IsDate(.StartDate) Then
            Insert = False
            On Error GoTo 0
            Err.Raise icErrReqFld + vbObjectError, Trim$(App.Title) & LoadResString(scInsEmpl), LoadResString(icErrReqFld) & "StartDate"
            Exit Function
        End If
        If Not CBool(Len(.UserName)) Then
            Insert = False
            On Error GoTo 0
            Err.Raise icErrReqFld + vbObjectError, Trim$(App.Title) & LoadResString(scInsEmpl), LoadResString(icErrReqFld) & "UserName"
            Exit Function
        End If
    End With
    If m_bUseODBC Then
        sQry = LoadResString(scExecute) & LoadResString(scIns_Employee) & scSpace
        With oEmployee
            sQry = sQry & CStr(.Bank.PKId) & scSpaceCommaSpace
            sQry = sQry & scSQ & Trim$(Left$(DoQuotes(.Address1), 60)) & scSQ & scSpaceCommaSpace
            sQry = sQry & scSQ & Trim$(Left$(DoQuotes(.Address2), 60)) & scSQ & scSpaceCommaSpace
            sQry = sQry & scSQ & Trim$(Left$(DoQuotes(.BankAccount), 35)) & scSQ & scSpaceCommaSpace
            sQry = sQry & scSQ & Trim$(Left$(DoQuotes(.City), 50)) & scSQ & scSpaceCommaSpace
            If IsDate(.DateOfBirth) Then
                sQry = sQry & scSQ & Format$(.DateOfBirth, scDateFormat) & scSQ & scSpaceCommaSpace
            Else
                sQry = sQry & scNull & scSpaceCommaSpace
            End If
            If IsDate(.EndDate) Then
                sQry = sQry & scSQ & Format$(.EndDate, scDateFormat) & scSQ & scSpaceCommaSpace
            Else
                sQry = sQry & scNull & scSpaceCommaSpace
            End If
            sQry = sQry & scSQ & Trim$(Left$(DoQuotes(.FirstName), 35)) & scSQ & scSpaceCommaSpace
            If .Gender Then
                sQry = sQry & scZeroStr & scSpaceCommaSpace
            Else
                sQry = sQry & scOneStr & scSpaceCommaSpace
            End If
            sQry = sQry & scSQ & Trim$(Left$(DoQuotes(.LastName), 35)) & scSQ & scSpaceCommaSpace
            sQry = sQry & scSQ & Trim$(Left$(DoQuotes(.Position), 50)) & scSQ & scSpaceCommaSpace
            sQry = sQry & CStr(.Salary) & scSpaceCommaSpace
            sQry = sQry & scSQ & Format$(.StartDate, scDateFormat) & scSQ & scSpaceCommaSpace
            sQry = sQry & scSQ & Trim$(Left$(DoQuotes(.UserName), 16)) & scSQ
        End With
        If Not ExecBooleanSQL(sQry) Then
            Insert = False
            Exit Function
        End If
    Else
        ' Create CList with parameter information
        Set lstParms = New CList
        lstParms.Add "@BankID", oEmployee.Bank.PKId
        lstParms.Add "@Address1", Trim$(Left$(oEmployee.Address1, 60))
        lstParms.Add "@Address2", Trim$(Left$(oEmployee.Address2, 60))
        lstParms.Add "@BankAccount", Trim$(Left$(oEmployee.BankAccount, 35))
        lstParms.Add "@City", Trim$(Left$(oEmployee.City, 50))
        lstParms.Add "@DateOfBirth", oEmployee.DateOfBirth
        lstParms.Add "@EndDate", oEmployee.EndDate
        lstParms.Add "@FirstName", Trim$(Left$(oEmployee.FirstName, 35))
        lstParms.Add "@Gender", oEmployee.Gender
        lstParms.Add "@LastName", Trim$(Left$(oEmployee.LastName, 35))
        lstParms.Add "@Position", Trim$(Left$(oEmployee.Position, 50))
        lstParms.Add "@Salary", oEmployee.Salary
        lstParms.Add "@StartDate", oEmployee.StartDate
        lstParms.Add "@UserName", Trim$(Left$(oEmployee.UserName, 16))
        If Not ExecBooleanMDB(LoadResString(scIns_Employee), lstParms) Then
            Insert = False
            Exit Function
        End If
        Set lstParms = Nothing
    End If
    oEmployee.Term
    Set oEmployee = Nothing
    FillListEmployees
End If
Insert = True
Exit Function

errInsert:
    Insert = False
    m_lErrNo = Err
    If Err < vbObjectError Then m_lErrNo = m_lErrNo + vbObjectError
    Err.Raise m_lErrNo, Trim$(App.Title) & LoadResString(scIns), Err.Description
    Exit Function
    
End Function

'IsUserValid
'Purpose     - Validates user
'Inputs      - (String)  sUserName
'Outputs     - (Integer) Access level
Public Function IsUserValid(ByVal sUserName As String) As Integer
Dim sData()     As String
Dim sQry        As String

On Error Resume Next

If UCase(Trim$(sUserName)) = "ADMIN" Then
    IsUserValid = icValidAdmin
    Exit Function
End If

If m_iSrcType <> icDoDataDirect Then
    Err = False
    IsUserValid = m_oRemote.IsUserValid(sUserName)
    If CBool(Err) Then
        On Error GoTo 0
        Err.Raise Err, Trim$(App.Title) & LoadResString(scIsUserValid), Err.Description
        Exit Function
    End If
Else
    sQry = LoadResString(scCHKUSER) & scSQ & Trim$(Left$(DoQuotes(sUserName), 16)) & scSQ
    Err = False
    If ExecFillArray(sQry, sData) Then
        If CBool(Err) Then
            IsUserValid = False
            m_lErrNo = Err
            If Err < vbObjectError Then m_lErrNo = m_lErrNo + vbObjectError
            On Error GoTo 0
            Err.Raise m_lErrNo, Trim$(App.Title) & LoadResString(scIsUserValid), Err.Description
            Exit Function
        End If
        If CBool(sData(0, 0)) Then
            IsUserValid = icValidUser
            Exit Function
        Else
            IsUserValid = icInvalidUser
            Exit Function
        End If
    Else
        IsUserValid = icInvalidUser
        Exit Function
    End If
End If
End Function

'Term
'Purpose     - Terminate an object, release all object references
Public Sub Term()
Dim oBank       As CBank
Dim oBenefit    As CBenefit
Dim oEmployee   As CEmployee

On Error Resume Next

'Clear the Benefits List
Set m_lstBenefits = Nothing
    
'Clear the Banks List
Set m_lstBanks = Nothing

'Clear the Employees List
Set m_lstEmployees = Nothing

m_oRemote.Term
Set m_oRemote = Nothing

m_dbBenefit.Close
Set m_dbBenefit = Nothing
Set m_wsBenefit = Nothing

End Sub

'Update
'Purpose     - Update data object
'Inputs      - (CBank)        Optional oBank
'              (CBenefit)     Optional oBenefit
'              (CBenefitLink) Optional oBenefitLink
'              (CDependant)   Optional oDependant
'              (CEmployee)    Optional oEmployee
'Outputs     - (Boolean) Success
Public Function Update(Optional oBank As CBank, Optional oBenefit As CBenefit, Optional oBenefitLink As CBenefitLink, Optional oDependant As CDependant, Optional oEmployee As CEmployee) As Boolean
Dim lstParms     As CList
Dim iParameters  As Integer
Dim sQry         As String

On Error GoTo errUpdate

If Not (oBank Is Nothing) Then iParameters = iParameters + 1
If Not (oBenefit Is Nothing) Then iParameters = iParameters + 1
If Not (oBenefitLink Is Nothing) Then iParameters = iParameters + 1
If Not (oDependant Is Nothing) Then iParameters = iParameters + 1
If Not (oEmployee Is Nothing) Then iParameters = iParameters + 1

If iParameters <> 1 Then
    Update = False
    On Error GoTo 0
    Err.Raise vbObjectError + icErrObjectOp, Trim$(App.Title), LoadResString(icErrObjectOp)
    Exit Function
End If

'Updating a Bank?
If Not (oBank Is Nothing) Then
    ' validate oBank object
    With oBank
                
        If Not CBool(.PKId) Then
            Update = False
            On Error GoTo 0
            Err.Raise icErrReqFld + vbObjectError, Trim$(App.Title) & LoadResString(scUpdBank), LoadResString(icErrReqFld) & "PKId"
            Exit Function
        End If
        If Not CBool(Len(.Code)) Then
            Update = False
            On Error GoTo 0
            Err.Raise icErrReqFld + vbObjectError, Trim$(App.Title) & LoadResString(scUpdBank), LoadResString(icErrReqFld) & "Code"
            Exit Function
        End If
        If Not CBool(Len(.Name)) Then
            Update = False
            On Error GoTo 0
            Err.Raise icErrReqFld + vbObjectError, Trim$(App.Title) & LoadResString(scUpdBank), LoadResString(icErrReqFld) & "Name"
            Exit Function
        End If
    End With
    If m_bUseODBC Then
        sQry = LoadResString(scExecute) & LoadResString(scUpd_Bank) & scSpace
        With oBank
            sQry = sQry & CStr(.PKId) & scSpaceCommaSpace
            sQry = sQry & scSQ & Trim$(Left$(DoQuotes(.Address1), 60)) & scSQ & scSpaceCommaSpace
            sQry = sQry & scSQ & Trim$(Left$(DoQuotes(.Address2), 60)) & scSQ & scSpaceCommaSpace
            sQry = sQry & scSQ & Trim$(Left$(DoQuotes(.City), 50)) & scSQ & scSpaceCommaSpace
            sQry = sQry & scSQ & Trim$(Left$(DoQuotes(.Code), 16)) & scSQ & scSpaceCommaSpace
            sQry = sQry & scSQ & Trim$(Left$(DoQuotes(.Name), 80)) & scSQ
        End With
        If Not ExecBooleanSQL(sQry) Then
            Update = False
            Exit Function
        End If
    Else
        ' Create CList with parameter information
        Set lstParms = New CList
        lstParms.Add "@PKId", oBank.PKId
        lstParms.Add "@Address1", Trim$(Left$(oBank.Address1, 60))
        lstParms.Add "@Address2", Trim$(Left$(oBank.Address2, 60))
        lstParms.Add "@City", Trim$(Left$(oBank.City, 15))
        lstParms.Add "@Code", Trim$(Left$(oBank.Code, 16))
        lstParms.Add "@Name", Trim$(Left$(oBank.Name, 80))
        If Not ExecBooleanMDB(LoadResString(scUpd_Bank), lstParms) Then
            Update = False
            Exit Function
        End If
        Set lstParms = Nothing
    End If
    oBank.Term
End If

' Updating a benefit?
If Not (oBenefit Is Nothing) Then
    ' validate oBenefit object
    With oBenefit
        If Not CBool(.PKId) Then
            Update = False
            On Error GoTo 0
            Err.Raise icErrReqFld + vbObjectError, Trim$(App.Title) & LoadResString(scUpdBnf), LoadResString(icErrReqFld) & "PKId"
            Exit Function
        End If
        If Not CBool(.BnfTaxId) Then
            Update = False
            On Error GoTo 0
            Err.Raise icErrReqFld + vbObjectError, Trim$(App.Title) & LoadResString(scUpdBnf), LoadResString(icErrReqFld) & "BnfTaxId"
            Exit Function
        End If
        If Not CBool(.BnfTypeId) Then
            Update = False
            On Error GoTo 0
            Err.Raise icErrReqFld + vbObjectError, Trim$(App.Title) & LoadResString(scUpdBnf), LoadResString(icErrReqFld) & "BnfTypeId"
            Exit Function
        End If
        If Not CBool(Len(.Description)) Then
            Update = False
            On Error GoTo 0
            Err.Raise icErrReqFld + vbObjectError, Trim$(App.Title) & LoadResString(scUpdBnf), LoadResString(icErrReqFld) & "Description"
            Exit Function
        End If
    End With
    
    If m_bUseODBC Then
        sQry = LoadResString(scExecute) & LoadResString(scUpd_Benefit) & scSpace
        With oBenefit
            sQry = sQry & CStr(.PKId) & scSpaceCommaSpace
            sQry = sQry & CStr(.BnfTaxId) & scSpaceCommaSpace
            sQry = sQry & CStr(.BnfTypeId) & scSpaceCommaSpace
            sQry = sQry & CStr(.Cost) & scSpaceCommaSpace
            sQry = sQry & scSQ & Trim$(Left$(DoQuotes(.Description), 80)) & scSQ & scSpaceCommaSpace
            sQry = sQry & CStr(.Increment) & scSpaceCommaSpace
            sQry = sQry & CStr(.Maximum) & scSpaceCommaSpace
            sQry = sQry & CStr(.Minimum) & scSpaceCommaSpace
            sQry = sQry & CStr(.Multiples)
        End With
        If Not ExecBooleanSQL(sQry) Then
            Update = False
            Exit Function
        End If
    Else
        ' Create CList with parameter information
        Set lstParms = New CList
        lstParms.Add "@PKId", oBenefit.PKId
        lstParms.Add "@BnfTaxId", oBenefit.BnfTaxId
        lstParms.Add "@BnfTypeId", oBenefit.BnfTypeId
        lstParms.Add "@Cost", oBenefit.Cost
        lstParms.Add "@Description", Trim$(Left$(oBenefit.Description, 80))
        lstParms.Add "@Increment", oBenefit.Increment
        lstParms.Add "@Maximum", oBenefit.Maximum
        lstParms.Add "@Minimum", oBenefit.Minimum
        lstParms.Add "@Multiples", oBenefit.Multiples
        If Not ExecBooleanMDB(LoadResString(scUpd_Benefit), lstParms) Then
            Update = False
            Exit Function
        End If
        Set lstParms = Nothing
    End If
    oBenefit.Term
End If

'Updating a BenefitLink?
If Not (oBenefitLink Is Nothing) Then
    ' validate oEmployee object
    With oBenefitLink
        If Not CBool(.PKId) Then
            Update = False
            On Error GoTo 0
            Err.Raise icErrReqFld + vbObjectError, Trim$(App.Title) & LoadResString(scUpdBnfLnk), LoadResString(icErrReqFld) & "BenefitLink.PKId"
            Exit Function
        End If
        If .Employee Is Nothing Then
                Update = False
                On Error GoTo 0
                Err.Raise icErrReqFld + vbObjectError, Trim$(App.Title) & LoadResString(scUpdBnfLnk), LoadResString(icErrReqFld) & "Employee Object"
                Exit Function
        Else
            If Not CBool(.Employee.PKId) Then
                Update = False
                On Error GoTo 0
                Err.Raise icErrReqFld + vbObjectError, Trim$(App.Title) & LoadResString(scUpdBnfLnk), LoadResString(icErrReqFld) & "Employee.PKId"
                Exit Function
            End If
        End If
        If .Benefit Is Nothing Then
                Update = False
                On Error GoTo 0
                Err.Raise icErrReqFld + vbObjectError, Trim$(App.Title) & LoadResString(scUpdBnfLnk), LoadResString(icErrReqFld) & "Benefit"
                Exit Function
        Else
            If Not CBool(.Benefit.PKId) Then
                Update = False
                On Error GoTo 0
                Err.Raise icErrReqFld + vbObjectError, Trim$(App.Title) & LoadResString(scUpdBnfLnk), LoadResString(icErrReqFld) & "Benefit.PKId"
                Exit Function
            End If
        End If
        If Not CBool(.CoverageId) Then
            Update = False
            On Error GoTo 0
            Err.Raise icErrReqFld + vbObjectError, Trim$(App.Title) & LoadResString(scUpdBnfLnk), LoadResString(icErrReqFld) & "Benefit.CoverageID"
            Exit Function
        End If
        If Not CBool(.TaxId) Then
            Update = False
            On Error GoTo 0
            Err.Raise icErrReqFld + vbObjectError, Trim$(App.Title) & LoadResString(scUpdBnfLnk), LoadResString(icErrReqFld) & "Benefit.TaxId"
            Exit Function
        End If
        If m_bUseODBC Then
            sQry = LoadResString(scExecute) & LoadResString(scUpd_EmplBnf) & scSpace
            sQry = sQry & CStr(.PKId) & scSpaceCommaSpace
            sQry = sQry & CStr(.Benefit.PKId) & scSpaceCommaSpace
            sQry = sQry & CStr(.CoverageId) & scSpaceCommaSpace
            sQry = sQry & CStr(.Employee.PKId) & scSpaceCommaSpace
            sQry = sQry & CStr(.TaxId) & scSpaceCommaSpace
            sQry = sQry & CStr(.Amount)
            If Not ExecBooleanSQL(sQry) Then
                Update = False
                Exit Function
            End If
        Else
            ' Create CList with parameter information
            Set lstParms = New CList
            lstParms.Add "@PKId", oBenefitLink.PKId
            lstParms.Add "@BenefitId", oBenefitLink.Benefit.PKId
            lstParms.Add "@CoverageId", oBenefitLink.CoverageId
            lstParms.Add "@EmployeeId", oBenefitLink.Employee.PKId
            lstParms.Add "@TaxId", oBenefitLink.TaxId
            lstParms.Add "@Amount", oBenefitLink.Amount
            If Not ExecBooleanMDB(LoadResString(scUpd_EmplBnf), lstParms) Then
                Update = False
                Exit Function
            End If
            Set lstParms = Nothing
        End If
    End With
    oBenefitLink.Term
End If

'Updating a Dependant?
If Not (oDependant Is Nothing) Then
    ' validate oDependant object
    With oDependant
        If Not CBool(.PKId) Then
            Update = False
            On Error GoTo 0
            Err.Raise icErrReqFld + vbObjectError, Trim$(App.Title) & LoadResString(scUpdDep), LoadResString(icErrReqFld) & "PKId"
            Exit Function
        End If
        If Not CBool(.DepTypeId) Then
            Update = False
            On Error GoTo 0
            Err.Raise icErrReqFld + vbObjectError, Trim$(App.Title) & LoadResString(scUpdDep), LoadResString(icErrReqFld) & "DepTypeId"
            Exit Function
        End If
        If .Employee Is Nothing Then
                Update = False
                On Error GoTo 0
                Err.Raise icErrReqFld + vbObjectError, Trim$(App.Title) & LoadResString(scUpdDep), LoadResString(icErrReqFld) & "Employee Object"
                Exit Function
        Else
            If Not CBool(.Employee.PKId) Then
                Update = False
                On Error GoTo 0
                Err.Raise icErrReqFld + vbObjectError, Trim$(App.Title) & LoadResString(scUpdDep), LoadResString(icErrReqFld) & "Employee.PKId"
                Exit Function
            End If
        End If
        If Not CBool(Len(.FirstName)) Then
            Update = False
            On Error GoTo 0
            Err.Raise icErrReqFld + vbObjectError, Trim$(App.Title) & LoadResString(scUpdDep), LoadResString(icErrReqFld) & "FirstName"
            Exit Function
        End If
        If Not CBool(Len(.LastName)) Then
            Update = False
            On Error GoTo 0
            Err.Raise icErrReqFld + vbObjectError, Trim$(App.Title) & LoadResString(scUpdDep), LoadResString(icErrReqFld) & "LastName"
            Exit Function
        End If
    End With
    If m_bUseODBC Then
        sQry = LoadResString(scExecute) & LoadResString(scUpd_Dependant) & scSpace
        With oDependant
            sQry = sQry & CStr(.PKId) & scSpaceCommaSpace
            sQry = sQry & CStr(.DepTypeId) & scSpaceCommaSpace
            sQry = sQry & CStr(.Employee.PKId) & scSpaceCommaSpace
            If IsDate(.DateOfBirth) Then
                sQry = sQry & scSQ
                sQry = sQry & Format$(.DateOfBirth, scDateFormat) & scSQ & scSpaceCommaSpace
            Else
                sQry = sQry & scNull & scSpaceCommaSpace
            End If
            sQry = sQry & scSQ & Trim$(Left$(DoQuotes(.FirstName), 35)) & scSQ & scSpaceCommaSpace
            If .Gender Then
                sQry = sQry & scZeroStr & scSpaceCommaSpace
            Else
                sQry = sQry & scOneStr & scSpaceCommaSpace
            End If
            sQry = sQry & scSQ & Trim$(Left$(DoQuotes(.LastName), 35)) & scSQ & scSpaceCommaSpace
            If .Status Then
                sQry = sQry & scOneStr
            Else
                sQry = sQry & scZeroStr
            End If
        End With
        If Not ExecBooleanSQL(sQry) Then
            Update = False
            Exit Function
        End If
    Else
        'Create CList with parameter information
        Set lstParms = New CList
        lstParms.Add "@PKId", oDependant.PKId
        lstParms.Add "@DepTypeId", oDependant.DepTypeId
        lstParms.Add "@EmployeeId", oDependant.Employee.PKId
        lstParms.Add "@DateOfBirth", oDependant.DateOfBirth
        lstParms.Add "@FirstName", Trim$(Left$(oDependant.FirstName, 35))
        lstParms.Add "@Gender", oDependant.Gender
        lstParms.Add "@LastName", Trim$(Left$(oDependant.LastName, 35))
        lstParms.Add "@Status", oDependant.Status
        If Not ExecBooleanMDB(LoadResString(scUpd_Dependant), lstParms) Then
            Update = False
            Exit Function
        End If
        Set lstParms = Nothing
    End If
#If UseEvents Then
    RaiseEvent DependantChanged(oDependant.Employee.PKId)
#End If
    oDependant.Term
End If

'Updating an Employee?
If Not (oEmployee Is Nothing) Then
    ' validate oEmployee object
    With oEmployee
        If Not CBool(.PKId) Then
            Update = False
            On Error GoTo 0
            Err.Raise icErrReqFld + vbObjectError, Trim$(App.Title) & LoadResString(scUpdEmpl), LoadResString(icErrReqFld) & "PKId"
            Exit Function
        End If
        If .Bank Is Nothing Then
            Update = False
            On Error GoTo 0
            Err.Raise icErrReqFld + vbObjectError, Trim$(App.Title) & LoadResString(scUpdEmpl), LoadResString(icErrReqFld) & "Bank Object"
            Exit Function
        Else
            If Not CBool(.Bank.PKId) Then
            Update = False
            On Error GoTo 0
            Err.Raise icErrReqFld + vbObjectError, Trim$(App.Title) & LoadResString(scUpdEmpl), LoadResString(icErrReqFld) & "Bank.PKId"
            Exit Function
            End If
        End If
        If Not CBool(Len(.FirstName)) Then
            Update = False
            On Error GoTo 0
            Err.Raise icErrReqFld + vbObjectError, Trim$(App.Title) & LoadResString(scUpdEmpl), LoadResString(icErrReqFld) & "FirstName"
            Exit Function
        End If
        If Not CBool(Len(.LastName)) Then
            Update = False
            On Error GoTo 0
            Err.Raise icErrReqFld + vbObjectError, Trim$(App.Title) & LoadResString(scUpdEmpl), LoadResString(icErrReqFld) & "LastName"
            Exit Function
        End If
        If Not CBool(Len(.Position)) Then
            Update = False
            On Error GoTo 0
            Err.Raise icErrReqFld + vbObjectError, Trim$(App.Title) & LoadResString(scUpdEmpl), LoadResString(icErrReqFld) & "Position"
            Exit Function
        End If
        If Not IsDate(.StartDate) Then
            Update = False
            On Error GoTo 0
            Err.Raise icErrReqFld + vbObjectError, Trim$(App.Title) & LoadResString(scUpdEmpl), LoadResString(icErrReqFld) & "StartDate"
            Exit Function
        End If
        If Not CBool(Len(.UserName)) Then
            Update = False
            On Error GoTo 0
            Err.Raise icErrReqFld + vbObjectError, Trim$(App.Title) & LoadResString(scUpdEmpl), LoadResString(icErrReqFld) & "UserName"
            Exit Function
        End If
    End With
    If m_bUseODBC Then
        sQry = LoadResString(scExecute) & LoadResString(scUpd_Employee) & scSpace
        With oEmployee
            sQry = sQry & CStr(.PKId) & scSpaceCommaSpace
            sQry = sQry & CStr(.Bank.PKId) & scSpaceCommaSpace
            sQry = sQry & scSQ & Trim$(Left$(DoQuotes(.Address1), 60)) & scSQ & scSpaceCommaSpace
            sQry = sQry & scSQ & Trim$(Left$(DoQuotes(.Address2), 60)) & scSQ & scSpaceCommaSpace
            sQry = sQry & scSQ & Trim$(Left$(DoQuotes(.BankAccount), 35)) & scSQ & scSpaceCommaSpace
            sQry = sQry & scSQ & Trim$(Left$(DoQuotes(.City), 50)) & scSQ & scSpaceCommaSpace
            If IsDate(.DateOfBirth) Then
                sQry = sQry & scSQ & Format$(.DateOfBirth, scDateFormat) & scSQ & scSpaceCommaSpace
            Else
                sQry = sQry & scNull & scSpaceCommaSpace
            End If
            If IsDate(.EndDate) Then
                sQry = sQry & scSQ & Format$(.EndDate, scDateFormat) & scSQ & scSpaceCommaSpace
            Else
                sQry = sQry & scNull & scSpaceCommaSpace
            End If
            sQry = sQry & scSQ & Trim$(Left$(DoQuotes(.FirstName), 35)) & scSQ & scSpaceCommaSpace
            If .Gender Then
                sQry = sQry & scZeroStr & scSpaceCommaSpace
            Else
                sQry = sQry & scOneStr & scSpaceCommaSpace
            End If
            sQry = sQry & scSQ & Trim$(Left$(DoQuotes(.LastName), 35)) & scSQ & scSpaceCommaSpace
            sQry = sQry & scSQ & Trim$(Left$(DoQuotes(.Position), 50)) & scSQ & scSpaceCommaSpace
            sQry = sQry & CStr(.Salary) & scSpaceCommaSpace
            sQry = sQry & scSQ & Format$(.StartDate, scDateFormat) & scSQ & scSpaceCommaSpace
            sQry = sQry & scSQ & Trim$(Left$(DoQuotes(.UserName), 16)) & scSQ
        End With
        If Not ExecBooleanSQL(sQry) Then
            Update = False
            Exit Function
        End If
    Else
        ' Create CList with parameter information
        Set lstParms = New CList
        lstParms.Add "@PKId", oEmployee.PKId
        lstParms.Add "@BankID", oEmployee.Bank.PKId
        lstParms.Add "@Address1", Trim$(Left$(oEmployee.Address1, 60))
        lstParms.Add "@Address2", Trim$(Left$(oEmployee.Address2, 60))
        lstParms.Add "@BankAccount", Trim$(Left$(oEmployee.BankAccount, 35))
        lstParms.Add "@City", Trim$(Left$(oEmployee.City, 50))
        lstParms.Add "@DateOfBirth", oEmployee.DateOfBirth
        lstParms.Add "@EndDate", oEmployee.EndDate
        lstParms.Add "@FirstName", Trim$(Left$(oEmployee.FirstName, 35))
        lstParms.Add "@Gender", oEmployee.Gender
        lstParms.Add "@LastName", Trim$(Left$(oEmployee.LastName, 35))
        lstParms.Add "@Position", Trim$(Left$(oEmployee.Position, 50))
        lstParms.Add "@Salary", oEmployee.Salary
        lstParms.Add "@StartDate", oEmployee.StartDate
        lstParms.Add "@UserName", Trim$(Left$(oEmployee.UserName, 16))
        If Not ExecBooleanMDB(LoadResString(scUpd_Employee), lstParms) Then
            Update = False
            Exit Function
        End If
        Set lstParms = Nothing
    End If
    oEmployee.Term
End If
Update = True
Exit Function

errUpdate:
    Update = False
    m_lErrNo = Err
    If Err < vbObjectError Then m_lErrNo = m_lErrNo + vbObjectError
    Err.Raise m_lErrNo, Trim$(App.Title) & LoadResString(scUpd), Err.Description
    Exit Function
    
End Function

Public Property Get Banks() As CListBank
If m_lstBanks Is Nothing Then Set m_lstBanks = New CListBank
Set Banks = m_lstBanks
End Property

Public Property Set Banks(lstBanks As CListBank)
Set m_lstBanks = lstBanks
End Property

Public Property Get Benefits() As CListBenefit
If m_lstBenefits Is Nothing Then Set m_lstBenefits = New CListBenefit
Set Benefits = m_lstBenefits
End Property

Public Property Set Benefits(lstBenefits As CListBenefit)
Set m_lstBenefits = lstBenefits
End Property

Public Property Get Employees() As CListEmployee
If m_lstEmployees Is Nothing Then Set m_lstEmployees = New CListEmployee
Set Employees = m_lstEmployees
End Property

Public Property Set Employees(lstEmployees As CListEmployee)
Set m_lstEmployees = lstEmployees
End Property

Public Property Get UseODBC() As Boolean
UseODBC = m_bUseODBC
End Property

Private Sub Class_Terminate()
On Error Resume Next
Term
End Sub

