//cadmin.cpp  - Copyright (C) 1996 Microsoft Corporation
//            - All rights reserved.
//Date        - 10/19/96
//Authors     - Ken Bergmann, Robert Coleridge
//Purpose     - Implementation file for the CAdmin object.

#pragma warning (disable : 4800)	// turns off BOOL warning

#include "support.h"
#include "cstrfunc.h"
#include "csaferay.h"
#include "cienum.h"

#include "cadmin.h"
#include "cbank.h"
#include "cbenefit.h"
#include "cemployee.h"

#include "clist.h"

#include <dbdaoerr.h>

BSTR sKey = TEXT("mbmabptebkjcdlgtjmskjwtsdhjbmkmwtrak");

extern long g_oid;

struct OleInit
{
	OleInit()	{ OleInitialize(NULL);	}
	~OleInit()	{ OleUninitialize();	}
} Init_Ole;


OLECHAR * RemoteServer_Methods[] =
{
	TEXT("Init"),
		TEXT("ExecBooleanSQL"),
		TEXT("ExecBooleanMDB"),
		TEXT("GetListBanks"),
		TEXT("GetListBenefits"),
		TEXT("GetListEmployees"),
		TEXT("FillListBanks"),
		TEXT("FillListBenefits"),
		TEXT("FillListEmployees"),
		TEXT("IsUserValid"),
		TEXT("ExecFillArray"),
		TEXT("Term")
};

#define REMOTESERVER_METHOD_COUNT 12
DISPID RemoteServer_Dispatch_Ids[REMOTESERVER_METHOD_COUNT];
#define REMOTESERVER_METHOD_INIT 0
#define REMOTESERVER_METHOD_EXECBOOLEANSQL 1
#define REMOTESERVER_METHOD_EXECBOOLEANMDB 2
#define REMOTESERVER_METHOD_GETLISTBANKS 3
#define REMOTESERVER_METHOD_GETLISTBENEFITS 4
#define REMOTESERVER_METHOD_GETLISTEMPLOYEES 5
#define REMOTESERVER_METHOD_FILLLISTBANKS 6
#define REMOTESERVER_METHOD_FILLLISTBENEFITS 7
#define REMOTESERVER_METHOD_FILLLISTEMPLOYEES 8
#define REMOTESERVER_METHOD_ISUSERVALID 9
#define REMOTESERVER_METHOD_EXECFILLARRAY 10
#define REMOTESERVER_METHOD_TERM 11

//globals used to determine when to unload a in-proc or local server.
extern ULONG       g_lObjectCount;
extern ULONG       g_lLockedObjects;

//Delete
//Purpose		- Deletes specified object
//Assumptions	- Only one input object can be specified
//Inputs		- (pvEmployee)		- object with delete information (can be null)
//				- (pvBank)			- object with delete information (can be null)
//				- (pvDependant)		- object with delete information (can be null)
//				- (pvBenefit)		- object with delete information (can be null)
//				- (pvBenefitLink)	- object with delete information (can be null)
//				- (pbSuccess)		- boolean output target
//Outputs		- (pbSuccess)		- boolean success value
STDMETHODIMP CAdmin::Delete(VARIANT * pvEmployee, VARIANT * pvBank, VARIANT * pvDependant, VARIANT * pvBenefit, VARIANT * pvBenefitLink, BOOL * pbSuccess)
{
	HRESULT		hResult;
	_bstr_t		sSQL;
	int			iParameterCount;
	IEmployee *	poEmployee;
	IBank*		poBank;
	IDependant*	poDependant;
	IBenefit *	poBenefit;
	IDispatch * IDisp;
	IBenefitLink * poBenefitLink;
	

	hResult		= NOERROR;
	iParameterCount = 0;
	poEmployee	= NULL;
	poBank		= NULL;
	poDependant = NULL;
	poBenefit	= NULL;
	poBenefitLink = NULL;

	if (!IsOptionalMissing(pvEmployee) && (pvEmployee->vt & VT_DISPATCH == VT_DISPATCH))
	{
		IDisp = *(V_DISPATCHREF(pvEmployee));
		hResult = IDisp->QueryInterface(IID_IEmployee, (void **)&poEmployee);
		if (hResult == NOERROR)
			iParameterCount++;
		else
			return (FlagException(IID_IAdmin, scDel, icErrObjectOp));
	}
	if (!IsOptionalMissing(pvBank) && (pvBank->vt & VT_DISPATCH == VT_DISPATCH))
	{
		IDisp = *(V_DISPATCHREF(pvBank));
		hResult = IDisp->QueryInterface(IID_IBank, (void **)&poBank);
		if (hResult == NOERROR)
			iParameterCount++;
		else
			return (FlagException(IID_IAdmin, scDel, icErrObjectOp));
	}
	if (!IsOptionalMissing(pvDependant) && (pvDependant->vt & VT_DISPATCH == VT_DISPATCH))
	{
		IDisp = *(V_DISPATCHREF(pvDependant));
		hResult = IDisp->QueryInterface(IID_IDependant, (void **)&poDependant);
		if (hResult == NOERROR)
			iParameterCount++;
		else
			return (FlagException(IID_IAdmin, scDel, icErrObjectOp));
	}
	if (!IsOptionalMissing(pvBenefit) && (pvBenefit->vt & VT_DISPATCH == VT_DISPATCH))
	{
		IDisp = *(V_DISPATCHREF(pvBenefit));
		hResult = IDisp->QueryInterface(IID_IBenefit, (void **)&poBenefit);
		if (hResult == NOERROR)
			iParameterCount++;
		else
			return (FlagException(IID_IAdmin, scDel, icErrObjectOp));
	}
	if (!IsOptionalMissing(pvBenefitLink) && (pvBenefitLink->vt & VT_DISPATCH == VT_DISPATCH))
	{
		IDisp = *(V_DISPATCHREF(pvBenefitLink));
		hResult = IDisp->QueryInterface(IID_IBenefitLink, (void **)&poBenefitLink);
		if (hResult == NOERROR)
			iParameterCount++;
		else
			return (FlagException(IID_IAdmin, scDel, icErrObjectOp));
	}

	if (iParameterCount != 1)
		return (FlagException(IID_IAdmin, scDel, icErrObjectOp));
	
	if (poBank)
		hResult = Delete_Bank(poBank, pbSuccess);
	else
	{
		if (poBenefit)
			hResult = Delete_Benefit(poBenefit, pbSuccess);
		else
		{
			if (poBenefitLink)
				hResult = Delete_BenefitLink(poBenefitLink, pbSuccess);
			else
			{
				if (poDependant)
					hResult = Delete_Dependant(poDependant, pbSuccess);
				else 
				{
					if (poEmployee)
						hResult = Delete_Employee(poEmployee, pbSuccess);
				}
			}
		}
	}
	
	return (hResult);
}

//ExecBooleanMDB
//Purpose	- Execute SQL statement with parameter list
//Inputs	- (sQry)			- SQL statement to execute
//			- (lstParameters)	- IList of paramters for SQL
//			- (pbSuccess)		- boolean output target
//Outputs	- (pbSuccess)		- boolean success value
STDMETHODIMP CAdmin::ExecBooleanMDB(BSTR sQry, IDispatch * lstParameters, BOOL * pbSuccess)
{	
	HRESULT			hResult;
	int				iCount;
	int				iParam;
	VARIANT			vTemp;
	VARIANT			vResult;
	BSTR			sParam;
	DISPPARAMS		pDispParams;
	VARIANTARG		vArgs[1];
	_QueryDefPtr	pqryDef;
	ParametersPtr	pParameters;
	ParameterPtr	pParm;
	VARIANTARG		vTempArgs[2];

	hResult		= NOERROR;
	*pbSuccess	= FALSE;

	if (m_iSrcType != icDoDataDirect)
	{
		if (m_pIRemoteBenefitsServer == NULL)
			return (E_FAIL);
		
		for (iParam = 0; iParam < 2; iParam++)
			::VariantInit(&vTempArgs[iParam]);
		
		vTempArgs[0].vt			= VT_DISPATCH | VT_BYREF;
		vTempArgs[0].ppdispVal	= &lstParameters;
		vTempArgs[1].vt			= VT_BSTR;
		vTempArgs[1].bstrVal	= sQry;
		pDispParams.cArgs		= 2;
		pDispParams.cNamedArgs	= 0;
		pDispParams.rgvarg		= vTempArgs;
		
		::VariantInit(&vResult);
		vResult.vt = VT_BOOL;
		
		hResult = m_pIRemoteBenefitsServer->Invoke(
					RemoteServer_Dispatch_Ids[REMOTESERVER_METHOD_EXECBOOLEANMDB],
					IID_NULL,
					LOCALE_USER_DEFAULT,
					DISPATCH_METHOD,
					&pDispParams,
					&vResult,
					&m_RemoteServer_ExceptionInfo,
					NULL);
		
		if (hResult == NOERROR)
			*pbSuccess = vResult.boolVal;
		
		return (hResult);
	}
	
	hResult = m_Database->BeginTrans();

	try
	{
		pqryDef = m_Database->OpenQueryDef(sQry);
		pParameters = pqryDef->GetParameters();
		iCount = pParameters->GetCount();
	}
	catch (...)
	{
		return (NOERROR);
	}
			
	for (iParam = 0; iParam < iCount; iParam++)
	{

		::VariantInit(&vTemp);
		::VariantInit(&vArgs[0]);

		vTemp.vt = VT_I2;
		V_I2(&vTemp) = iParam;
		pParm = pParameters->GetItem(vTemp);

		::VariantClear(&vTemp);
		sParam = ::SysAllocString(pParm->GetName());
	
		vArgs[0].vt				= VT_BSTR;
		vArgs[0].bstrVal		= sParam;
		pDispParams.cArgs		= 1;
		pDispParams.cNamedArgs	= 0;
		pDispParams.rgvarg		= vArgs;
	
		::VariantInit(&vResult);
		vResult.vt = VT_EMPTY;

		hResult = lstParameters->Invoke(
							0,			// DISID 0 - List Property Item
							IID_NULL,
							LOCALE_USER_DEFAULT,
							DISPATCH_PROPERTYGET,
							&pDispParams,
							&vResult,
							&m_RemoteServer_ExceptionInfo,
							NULL);
		
		pParm->Value = vResult;
		pParm->Release();
		::VariantClear(&vResult);
		
		::SysFreeString(sParam);
	}
	
	::VariantInit(&vTemp);
	vTemp.vt = VT_I4;
	V_I4(&vTemp) = dbFailOnError + dbSeeChanges;
	
	try
	{
		hResult = pqryDef->Execute(vTemp);
	}
	catch (...)
	{
		return (NOERROR);
	}
	
	if (hResult != NOERROR)
		hResult = m_Database->Rollback();
	else
	{
		hResult = m_Database->CommitTrans(dbForceOSFlush);
		*pbSuccess = TRUE;
	}
	
	if (hResult != NOERROR)
		return (hResult);

	return (NOERROR);
}

//ExecBooleanSQL
//Purpose	- Execute SQL statement
//Inputs	- (sQry)			- SQL statement to execute
//			- (pbSuccess)		- boolean output target
//Outputs	- (pbSuccess)		- boolean success value
STDMETHODIMP CAdmin::ExecBooleanSQL(BSTR sQry, BOOL * pbSuccess)
{
	HRESULT			hResult;
	int				iParam;
	VARIANT			vTemp;
	VARIANT			vResult;
	BSTR			sExecute;
	DISPPARAMS		pDispParams;
	VARIANTARG		vArgs[1];
	
	hResult		= NOERROR;
	*pbSuccess	= FALSE;

	if (m_iSrcType != icDoDataDirect)
	{
		if (m_pIRemoteBenefitsServer == NULL)
			return (E_FAIL);

		for (iParam = 0; iParam < 1; iParam++)
			::VariantInit(&vArgs[iParam]);
		
		vArgs[0].vt = VT_BSTR;
		vArgs[0].bstrVal = sQry;
		pDispParams.cArgs = 1;
		pDispParams.cNamedArgs = 0;
		pDispParams.rgvarg = vArgs;
		
		::VariantInit(&vResult);
		vResult.vt = VT_BOOL;
		
		hResult = m_pIRemoteBenefitsServer->Invoke(
							RemoteServer_Dispatch_Ids[REMOTESERVER_METHOD_EXECBOOLEANSQL],
							IID_NULL,
							LOCALE_USER_DEFAULT,
							DISPATCH_METHOD,
							&pDispParams,
							&vResult,
							&m_RemoteServer_ExceptionInfo,
							NULL);
		
		if (hResult == NOERROR)
			*pbSuccess = vResult.boolVal;

		return (hResult);
	}
	
	hResult = m_Database->BeginTrans();
	if (hResult != NOERROR)
		return (hResult);

	::VariantInit(&vTemp);
	vTemp.vt = VT_I4;
	V_I4(&vTemp) = dbFailOnError + dbSQLPassThrough;
	
	sExecute = ::SysAllocString(sQry);

	hResult = m_Database->Execute(sExecute, vTemp);
	
	::SysFreeString(sExecute);

	if (hResult != NOERROR)
		hResult = m_Database->Rollback();
	else
	{
		hResult = m_Database->CommitTrans(dbForceOSFlush);
		*pbSuccess = TRUE;
	}
	
	if (hResult != NOERROR)
		return (hResult);

	return (NOERROR);
}

//ExecFillArray
//Purpose		- Used to execute a SQL statement and return an array filled with results
//Inputs		- (sQry)			- SQL statement to execute
//				- {pUserSafeArray)	- pointer to SAFEARRAY to be filled
//				- {bAddColumnName)	- boolean indicating whether or not to include column names in results
//				- (pbSuccess)	- boolean output target
//Outputs		- (pbSuccess)	- boolean success value
STDMETHODIMP CAdmin::ExecFillArray(BSTR sQry, VARIANT * pUserSafeArray, BOOL bAddColumnName, BOOL * pbSuccess)
{
	RecordsetPtr	pRecordset;
	FieldsPtr		pFields;

	HRESULT			hResult;
	CSafeArray	  * pSafeArray;
	BOOL			bArrayCreated;
	BSTR			bstrResult;
	BOOL			bBOF;
	BOOL			bEOF;
	long			lCols;
	long			lCurCol;
	long			lCurRow;
	BOOL			bFirstRecord;

	DISPPARAMS		pDispParams;
	VARIANTARG		vArgs[4];
	int				iParam;
	VARIANT			vResult;
	long			lBounds[2];
	SAFEARRAY	  * psa;

	VARIANT			Type;
	VARIANT			Options;
	
	hResult			= NOERROR;
	*pbSuccess		= FALSE;
	bArrayCreated	= FALSE;
	bFirstRecord	= TRUE;

	if (sQry == NULL)
		return (NOERROR);

	if (::SysStringLen(sQry) == 0)
		return (NOERROR);

	if (m_iSrcType != icDoDataDirect)
	{
		if (m_pIRemoteBenefitsServer == NULL)
			return (E_FAIL);
				
		for (iParam = 0; iParam < 4; iParam++)
			::VariantInit(&vArgs[iParam]);
		
		pDispParams.rgvarg = vArgs;
		pDispParams.cNamedArgs = 0;
		pDispParams.cArgs = 3;
		
		vArgs[2].vt = VT_BSTR;
		vArgs[2].bstrVal = sQry;
		
		lBounds[0] = 0;
		lBounds[1] = 0;
		pSafeArray = new CSafeArray(VT_BSTR, 2, lBounds);
		psa = pSafeArray->GetPSA();
		
		vArgs[1].vt = VT_BSTR | VT_ARRAY | VT_BYREF;
		V_ARRAYREF(&vArgs[1]) = &psa;
		
		vArgs[0].vt = VT_BOOL;
		vArgs[0].boolVal = bAddColumnName;
		
		::VariantInit(&vResult);
		vResult.vt = VT_EMPTY;
		
		hResult = m_pIRemoteBenefitsServer->Invoke(
			RemoteServer_Dispatch_Ids[REMOTESERVER_METHOD_EXECFILLARRAY],
			IID_NULL,
			LOCALE_USER_DEFAULT,
			DISPATCH_METHOD,
			&pDispParams,
			&vResult,
			&m_RemoteServer_ExceptionInfo,
			NULL);
		
		if (hResult != NOERROR)
			return (hResult);
		
		if (pUserSafeArray->vt & VT_BYREF)
			hResult = ::SafeArrayCopy(psa, V_ARRAYREF(pUserSafeArray));
		else
			hResult = ::SafeArrayCopy(psa, &V_ARRAY(pUserSafeArray));
		
		delete pSafeArray;
		
		if (hResult != NOERROR)
			return (hResult);
		
		*pbSuccess = vResult.boolVal;
		return (NOERROR);
	}

	try {
		if (m_bUseODBC)
		{
			Type.vt = VT_I4;
			V_I4(&Type) = dbOpenSnapshot;
			Options.vt = VT_I4;
			V_I4(&Options) = dbForwardOnly + dbSQLPassThrough;
			pRecordset = m_Database->OpenRecordset(sQry, Type, Options);
		}
		else
		{
			pRecordset = m_Database->OpenRecordset(sQry);
		}
	}
	catch (...)
	{
		pRecordset = NULL;
	}

	// if we have a recordset
	if (pRecordset)
	{
		// get flags
		bBOF = pRecordset->GetBOF();
		bEOF = pRecordset->GetEOF();
		if (!(bBOF && bEOF))
		{
			// find size of data in recordset
			pRecordset->MoveLast(0);
			pRecordset->MoveFirst();
			lCurRow = pRecordset->GetRecordCount();
			
			// if we have data
			if (lCurRow)
			{
				// get fields object
				pFields = pRecordset->GetFields();
				
				// get # of fields
				lCols = pFields->GetCount();
				
				// create array (rows * columns) to hold data
				{
					lBounds[0] = lCols;
					lBounds[1] = lCurRow;
					if (bAddColumnName)
						lBounds[1]++;
					pSafeArray = new CSafeArray(VT_BSTR, 2, lBounds);
					bArrayCreated = TRUE;
				}
				
				// fill array starting at first row
				lCurRow = 0;
			
				// process recordset
				do
				{
					pFields = pRecordset->GetFields();
					{
						// for each field
						for (lCurCol = 0; lCurCol < lCols; lCurCol++) 
						{
							// get field item object
							_FieldPtr pFieldItem = pFields->GetItem(lCurCol);
							{
								// get value fo field
								vResult = pFieldItem->GetValue();
								
								// convert everything to string
								bstrResult = ::SysAllocString(this->VariantXToBstr(&vResult));
								if (bstrResult)
								{
									// store in array
									pSafeArray->bstrPut(bstrResult, (long)lCurCol, lCurRow);
									::SysFreeString(bstrResult);
								}
								if (bAddColumnName && bFirstRecord)
								{
									// get value fo field
									bstrResult = ::SysAllocString(pFieldItem->GetName());
									if (bstrResult)
									{
										// store in array
										pSafeArray->bstrPut(bstrResult, (long)lCurCol, 1);
										::SysFreeString(bstrResult);
									}
								}
							}
						}
					}
					// delete resource
					// move recordset cursor to next record
					pRecordset->MoveNext();
					bFirstRecord = FALSE;
					if (bAddColumnName)
						break;
					
					// get eof flag
					bEOF = pRecordset->GetEOF();
					lCurRow++;
				} while (!bEOF);
			
				// copy array to user and destroy local copy
				if (pUserSafeArray->vt & VT_BYREF)
					hResult = ::SafeArrayCopy(pSafeArray->GetPSA(), V_ARRAYREF(pUserSafeArray));
				else
					hResult = ::SafeArrayCopy(pSafeArray->GetPSA(), &V_ARRAY(pUserSafeArray));
				
				if (hResult == NOERROR)
					*pbSuccess = TRUE;
				
				// delete resource
				delete pSafeArray;
			}
		}
	
		// if no data then create empty array to pass back
		if (!bArrayCreated)
		{
			lBounds[0] = 0;
			lBounds[1] = 0;
			pSafeArray = new CSafeArray(VT_BSTR, 2, lBounds);
			if (pUserSafeArray->vt & VT_BYREF)
				hResult = ::SafeArrayCopy(pSafeArray->GetPSA(), V_ARRAYREF(pUserSafeArray));
			else
				hResult = ::SafeArrayCopy(pSafeArray->GetPSA(), &V_ARRAY(pUserSafeArray));
			
			delete pSafeArray;
			
			if (hResult == NOERROR)
				*pbSuccess = TRUE;
		}
		
		// close recordset and resource
		pRecordset->Close();
	}
	
	return (NOERROR);
}
	
//ExecFillList
//Purpose		- Used to execute a SQL statement and return an IList filled with results
//Inputs		- (sQry)		- SQL statement to execute
//				- {ppcOutput)	- pointer to IList object to be filled
//				- (pbSuccess)	- boolean output target
//Outputs		- (pbSuccess)	- boolean success value
STDMETHODIMP CAdmin::ExecFillList(BSTR sQry, IList ** ppcOutput, BOOL * pbSuccess)
{
	ASSERT(pbSuccess != NULL);
	
	HRESULT hResult;
	VARIANT v;
	VARIANT vValue;
	BSTR	bstrName;
	BSTR	bstrValue;
	long	lMin;
	long	lMax;
	BOOL	bValid;
	int		i;

	hResult		= NOERROR;
	*pbSuccess	= FALSE;
	
	if (*ppcOutput)
	try
	{
		(*ppcOutput)->Release();
	}
	catch (...)
	{
	}

	hResult = ::CoCreateInstance(CLSID_CList, NULL, THISCLSCTX, IID_IList, (void **)ppcOutput);

	if (hResult != NOERROR)
		return (hResult);
	
	// fill the array
	::VariantInit(&v);
	v.vt = VT_BSTR | VT_ARRAY;
	hResult = ExecFillArray(sQry, &v, TRUE, &bValid);
	if (hResult != NOERROR)
		return (hResult);
	
	if (bValid)
	{
		CSafeArray SA(V_ARRAY(&v));
		lMin = SA.LBound(1);
		lMax = SA.UBound(1);
		::VariantInit(&vValue);
		for (i = lMin; i < lMax; i++)
		{
			bstrName	= ::SysAllocString(SA.bstrGet(i, 1));
			bstrValue	= ::SysAllocString(SA.bstrGet(i, 0));
			vValue.vt	= VT_BSTR;
			V_BSTR(&vValue) = bstrValue;
			hResult = (*ppcOutput)->Add(bstrName, &vValue);
			::VariantClear(&vValue);
			::SysFreeString(bstrName);
			if (hResult != NOERROR)
				return (hResult);
		}
		*pbSuccess = TRUE;
	}
	
	return (NOERROR);
}

//FillListBanks
//Purpose		- Used to fill a list object with a current list of banks
STDMETHODIMP CAdmin::FillListBanks()
{
	BSTR		sSQL;
	HRESULT		hResult;
	BOOL		bValid;
	SAFEARRAY * psa;
	VARIANT		vTemp;
	VARIANT		vItem;
	IBank *		pBank;
	ULONG		uCount;
	long		iIndex;
	long		lPKId;
	long		lMin;
	long		lMax;
	CStrFnc *	bstrTemp;
	DISPPARAMS	dispparamsNoArgs;
	CIEnumVariant * pForEach;

	dispparamsNoArgs.rgvarg = NULL;				// Array of arguments.
	dispparamsNoArgs.rgdispidNamedArgs = NULL;	// Dispatch IDs of named arguments.
	dispparamsNoArgs.cArgs = 0;					// Number of arguments.
	dispparamsNoArgs.cNamedArgs = 0;
	
	hResult = NOERROR;
	uCount = 0;
	
	// invoke remote server if doing remoting
	if (m_iSrcType == icDoDataDirect)
	{
		// initialize variant to BSTR array
		::VariantInit(&vTemp);
		vTemp.vt = VT_BSTR | VT_ARRAY;
		// get SQL
		if (m_bUseODBC)
			sSQL = CStrFnc::LoadResString(scLvw_Bank_SQL);
		else
			sSQL = CStrFnc::LoadResString(scLvw_Bank_MDB);
		
		hResult = ExecFillArray(sSQL, &vTemp, false, &bValid);
		if (hResult != NOERROR)
			return (hResult);
		
		::SysFreeString(sSQL);
		
		if (!bValid)
			return (NOERROR);
		
		psa = V_ARRAY(&vTemp);
	}
	else
	{
		if (m_pIRemoteBenefitsServer == NULL)
			return (E_FAIL);

		::VariantInit(&vTemp);
		vTemp.vt = VT_EMPTY;
		
		hResult = m_pIRemoteBenefitsServer->Invoke(
			RemoteServer_Dispatch_Ids[REMOTESERVER_METHOD_GETLISTBANKS],
			IID_NULL,
			LOCALE_USER_DEFAULT,
			DISPATCH_METHOD,
			&dispparamsNoArgs,
			&vTemp,
			&m_RemoteServer_ExceptionInfo,
			NULL);
		
		if (hResult != NOERROR)
			return (hResult);
		
		psa = V_ARRAY(&vTemp);
	}
	
	::VariantInit(&vItem);

	if (m_Banks)
	{
		hResult = m_Banks->get__NewEnum((IUnknown **)&pForEach);
		::VariantClear(&vItem);
		hResult = pForEach->Next(1, &vItem, &uCount);
		while (hResult == NOERROR)
		{
			pBank = (IBank *)V_DISPATCH(&vItem);
			pBank->Term();
			pBank->Release();
			::VariantClear(&vItem);
			hResult = pForEach->Next(1, &vItem, &uCount);
		}
		m_Banks->Release();
		m_Banks = NULL;
	}
	
	CSafeArray SA(psa);
	
	hResult = ::CoCreateInstance(CLSID_CListBank, NULL, THISCLSCTX, IID_IListBank, (void **)&m_Banks);
	if (hResult != NOERROR)
		return (hResult);
	
	lMin = SA.LBound(2);
	lMax = SA.UBound(2);
	for (iIndex = lMin; iIndex < lMax; iIndex++)
	{
		hResult = ::CoCreateInstance(CLSID_CBank, NULL, THISCLSCTX, IID_IBank, (void **)&pBank);
		if (hResult == NOERROR)
		{
			bstrTemp = new CStrFnc(SA.bstrGet((long)icLvwPKId, iIndex));
			lPKId = bstrTemp->Val();
			delete bstrTemp;
			
			pBank->putref_Admin(this);
			pBank->put_PKId(lPKId);
			m_Banks->Add(pBank);
			pBank->Release();
		}
	}
	
	return (NOERROR);
}

//FillListBenefits
//Purpose		- Used to fill a list object with a current list of benefits
STDMETHODIMP CAdmin::FillListBenefits()
{
	BSTR		sSQL;
	HRESULT		hResult;
	BOOL		bValid;
	SAFEARRAY * psa;
	VARIANT		v;
	IBenefit *	pBenefit;
	long		lPKId;
	DISPPARAMS	dispparamsNoArgs;
	long		lCount;
	VARIANT		varTemp;
	long		i;
	long		lMin;
	long		lMax;
	CStrFnc *	bstrTemp;

	hResult	= NOERROR;
	lCount	= 0;
	dispparamsNoArgs.rgvarg = NULL;				// Array of arguments.
	dispparamsNoArgs.rgdispidNamedArgs = NULL;	// Dispatch IDs of named arguments.
	dispparamsNoArgs.cArgs = 0;					// Number of arguments.
	dispparamsNoArgs.cNamedArgs = 0;
	

	::VariantInit(&v);
	if (m_iSrcType == icDoDataDirect)
	{
		v.vt = VT_BSTR | VT_ARRAY;
		
		if (m_bUseODBC)
			sSQL = CStrFnc::LoadResString(scLvw_Benefit_SQL);
		else
			sSQL = CStrFnc::LoadResString(scLvw_Benefit_MDB);
		
		hResult = ExecFillArray(sSQL, &v, false, &bValid);
		if (hResult != NOERROR)
			return (hResult);
		
		::SysFreeString(sSQL);
		
		if (!bValid)
			return (NOERROR);
		
		psa = V_ARRAY(&v);
	}
	else
	{
		if (m_pIRemoteBenefitsServer == NULL)
			return (E_FAIL);

		::VariantInit(&v);
		v.vt = VT_EMPTY;
		
		hResult = m_pIRemoteBenefitsServer->Invoke(
			RemoteServer_Dispatch_Ids[REMOTESERVER_METHOD_GETLISTBENEFITS],
			IID_NULL,
			LOCALE_USER_DEFAULT,
			DISPATCH_METHOD,
			&dispparamsNoArgs,
			&v,
			&m_RemoteServer_ExceptionInfo,
			NULL);
		
		if (hResult != NOERROR)
			return (hResult);
		
		psa = V_ARRAY(&v);
	}
	
	::VariantInit(&varTemp);

	if (m_Benefits)
	{
		m_Benefits->Release();
		m_Benefits = NULL;
	}
	
	CSafeArray SA(psa);
	
	hResult = ::CoCreateInstance(CLSID_CListBenefit, NULL, THISCLSCTX, IID_IListBenefit, (void **)&m_Benefits);
	if (hResult != NOERROR)
		return (hResult);

	lMin = SA.LBound(2);
	lMax = SA.UBound(2);
	for (i = lMin; i < lMax; i++)
	{
		hResult = ::CoCreateInstance(CLSID_CBenefit, NULL, THISCLSCTX, IID_IBenefit, (void **)&pBenefit);

		bstrTemp = new CStrFnc(SA.bstrGet(icLvwPKId, i));
		lPKId = bstrTemp->Val();
		delete bstrTemp;
		
		pBenefit->putref_Admin(this);
		pBenefit->put_PKId(lPKId);
		m_Benefits->Add(pBenefit);
		pBenefit->Release();
	}
	
	return (NOERROR);
}

//FillListEmployees
//Purpose		- Used to fill a list object with a current list of employees
STDMETHODIMP CAdmin::FillListEmployees()
{
	BSTR		sSQL;
	HRESULT		hResult;
	BOOL		bValid;
	SAFEARRAY *	psa;
	VARIANT		v;
	IEmployee *	pEmployee;
	DISPPARAMS	dispparamsNoArgs;
	CIEnumVariant * pForEach;
	long		lCount;
	ULONG		uCount;
	VARIANT		varTemp;
	long		i;
	long		lPKId;
	long		lMin;
	long		lMax;
	CStrFnc *	bstrTemp;
	
	hResult = NOERROR;
	lCount	= 0;
	dispparamsNoArgs.rgvarg = NULL;				// Array of arguments.
	dispparamsNoArgs.rgdispidNamedArgs = NULL;	// Dispatch IDs of named arguments.
	dispparamsNoArgs.cArgs = 0;					// Number of arguments.
	dispparamsNoArgs.cNamedArgs = 0;

	if (m_iSrcType == icDoDataDirect)
	{
		::VariantInit(&v);
		v.vt = VT_BSTR | VT_ARRAY;
		
		if (m_bUseODBC)
			sSQL = CStrFnc::LoadResString(scLvw_Employee_SQL);
		else
			sSQL = CStrFnc::LoadResString(scLvw_Employee_MDB);
		
		hResult = ExecFillArray(sSQL, &v, false, &bValid);
		if (hResult != NOERROR)
			return (hResult);
		
		::SysFreeString(sSQL);
		
		if (!bValid)
			return (NOERROR);
		
		psa = V_ARRAY(&v);
	}
	else
	{
		if (m_pIRemoteBenefitsServer == NULL)
			return (E_FAIL);

		::VariantInit(&v);
		v.vt = VT_EMPTY;

		hResult = m_pIRemoteBenefitsServer->Invoke(
			RemoteServer_Dispatch_Ids[REMOTESERVER_METHOD_GETLISTEMPLOYEES],
			IID_NULL,
			LOCALE_USER_DEFAULT,
			DISPATCH_METHOD,
			&dispparamsNoArgs,
			&v,
			&m_RemoteServer_ExceptionInfo,
			NULL);
		
		if (hResult != NOERROR)
			return (hResult);
		
		psa = V_ARRAY(&v);
	}
	
	::VariantInit(&varTemp);

	if (m_Employees)
	{
		hResult = m_Employees->get__NewEnum((IUnknown **)&pForEach);
		::VariantClear(&varTemp);
		hResult = pForEach->Next(1, &varTemp, &uCount);
		while (hResult == NOERROR)
		{
			pEmployee = (IEmployee *)V_DISPATCH(&varTemp);
			pEmployee->Term();
			pEmployee->Release();
			::VariantClear(&varTemp);
			hResult = pForEach->Next(1, &varTemp, &uCount);
		}
		m_Employees->Release();
		m_Employees = NULL;
	}
	
	CSafeArray SA(psa);
	
	hResult = ::CoCreateInstance(CLSID_CListEmployee, NULL, THISCLSCTX, IID_IListEmployee, (void **)&m_Employees);
	if (hResult != NOERROR)
		return (hResult);
	
	lMin = SA.LBound(2);
	lMax = SA.UBound(2);
	for (i = lMin; i < lMax; i++)
	{
		hResult = ::CoCreateInstance(CLSID_CEmployee, NULL, THISCLSCTX, IID_IEmployee, (void **)&pEmployee);

		bstrTemp = new CStrFnc(SA.bstrGet(icLvwPKId, i));
		lPKId = bstrTemp->Val();
		delete bstrTemp;
		
		pEmployee->putref_Admin(this);
		pEmployee->put_PKId(lPKId);
		m_Employees->Add(pEmployee);
		pEmployee->Release();
	}
	
	return (NOERROR);
}

//GetBenefitTaxTypes
//Purpose	- Fills an IList object, according to predetermined rules,  with a particular recordset
//Inputs	- (ppcOutput)		- pointer to IList object to be filled
//			- (pbSuccess)		- boolean output target
//Outputs	- (pbSuccess)		- boolean success value
STDMETHODIMP CAdmin::GetBenefitTaxTypes(IList * * ppcOutput, BOOL * pbSuccess)
{
	BSTR	sSQL;
	BOOL	bRetVal;
	HRESULT	hResult;
	
	hResult		= NOERROR;
	*pbSuccess	= FALSE;

	if (m_bUseODBC)
		sSQL = CStrFnc::LoadResString(scLvw_BnfTax_SQL);
	else
		sSQL = CStrFnc::LoadResString(scLvw_BnfTax_MDB);
	
	if (*ppcOutput)
		try
		{
			(*ppcOutput)->Release();
		}
		catch (...)
		{
		}
		
	hResult = ::CoCreateInstance(CLSID_CList, NULL, THISCLSCTX, IID_IList, (void **)ppcOutput);
	if (hResult == NOERROR)
		hResult = GetDomain(sSQL, ppcOutput, &bRetVal);

	::SysFreeString(sSQL);
	
	if (hResult != NOERROR)
		return (hResult);

	*pbSuccess = bRetVal;
	return (NOERROR);
}

//GetBenefitTypes
//Purpose	- Fills an IList object, according to predetermined rules,  with a particular recordset
//Inputs	- (ppcOutput)		- pointer to IList object to be filled
//			- (pbSuccess)		- boolean output target
//Outputs	- (pbSuccess)		- boolean success value
STDMETHODIMP CAdmin::GetBenefitTypes(IList * * ppcOutput, BOOL * pbSuccess)
{
	BSTR	sSQL;
	BOOL	bRetVal;
	HRESULT	hResult;
	
	hResult = NOERROR;
	*pbSuccess = FALSE;

	if (m_bUseODBC)
		sSQL = CStrFnc::LoadResString(scLvw_BnfType_SQL);
	else
		sSQL = CStrFnc::LoadResString(scLvw_BnfType_MDB);
	
	if (*ppcOutput)
		try
		{
			(*ppcOutput)->Release();
		}
		catch (...)
		{
		}
		
	hResult = ::CoCreateInstance(CLSID_CList, NULL, THISCLSCTX, IID_IList, (void **)ppcOutput);
	if (hResult == NOERROR)
		hResult = GetDomain(sSQL, ppcOutput, &bRetVal);

	::SysFreeString(sSQL);
	
	if (hResult != NOERROR)
		return (hResult);

	*pbSuccess = bRetVal;
	return (NOERROR);
}

//GetDependantTypes
//Purpose	- Fills an IList object, according to predetermined rules,  with a particular recordset
//Inputs	- (ppcOutput)		- pointer to IList object to be filled
//			- (pbSuccess)		- boolean output target
//Outputs	- (pbSuccess)		- boolean success value
STDMETHODIMP CAdmin::GetDependantTypes(IList ** ppcOutput, BOOL * pbSuccess)
{
	BSTR sSQL;
	BOOL bRetVal;
	HRESULT hResult;
	
	hResult = NOERROR;

	if (m_bUseODBC)
		sSQL = CStrFnc::LoadResString(scLvw_DepType_SQL);
	else
		sSQL = CStrFnc::LoadResString(scLvw_DepType_MDB);
	
	if (*ppcOutput)
		try
		{
			(*ppcOutput)->Release();
		}
		catch (...)
		{
		}

	hResult = ::CoCreateInstance(CLSID_CList, NULL, THISCLSCTX, IID_IList, (void **)ppcOutput);
	if (hResult == NOERROR)
		hResult = GetDomain(sSQL, ppcOutput, &bRetVal);
	
	::SysFreeString(sSQL);
	
	*pbSuccess = bRetVal;
	return (NOERROR);
}

//GetDomain
//Purpose	- Execute SQL statement and fill list with results
//Inputs	- (sQry)			- SQL statement to execute
//			- (ppOut)			- pointer to variable to hold new list
//			- (pbSuccess)		- boolean output target
//Outputs	- (pbSuccess)		- boolean success value
STDMETHODIMP CAdmin::GetDomain(BSTR sSQL, IList ** ppOut, BOOL * pbSuccess)
{
	BSTR	sCapDesc;
	BSTR	sCapPKId;
	BSTR	sCapCode;
	VARIANT	v;
	BOOL	bValid;
	HRESULT hResult;
	IList * plstItem;
	long	lMin;
	long	lMax;
	long	i;
	long	lPKId;
	CStrFnc * bstrTemp;
	
	hResult		= NOERROR;
	*pbSuccess	= FALSE;
	sCapDesc = CStrFnc::LoadResString(scLvwDesc);
	sCapPKId = CStrFnc::LoadResString(scLvwPKId);
	sCapCode = CStrFnc::LoadResString(scLvwCode);
	
	::VariantInit(&v);
	v.vt = VT_BSTR | VT_ARRAY;
	
	hResult = ExecFillArray(sSQL, &v, false, &bValid);
	if (hResult != NOERROR)
		return (hResult);
	
	if (bValid)
	{
		if (*ppOut)
			try
			{
				(*ppOut)->Release();
			}
			catch (...)
			{
			}

		hResult = ::CoCreateInstance(CLSID_CList, NULL, THISCLSCTX, IID_IList, (void **)ppOut);
		
		CSafeArray SA(V_ARRAY(&v));
		lMin = SA.LBound(2);
		lMax = SA.UBound(2);
		for (i = lMin; i < lMax; i++)
		{
			hResult = ::CoCreateInstance(CLSID_CList, NULL, THISCLSCTX, IID_IList, (void **)&plstItem);
			
			::VariantInit(&v);
			v.vt = VT_BSTR;
			V_BSTR(&v) = SA.bstrGet((long)icLvwDesc, i);
			plstItem->Add(sCapDesc,&v);
			
			bstrTemp = new CStrFnc(SA.bstrGet(icLvwPKId, i));
			lPKId = bstrTemp->Val();
			delete bstrTemp;
			::VariantInit(&v);
			v.vt = VT_I4;
			V_I4(&v) = lPKId;
			plstItem->Add(sCapPKId,&v);
			
			::VariantInit(&v);
			v.vt = VT_BSTR;
			V_BSTR(&v) = SA.bstrGet(icLvwCode, i);
			plstItem->Add(sCapCode,&v);
			
			bstrTemp = new CStrFnc(CList::GenListKey(lPKId));
			::VariantInit(&v);
			v.vt = VT_DISPATCH;
			V_DISPATCH(&v) = plstItem;
			(*ppOut)->Add(bstrTemp->GetBstr(), &v);
			delete bstrTemp;
			
			plstItem->Release();
		}
		*pbSuccess = TRUE;
	}
	::SysFreeString(sCapDesc);
	::SysFreeString(sCapPKId);
	::SysFreeString(sCapCode);
	
	return (NOERROR);
}

//GetEmployeeBenefitCoverageTypes
//Purpose	- Fills an IList object, according to predetermined rules,  with a particular recordset
//Inputs	- (ppcOutput)		- pointer to IList object to be filled
//			- (pbSuccess)		- boolean output target
//Outputs	- (pbSuccess)		- boolean success value
STDMETHODIMP CAdmin::GetEmployeeBenefitCoverageTypes(IList * * ppcOutput, BOOL * pbSuccess)
{
	BSTR	sSQL;
	BOOL	bRetVal;
	HRESULT hResult;
	
	hResult = NOERROR;

	if (m_bUseODBC)
		sSQL = CStrFnc::LoadResString(scLvw_EmplBnfCoverage_SQL);
	else
		sSQL = CStrFnc::LoadResString(scLvw_EmplBnfCoverage_MDB);
	
	if (*ppcOutput)
		try
		{
			(*ppcOutput)->Release();
		}
		catch (...)
		{
		}

	hResult = ::CoCreateInstance(CLSID_CList, NULL, THISCLSCTX, IID_IList, (void **)ppcOutput);
	if (hResult == NOERROR)
		hResult = GetDomain(sSQL, ppcOutput, &bRetVal);
	
	::SysFreeString(sSQL);
	
	*pbSuccess = bRetVal;
	return (NOERROR);
}

//GetEmployeeBenefitTaxTypes
//Purpose	- Fills an IList object, according to predetermined rules,  with a particular recordset
//Inputs	- (ppcOutput)		- pointer to IList object to be filled
//			- (pbSuccess)		- boolean output target
//Outputs	- (pbSuccess)		- boolean success value
STDMETHODIMP CAdmin::GetEmployeeBenefitTaxTypes(IList ** ppcOutput, BOOL * pbSuccess)
{
	BSTR	sSQL;
	BOOL	bRetVal;
	HRESULT hResult;
	
	hResult = NOERROR;

	if (m_bUseODBC)
		sSQL = CStrFnc::LoadResString(scLvw_EmplBnfTax_SQL);
	else
		sSQL = CStrFnc::LoadResString(scLvw_EmplBnfTax_MDB);
	
	if (*ppcOutput)
		try
		{
			(*ppcOutput)->Release();
		}
		catch (...)
		{
		}

	hResult = ::CoCreateInstance(CLSID_CList, NULL, THISCLSCTX, IID_IList, (void **)ppcOutput);
	if (hResult == NOERROR)
		hResult = GetDomain(sSQL, ppcOutput, &bRetVal);
	
	::SysFreeString(sSQL);
	
	*pbSuccess = bRetVal;
	return (NOERROR);
}

//GetListBanks
//Purpose	- Used to fill a given SAFEARRAY with a list of banks
//Inputs	- (pvSARRAY)	- pointer to SAEARRAY to hold results
STDMETHODIMP CAdmin::GetListBanks(VARIANT * pvSARRAY)
{
	BSTR	sSQL;
	HRESULT hResult = 0;
	BOOL	bValid;
	VARIANT v;
	DISPPARAMS dispparamsNoArgs;
	
	hResult = NOERROR;
	dispparamsNoArgs.rgvarg = NULL;				// Array of arguments.
	dispparamsNoArgs.rgdispidNamedArgs = NULL;	// Dispatch IDs of named arguments.
	dispparamsNoArgs.cArgs = 0;					// Number of arguments.
	dispparamsNoArgs.cNamedArgs = 0;

	::VariantInit(&v);
	if (m_iSrcType == icDoDataDirect)
	{
		v.vt = VT_BSTR | VT_ARRAY;
		
		// get SQL
		if (m_bUseODBC)
			sSQL = CStrFnc::LoadResString(scLvw_Bank_SQL);
		else
			sSQL = CStrFnc::LoadResString(scLvw_Bank_MDB);
		
		hResult = ExecFillArray(sSQL, &v, false, &bValid);
		
		::SysFreeString(sSQL);
		
		if (hResult != NOERROR)
			return (hResult);

		if (!bValid)
			return (NOERROR);
	}
	else
	{
		if (m_pIRemoteBenefitsServer == NULL)
			return (E_FAIL);

		::VariantInit(&v);
		v.vt = VT_EMPTY;
		
		hResult = m_pIRemoteBenefitsServer->Invoke(
			RemoteServer_Dispatch_Ids[REMOTESERVER_METHOD_GETLISTBANKS],
			IID_NULL,
			LOCALE_USER_DEFAULT,
			DISPATCH_METHOD,
			&dispparamsNoArgs,
			&v,
			&m_RemoteServer_ExceptionInfo,
			NULL);
		
		if (hResult != NOERROR)
			return (hResult);
	}
	
	if (pvSARRAY->vt & VT_BYREF)
		hResult = ::SafeArrayCopy(V_ARRAY(&v), V_ARRAYREF(pvSARRAY));
	else
		hResult = ::SafeArrayCopy(V_ARRAY(&v), &V_ARRAY(pvSARRAY));

	if (hResult != NOERROR)
		return (hResult);

	pvSARRAY->vt = v.vt;
	
	return (NOERROR);
}

//GetListBenefits
//Purpose	- Used to fill a given SAFEARRAY with a list of benefits
//Inputs	- (pvSARRAY)	- pointer to SAEARRAY to hold results
STDMETHODIMP CAdmin::GetListBenefits(VARIANT * pvSARRAY)
{
	BSTR	sSQL;
	HRESULT	hResult;
	BOOL	bValid;
	VARIANT	v;
	DISPPARAMS dispparamsNoArgs;
	
	hResult = NOERROR;
	dispparamsNoArgs.rgvarg = NULL;				// Array of arguments.
	dispparamsNoArgs.rgdispidNamedArgs = NULL;	// Dispatch IDs of named arguments.
	dispparamsNoArgs.cArgs = 0;					// Number of arguments.
	dispparamsNoArgs.cNamedArgs = 0;

	::VariantInit(&v);
	if (m_iSrcType == icDoDataDirect)
	{
		v.vt = VT_BSTR | VT_ARRAY ;
		
		// get SQL
		if (m_bUseODBC)
			sSQL = CStrFnc::LoadResString(scLvw_Benefit_SQL);
		else
			sSQL = CStrFnc::LoadResString(scLvw_Benefit_MDB);
		
		hResult = ExecFillArray(sSQL, &v, false, &bValid);
		if (hResult != NOERROR)
			return (hResult);
		
		::SysFreeString(sSQL);
		
		if (!bValid)
			return (NOERROR);
	}
	else
	{
		if (m_pIRemoteBenefitsServer == NULL)
			return (E_FAIL);

		::VariantInit(&v);
		v.vt = VT_EMPTY;
		
		hResult = m_pIRemoteBenefitsServer->Invoke(
			RemoteServer_Dispatch_Ids[REMOTESERVER_METHOD_GETLISTBENEFITS],
			IID_NULL,
			LOCALE_USER_DEFAULT,
			DISPATCH_METHOD,
			&dispparamsNoArgs,
			&v,
			&m_RemoteServer_ExceptionInfo,
			NULL);
		
		if (hResult != NOERROR)
			return (hResult);
	}
	
	if (pvSARRAY->vt & VT_BYREF)
		hResult = ::SafeArrayCopy(V_ARRAY(&v), V_ARRAYREF(pvSARRAY));
	else
		hResult = ::SafeArrayCopy(V_ARRAY(&v), &V_ARRAY(pvSARRAY));
	pvSARRAY->vt = v.vt;
	
	return (NOERROR);
}

//GetListEmployees
//Purpose	- Used to fill a given SAFEARRAY with a list of employees
//Inputs	- (pvSARRAY)	- pointer to SAEARRAY to hold results
STDMETHODIMP CAdmin::GetListEmployees(VARIANT * pvSARRAY)
{
	BSTR	sSQL;
	HRESULT	hResult;
	BOOL	bValid;
	VARIANT	v;
	DISPPARAMS dispparamsNoArgs;
	
	hResult = NOERROR;
	dispparamsNoArgs.rgvarg = NULL;				// Array of arguments.
	dispparamsNoArgs.rgdispidNamedArgs = NULL;	// Dispatch IDs of named arguments.
	dispparamsNoArgs.cArgs = 0;					// Number of arguments.
	dispparamsNoArgs.cNamedArgs = 0;

	::VariantInit(&v);
	if (m_iSrcType == icDoDataDirect)
	{
		v.vt = VT_BSTR | VT_ARRAY;
		
		if (m_bUseODBC)
			sSQL = CStrFnc::LoadResString(scLvw_Employee_SQL);
		else
			sSQL = CStrFnc::LoadResString(scLvw_Employee_MDB);
		
		hResult = ExecFillArray(sSQL, &v, false, &bValid);
		
		::SysFreeString(sSQL);
		
		if (hResult != NOERROR)
			return (hResult);

		if (!bValid)
			return (NOERROR);
	}
	else
	{
		if (m_pIRemoteBenefitsServer == NULL)
			return (E_FAIL);

		::VariantInit(&v);
		v.vt = VT_EMPTY;
		
		hResult = m_pIRemoteBenefitsServer->Invoke(
			RemoteServer_Dispatch_Ids[REMOTESERVER_METHOD_GETLISTEMPLOYEES],
			IID_NULL,
			LOCALE_USER_DEFAULT,
			DISPATCH_METHOD,
			&dispparamsNoArgs,
			&v,
			&m_RemoteServer_ExceptionInfo,
			NULL);
		
		if (hResult != NOERROR)
			return (hResult);
	}
	
	if (pvSARRAY->vt & VT_BYREF)
		hResult = ::SafeArrayCopy(V_ARRAY(&v), V_ARRAYREF(pvSARRAY));
	else
		hResult = ::SafeArrayCopy(V_ARRAY(&v), &V_ARRAY(pvSARRAY));
	pvSARRAY->vt = v.vt;

	return (NOERROR);
}

//GetNumEmplsPerBenefit
//Purpose	- Fills an IList object, according to predetermined rules,  with a particular recordset
//Inputs	- (ppcOutput)		- pointer to IList object to be filled
//			- (lCoverageID)		- data value used to construct SQL statement to extract data
//			- (lTaxID			- "	"	"	"	"	"	"	"	"	"	"	"	"	"	"	"
//			- (bGender			- "	"	"	"	"	"	"	"	"	"	"	"	"	"	"	"
//			- (pbSuccess)		- boolean output target
//Outputs	- (pbSuccess)		- boolean success value
STDMETHODIMP CAdmin::GetNumEmplsPerBenefit(IList ** ppcOutput, long lCoverageID, long lTaxID, VARIANT * bGender, BOOL * pbSuccess)
{
	ASSERT(ppcOutput != NULL);
	ASSERT(pbSuccess != NULL);
	ASSERT(bGender != NULL);
	
	_bstr_t sSQL;
	CStrFnc sTemp;
	VARIANT v;
	BOOL	bValid;
	HRESULT hResult;
	IList * plstItem;
	long	lMax;
	long	lMin;
	long	i;
	CStrFnc * bstrTemp;
	BSTR BenefitName;
	BSTR EmployeeCount;
	
	hResult = NOERROR;
	*pbSuccess = FALSE;
	
	if (*ppcOutput)
		try
		{
			(*ppcOutput)->Release();
		}
		catch (...)
		{
		}

	hResult = ::CoCreateInstance(CLSID_CList, NULL, THISCLSCTX, IID_IList, (void **)ppcOutput);
	
	sSQL = CStrFnc::LoadResString(scGNEPBSQL1);
	if (!IsOptionalMissing(bGender))
	{
		sTemp = (V_I2(bGender) || V_I4(bGender)) ? 
			CStrFnc::LoadResString(scGNEPBSQL2) :
		CStrFnc::LoadResString(scGNEPBSQL3);
		sSQL += sTemp.GetBstr();
	}
	
	if (lTaxID)
	{
		sSQL += CStrFnc::LoadResString(scGNEPBSQL4);
		sTemp.Format(TEXT("%ld"), lTaxID);
		sSQL += sTemp.GetBstr();
	}
	
	if (lCoverageID)
	{
		sSQL += CStrFnc::LoadResString(scGNEPBSQL5);
		sTemp.Format(TEXT("%ld"), lCoverageID);
		sSQL += sTemp.GetBstr();
	}
	
	sSQL += CStrFnc::LoadResString(scGNEPBSQL6);
	::VariantInit(&v);
	v.vt = VT_BSTR | VT_ARRAY;
	
	hResult = ExecFillArray(sSQL, &v, false, &bValid);
	
	::SysFreeString(sSQL);
	
	if (hResult != NOERROR)
		return (hResult);

	if (bValid)
	{
		BenefitName = CStrFnc::LoadResString(scBnfName);
		EmployeeCount = CStrFnc::LoadResString(scEmplCnt);
		CSafeArray SA(V_ARRAY(&v));
		lMin = SA.LBound(2);
		lMax = SA.UBound(2);
		for (i = lMin; i < lMax; i++)
		{
			hResult = ::CoCreateInstance(CLSID_CList, NULL, THISCLSCTX, IID_IList, (void **)&plstItem);
			
			::VariantInit(&v);
			v.vt = VT_BSTR;
			V_BSTR(&v) = SA.bstrGet((long)icLvwDesc, i);
			plstItem->Add(BenefitName,&v);
			
			::VariantInit(&v);
			v.vt = VT_I4;
			bstrTemp = new CStrFnc(SA.bstrGet(icLvwPKId, i));
			V_I4(&v) = bstrTemp->Val();
			delete bstrTemp;
			plstItem->Add(EmployeeCount, &v);
			
			
			sTemp = CList::GenListKey(i);
			::VariantInit(&v);
			v.vt = VT_DISPATCH;
			V_DISPATCH(&v) = plstItem;
			(*ppcOutput)->Add(sTemp.GetBstr(), &v);

			plstItem->Release();
		}
		*pbSuccess = TRUE;
	}
	return (NOERROR);
}

//GetNumEmplsPerBenefitType
//Purpose	- Fills an IList object, according to predetermined rules,  with a particular recordset
//Inputs	- (ppcOutput)		- pointer to IList object to be filled
//			- (lCoverageID)		- data value used to construct SQL statement to extract data
//			- (lTaxID			- "	"	"	"	"	"	"	"	"	"	"	"	"	"	"	"
//			- (bGender			- "	"	"	"	"	"	"	"	"	"	"	"	"	"	"	"
//			- (pbSuccess)		- boolean output target
//Outputs	- (pbSuccess)		- boolean success value
STDMETHODIMP CAdmin::GetNumEmplsPerBenefitType(IList * * ppcOutput, long lCoverageID, long lTaxID, VARIANT * bGender, BOOL * pbSuccess)
{
	ASSERT(pbSuccess != NULL);
	ASSERT(bGender != NULL);
	
	_bstr_t sSQL;
	CStrFnc * pTemp;
	CStrFnc sTemp;
	VARIANT v;
	BOOL	bSuccess;
	HRESULT hResult;
	long	lMin;
	long	lMax;
	long	i;
	long	n;
	_bstr_t BenefitName;
	_bstr_t EmployeeCount;
	BSTR	bstrTemp;
	IList * plstInner;

	
	hResult = NOERROR;
	lMin = 0;
	lMax = 0;
	i	 = 0;
	n	 = 0;
	*pbSuccess = FALSE;
	
	BenefitName		= CStrFnc::LoadResString(scBnfTypeName);
	EmployeeCount	= CStrFnc::LoadResString(scEmplCnt);
	sSQL			= CStrFnc::LoadResString(scGNEPBTSQL1);
	
	if (!IsOptionalMissing(bGender))
	{
		sTemp = (V_I2(bGender) || V_I4(bGender)) ?
			CStrFnc::LoadResString(scGNEPBSQL2) :
		CStrFnc::LoadResString(scGNEPBSQL3);
		sSQL += sTemp.GetBstr();
	}
	
	if (lTaxID)
	{
		sTemp = CStrFnc::LoadResString(scGNEPBSQL4);
		sSQL += sTemp.GetBstr();
		sTemp.Format(TEXT("%ld"), lTaxID);
		sSQL += sTemp.GetBstr();
	}
	
	if (lCoverageID)
	{
		sTemp = CStrFnc::LoadResString(scGNEPBSQL5);
		sSQL += sTemp.GetBstr();
		sTemp.Format(TEXT("%ld"), lCoverageID);
		sSQL += sTemp.GetBstr();
	}
	sTemp = CStrFnc::LoadResString(scGNEPBTSQL2);
	sSQL += sTemp.GetBstr();
	
	::VariantInit(&v);
	v.vt = VT_BSTR | VT_ARRAY;
	
	hResult = ExecFillArray(sSQL, &v, false, &bSuccess);
	
	::SysFreeString(sSQL);
	
	if (hResult != NOERROR)
		return (hResult);

	if (bSuccess)
	{
		if (*ppcOutput)
			try
			{
				(*ppcOutput)->Release();
			}
			catch (...)
			{
			}

		hResult = ::CoCreateInstance(CLSID_CList, NULL, THISCLSCTX, IID_IList, (void **)ppcOutput);
		
		CSafeArray SA(V_ARRAY(&v));
		lMin = SA.LBound(2);
		lMax = SA.UBound(2);
		for (i = lMin; i < lMax; i++)
		{
			hResult = ::CoCreateInstance(CLSID_CList, NULL, THISCLSCTX, IID_IList, (void **)&plstInner);
			
			::VariantInit(&v);
			v.vt = VT_BSTR;
			bstrTemp = ::SysAllocString(SA.bstrGet((long)icLvwDesc, i));
			V_BSTR(&v) = bstrTemp;
			plstInner->Add(BenefitName,&v);
			::SysFreeString(bstrTemp);
			
			bstrTemp = SA.bstrGet(icLvwPKId, i);
			pTemp = new CStrFnc(bstrTemp);
			n = pTemp->Val();
			delete pTemp;
			::VariantInit(&v);
			v.vt = VT_I4;
			V_I4(&v) = n;
			plstInner->Add(EmployeeCount,&v);
			
			sTemp = CList::GenListKey(i);
			::VariantInit(&v);
			v.vt = VT_DISPATCH;
			V_DISPATCH(&v) = plstInner;
			(*ppcOutput)->Add(sTemp.GetBstr(), &v);
			
			plstInner->Release();
		}
		*pbSuccess = TRUE;
	}
	return (NOERROR);
}

//Init
//Purpose		- Function used to initialize DB and call remote server if required
//Inputs		- (bUseODBC)	- boolean to indicate DB method
//				- {DataSource)	- location of data source
//				- {iSourceType)	- flag indicating how data is to be accessed
//				- {iRemoteSvr)	- "	"	"	"	"	"	"	"	"	"	"	"
//				- {iLinkSvr)	- "	"	"	"	"	"	"	"	"	"	"	"
//				- (pbSuccess)	- boolean output target
//Outputs		- (pbSuccess)	- boolean success value
STDMETHODIMP CAdmin::Init(BOOL bUseODBC, BSTR DataSource, int iSourceType, int iRemoteSvr, int iLinkSvr, BOOL * pbSuccess)
{
	HRESULT		hResult;
	_DBEngine *	pTempDBEngine;
	VARIANT		ReadOnly;
	VARIANT		Options;
	
	hResult = NOERROR;
	*pbSuccess = FALSE;

	Term();
	
	m_iSrcType = iSourceType;
	m_bUseODBC = bUseODBC;
	
	switch (iSourceType)
	{
	case icDoDataDirect:
		{
			// create COM link to DAO engine
			try
			{
				if (m_pClassFactory2)
				{
					m_pClassFactory2->Release();
					m_pClassFactory2 = NULL;
				}

				hResult = CoGetClassObject(__uuidof(DBEngine), CLSCTX_INPROC_SERVER, NULL, IID_IClassFactory2, (void**)&m_pClassFactory2);
				if (hResult != NOERROR)
					return (hResult);

				hResult = m_pClassFactory2->CreateInstanceLic(NULL, NULL, __uuidof(_DBEngine), sKey,(LPVOID *)&pTempDBEngine);
				if (hResult != NOERROR)
					return (hResult);
				    
				m_DbEngine = pTempDBEngine;

				pTempDBEngine->Release();
			}
			catch (...)
			{
				return (NOERROR);
			}
			
			try {
				if (bUseODBC)
				{
					Options.vt = VT_I4;
					V_I4(&Options) = FALSE;
					ReadOnly.vt = VT_I4;
					V_I4(&ReadOnly) = FALSE;
					m_Database = m_DbEngine->OpenDatabase(scBlank, Options, ReadOnly, DataSource);
				}
				else
				{
					m_Database = m_DbEngine->OpenDatabase(DataSource);
				}
			}
			catch (...)
			{
				return (NOERROR);
			}
			
			m_bstrDataSource = ::SysAllocString(DataSource);			
			break;
		}
		
	case icDoDataRemote:
		{
			switch (iRemoteSvr)
			{
			case icRemCPP:
				hResult = StartRemoteObject(TEXT("BenefitsCPPRemote.CAdmin"), bUseODBC, DataSource, icDoDataDirect, 0,0);
				break;
			case icRemVB:
				hResult = StartRemoteObject(TEXT("BenefitsVBRemote.CAdmin"), bUseODBC, DataSource, icDoDataDirect, 0,0);
				break;
			default:
				return (FlagException(IID_IAdmin, scInit, icErrSrcType));
			}
			break;
		}
		
	case icDoDataLinked:
		{
			switch (iRemoteSvr)
			{
			case icLnkCPP:
				hResult = StartRemoteObject(TEXT("BenefitsCPPRemote.CAdmin"), bUseODBC, DataSource, icDoDataRemote, 0,0);
				break;
			case icLnkVB:
				hResult = StartRemoteObject(TEXT("BenefitsVBRemote.CAdmin"), bUseODBC, DataSource, icDoDataRemote, 0,0);
				break;
			default:
				return (FlagException(IID_IAdmin, scInit, icErrSrcType));
			}
			break;
		}

	default:
		return (FlagException(IID_IAdmin, scInit, icErrSrcType));
	}
	
	if (hResult == NOERROR)
		*pbSuccess = TRUE;

	return (hResult);
}

//Insert
//Purpose		- Inserts specified object
//Assumptions	- Only one input object can be specified
//Inputs		- (pvEmployee)		- object with insert information (can be null)
//				- (pvBank)			- object with insert information (can be null)
//				- (pvDependant)		- object with insert information (can be null)
//				- (pvBenefit)		- object with insert information (can be null)
//				- (pvBenefitLink)	- object with insert information (can be null)
//				- (pbSuccess)		- boolean output target
//Outputs		- (pbSuccess)		- boolean success value
STDMETHODIMP CAdmin::Insert(VARIANT * pvEmployee, VARIANT * pvBank, VARIANT * pvDependant, VARIANT * pvBenefit, VARIANT * pvBenefitLink, BOOL * pbSuccess)
{
	HRESULT			hResult;
	int				iParameterCount;
	IEmployee *		poEmployee;
	IBank *			poBank;
	IDependant *	poDependant;
	IBenefit *		poBenefit;
	IBenefitLink *	poBenefitLink;
	IDispatch *		IDisp;
	
	hResult		= NOERROR;
	poEmployee	= NULL;
	poBank		= NULL;
	poDependant = NULL;
	poBenefit	= NULL;
	poBenefitLink	= NULL;
	iParameterCount = 0;

	if (!IsOptionalMissing(pvEmployee) && (pvEmployee->vt & VT_DISPATCH == VT_DISPATCH))
	{
		IDisp = *(V_DISPATCHREF(pvEmployee));
		hResult = IDisp->QueryInterface(IID_IEmployee, (void **)&poEmployee);
		if (hResult == NOERROR)
			iParameterCount++;
		else
			return (FlagException(IID_IAdmin, scIns, icErrObjectOp));
	}
	if (!IsOptionalMissing(pvBank) && (pvBank->vt & VT_DISPATCH == VT_DISPATCH))
	{
		IDisp = *(V_DISPATCHREF(pvBank));
		hResult = IDisp->QueryInterface(IID_IBank, (void **)&poBank);
		if (hResult == NOERROR)
			iParameterCount++;
		else
			return (FlagException(IID_IAdmin, scIns, icErrObjectOp));
	}
	if (!IsOptionalMissing(pvDependant) && (pvDependant->vt & VT_DISPATCH == VT_DISPATCH))
	{
		IDisp = *(V_DISPATCHREF(pvDependant));
		hResult = IDisp->QueryInterface(IID_IDependant, (void **)&poDependant);
		if (hResult == NOERROR)
			iParameterCount++;
		else
			return (FlagException(IID_IAdmin, scIns, icErrObjectOp));
	}
	if (!IsOptionalMissing(pvBenefit) && (pvBenefit->vt & VT_DISPATCH == VT_DISPATCH))
	{
		IDisp = *(V_DISPATCHREF(pvBenefit));
		hResult = IDisp->QueryInterface(IID_IBenefit, (void **)&poBenefit);
		if (hResult == NOERROR)
			iParameterCount++;
		else
			return (FlagException(IID_IAdmin, scIns, icErrObjectOp));
	}
	if (!IsOptionalMissing(pvBenefitLink) && (pvBenefitLink->vt & VT_DISPATCH == VT_DISPATCH))
	{
		IDisp = *(V_DISPATCHREF(pvBenefitLink));
		hResult = IDisp->QueryInterface(IID_IBenefitLink, (void **)&poBenefitLink);
		if (hResult == NOERROR)
			iParameterCount++;
		else
			return (FlagException(IID_IAdmin, scIns, icErrObjectOp));
	}
	if (iParameterCount != 1)
		return (FlagException(IID_IAdmin, scIns, icErrObjectOp));
	
	
	if (poEmployee)
		hResult = Insert_Employee(poEmployee, pbSuccess);
	else
	{
		if (poBank)
			hResult = Insert_Bank(poBank, pbSuccess);
		else
		{
			if (poDependant)
				hResult = Insert_Dependant(poDependant, pbSuccess);
			else
			{
				if (poBenefit)
					hResult = Insert_Benefit(poBenefit, pbSuccess);
				else
				{
					if (poBenefitLink)
						hResult = Insert_BenefitLink(poBenefitLink, pbSuccess);
				}
			}
		}
	}
					
	return (hResult);
}

//IsUserValid
//Purpose		- Validates a specified userid
//Inputs		- (bstrUserId)	- User Id to be validated
//				- (pbSuccess)	- boolean output target
//Outputs		- (pbSuccess)	- boolean success value
STDMETHODIMP CAdmin::IsUserValid(BSTR bstrUserId, int * piSuccess)
{
	_bstr_t sSQL;
	CStrFnc sUserName(bstrUserId);
	HRESULT hResult;
	BOOL	bValid;
	VARIANT v;
	BSTR	bstrTemp;
	int		n;
	VARIANTARG vArgs[1];
	DISPPARAMS pDispParams;
	
	hResult = NOERROR;
	*piSuccess = icInvalidUser;
	
	if (m_iSrcType == icDoDataDirect)
	{
		
		sUserName.Trim();
		
		if (_wcsicmp(sUserName.GetBstr(), TEXT("admin")) == 0)
		{
			*piSuccess = icValidAdmin;
			return (NOERROR);
		}
		
		sUserName.Left(16);
		sSQL = CStrFnc::LoadResString(scCHKUSER);
		sSQL += scSQ;
		sSQL += sUserName.GetBstr();
		sSQL += scSQ;
		
		::VariantInit(&v);
		v.vt = VT_BSTR | VT_ARRAY;
		
		hResult = ExecFillArray(sSQL, &v, false, &bValid);
		if (hResult != NOERROR)
			return (hResult);
		
		if (!bValid)
			return (NOERROR);
		
		CSafeArray psa(V_ARRAY(&v));
		bstrTemp = psa.bstrGet((long)0, (long)0);
		if (CStrFnc::Val(bstrTemp))
			*piSuccess = icValidUser;
	}
	else
	{
		if (m_pIRemoteBenefitsServer == NULL)
			return (E_FAIL);

		::VariantInit(&v);
		v.vt = VT_EMPTY;
		
		for (n = 0; n < 1; n++)
			::VariantInit(&vArgs[n]);
		
		vArgs[0].vt = VT_BSTR;
		vArgs[0].bstrVal = bstrUserId;
		pDispParams.cArgs = 1;
		pDispParams.cNamedArgs = 0;
		pDispParams.rgvarg = vArgs;
		
		hResult = m_pIRemoteBenefitsServer->Invoke(
			RemoteServer_Dispatch_Ids[REMOTESERVER_METHOD_ISUSERVALID],
			IID_NULL,
			LOCALE_USER_DEFAULT,
			DISPATCH_METHOD,
			&pDispParams,
			&v,
			&m_RemoteServer_ExceptionInfo,
			NULL);
		
		if (hResult != NOERROR)
			return (hResult);

		*piSuccess = v.iVal;
	}
	
	return (NOERROR);
}

//Term
//Purpose		- Used to clean up object and release resources
STDMETHODIMP CAdmin::Term()
{
	HRESULT		hResult;
	DISPPARAMS	dispparamsNoArgs;
	CIEnumVariant * pForEach;
	IEmployee * pEmp;
	IBank *		pBnk;
	long		lCount;
	ULONG		uCount;
	VARIANT		varTemp;
	
	hResult = NOERROR;
	lCount	= 0;
	dispparamsNoArgs.rgvarg = NULL;				// Array of arguments.
	dispparamsNoArgs.rgdispidNamedArgs = NULL;	// Dispatch IDs of named arguments.
	dispparamsNoArgs.cArgs = 0;					// Number of arguments.
	dispparamsNoArgs.cNamedArgs = 0;

	
	::VariantInit(&varTemp);

	if (m_Benefits)
	{
		m_Benefits->Release();
		m_Benefits = NULL;
	}

	::VariantInit(&varTemp);

	if (m_Banks)
	{
		hResult = m_Banks->get__NewEnum((IUnknown **)&pForEach);
		::VariantInit(&varTemp);
		hResult = pForEach->Next(1, &varTemp, &uCount);
		while (hResult == NOERROR)
		{
			pBnk = (IBank *)V_DISPATCH(&varTemp);
			pBnk->Term();
			pBnk->Release();
			::VariantClear(&varTemp);
			hResult = pForEach->Next(1, &varTemp, &uCount);
		}
    	m_Banks->Release();
		m_Banks = NULL;
	}
	
	::VariantInit(&varTemp);

	if (m_Employees)
	{
		hResult = m_Employees->get__NewEnum((IUnknown **)&pForEach);
		::VariantInit(&varTemp);
		hResult = pForEach->Next(1, &varTemp, &uCount);
		while (hResult == NOERROR)
		{
			pEmp = (IEmployee *)V_DISPATCH(&varTemp);
			pEmp->Term();
			pEmp->Release();
			::VariantClear(&varTemp);
			hResult = pForEach->Next(1, &varTemp, &uCount);
		}
		m_Employees->Release();
		m_Employees = NULL;
	}
	
	if (m_Database)
	{
		hResult = m_Database->Close();
		m_Database = NULL;
	}
	
	if (m_DbEngine)
		m_DbEngine = NULL;
	
	if (m_iSrcType != icDoDataDirect)
	{
		if (m_pIRemoteBenefitsServer)
		{
			hResult = m_pIRemoteBenefitsServer->Invoke(
				RemoteServer_Dispatch_Ids[REMOTESERVER_METHOD_TERM],
				IID_NULL,
				LOCALE_USER_DEFAULT,
				DISPATCH_METHOD,
				&dispparamsNoArgs,
				NULL,
				&m_RemoteServer_ExceptionInfo,
				NULL);

			m_pIRemoteBenefitsServer->Release();
			m_pIRemoteBenefitsServer = NULL;
		}
	}
	
	return (hResult);
}

//Update
//Purpose		- Update specified object
//Assumptions	- Only one input object can be specified
//Inputs		- (pvEmployee)		- object with update information (can be null)
//				- (pvBank)			- object with update information (can be null)
//				- (pvDependant)		- object with update information (can be null)
//				- (pvBenefit)		- object with update information (can be null)
//				- (pvBenefitLink)	- object with update information (can be null)
//				- (pbSuccess)		- boolean output target
//Outputs		- (pbSuccess)		- boolean success value
STDMETHODIMP CAdmin::Update(VARIANT * pvEmployee, VARIANT * pvBank, VARIANT * pvDependant, VARIANT * pvBenefit, VARIANT * pvBenefitLink, BOOL * pbSuccess)
{	
	HRESULT hResult;
	_bstr_t	sSQL;
	_bstr_t sErr;
	CStrFnc sTemp;
	int		iParameterCount;
	IEmployee *		poEmployee;
	IBank *			poBank;
	IDependant *	poDependant;
	IBenefit *		poBenefit;
	IBenefitLink *	poBenefitLink;
	IDispatch *		IDisp;
	
	hResult			= NOERROR;
	poEmployee		= NULL;
	poBank			= NULL;
	poDependant		= NULL;
	poBenefit		= NULL;
	poBenefitLink	= NULL;
	iParameterCount = 0;

	if (!IsOptionalMissing(pvEmployee) && (pvEmployee->vt & VT_DISPATCH == VT_DISPATCH))
	{
		IDisp = *(V_DISPATCHREF(pvEmployee));
		hResult = IDisp->QueryInterface(IID_IEmployee, (void **)&poEmployee);
		if (hResult == NOERROR)
			iParameterCount++;
		else
			return (FlagException(IID_IAdmin, scUpd, icErrObjectOp));
	}
	if (!IsOptionalMissing(pvBank) && (pvBank->vt & VT_DISPATCH == VT_DISPATCH))
	{
		IDisp = *(V_DISPATCHREF(pvBank));
		hResult = IDisp->QueryInterface(IID_IBank, (void **)&poBank);
		if (hResult == NOERROR)
			iParameterCount++;
		else
			return (FlagException(IID_IAdmin, scUpd, icErrObjectOp));
	}
	if (!IsOptionalMissing(pvDependant) && (pvDependant->vt & VT_DISPATCH == VT_DISPATCH))
	{
		IDisp = *(V_DISPATCHREF(pvDependant));
		hResult = IDisp->QueryInterface(IID_IDependant, (void **)&poDependant);
		if (hResult == NOERROR)
			iParameterCount++;
		else
			return (FlagException(IID_IAdmin, scUpd, icErrObjectOp));
	}
	if (!IsOptionalMissing(pvBenefit) && (pvBenefit->vt & VT_DISPATCH == VT_DISPATCH))
	{
		IDisp = *(V_DISPATCHREF(pvBenefit));
		hResult = IDisp->QueryInterface(IID_IBenefit, (void **)&poBenefit);
		if (hResult == NOERROR)
			iParameterCount++;
		else
			return (FlagException(IID_IAdmin, scUpd, icErrObjectOp));
	}
	if (!IsOptionalMissing(pvBenefitLink) && (pvBenefitLink->vt & VT_DISPATCH == VT_DISPATCH))
	{
		IDisp = *(V_DISPATCHREF(pvBenefitLink));
		hResult = IDisp->QueryInterface(IID_IBenefitLink, (void **)&poBenefitLink);
		if (hResult == NOERROR)
			iParameterCount++;
		else
			return (FlagException(IID_IAdmin, scUpd, icErrObjectOp));
	}
	if (iParameterCount != 1)
		return (FlagException(IID_IAdmin, scUpd, icErrObjectOp));
	
	
	if (poEmployee)
		hResult = Update_Employee(poEmployee, pbSuccess);
	else
	{
		if (poBank)
			hResult = Update_Bank(poBank, pbSuccess);
		else
		{
			if (poDependant)
				hResult = Update_Dependant(poDependant, pbSuccess);
			else
			{
				if (poBenefit)
					hResult = Update_Benefit(poBenefit, pbSuccess);
				else
				{
					if (poBenefitLink)
						hResult = Update_BenefitLink(poBenefitLink, pbSuccess);
				}
			}
		}
	}
					
	return (hResult);
}

STDMETHODIMP CAdmin::get_Banks(IListBank ** ppOut)
{
	HRESULT hResult;

	hResult = NOERROR;
	
	if (m_Banks == NULL)
	{
		hResult = ::CoCreateInstance(CLSID_CListBank, NULL, THISCLSCTX, IID_IListBank, (void **)&m_Banks);
		if (hResult != NOERROR)
			return (hResult);
	}
	*ppOut = m_Banks;
	
	if (m_Banks)
		m_Banks->AddRef();

	return (hResult);
}

STDMETHODIMP CAdmin::get_Benefits(IListBenefit ** ppOut)
{
	HRESULT hResult;

	hResult = NOERROR;

	if (m_Benefits == NULL)
	{
		hResult = ::CoCreateInstance(CLSID_CListBenefit, NULL, THISCLSCTX, IID_IListBenefit, (void **)&m_Benefits);
		if (hResult != NOERROR)
			return (hResult);
	}
	*ppOut = m_Benefits;
	
	if (m_Benefits)
		m_Benefits->AddRef();

	return (hResult);
}

STDMETHODIMP CAdmin::get_Employees(IListEmployee ** ppOut)
{	
	HRESULT hResult;

	hResult = NOERROR;

	if (m_Employees == NULL)
	{
		hResult = ::CoCreateInstance(CLSID_CListEmployee, NULL, THISCLSCTX, IID_IListEmployee, (void **)&m_Employees);
		if (hResult != NOERROR)
			return (hResult);
	}
	*ppOut = m_Employees;
	
	if (m_Employees)
		m_Employees->AddRef();

	return (hResult);
}

STDMETHODIMP CAdmin::get_UseODBC(BOOL * pbUseODBC)
{
	*pbUseODBC = m_bUseODBC;
	return (NOERROR);
}

//Delete_Bank
//Purpose		- Deletes specified bank object
//Inputs		- (poBank)			- bank object with delete information
//				- (pbSuccess)		- boolean output target
//Outputs		- (pbSuccess)		- boolean success value
HRESULT CAdmin::Delete_Bank(IBank * poBank, BOOL * pbSuccess)
{
	HRESULT hResult;
	_bstr_t sSQL;
	_bstr_t sErr;
	VARIANT vResult;
	long lPKId;
	CStrFnc sTemp;
	IList * ParameterList;
	VARIANT v;
	IDispatch* pIDispatch;
	
	hResult = NOERROR;
    *pbSuccess = FALSE;
	
	hResult = poBank->get_PKId(&lPKId);
	if (lPKId == 0)
	{
		poBank->Term();
		poBank->Release();
		sErr = CStrFnc::LoadResString(scDelBank);
		sErr += TEXT("\n\r");
		sErr += CStrFnc::LoadResString(icErrReqFld);
		sErr += TEXT("PKId");
		return (FlagExceptionMsg(IID_IAdmin, scDelBank, sErr));
	}
	
	if (m_bUseODBC)
	{
		sSQL = CStrFnc::LoadResString(scExecute);
		sSQL += CStrFnc::LoadResString(scDel_Bank);
		sSQL += scSpace;
		sTemp.Format(TEXT("%ld"), lPKId);
		sSQL += sTemp.GetBstr();
		if (ExecBooleanSQL(sSQL, pbSuccess) != NOERROR)
		{
			poBank->Term();
			poBank->Release();
			return (NOERROR);
		}
	}
	else
	{
		hResult = ::CoCreateInstance(CLSID_CList, NULL, THISCLSCTX, IID_IList, (void **)&ParameterList);
		::VariantInit(&v);
		vResult.vt = VT_I4;
		V_I4(&vResult) = lPKId;
		hResult = ParameterList->Add(TEXT("@PKId"),&vResult);
		sSQL = CStrFnc::LoadResString(scDel_Bank);
		hResult = ParameterList->QueryInterface(IID_IDispatch, (void **)&pIDispatch);
		if (ExecBooleanMDB(sSQL, pIDispatch, pbSuccess) != NOERROR)
		{
			poBank->Term();
			poBank->Release();
			pIDispatch->Release();
			ParameterList->Release();
			sSQL = CStrFnc::LoadResString(scDelBank);
			FlagExceptionMsg(IID_IAdmin, scDelBank, sSQL);
			return (NOERROR);
		}
		pIDispatch->Release();
		ParameterList->Release();
	}
	
	poBank->Term();
	poBank->Release();
	FillListBanks();
	
	return (NOERROR);
}

//Delete_Benefit
//Purpose		- Deletes specified benefit object
//Inputs		- (poBenefit)			- benefit object with delete information
//				- (pbSuccess)		- boolean output target
//Outputs		- (pbSuccess)		- boolean success value
HRESULT CAdmin::Delete_Benefit(IBenefit * poBenefit, BOOL * pbSuccess)
{
	HRESULT hResult;
	_bstr_t sSQL;
	_bstr_t sErr;
	VARIANT vResult;
	long	lPKId;
	CStrFnc sTemp;
	IList * ParameterList;
	IDispatch * pIDispatch;
	VARIANT v;
	
	hResult = NOERROR;
    *pbSuccess = FALSE;
	
	hResult = poBenefit->get_PKId(&lPKId);
	if (lPKId == 0)
	{
		poBenefit ->Term();
		poBenefit ->Release();
		sErr = CStrFnc::LoadResString(scDelBnf);
		sErr += TEXT("\n\r");
		sErr += CStrFnc::LoadResString(icErrReqFld);
		sErr += TEXT("PKId");
		return (FlagExceptionMsg(IID_IAdmin, scDelBnf, sErr));
	}
	
	if (m_bUseODBC)
	{
		sSQL = CStrFnc::LoadResString(scExecute);
		sSQL += CStrFnc::LoadResString(scDel_Benefit);
		sSQL += scSpace;
		sTemp.Format(TEXT("%ld"), lPKId);
		sSQL += sTemp.GetBstr();
		if (ExecBooleanSQL(sSQL, pbSuccess) != NOERROR)
		{
			poBenefit ->Term();
			poBenefit ->Release();
			return (NOERROR);
		}
	}
	else
	{
		hResult = ::CoCreateInstance(CLSID_CList, NULL, THISCLSCTX, IID_IList, (void **)&ParameterList);
		::VariantInit(&v);
		vResult.vt = VT_I4;
		V_I4(&vResult) = lPKId;
		hResult = ParameterList->Add(TEXT("@PKId"),&vResult);
		sSQL = CStrFnc::LoadResString(scDel_Benefit);
		hResult = ParameterList->QueryInterface(IID_IDispatch, (void **)&pIDispatch);
		if (ExecBooleanMDB(sSQL, pIDispatch, pbSuccess) != NOERROR)
		{
			poBenefit ->Term();
			poBenefit ->Release();
			pIDispatch->Release();
			ParameterList->Release();
			sSQL = CStrFnc::LoadResString(scDelBnf);
			FlagExceptionMsg(IID_IAdmin, scDelBnf, sSQL);
			return (NOERROR);
		}
		pIDispatch->Release();
		ParameterList->Release();
	}
	
	poBenefit ->Term();
	poBenefit ->Release();
	FillListBenefits();
	
	return (NOERROR);
}

//Delete_BenefitLink
//Purpose		- Deletes specified BenefitLink object
//Inputs		- (poBenefitLink)	- BenefitLink object with delete information
//				- (pbSuccess)		- boolean output target
//Outputs		- (pbSuccess)		- boolean success value
HRESULT CAdmin::Delete_BenefitLink(IBenefitLink * poBenefitLink, BOOL * pbSuccess)
{
	HRESULT hResult;
	_bstr_t sSQL;
	_bstr_t sErr;
	VARIANT vResult;
	long lPKId;
	CStrFnc sTemp;
	IDispatch* pIDispatch;
	IList * ParameterList;
	VARIANT v;
	
	hResult = NOERROR;
    *pbSuccess = FALSE;
	
	hResult = poBenefitLink->get_PKId(&lPKId);
	if (lPKId == 0)
	{
		poBenefitLink->Term();
		poBenefitLink->Release();
		sErr = CStrFnc::LoadResString(scDelBnfLnk);
		sErr += TEXT("\n\r");
		sErr += CStrFnc::LoadResString(icErrReqFld);
		sErr += TEXT("PKId");
		return (FlagExceptionMsg(IID_IAdmin, scDelBnfLnk, sErr));
	}
	
	if (m_bUseODBC)
	{
		sSQL = CStrFnc::LoadResString(scExecute);
		sSQL += CStrFnc::LoadResString(scDel_EmplBnf);
		sSQL += scSpace;
		sTemp.Format(TEXT("%ld"), lPKId);
		sSQL += sTemp.GetBstr();
		if (ExecBooleanSQL(sSQL, pbSuccess) != NOERROR)
		{
			poBenefitLink->Term();
			poBenefitLink->Release();
			return (NOERROR);
		}
	}
	else
	{
		hResult = ::CoCreateInstance(CLSID_CList, NULL, THISCLSCTX, IID_IList, (void **)&ParameterList);
		if (hResult == NOERROR)
		{		
			::VariantInit(&v);
			vResult.vt = VT_I4;
			V_I4(&vResult) = lPKId;
			hResult = ParameterList->Add(TEXT("@PKId"),&vResult);
			if (hResult == NOERROR)
			{		
				sSQL = CStrFnc::LoadResString(scDel_EmplBnf);
				hResult = ParameterList->QueryInterface(IID_IDispatch, (void **)&pIDispatch);
				if (ExecBooleanMDB(sSQL, pIDispatch, pbSuccess) != NOERROR)
				{
					poBenefitLink->Term();
					poBenefitLink->Release();
					pIDispatch->Release();
					ParameterList->Release();
					sSQL = CStrFnc::LoadResString(scDelBnfLnk);
					FlagExceptionMsg(IID_IAdmin, scDelBnfLnk, sSQL);
					return (NOERROR);
				}
			}
			pIDispatch->Release();
			ParameterList->Release();
		}
	}
	
	poBenefitLink->Term();
	poBenefitLink->Release();	
	return (NOERROR);
}

//Delete_Dependant
//Purpose		- Deletes specified dependant object
//Inputs		- (poDependant)		- dependant object with delete information
//				- (pbSuccess)		- boolean output target
//Outputs		- (pbSuccess)		- boolean success value
HRESULT CAdmin::Delete_Dependant(IDependant * poDependant, BOOL * pbSuccess)
{
	HRESULT hResult;
	_bstr_t sSQL;
	_bstr_t sErr;
	VARIANT vResult;
	long	lPKId;
	CStrFnc sTemp;
	VARIANT v;
	IList *		ParameterList;
	IDispatch*	pIDispatch;
	
	hResult = NOERROR;
    *pbSuccess = FALSE;
	
	hResult = poDependant->get_PKId(&lPKId);
	if (lPKId == 0)
	{
		poDependant->Term();
		poDependant->Release();	
		sErr = CStrFnc::LoadResString(scDelDep);
		sErr += TEXT("\n\r");
		sErr += CStrFnc::LoadResString(icErrReqFld);
		sErr += TEXT("PKId");
		return (FlagExceptionMsg(IID_IAdmin, scDelDep, sErr));
	}
	
	if (m_bUseODBC)
	{
		sSQL = CStrFnc::LoadResString(scExecute);
		sSQL += CStrFnc::LoadResString(scDel_Dependant);
		sSQL += scSpace;
		sTemp.Format(TEXT("%ld"), lPKId);
		sSQL += sTemp.GetBstr();
		if (ExecBooleanSQL(sSQL, pbSuccess) != NOERROR)
		{
			poDependant->Term();
			poDependant->Release();	
			return (NOERROR);
		}
	}
	else
	{
		hResult = ::CoCreateInstance(CLSID_CList, NULL, THISCLSCTX, IID_IList, (void **)&ParameterList);
		::VariantInit(&v);
		vResult.vt = VT_I4;
		V_I4(&vResult) = lPKId;
		hResult = ParameterList->Add(TEXT("@PKId"),&vResult);
		sSQL = CStrFnc::LoadResString(scDel_Dependant);
		hResult = ParameterList->QueryInterface(IID_IDispatch, (void **)&pIDispatch);
		if (ExecBooleanMDB(sSQL, pIDispatch, pbSuccess) != NOERROR)
		{
			poDependant->Term();
			poDependant->Release();	
			pIDispatch->Release();
			ParameterList->Release();
			sSQL = CStrFnc::LoadResString(scDelDep);
			FlagExceptionMsg(IID_IAdmin, scDelDep, sSQL);
			return (NOERROR);
		}
		pIDispatch->Release();
		ParameterList->Release();
	}
	
	poDependant->Term();
	poDependant->Release();	
	return (NOERROR);
}

//Delete_Employee
//Purpose		- Deletes specified employee object
//Inputs		- (poEmployee)		- employee object with delete information
//				- (pbSuccess)		- boolean output target
//Outputs		- (pbSuccess)		- boolean success value
HRESULT CAdmin::Delete_Employee(IEmployee * poEmployee, BOOL * pbSuccess)
{
	HRESULT hResult;
	_bstr_t sSQL;
	_bstr_t sErr;
	VARIANT vResult;
	long	lPKId;
	CStrFnc sTemp;
	VARIANT v;
	IList * ParameterList;
	IDispatch* pIDispatch;
	
	hResult = NOERROR;
    *pbSuccess = FALSE;
	
	hResult = poEmployee->get_PKId(&lPKId);
	if (lPKId == 0)
	{
		poEmployee->Term();
		poEmployee->Release();
		sErr = CStrFnc::LoadResString(scDelEmpl);
		sErr += TEXT("\n\r");
		sErr += CStrFnc::LoadResString(icErrReqFld);
		sErr += TEXT("PKId");
		return (FlagExceptionMsg(IID_IAdmin, scDelEmpl, sErr));
	}
	
	if (m_bUseODBC)
	{
		sSQL = CStrFnc::LoadResString(scExecute);
		sSQL += CStrFnc::LoadResString(scDel_Employee);
		sSQL += scSpace;
		sTemp.Format(TEXT("%ld"), lPKId);
		sSQL += sTemp.GetBstr();
		if (ExecBooleanSQL(sSQL, pbSuccess) != NOERROR)
		{
			poEmployee->Term();
			poEmployee->Release();
			return (NOERROR);
		}
	}
	else
	{
		hResult = ::CoCreateInstance(CLSID_CList, NULL, THISCLSCTX, IID_IList, (void **)&ParameterList);
		::VariantInit(&v);
		vResult.vt = VT_I4;
		V_I4(&vResult) = lPKId;
		hResult = ParameterList->Add(TEXT("@PKId"),&vResult);
		sSQL = CStrFnc::LoadResString(scDel_Employee);
		hResult = ParameterList->QueryInterface(IID_IDispatch, (void **)&pIDispatch);
		if (ExecBooleanMDB(sSQL, pIDispatch, pbSuccess) != NOERROR)
		{
			poEmployee->Term();
			poEmployee->Release();
			pIDispatch->Release();
			ParameterList->Release();
			sSQL = CStrFnc::LoadResString(scDelEmpl);
			FlagExceptionMsg(IID_IAdmin, scDelEmpl, sSQL);
			return (NOERROR);
		}
		pIDispatch->Release();
		ParameterList->Release();
	}
	
	poEmployee->Term();
	poEmployee->Release();

	FillListEmployees();
	
	return (NOERROR);
}

//Insert_Bank
//Purpose		- Inserts specified bank object
//Inputs		- (poBank)		- bank object with insert information
//				- (pbSuccess)		- boolean output target
//Outputs		- (pbSuccess)		- boolean success value
HRESULT CAdmin::Insert_Bank(IBank * poBank, BOOL * pbSuccess)
{
	HRESULT hResult;
	_bstr_t sSQL;
	_bstr_t sErr;
	CStrFnc sTemp;
	VARIANT vResult;
	BSTR	bstrGet;
	IList * ParameterList;
	IDispatch* pIDispatch;
	
	hResult		= NOERROR;
	bstrGet		= NULL;
    *pbSuccess	= FALSE;
	
	_bstr_t bstrTemp;
	
	NULLBSTR(bstrGet);
	hResult = poBank->get_Code(&bstrGet);
	if (CStrFnc::IsEmpty(bstrGet))
	{
		poBank->Term();
		poBank->Release();
		sErr = CStrFnc::LoadResString(scInsBank);
		sErr += TEXT("\n\r");
		sErr += CStrFnc::LoadResString(icErrReqFld);
		sErr += TEXT("Code");
		return (FlagExceptionMsg(IID_IAdmin, scInsBank, sErr));
	}
	
	NULLBSTR(bstrGet);
	hResult = poBank->get_Name(&bstrGet);
	if (CStrFnc::IsEmpty(bstrGet))
	{
		poBank->Term();
		poBank->Release();
		sErr = CStrFnc::LoadResString(scInsBank);
		sErr += TEXT("\n\r");
		sErr += CStrFnc::LoadResString(icErrReqFld);
		sErr += TEXT("Name");
		return (FlagExceptionMsg(IID_IAdmin, scInsBank, sErr));
	}
	
	if (m_bUseODBC)
	{
		sSQL = CStrFnc::LoadResString(scExecute);
		sSQL += CStrFnc::LoadResString(scIns_Bank);
		sSQL += scSpace;
		
		NULLBSTR(bstrGet);
		hResult = poBank->get_Address1(&bstrGet);
		sTemp = bstrGet;
		sTemp.Trim(sTemp.Left(60));
		sSQL += scSQ;
		sSQL += sTemp.GetBstr();
		sSQL += scSQ;
		sSQL += scSpaceCommaSpace;
		
		NULLBSTR(bstrGet);
		hResult = poBank->get_Address2(&bstrGet);
		sTemp = bstrGet;
		sTemp.Trim(sTemp.Left(60));
		sSQL += scSQ;
		sSQL += sTemp.GetBstr();
		sSQL += scSQ;
		sSQL += scSpaceCommaSpace;
		
		NULLBSTR(bstrGet);
		hResult = poBank->get_City(&bstrGet);
		sTemp = bstrGet;
		sTemp.Trim(sTemp.Left(50));
		sSQL += scSQ;
		sSQL += sTemp.GetBstr();
		sSQL += scSQ;
		sSQL += scSpaceCommaSpace;
		
		NULLBSTR(bstrGet);
		hResult = poBank->get_Code(&bstrGet);
		sTemp = bstrGet;
		sTemp.Trim(sTemp.Left(16));
		sSQL += scSQ;
		sSQL += sTemp.GetBstr();
		sSQL += scSQ;
		sSQL += scSpaceCommaSpace;
		
		NULLBSTR(bstrGet);
		hResult = poBank->get_Name(&bstrGet);
		sTemp = bstrGet;
		sTemp.Trim(sTemp.Left(80));
		sSQL += scSQ;
		sSQL += sTemp.GetBstr();
		sSQL += scSQ;
		sSQL += scSpaceCommaSpace;
		
		if (ExecBooleanSQL(sSQL, pbSuccess) != NOERROR)
		{
			poBank->Term();
			poBank->Release();
			sErr = CStrFnc::LoadResString(scInsBank);
			FlagExceptionMsg(IID_IAdmin, scInsBank, sErr);
			return (NOERROR);
		}
	}
	else
	{
		hResult = ::CoCreateInstance(CLSID_CList, NULL, THISCLSCTX, IID_IList, (void **)&ParameterList);
		
		NULLBSTR(bstrGet);
		hResult = poBank->get_Address1(&bstrGet);
		::VariantInit(&vResult);
		vResult.vt = VT_BSTR;
		V_BSTR(&vResult) = bstrGet;
		hResult = ParameterList->Add(TEXT("@Address1"),&vResult);
		
		NULLBSTR(bstrGet);
		hResult = poBank->get_Address2(&bstrGet);
		::VariantInit(&vResult);
		vResult.vt = VT_BSTR;
		V_BSTR(&vResult) = bstrGet;
		hResult = ParameterList->Add(TEXT("@Address2"),&vResult);
		
		NULLBSTR(bstrGet);
		hResult = poBank->get_City(&bstrGet);
		::VariantInit(&vResult);
		vResult.vt = VT_BSTR;
		V_BSTR(&vResult) = bstrGet;
		hResult = ParameterList->Add(TEXT("@City"),&vResult);
		
		NULLBSTR(bstrGet);
		hResult = poBank->get_Code(&bstrGet);
		::VariantInit(&vResult);
		vResult.vt = VT_BSTR;
		V_BSTR(&vResult) = bstrGet;
		hResult = ParameterList->Add(TEXT("@Code"),&vResult);
		
		NULLBSTR(bstrGet);
		hResult = poBank->get_Name(&bstrGet);
		::VariantInit(&vResult);
		vResult.vt = VT_BSTR;
		V_BSTR(&vResult) = bstrGet;
		hResult = ParameterList->Add(TEXT("@Name"),&vResult);
		
		sSQL = CStrFnc::LoadResString(scIns_Bank);
		hResult = ParameterList->QueryInterface(IID_IDispatch, (void **)&pIDispatch);
		if (ExecBooleanMDB(sSQL, pIDispatch, pbSuccess) != NOERROR)
		{
			poBank->Term();
			poBank->Release();
			pIDispatch->Release();
			ParameterList->Release();
			sErr = CStrFnc::LoadResString(scInsBank);
			FlagExceptionMsg(IID_IAdmin, scInsBank, sErr);
			return (NOERROR);
		}
		pIDispatch->Release();
		ParameterList->Release();
	}
	
	poBank->Term();
	poBank->Release();
	FillListBanks();
    *pbSuccess = TRUE;
	
	return (NOERROR);
}
	
//Insert_Benefit
//Purpose		- Inserts specified benefit object
//Inputs		- (poBenefit)		- benefit object with insert information
//				- (pbSuccess)		- boolean output target
//Outputs		- (pbSuccess)		- boolean success value
HRESULT CAdmin::Insert_Benefit(IBenefit * poBenefit, BOOL * pbSuccess)
{
	HRESULT hResult;
	_bstr_t sSQL;
	_bstr_t sErr;
	CStrFnc sTemp;
	VARIANT vResult;
	BSTR	bstrGet;
	long	lBnfTaxId;
	long	lBnfTypeId;
	short	iIncrement;
	long	lMaximum;
	long	lMinimum;
	short	iMultiples;
	long	lCost;
	_bstr_t bstrTemp;
	IList * ParameterList;
	IDispatch * pIDispatch;
	
	hResult		= NOERROR;
	bstrGet		= NULL;
	*pbSuccess	= FALSE;
	
	
	hResult = poBenefit->get_BnfTaxId(&lBnfTaxId);
	if (lBnfTaxId == 0)
	{
		poBenefit->Term();
		poBenefit->Release();
		sErr = CStrFnc::LoadResString(scInsBnf);
		sErr += TEXT("\n\r");
		sErr += CStrFnc::LoadResString(icErrReqFld);
		sErr += TEXT("BnfTaxId");
		return (FlagExceptionMsg(IID_IAdmin, scInsBnf, sErr));
	}
	
	hResult = poBenefit->get_BnfTypeId(&lBnfTypeId);
	if (lBnfTypeId == 0)
	{
		poBenefit->Term();
		poBenefit->Release();
		sErr = CStrFnc::LoadResString(scInsBnf);
		sErr += TEXT("\n\r");
		sErr += CStrFnc::LoadResString(icErrReqFld);
		sErr += TEXT("BnfTypeId");
		return (FlagExceptionMsg(IID_IAdmin, scInsBnf, sErr));
	}
	
	NULLBSTR(bstrGet);
	hResult = poBenefit->get_Description(&bstrGet);
	if (CStrFnc::IsEmpty(bstrGet))
	{
		poBenefit->Term();
		poBenefit->Release();
		sErr = CStrFnc::LoadResString(scInsBnf);
		sErr += TEXT("\n\r");
		sErr += CStrFnc::LoadResString(icErrReqFld);
		sErr += TEXT("Description");
		return (FlagExceptionMsg(IID_IAdmin, scInsBnf, sErr));
	}
	
	hResult = poBenefit->get_Cost(&lCost);
	if (lCost == 0)
	{
		poBenefit->Term();
		poBenefit->Release();
		sErr = CStrFnc::LoadResString(scInsBnf);
		sErr += TEXT("\n\r");
		sErr += CStrFnc::LoadResString(icErrReqFld);
		sErr += TEXT("Cost");
		return (FlagExceptionMsg(IID_IAdmin, scInsBnf, sErr));
	}
	
	hResult = poBenefit->get_Increment(&iIncrement);
	if (iIncrement == 0)
	{
		poBenefit->Term();
		poBenefit->Release();
		sErr = CStrFnc::LoadResString(scInsBnf);
		sErr += TEXT("\n\r");
		sErr += CStrFnc::LoadResString(icErrReqFld);
		sErr += TEXT("Increment");
		return (FlagExceptionMsg(IID_IAdmin, scInsBnf, sErr));
	}
	
	hResult = poBenefit->get_Maximum(&lMaximum);
	if (lMaximum == 0)
	{
		poBenefit->Term();
		poBenefit->Release();
		sErr = CStrFnc::LoadResString(scInsBnf);
		sErr += TEXT("\n\r");
		sErr += CStrFnc::LoadResString(icErrReqFld);
		sErr += TEXT("Maximum");
		return (FlagExceptionMsg(IID_IAdmin, scInsBnf, sErr));
	}
	
	hResult = poBenefit->get_Minimum(&lMinimum);
	if (lMinimum == 0)
	{
		poBenefit->Term();
		poBenefit->Release();
		sErr = CStrFnc::LoadResString(scInsBnf);
		sErr += TEXT("\n\r");
		sErr += CStrFnc::LoadResString(icErrReqFld);
		sErr += TEXT("Minimum");
		return (FlagExceptionMsg(IID_IAdmin, scInsBnf, sErr));
	}
	
	hResult = poBenefit->get_Multiples(&iMultiples);
	if (iMultiples == 0)
	{
		poBenefit->Term();
		poBenefit->Release();
		sErr = CStrFnc::LoadResString(scInsBnf);
		sErr += TEXT("\n\r");
		sErr += CStrFnc::LoadResString(icErrReqFld);
		sErr += TEXT("Multiples");
		return (FlagExceptionMsg(IID_IAdmin, scInsBnf, sErr));
	}
	
	if (m_bUseODBC)
	{
		sSQL = CStrFnc::LoadResString(scExecute);
		sSQL += CStrFnc::LoadResString(scIns_Benefit);
		sSQL += scSpace;
		
		sTemp.Format(TEXT("%ld"), lBnfTaxId);
		sSQL += sTemp.GetBstr();
		sSQL += scSpaceCommaSpace;
		
		sTemp.Format(TEXT("%ld"), lBnfTypeId);
		sSQL += sTemp.GetBstr();
		sSQL += scSpaceCommaSpace;
		hResult = poBenefit->get_Cost(&lCost);
		
		sTemp.Format(TEXT("%ld"), lCost);
		sSQL += sTemp.GetBstr();
		sSQL += scSpaceCommaSpace;
		
		NULLBSTR(bstrGet);
		
		hResult = poBenefit->get_Description(&bstrGet);
		sTemp = bstrTemp;
		sSQL += bstrGet;
		sSQL += scSQ;
		sSQL += scSpaceCommaSpace;
		
		sTemp.Format(TEXT("%ld"), iIncrement);
		sSQL += sTemp.GetBstr();
		sSQL += scSpaceCommaSpace;
		
		sTemp.Format(TEXT("%ld"), lMaximum);
		sSQL += sTemp.GetBstr();
		sSQL += scSpaceCommaSpace;
		
		sTemp.Format(TEXT("%ld"), lMinimum);
		sSQL += sTemp.GetBstr();
		sSQL += scSpaceCommaSpace;
		
		sTemp.Format(TEXT("%d"), iMultiples);
		sSQL += sTemp.GetBstr();
		
		if (ExecBooleanSQL(sSQL, pbSuccess) != NOERROR)
		{
			poBenefit->Term();
			poBenefit->Release();
			sErr = CStrFnc::LoadResString(scInsBnf);
			FlagExceptionMsg(IID_IAdmin, scInsBnf, sErr);
			return (NOERROR);
		}
	}
	else
	{
		hResult = ::CoCreateInstance(CLSID_CList, NULL, THISCLSCTX, IID_IList, (void **)&ParameterList);
		
		::VariantInit(&vResult);
		vResult.vt = VT_I4;
		V_I4(&vResult) = lBnfTaxId;
		hResult = ParameterList->Add(TEXT("@BnfTaxId"),&vResult);
		
		::VariantInit(&vResult);
		vResult.vt = VT_I4;
		V_I4(&vResult) = lBnfTypeId;
		hResult = ParameterList->Add(TEXT("@BnfTypeId"),&vResult);
		
		::VariantInit(&vResult);
		vResult.vt = VT_I4;
		V_I4(&vResult) = lCost;
		hResult = ParameterList->Add(TEXT("@Cost"),&vResult);
		
		NULLBSTR(bstrGet);
		hResult = poBenefit->get_Description(&bstrGet);
		::VariantInit(&vResult);
		vResult.vt = VT_BSTR;
		V_BSTR(&vResult) = bstrGet;
		hResult = ParameterList->Add(TEXT("@Description"),&vResult);
		
		::VariantInit(&vResult);
		vResult.vt = VT_I2;
		V_I2(&vResult) = iIncrement;
		hResult = ParameterList->Add(TEXT("@Increment"),&vResult);
		
		::VariantInit(&vResult);
		vResult.vt = VT_I4;
		V_I4(&vResult) = lMaximum;
		hResult = ParameterList->Add(TEXT("@Maximum"),&vResult);
		
		::VariantInit(&vResult);
		vResult.vt = VT_I4;
		V_I4(&vResult) = lMinimum;
		hResult = ParameterList->Add(TEXT("@Minimum"),&vResult);
		
		::VariantInit(&vResult);
		vResult.vt = VT_I2;
		V_I2(&vResult) = iMultiples;
		hResult = ParameterList->Add(TEXT("@Multiples"),&vResult);
		
		sSQL = CStrFnc::LoadResString(scIns_Benefit);
		hResult = ParameterList->QueryInterface(IID_IDispatch, (void **)&pIDispatch);
		if (ExecBooleanMDB(sSQL, pIDispatch, pbSuccess) != NOERROR)
		{
			poBenefit->Term();
			poBenefit->Release();
			pIDispatch->Release();
			ParameterList->Release();
			sErr = CStrFnc::LoadResString(scInsBnf);
			FlagExceptionMsg(IID_IAdmin, scInsBnf, sErr);
			return (NOERROR);
		}
		pIDispatch->Release();
		ParameterList->Release();
	}
		
	poBenefit->Term();
	poBenefit->Release();
	FillListBenefits();
	*pbSuccess = TRUE;
	
	return (NOERROR);
}
	
//Insert_BenefitLink
//Purpose		- Inserts specified BenefitLink object
//Inputs		- (poBenefitLink)	- BenefitLink object with insert information
//				- (pbSuccess)		- boolean output target
//Outputs		- (pbSuccess)		- boolean success value
HRESULT CAdmin::Insert_BenefitLink(IBenefitLink * poBenefitLink, BOOL * pbSuccess)
{
	HRESULT hResult;
	_bstr_t sSQL;
	_bstr_t sErr;
	_bstr_t sTemp;
	VARIANT vResult;
	long	lCoverageId;
	long	lTaxId;
	long	lEmployeePKId;
	long	lBenefitPKId;
	IBenefit *	pBenefit;
	IEmployee * pEmployee;
	CURRENCY	cyAmount;
	IList *		ParameterList;
	IDispatch *	pIDispatch;
	
	*pbSuccess = FALSE;
	
	hResult = poBenefitLink->get_Employee(&pEmployee);
	if (hResult != NOERROR)
	{
		poBenefitLink->Term();
		poBenefitLink->Release();
		sErr = CStrFnc::LoadResString(scInsBnfLnk);
		sErr += TEXT("\n\r");
		sErr += CStrFnc::LoadResString(icErrReqFld);
		sErr += TEXT("Bank.PKId");
		return (FlagExceptionMsg(IID_IAdmin, scInsBnfLnk, sErr));
	}
	else
	{
		if (pEmployee)
		{
			hResult = pEmployee->get_PKId(&lEmployeePKId);
			hResult = pEmployee->Release();
			if (lEmployeePKId == 0)
			{
				poBenefitLink->Term();
				poBenefitLink->Release();
				sErr = CStrFnc::LoadResString(scInsBnfLnk);
				sErr += TEXT("\n\r");
				sErr += CStrFnc::LoadResString(icErrReqFld);
				sErr += TEXT("Employee.PKId");
				return (FlagExceptionMsg(IID_IAdmin, scInsBnfLnk, sErr));
			}
		}
		else
		{
			poBenefitLink->Term();
			poBenefitLink->Release();
			sErr = CStrFnc::LoadResString(scInsBnfLnk);
			sErr += TEXT("\n\r");
			sErr += CStrFnc::LoadResString(icErrReqFld);
			sErr += TEXT("Bank.PKId");
			return (FlagExceptionMsg(IID_IAdmin, scInsBnfLnk, sErr));
		}
	}
		
	hResult = poBenefitLink->get_Benefit(&pBenefit);
	if (hResult != NOERROR)
	{
		poBenefitLink->Term();
		poBenefitLink->Release();
		sErr = CStrFnc::LoadResString(scInsBnfLnk);
		sErr += TEXT("\n\r");
		sErr += CStrFnc::LoadResString(icErrReqFld);
		sErr += TEXT("Bank.PKId");
		return (FlagExceptionMsg(IID_IAdmin, scInsBnfLnk, sErr));
	}
	else
	{
		if (pBenefit)
		{
			hResult = pBenefit->get_PKId(&lBenefitPKId);
			hResult = pBenefit->Release();
			if (lBenefitPKId == 0)
			{
				poBenefitLink->Term();
				poBenefitLink->Release();
				sErr = CStrFnc::LoadResString(scInsBnfLnk);
				sErr += TEXT("\n\r");
				sErr += CStrFnc::LoadResString(icErrReqFld);
				sErr += TEXT("Benefit.PKId");
				return (FlagExceptionMsg(IID_IAdmin, scInsBnfLnk, sErr));
			}
		}
		else
		{
			poBenefitLink->Term();
			poBenefitLink->Release();
			sErr = CStrFnc::LoadResString(scInsBnfLnk);
			sErr += TEXT("\n\r");
			sErr += CStrFnc::LoadResString(icErrReqFld);
			sErr += TEXT("Bank.PKId");
			return (FlagExceptionMsg(IID_IAdmin, scInsBnfLnk, sErr));
		}
	}
		
	hResult = poBenefitLink->get_CoverageId(&lCoverageId);
	if (lCoverageId == 0)
	{
		poBenefitLink->Term();
		poBenefitLink->Release();
		sErr = CStrFnc::LoadResString(scInsBnfLnk);
		sErr += TEXT("\n\r");
		sErr += CStrFnc::LoadResString(icErrReqFld);
		sErr += TEXT("Benefit.CoverageId");
		return (FlagExceptionMsg(IID_IAdmin, scInsBnfLnk, sErr));
	}
	
	hResult = poBenefitLink->get_TaxId(&lTaxId);
	if (lTaxId == 0)
	{
		poBenefitLink->Term();
		poBenefitLink->Release();
		sErr = CStrFnc::LoadResString(scInsBnfLnk);
		sErr += TEXT("\n\r");
		sErr += CStrFnc::LoadResString(icErrReqFld);
		sErr += TEXT("Benefit.TaxId");
		return (FlagExceptionMsg(IID_IAdmin, scInsBnfLnk, sErr));
	}
	
	if (m_bUseODBC)
	{
		if (ExecBooleanSQL(sSQL, pbSuccess) != NOERROR)
		{
			poBenefitLink->Term();
			poBenefitLink->Release();
			sErr = CStrFnc::LoadResString(scInsBnfLnk);
			FlagExceptionMsg(IID_IAdmin, scInsBnfLnk, sErr);
			return (NOERROR);
		}
	}
	else
	{
		hResult = ::CoCreateInstance(CLSID_CList, NULL, THISCLSCTX, IID_IList, (void **)&ParameterList);
		
		::VariantInit(&vResult);
		vResult.vt = VT_I4;
		V_I4(&vResult) = lBenefitPKId;
		hResult = ParameterList->Add(TEXT("@BenefitId"),&vResult);
		
		::VariantInit(&vResult);
		vResult.vt = VT_I4;
		V_I4(&vResult) = lCoverageId;
		hResult = ParameterList->Add(TEXT("@CoverageId"),&vResult);
		
		::VariantInit(&vResult);
		vResult.vt = VT_I4;
		V_I4(&vResult) = lEmployeePKId;
		hResult = ParameterList->Add(TEXT("@EmployeeId"),&vResult);
		
		::VariantInit(&vResult);
		vResult.vt = VT_I4;
		V_I4(&vResult) = lTaxId;
		hResult = ParameterList->Add(TEXT("@TaxId"),&vResult);
		
		hResult = poBenefitLink->get_Amount(&cyAmount);
		::VariantInit(&vResult);
		vResult.vt = VT_CY;
		V_CY(&vResult) = cyAmount;
		hResult = ParameterList->Add(TEXT("@Amount"),&vResult);
		
		sSQL = CStrFnc::LoadResString(scIns_EmplBnf);
		hResult = ParameterList->QueryInterface(IID_IDispatch, (void **)&pIDispatch);
		if (ExecBooleanMDB(sSQL, pIDispatch, pbSuccess) != NOERROR)
		{
			poBenefitLink->Term();
			poBenefitLink->Release();
			pIDispatch->Release();
			ParameterList->Release();
			sErr = CStrFnc::LoadResString(scInsBnfLnk);
			FlagExceptionMsg(IID_IAdmin, scInsBnfLnk, sErr);
			return (NOERROR);
		}
		pIDispatch->Release();
		ParameterList->Release();
	}
	
	poBenefitLink->Term();
	poBenefitLink->Release();
	*pbSuccess = TRUE;
	
	return (NOERROR);
}
	
//Insert_Dependant
//Purpose		- Inserts specified dependant object
//Inputs		- (poDependant)		- dependant object with insert information
//				- (pbSuccess)		- boolean output target
//Outputs		- (pbSuccess)		- boolean success value
HRESULT CAdmin::Insert_Dependant(IDependant * poDependant, BOOL * pbSuccess)
{
	HRESULT hResult;
	_bstr_t sSQL;
	_bstr_t sErr;
	CStrFnc	sTemp;
	VARIANT	vResult;
	BSTR	bstrGet;
	long	lDepTypeId;
	BOOL	bGender;
	BOOL	bStatus;
	_bstr_t	bstrTemp;
	DATE	TempDate;
	long	lEmployeePKId;
	IEmployee * pEmployee;
	IDispatch *	pIDispatch;
	IList *		ParameterList;
	

	bstrGet = NULL;
	hResult = NOERROR;
	*pbSuccess = FALSE;
	
	hResult = poDependant->get_DepTypeId(&lDepTypeId);
	if (lDepTypeId == 0)
	{
		poDependant->Term();
		poDependant->Release();
		sErr = CStrFnc::LoadResString(scInsDep);
		sErr += TEXT("\n\r");
		sErr += CStrFnc::LoadResString(icErrReqFld);
		sErr += TEXT("DepTypeId");
		return (FlagExceptionMsg(IID_IAdmin, scInsDep, sErr));
	}
	
	hResult = poDependant->get_Employee(&pEmployee);
	if (hResult != NOERROR)
	{
		poDependant->Term();
		poDependant->Release();
		sErr = CStrFnc::LoadResString(scInsDep);
		sErr += TEXT("\n\r");
		sErr += CStrFnc::LoadResString(icErrReqFld);
		sErr += TEXT("Employee.PKId");
		return (FlagExceptionMsg(IID_IAdmin, scInsDep, sErr));
	}
	else
	{
		if (pEmployee)
		{
			hResult = pEmployee->get_PKId(&lEmployeePKId);
			hResult = pEmployee->Release();
			if (lEmployeePKId == 0)
			{
				poDependant->Term();
				poDependant->Release();
				sErr = CStrFnc::LoadResString(scInsDep);
				sErr += TEXT("\n\r");
				sErr += CStrFnc::LoadResString(icErrReqFld);
				return (FlagExceptionMsg(IID_IAdmin, scInsDep, sErr));
			}
		}
		else
		{
			poDependant->Term();
			poDependant->Release();
			sErr = CStrFnc::LoadResString(scInsDep);
			sErr += TEXT("\n\r");
			sErr += CStrFnc::LoadResString(icErrReqFld);
			sErr += TEXT("Employee.PKId");
			return (FlagExceptionMsg(IID_IAdmin, scInsDep, sErr));
		}
	}
	
	NULLBSTR(bstrGet);
	hResult = poDependant->get_FirstName(&bstrGet);
	if (CStrFnc::IsEmpty(bstrGet))
	{
		poDependant->Term();
		poDependant->Release();
		sErr = CStrFnc::LoadResString(scInsDep);
		sErr += TEXT("\n\r");
		sErr += CStrFnc::LoadResString(icErrReqFld);
		sErr += TEXT("FirstName");
		return (FlagExceptionMsg(IID_IAdmin, scInsDep, sErr));
	}
	
	NULLBSTR(bstrGet);
	hResult = poDependant->get_LastName(&bstrGet);
	if (CStrFnc::IsEmpty(bstrGet))
	{
		poDependant->Term();
		poDependant->Release();
		sErr = CStrFnc::LoadResString(scInsDep);
		sErr += TEXT("\n\r");
		sErr += CStrFnc::LoadResString(icErrReqFld);
		sErr += TEXT("LastName");
		return (FlagExceptionMsg(IID_IAdmin, scInsDep, sErr));
	}
	
	hResult = poDependant->get_Gender(&bGender);
	if (hResult != NOERROR)
	{
		poDependant->Term();
		poDependant->Release();
		sErr = CStrFnc::LoadResString(scInsDep);
		sErr += TEXT("\n\r");
		sErr += CStrFnc::LoadResString(icErrReqFld);
		sErr += TEXT("Gender");
		return (FlagExceptionMsg(IID_IAdmin, scInsDep, sErr));
	}

	hResult = poDependant->get_Status(&bStatus);
	if (hResult != NOERROR)
	{
		poDependant->Term();
		poDependant->Release();
		sErr = CStrFnc::LoadResString(scInsDep);
		sErr += TEXT("\n\r");
		sErr += CStrFnc::LoadResString(icErrReqFld);
		sErr += TEXT("Status");
		return (FlagExceptionMsg(IID_IAdmin, scInsDep, sErr));
	}

	if (m_bUseODBC)
	{
		sSQL = CStrFnc::LoadResString(scExecute);
		sSQL += CStrFnc::LoadResString(scIns_Dependant);
		sSQL += scSpace;
		sTemp.Format(TEXT("%ld"), lDepTypeId);
		sSQL += sTemp.GetBstr();
		sSQL += scSpaceCommaSpace;
		sTemp.Format(TEXT("%ld"), lEmployeePKId);
		sSQL += sTemp.GetBstr();
		sSQL += scSpaceCommaSpace;
		
		hResult = poDependant->get_DateOfBirth(&TempDate);
		if (CStrFnc::IsValidDate(TempDate))
		{
			sSQL += scSQ;
			sSQL += CStrFnc::FormatDate(TempDate, scDateFormat);
			sSQL += scSQ;
			sSQL += scSpaceCommaSpace;
		}
		else
		{
			sSQL += scNull;
			sSQL += scSpaceCommaSpace;
		}
		
		NULLBSTR(bstrGet);
		hResult = poDependant->get_FirstName(&bstrGet);
		sSQL += scSQ;
		sSQL += bstrGet;
		sSQL += scSQ;
		sSQL += scSpaceCommaSpace;
		hResult = poDependant->get_Gender(&bGender);
		if (bGender)
		{
			sSQL += scZeroStr;
			sSQL += scSpaceCommaSpace;
		}
		else
		{
			sSQL += scOneStr;
			sSQL += scSpaceCommaSpace;
		}
		
		NULLBSTR(bstrGet);
		hResult = poDependant->get_LastName(&bstrGet);
		sSQL += scSQ;
		sSQL += bstrGet;
		sSQL += scSQ;
		sSQL += scSpaceCommaSpace;
		hResult = poDependant->get_Status(&bStatus);
		if (bStatus)
		{
			sSQL += scZeroStr;
			sSQL += scSpaceCommaSpace;
		}
		else
		{
			sSQL += scOneStr;
			sSQL += scSpaceCommaSpace;
		}
		if (ExecBooleanSQL(sSQL, pbSuccess) != NOERROR)
		{
			poDependant->Term();
			poDependant->Release();
			sErr = CStrFnc::LoadResString(scInsDep);
			FlagExceptionMsg(IID_IAdmin, scInsDep, sErr);
			return (NOERROR);
		}
	}
	else
	{
		hResult = ::CoCreateInstance(CLSID_CList, NULL, THISCLSCTX, IID_IList, (void **)&ParameterList);
		
		::VariantInit(&vResult);
		vResult.vt = VT_I4;
		V_I4(&vResult) = lDepTypeId;
		hResult = ParameterList->Add(TEXT("@DepTypeId"),&vResult);
		
		::VariantInit(&vResult);
		vResult.vt = VT_I4;
		V_I4(&vResult) = lEmployeePKId;
		hResult = ParameterList->Add(TEXT("@EmployeeId"),&vResult);
		
		hResult = poDependant->get_DateOfBirth(&TempDate);
		::VariantInit(&vResult);
		vResult.vt = VT_DATE;
		V_DATE(&vResult) = TempDate;
		hResult = ParameterList->Add(TEXT("@DateOfBirth"),&vResult);
		
		NULLBSTR(bstrGet);
		hResult = poDependant->get_FirstName(&bstrGet);
		::VariantInit(&vResult);
		vResult.vt = VT_BSTR;
		V_BSTR(&vResult) = bstrGet;
		hResult = ParameterList->Add(TEXT("@FirstName"),&vResult);
		
		::VariantInit(&vResult);
		vResult.vt = VT_I2;
		V_I4(&vResult) = bGender;
		hResult = ParameterList->Add(TEXT("@Gender"),&vResult);
		
		NULLBSTR(bstrGet);
		hResult = poDependant->get_LastName(&bstrGet);
		::VariantInit(&vResult);
		vResult.vt = VT_BSTR;
		V_BSTR(&vResult) = bstrGet;
		hResult = ParameterList->Add(TEXT("@LastName"),&vResult);
		
		::VariantInit(&vResult);
		vResult.vt = VT_I2;
		V_I4(&vResult) = bStatus;
		hResult = ParameterList->Add(TEXT("@Status"),&vResult);
		
		sSQL = CStrFnc::LoadResString(scIns_Dependant);
		hResult = ParameterList->QueryInterface(IID_IDispatch, (void **)&pIDispatch);
		if (ExecBooleanMDB(sSQL, pIDispatch, pbSuccess) != NOERROR)
		{
			poDependant->Term();
			poDependant->Release();
			pIDispatch->Release();
			ParameterList->Release();
			sErr = CStrFnc::LoadResString(scInsDep);
			FlagExceptionMsg(IID_IAdmin, scInsDep, sErr);
			return (NOERROR);
		}
		pIDispatch->Release();
		ParameterList->Release();
	}
	
	poDependant->Term();
	poDependant->Release();
	*pbSuccess = TRUE;
	
	return (NOERROR);
}
	
//Insert_Employee
//Purpose		- Inserts specified employee object
//Inputs		- (poEmployee)		- employee object with insert information
//				- (pbSuccess)		- boolean output target
//Outputs		- (pbSuccess)		- boolean success value
HRESULT CAdmin::Insert_Employee(IEmployee * poEmployee, BOOL * pbSuccess)
{
	HRESULT hResult;
	_bstr_t sSQL;
	_bstr_t sErr;
	CStrFnc sTemp;
	VARIANT vResult;
	BSTR	bstrGet;
	long	lPKId;
	_bstr_t bstrTemp;
	BOOL	bGender;
	DATE	TempDate;
	CURRENCY Salary;
	IBank *		oBank;
	IDispatch *	pIDispatch;
	IList *		ParameterList;
	
	hResult = NOERROR;
	bstrGet = NULL;
	*pbSuccess = FALSE;
	
	hResult = poEmployee->get_Bank(&oBank);
	if (hResult != NOERROR)
	{
		poEmployee->Term();
		poEmployee->Release();
		sErr = CStrFnc::LoadResString(scInsEmpl);
		sErr += TEXT("\n\r");
		sErr += CStrFnc::LoadResString(icErrReqFld);
		sErr += TEXT("Bank.PKId");
		return (FlagExceptionMsg(IID_IAdmin, scInsEmpl, sErr));
	}
	else
	{
		if (oBank)
		{
			hResult = oBank->get_PKId(&lPKId);
			hResult = oBank->Release();
			if (lPKId == 0)
			{
				poEmployee->Term();
				poEmployee->Release();
				sErr = CStrFnc::LoadResString(scInsEmpl);
				sErr += TEXT("\n\r");
				sErr += CStrFnc::LoadResString(icErrReqFld);
				sErr += TEXT("Bank.PKId");
				return (FlagExceptionMsg(IID_IAdmin, scInsEmpl, sErr));
			}
		}
		else
		{
			poEmployee->Term();
			poEmployee->Release();
			sErr = CStrFnc::LoadResString(scInsEmpl);
			sErr += TEXT("\n\r");
			sErr += CStrFnc::LoadResString(icErrReqFld);
			sErr += TEXT("Bank.PKId");
			return (FlagExceptionMsg(IID_IAdmin, scInsEmpl, sErr));
		}
	}
	
	NULLBSTR(bstrGet);
	poEmployee->get_FirstName(&bstrGet);
	if (CStrFnc::IsEmpty(bstrGet))
	{
		poEmployee->Term();
		poEmployee->Release();
		sErr = CStrFnc::LoadResString(scInsEmpl);
		sErr += TEXT("\n\r");
		sErr += CStrFnc::LoadResString(icErrReqFld);
		sErr += TEXT("FirstName");
		return (FlagExceptionMsg(IID_IAdmin, scInsEmpl, sErr));
	}
	
	NULLBSTR(bstrGet);
	poEmployee->get_LastName(&bstrGet);
	if (CStrFnc::IsEmpty(bstrGet))
	{
		poEmployee->Term();
		poEmployee->Release();
		sErr = CStrFnc::LoadResString(scInsEmpl);
		sErr += TEXT("\n\r");
		sErr += CStrFnc::LoadResString(icErrReqFld);
		sErr += TEXT("LastName");
		return (FlagExceptionMsg(IID_IAdmin, scInsEmpl, sErr));
	}
	
	NULLBSTR(bstrGet);
	poEmployee->get_Position(&bstrGet);
	if (CStrFnc::IsEmpty(bstrGet))
	{
		poEmployee->Term();
		poEmployee->Release();
		sErr = CStrFnc::LoadResString(scInsEmpl);
		sErr += TEXT("\n\r");
		sErr += CStrFnc::LoadResString(icErrReqFld);
		sErr += TEXT("Position");
		return (FlagExceptionMsg(IID_IAdmin, scInsEmpl, sErr));
	}
	
	NULLBSTR(bstrGet);
	poEmployee->get_UserName(&bstrGet);
	if (CStrFnc::IsEmpty(bstrGet))
	{
		poEmployee->Term();
		poEmployee->Release();
		sErr = CStrFnc::LoadResString(scInsEmpl);
		sErr += TEXT("\n\r");
		sErr += CStrFnc::LoadResString(icErrReqFld);
		sErr += TEXT("UserName");
		return (FlagExceptionMsg(IID_IAdmin, scInsEmpl, sErr));
	}
	
	poEmployee->get_StartDate(&TempDate);
	if (!CStrFnc::IsValidDate(TempDate))
	{
		poEmployee->Term();
		poEmployee->Release();
		sErr = CStrFnc::LoadResString(scInsEmpl);
		sErr += TEXT("\n\r");
		sErr += CStrFnc::LoadResString(icErrReqFld);
		sErr += TEXT("StartDate");
		return (FlagExceptionMsg(IID_IAdmin, scInsEmpl, sErr));
	}
	
	hResult = poEmployee->get_Gender(&bGender);
	if (hResult != NOERROR)
	{
		poEmployee->Term();
		poEmployee->Release();
		sErr = CStrFnc::LoadResString(scInsEmpl);
		sErr += TEXT("\n\r");
		sErr += CStrFnc::LoadResString(icErrReqFld);
		sErr += TEXT("Gender");
		return (FlagExceptionMsg(IID_IAdmin, scInsEmpl, sErr));
	}

	if (m_bUseODBC)
	{
		sSQL = CStrFnc::LoadResString(scExecute);
		sSQL += CStrFnc::LoadResString(scIns_Employee);
		sSQL += scSpace;
		sTemp.Format(TEXT("%ld"), lPKId);
		sSQL += sTemp.GetBstr();
		sSQL += scSpaceCommaSpace;
		
		NULLBSTR(bstrGet);
		hResult = poEmployee->get_Address1(&bstrGet);
		sTemp = bstrGet;
		sTemp.Trim(sTemp.Left(60));
		sSQL += scSQ;
		sSQL += sTemp.GetBstr();
		sSQL += scSQ;
		sSQL += scSpaceCommaSpace;
		
		NULLBSTR(bstrGet);
		hResult = poEmployee->get_Address2(&bstrGet);
		sTemp = bstrGet;
		sTemp.Trim(sTemp.Left(60));
		sSQL += scSQ;
		sSQL += sTemp.GetBstr();
		sSQL += scSQ;
		sSQL += scSpaceCommaSpace;
		
		NULLBSTR(bstrGet);
		hResult = poEmployee->get_BankAccount(&bstrGet);
		sTemp = bstrGet;
		sTemp.Trim(sTemp.Left(60));
		sSQL += scSQ;
		sSQL += sTemp.GetBstr();
		sSQL += scSQ;
		sSQL += scSpaceCommaSpace;
		
		NULLBSTR(bstrGet);
		hResult = poEmployee->get_City(&bstrGet);
		sTemp = bstrGet;
		sTemp.Trim(sTemp.Left(60));
		sSQL += scSQ;
		sSQL += sTemp.GetBstr();
		sSQL += scSQ;
		sSQL += scSpaceCommaSpace;
		
		hResult = poEmployee->get_DateOfBirth(&TempDate);
		if (CStrFnc::IsValidDate(TempDate))
		{
			sSQL += scSQ;
			sSQL += CStrFnc::FormatDate(TempDate, scDateFormat);
			sSQL += scSQ;
			sSQL += scSpaceCommaSpace;
		}
		else
		{
			sSQL += scNull;
			sSQL += scSpaceCommaSpace;
		}
		
		hResult = poEmployee->get_EndDate(&TempDate);
		if (CStrFnc::IsValidDate(TempDate))
		{
			sSQL += scSQ;
			sSQL += CStrFnc::FormatDate(TempDate, scDateFormat);
			sSQL += scSQ;
			sSQL += scSpaceCommaSpace;
		}
		else
		{
			sSQL += scNull;
			sSQL += scSpaceCommaSpace;
		}
		
		NULLBSTR(bstrGet);
		hResult = poEmployee->get_FirstName(&bstrGet);
		sTemp = bstrGet;
		sTemp.Trim(sTemp.Left(35));
		sSQL += scSQ;
		sSQL += sTemp.GetBstr();
		sSQL += scSQ;
		sSQL += scSpaceCommaSpace;
		
		hResult = poEmployee->get_Gender(&bGender);
		if (bGender)
		{
			sSQL += scZeroStr;
			sSQL += scSpaceCommaSpace;
		}
		else
		{
			sSQL += scOneStr;
			sSQL += scSpaceCommaSpace;
		}
		
		NULLBSTR(bstrGet);
		hResult = poEmployee->get_LastName(&bstrGet);
		sTemp = bstrGet;
		sTemp.Trim(sTemp.Left(35));
		sSQL += scSQ;
		sSQL += sTemp.GetBstr();
		sSQL += scSQ;
		sSQL += scSpaceCommaSpace;
		
		NULLBSTR(bstrGet);
		hResult = poEmployee->get_Position(&bstrGet);
		sTemp = bstrGet;
		sTemp.Trim(sTemp.Left(35));
		sSQL += scSQ;
		sSQL += sTemp.GetBstr();
		sSQL += scSQ;
		sSQL += scSpaceCommaSpace;
		
		if (ExecBooleanSQL(sSQL, pbSuccess) != NOERROR)
		{
			poEmployee->Term();
			poEmployee->Release();
			sErr = CStrFnc::LoadResString(scInsEmpl);
			FlagExceptionMsg(IID_IAdmin, scInsEmpl, sErr);
			return (NOERROR);
		}
	}
	else
	{
		hResult = ::CoCreateInstance(CLSID_CList, NULL, THISCLSCTX, IID_IList, (void **)&ParameterList);
		
		::VariantInit(&vResult);
		vResult.vt = VT_I4;
		V_I4(&vResult) = lPKId;
		hResult = ParameterList->Add(TEXT("@BankId"),&vResult);
		
		NULLBSTR(bstrGet);
		hResult = poEmployee->get_Address1(&bstrGet);
		::VariantInit(&vResult);
		vResult.vt = VT_BSTR;
		V_BSTR(&vResult) = bstrGet;
		hResult = ParameterList->Add(TEXT("@Address1"),&vResult);
		
		NULLBSTR(bstrGet);
		hResult = poEmployee->get_Address2(&bstrGet);
		::VariantInit(&vResult);
		vResult.vt = VT_BSTR;
		V_BSTR(&vResult) = bstrGet;
		hResult = ParameterList->Add(TEXT("@Address2"),&vResult);
		
		NULLBSTR(bstrGet);
		hResult = poEmployee->get_BankAccount(&bstrGet);
		::VariantInit(&vResult);
		vResult.vt = VT_BSTR;
		V_BSTR(&vResult) = bstrGet;
		hResult = ParameterList->Add(TEXT("@BankAccount"),&vResult);
		
		NULLBSTR(bstrGet);
		hResult = poEmployee->get_City(&bstrGet);
		::VariantInit(&vResult);
		vResult.vt = VT_BSTR;
		V_BSTR(&vResult) = bstrGet;
		hResult = ParameterList->Add(TEXT("@City"),&vResult);
		
		hResult = poEmployee->get_DateOfBirth(&TempDate);
		::VariantInit(&vResult);
		vResult.vt = VT_DATE;
		V_DATE(&vResult) = TempDate;
		hResult = ParameterList->Add(TEXT("@DateOfBirth"),&vResult);
		
		hResult = poEmployee->get_EndDate(&TempDate);
		::VariantInit(&vResult);
		vResult.vt = VT_DATE;
		V_DATE(&vResult) = TempDate;
		hResult = ParameterList->Add(TEXT("@EndDate"),&vResult);
		
		NULLBSTR(bstrGet);
		hResult = poEmployee->get_FirstName(&bstrGet);
		::VariantInit(&vResult);
		vResult.vt = VT_BSTR;
		V_BSTR(&vResult) = bstrGet;
		hResult = ParameterList->Add(TEXT("@FirstName"),&vResult);
		
		hResult = poEmployee->get_Gender(&bGender);
		::VariantInit(&vResult);
		vResult.vt = VT_BOOL;
		V_BOOL(&vResult) = bGender;
		hResult = ParameterList->Add(TEXT("@Gender"),&vResult);
		
		NULLBSTR(bstrGet);
		hResult = poEmployee->get_LastName(&bstrGet);
		::VariantInit(&vResult);
		vResult.vt = VT_BSTR;
		V_BSTR(&vResult) = bstrGet;
		hResult = ParameterList->Add(TEXT("@LastName"),&vResult);
		
		NULLBSTR(bstrGet);
		hResult = poEmployee->get_Position(&bstrGet);
		::VariantInit(&vResult);
		vResult.vt = VT_BSTR;
		V_BSTR(&vResult) = bstrGet;
		hResult = ParameterList->Add(TEXT("@Position"),&vResult);
		
		hResult = poEmployee->get_Salary(&Salary);
		::VariantInit(&vResult);
		vResult.vt = VT_CY;
		V_CY(&vResult) = Salary;
		hResult = ParameterList->Add(TEXT("@Salary"),&vResult);
		
		hResult = poEmployee->get_StartDate(&TempDate);
		::VariantInit(&vResult);
		vResult.vt = VT_DATE;
		V_DATE(&vResult) = TempDate;
		hResult = ParameterList->Add(TEXT("@StartDate"),&vResult);
		
		NULLBSTR(bstrGet);
		hResult = poEmployee->get_UserName(&bstrGet);
		::VariantInit(&vResult);
		vResult.vt = VT_BSTR;
		V_BSTR(&vResult) = bstrGet;
		hResult = ParameterList->Add(TEXT("@UserName"),&vResult);
		
		sSQL = CStrFnc::LoadResString(scIns_Employee);
		hResult = ParameterList->QueryInterface(IID_IDispatch, (void**)&pIDispatch);
		if (ExecBooleanMDB(sSQL, pIDispatch, pbSuccess) != NOERROR)
		{
			poEmployee->Term();
			poEmployee->Release();
			pIDispatch->Release();
			ParameterList->Release();
			sErr = CStrFnc::LoadResString(scInsEmpl);
			FlagExceptionMsg(IID_IAdmin, scInsEmpl, sErr);
			return (NOERROR);
		}
		pIDispatch->Release();
		ParameterList->Release();
	}
	
	poEmployee->Term();
	poEmployee->Release();
	FillListEmployees();
	
	*pbSuccess = TRUE;
	return (NOERROR);
}

//StartRemoteObject
//Purpose	- To create a remote instance of a server and starts its 'Init' method
//Inputs	- (bstrRemoteProgId)	- registered progid as found in registry
//			- (bUseODBC)			- same flags as passed into local 'Init' routine
//			- (DataSource)			- "	"	"	"	"	"	"	"	"	"	"	"	"
//			- (iSourceType)			- "	"	"	"	"	"	"	"	"	"	"	"	"
//			- (iRemoteSvr)			- "	"	"	"	"	"	"	"	"	"	"	"	"
//			- (iLinkSvr))			- "	"	"	"	"	"	"	"	"	"	"	"	"
//Outputs	- (BOOL)	- bollean flag indicating success or failure
HRESULT CAdmin::StartRemoteObject(BSTR bstrRemoteProgId, BOOL bUseODBC, BSTR DataSource, int iSourceType, int iRemoteSvr, int iLinkSvr)
{
	CLSID	ClassId;
	HRESULT	hResult;
	VARIANT vResult;
	BSTR *	pName;
	DISPPARAMS pDispParams;
	VARIANTARG vArgs[12];
	int		n;
	
	hResult = NOERROR;

	hResult = CLSIDFromProgID(bstrRemoteProgId, &ClassId);
	if (hResult != NOERROR)
		return (hResult);
	
	hResult = CoCreateInstance(	ClassId,
		NULL,										// aggregate
		CLSCTX_ALL,//CLSCTX_LOCAL_SERVER,	// outproc or remote
		IID_IDispatch,
		(void **)&m_pIRemoteBenefitsServer);

	if (hResult != NOERROR)
		return (hResult);
	
	for (n = 0; n < REMOTESERVER_METHOD_COUNT; n++)
	{
		RemoteServer_Dispatch_Ids[n] = 0;
		pName = &RemoteServer_Methods[n];

		hResult = m_pIRemoteBenefitsServer->GetIDsOfNames(
									IID_NULL,
									pName,
									1,
									LOCALE_USER_DEFAULT,
									&RemoteServer_Dispatch_Ids[n]);
		if (hResult != NOERROR)
			return (hResult);
	}
	
	for (n = 0; n < 12; n++)
	{
		::VariantInit(&vArgs[n]);
	}
	
	vArgs[4].vt = VT_BOOL;
	vArgs[4].boolVal = bUseODBC;
	vArgs[3].vt = VT_BSTR;
	vArgs[3].bstrVal = DataSource;
	vArgs[2].vt = VT_I2;
	vArgs[2].iVal = iSourceType;
	vArgs[1].vt = VT_I2;
	vArgs[1].iVal = iRemoteSvr;
	vArgs[0].vt = VT_I2;
	vArgs[0].iVal = iLinkSvr;
	pDispParams.cArgs = 5;
	pDispParams.cNamedArgs = 0;
	pDispParams.rgvarg = vArgs;
	
	::VariantInit(&vResult);
	vResult.vt = VT_BOOL;
	
	hResult = m_pIRemoteBenefitsServer->Invoke(
						RemoteServer_Dispatch_Ids[REMOTESERVER_METHOD_INIT],
						IID_NULL,
						LOCALE_USER_DEFAULT,
						DISPATCH_METHOD,
						&pDispParams,
						&vResult,
						&m_RemoteServer_ExceptionInfo,
						NULL);
	
	if (hResult != NOERROR)
	{
		m_pIRemoteBenefitsServer->Release();
		m_pIRemoteBenefitsServer = NULL;
		return (hResult);
	}
	return (hResult);
}

//Update_Bank
//Purpose	- Updates a specified bank with given data
//Inputs	- (poBank)			- bank object containing information to be updated
//			- (pbSuccess)		- boolean output target
//Outputs	- (pbSuccess)		- boolean success value
HRESULT CAdmin::Update_Bank(IBank * poBank, BOOL * pbSuccess)
{
	ASSERT(poBank != NULL);
	ASSERT(pbSuccess != NULL);
	
	HRESULT hResult;
	_bstr_t sSQL;
	_bstr_t sErr;
	_bstr_t bstrTemp;
	CStrFnc sTemp;
	VARIANT vResult;
	BSTR	bstrGet;
	long	lPKId;
	IDispatch*	pIDispatch;
	IList * ParameterList;
	
	hResult		= NOERROR;
	bstrGet		= NULL;
	*pbSuccess	= FALSE;
	
	hResult = poBank->get_PKId(&lPKId);
	if (lPKId == 0)
	{
		sErr = CStrFnc::LoadResString(scUpdBank);
		sErr += TEXT("\n\r");
		sErr += CStrFnc::LoadResString(icErrReqFld);
		sErr += TEXT("PKId");
		return (FlagExceptionMsg(IID_IAdmin, scUpdBank, sErr));
	}
	
	NULLBSTR(bstrGet);
	hResult = poBank->get_Code(&bstrGet);
	if (CStrFnc::IsEmpty(bstrGet))
	{
		sErr = CStrFnc::LoadResString(scUpdBank);
		sErr += TEXT("\n\r");
		sErr += CStrFnc::LoadResString(icErrReqFld);
		sErr += TEXT("Code");
		return (FlagExceptionMsg(IID_IAdmin, scUpdBank, sErr));
	}
	
	NULLBSTR(bstrGet);
	hResult = poBank->get_Name(&bstrGet);
	if (CStrFnc::IsEmpty(bstrGet))
	{
		sErr = CStrFnc::LoadResString(scUpdBank);
		sErr += TEXT("\n\r");
		sErr += CStrFnc::LoadResString(icErrReqFld);
		sErr += TEXT("Name");
		return (FlagExceptionMsg(IID_IAdmin, scUpdBank, sErr));
	}
	
	if (m_bUseODBC)
	{
		sSQL = CStrFnc::LoadResString(scExecute);
		sSQL += CStrFnc::LoadResString(scUpd_Bank);
		sSQL += scSpace;
		
		sTemp.Format(TEXT("%ld"), lPKId);
		sSQL += sTemp.GetBstr();
		sSQL += scSpaceCommaSpace;
		
		NULLBSTR(bstrGet);
		hResult = poBank->get_Address1(&bstrGet);
		sTemp = bstrGet;
		sTemp.Trim(sTemp.Left(60));
		sSQL += scSQ;
		sSQL += sTemp.GetBstr();
		sSQL += scSQ;
		sSQL += scSpaceCommaSpace;
		
		NULLBSTR(bstrGet);
		hResult = poBank->get_Address2(&bstrGet);
		sTemp = bstrGet;
		sTemp.Trim(sTemp.Left(60));
		sSQL += scSQ;
		sSQL += sTemp.GetBstr();
		sSQL += scSQ;
		sSQL += scSpaceCommaSpace;
		
		NULLBSTR(bstrGet);
		hResult = poBank->get_City(&bstrGet);
		sTemp = bstrGet;
		sTemp.Trim(sTemp.Left(50));
		sSQL += scSQ;
		sSQL += sTemp.GetBstr();
		sSQL += scSQ;
		sSQL += scSpaceCommaSpace;
		
		NULLBSTR(bstrGet);
		hResult = poBank->get_Code(&bstrGet);
		sTemp = bstrGet;
		sTemp.Trim(sTemp.Left(16));
		sSQL += scSQ;
		sSQL += sTemp.GetBstr();
		sSQL += scSQ;
		sSQL += scSpaceCommaSpace;
		
		NULLBSTR(bstrGet);
		hResult = poBank->get_Name(&bstrGet);
		sTemp = bstrGet;
		sTemp.Trim(sTemp.Left(80));
		sSQL += scSQ;
		sSQL += sTemp.GetBstr();
		sSQL += scSQ;
		sSQL += scSpaceCommaSpace;
		
		if (ExecBooleanSQL(sSQL, pbSuccess) != NOERROR)
		{
			sErr = CStrFnc::LoadResString(scUpdBank);
			FlagExceptionMsg(IID_IAdmin, scUpdBank, sErr);
			return (NOERROR);
		}
	}
	else
	{
		hResult = ::CoCreateInstance(CLSID_CList, NULL, THISCLSCTX, IID_IList, (void **)&ParameterList);
		
		::VariantInit(&vResult);
		vResult.vt = VT_I4;
		V_I4(&vResult) = lPKId;
		hResult = ParameterList->Add(TEXT("@PKId"),&vResult);
		
		NULLBSTR(bstrGet);
		hResult = poBank->get_Address1(&bstrGet);
		::VariantInit(&vResult);
		vResult.vt = VT_BSTR;
		V_BSTR(&vResult) = bstrGet;
		hResult = ParameterList->Add(TEXT("@Address1"),&vResult);
		
		NULLBSTR(bstrGet);
		hResult = poBank->get_Address2(&bstrGet);
		::VariantInit(&vResult);
		vResult.vt = VT_BSTR;
		V_BSTR(&vResult) = bstrGet;
		hResult = ParameterList->Add(TEXT("@Address2"),&vResult);
		
		NULLBSTR(bstrGet);
		hResult = poBank->get_City(&bstrGet);
		::VariantInit(&vResult);
		vResult.vt = VT_BSTR;
		V_BSTR(&vResult) = bstrGet;
		hResult = ParameterList->Add(TEXT("@City"),&vResult);
		
		NULLBSTR(bstrGet);
		hResult = poBank->get_Code(&bstrGet);
		::VariantInit(&vResult);
		vResult.vt = VT_BSTR;
		V_BSTR(&vResult) = bstrGet;
		hResult = ParameterList->Add(TEXT("@Code"),&vResult);
		
		NULLBSTR(bstrGet);
		hResult = poBank->get_Name(&bstrGet);
		::VariantInit(&vResult);
		vResult.vt = VT_BSTR;
		V_BSTR(&vResult) = bstrGet;
		hResult = ParameterList->Add(TEXT("@Name"),&vResult);
		
		sSQL = CStrFnc::LoadResString(scUpd_Bank);
		hResult = ParameterList->QueryInterface(IID_IDispatch, (void **)&pIDispatch);
		if (ExecBooleanMDB(sSQL, pIDispatch, pbSuccess) != NOERROR)
		{
			pIDispatch->Release();
			ParameterList->Release();
			sErr = CStrFnc::LoadResString(scUpdBank);
			FlagExceptionMsg(IID_IAdmin, scUpdBank, sErr);
			return (NOERROR);
		}
		pIDispatch->Release();
		ParameterList->Release();
	}
	
	poBank->Term();
	poBank->Release();
	*pbSuccess = TRUE;
	return (NOERROR);
}

//Update_Benefit
//Purpose	- Updates an object
//Inputs	- (poBenefit)		- object with update information
//			- (pbSuccess)		- boolean output target
//Outputs	- (pbSuccess)		- boolean success value
HRESULT CAdmin::Update_Benefit(IBenefit * poBenefit, BOOL * pbSuccess)
{
	HRESULT hResult;
	_bstr_t sSQL;
	_bstr_t sErr;
	CStrFnc sTemp;
	VARIANT vResult;
	BSTR	bstrGet;
	long	lBnfTaxId;
	long	lBnfTypeId;
	short	iIncrement;
	long	lMaximum;
	long	lMinimum;
	long	lPKId;
	short	iMultiples;
	long	lCost;
	_bstr_t bstrTemp;
	IDispatch *	pIDispatch;
	IList *		ParameterList;
	

	hResult		= NOERROR;
	bstrGet		= NULL;
	*pbSuccess	= FALSE;
	
	hResult = poBenefit->get_PKId(&lPKId);
	if (lPKId == 0)
	{
		sErr = CStrFnc::LoadResString(scUpdBnf);
		sErr += TEXT("\n\r");
		sErr += CStrFnc::LoadResString(icErrReqFld);
		sErr += TEXT("PKId");
		return (FlagExceptionMsg(IID_IAdmin, scUpdBnf, sErr));
	}
	
	hResult = poBenefit->get_BnfTaxId(&lBnfTaxId);
	if (lBnfTaxId == 0)
	{
		sErr = CStrFnc::LoadResString(scUpdBnf);
		sErr += TEXT("\n\r");
		sErr += CStrFnc::LoadResString(icErrReqFld);
		sErr += TEXT("BnfTaxId");
		return (FlagExceptionMsg(IID_IAdmin, scUpdBnf, sErr));
	}
	
	hResult = poBenefit->get_BnfTypeId(&lBnfTypeId);
	if (lBnfTypeId == 0)
	{
		sErr = CStrFnc::LoadResString(scUpdBnf);
		sErr += TEXT("\n\r");
		sErr += CStrFnc::LoadResString(icErrReqFld);
		sErr += TEXT("BnfTypeId");
		return (FlagExceptionMsg(IID_IAdmin, scUpdBnf, sErr));
	}
	
	NULLBSTR(bstrGet);
	hResult = poBenefit->get_Description(&bstrGet);
	if (CStrFnc::IsEmpty(bstrGet))
	{
		sErr = CStrFnc::LoadResString(scUpdBnf);
		sErr += TEXT("\n\r");
		sErr += CStrFnc::LoadResString(icErrReqFld);
		sErr += TEXT("Description");
		return (FlagExceptionMsg(IID_IAdmin, scUpdBnf, sErr));
	}
	
	hResult = poBenefit->get_Cost(&lCost);
	if (lCost == 0)
	{
		sErr = CStrFnc::LoadResString(scUpdBnf);
		sErr += TEXT("\n\r");
		sErr += CStrFnc::LoadResString(icErrReqFld);
		sErr += TEXT("Cost");
		return (FlagExceptionMsg(IID_IAdmin, scUpdBnf, sErr));
	}
	
	hResult = poBenefit->get_Increment(&iIncrement);
	if (iIncrement == 0)
	{
		sErr = CStrFnc::LoadResString(scUpdBnf);
		sErr += TEXT("\n\r");
		sErr += CStrFnc::LoadResString(icErrReqFld);
		sErr += TEXT("Increment");
		return (FlagExceptionMsg(IID_IAdmin, scUpdBnf, sErr));
	}
	
	hResult = poBenefit->get_Maximum(&lMaximum);
	if (lMaximum == 0)
	{
		sErr = CStrFnc::LoadResString(scUpdBnf);
		sErr += TEXT("\n\r");
		sErr += CStrFnc::LoadResString(icErrReqFld);
		sErr += TEXT("Maximum");
		return (FlagExceptionMsg(IID_IAdmin, scUpdBnf, sErr));
	}
	
	hResult = poBenefit->get_Minimum(&lMinimum);
	if (lMinimum == 0)
	{
		sErr = CStrFnc::LoadResString(scUpdBnf);
		sErr += TEXT("\n\r");
		sErr += CStrFnc::LoadResString(icErrReqFld);
		sErr += TEXT("Minimum");
		return (FlagExceptionMsg(IID_IAdmin, scUpdBnf, sErr));
	}
	
	hResult = poBenefit->get_Multiples(&iMultiples);
	if (iMultiples == 0)
	{
		sErr = CStrFnc::LoadResString(scUpdBnf);
		sErr += TEXT("\n\r");
		sErr += CStrFnc::LoadResString(icErrReqFld);
		sErr += TEXT("Multiples");
		return (FlagExceptionMsg(IID_IAdmin, scUpdBnf, sErr));
	}
	
	hResult = poBenefit->get_Cost(&lCost);
	if (lCost == 0)
	{
		sErr = CStrFnc::LoadResString(scUpdBnf);
		sErr += TEXT("\n\r");
		sErr += CStrFnc::LoadResString(icErrReqFld);
		sErr += TEXT("Multiples");
		return (FlagExceptionMsg(IID_IAdmin, scUpdBnf, sErr));
	}
	
	if (m_bUseODBC)
	{
		sSQL = CStrFnc::LoadResString(scExecute);
		sSQL += CStrFnc::LoadResString(scUpd_Benefit);
		sSQL += scSpace;
		
		sTemp.Format(TEXT("%ld"), lPKId);
		sSQL += sTemp.GetBstr();
		sSQL += scSpaceCommaSpace;
		
		sTemp.Format(TEXT("%ld"), lBnfTaxId);
		sSQL += sTemp.GetBstr();
		sSQL += scSpaceCommaSpace;
		
		sTemp.Format(TEXT("%ld"), lBnfTypeId);
		sSQL += sTemp.GetBstr();
		sSQL += scSpaceCommaSpace;
		hResult = poBenefit->get_Cost(&lCost);
		
		sTemp.Format(TEXT("%ld"), lCost);
		sSQL += sTemp.GetBstr();
		sSQL += scSpaceCommaSpace;
		
		NULLBSTR(bstrGet);
		
		hResult = poBenefit->get_Description(&bstrGet);
		sTemp = bstrGet;
		sSQL += sTemp.GetBstr();
		sSQL += scSQ;
		sSQL += scSpaceCommaSpace;
		
		sTemp.Format(TEXT("%ld"), iIncrement);
		sSQL += sTemp.GetBstr();
		sSQL += scSpaceCommaSpace;
		
		sTemp.Format(TEXT("%ld"), lMaximum);
		sSQL += sTemp.GetBstr();
		sSQL += scSpaceCommaSpace;
		
		sTemp.Format(TEXT("%ld"), lMinimum);
		sSQL += sTemp.GetBstr();
		sSQL += scSpaceCommaSpace;
		
		sTemp.Format(TEXT("%d"), iMultiples);
		sSQL += sTemp.GetBstr();
		
		if (ExecBooleanSQL(sSQL, pbSuccess) != NOERROR)
		{
			sErr = CStrFnc::LoadResString(scUpdBnf);
			FlagExceptionMsg(IID_IAdmin, scUpdBnf, sErr);
			return (NOERROR);
		}
	}
	else
	{
		hResult = ::CoCreateInstance(CLSID_CList, NULL, THISCLSCTX, IID_IList, (void **)&ParameterList);
		
		::VariantInit(&vResult);
		vResult.vt = VT_I4;
		V_I4(&vResult) = lPKId;
		hResult = ParameterList->Add(TEXT("@PKId"),&vResult);
		
		::VariantInit(&vResult);
		vResult.vt = VT_I4;
		V_I4(&vResult) = lBnfTaxId;
		hResult = ParameterList->Add(TEXT("@BnfTaxId"),&vResult);
		
		::VariantInit(&vResult);
		vResult.vt = VT_I4;
		V_I4(&vResult) = lBnfTypeId;
		hResult = ParameterList->Add(TEXT("@BnfTypeId"),&vResult);
		
		::VariantInit(&vResult);
		vResult.vt = VT_I4;
		V_I4(&vResult) = lCost;
		hResult = ParameterList->Add(TEXT("@Cost"),&vResult);
		
		NULLBSTR(bstrGet);
		hResult = poBenefit->get_Description(&bstrGet);
		::VariantInit(&vResult);
		vResult.vt = VT_BSTR;
		V_BSTR(&vResult) = bstrGet;
		hResult = ParameterList->Add(TEXT("@Description"),&vResult);
		
		::VariantInit(&vResult);
		vResult.vt = VT_I2;
		V_I2(&vResult) = iIncrement;
		hResult = ParameterList->Add(TEXT("@Increment"),&vResult);
		
		::VariantInit(&vResult);
		vResult.vt = VT_I4;
		V_I4(&vResult) = lMaximum;
		hResult = ParameterList->Add(TEXT("@Maximum"),&vResult);
		
		::VariantInit(&vResult);
		vResult.vt = VT_I4;
		V_I4(&vResult) = lMinimum;
		hResult = ParameterList->Add(TEXT("@Minimum"),&vResult);
		
		::VariantInit(&vResult);
		vResult.vt = VT_I2;
		V_I2(&vResult) = iMultiples;
		hResult = ParameterList->Add(TEXT("@Multiples"),&vResult);
		
		sSQL = CStrFnc::LoadResString(scUpd_Benefit);
		hResult = ParameterList->QueryInterface(IID_IDispatch, (void **)&pIDispatch);
		if (ExecBooleanMDB(sSQL, pIDispatch, pbSuccess) != NOERROR)
		{
			pIDispatch->Release();
			ParameterList->Release();
			sErr = CStrFnc::LoadResString(scUpdBnf);
			FlagExceptionMsg(IID_IAdmin, scUpdBnf, sErr);
			return (NOERROR);
		}
		pIDispatch->Release();
		ParameterList->Release();
	}
	
	poBenefit->Term();
	poBenefit->Release();
    *pbSuccess = TRUE;
	
	return (NOERROR);
}
	
//Update_BenefitLink
//Purpose	- Updates an object
//Inputs	- (poBenefitLink)	- object with update information
//			- (pbSuccess)		- boolean output target
//Outputs	- (pbSuccess)		- boolean success value
HRESULT CAdmin::Update_BenefitLink(IBenefitLink * poBenefitLink, BOOL * pbSuccess)
{
	HRESULT hResult;
	_bstr_t sSQL;
	_bstr_t sErr;
	_bstr_t sTemp;
	VARIANT vResult;
	long	lCoverageId;
	long	lTaxId;
	long	lPKId;
	long	lEmployeePKId;
	long	lBenefitPKId;
	CURRENCY cyAmount;
	IBenefit *	pBenefit;
	IEmployee * pEmployee;
	IDispatch * pIDispatch;
	IList *		ParameterList;
	
	
	hResult = NOERROR;
	*pbSuccess = FALSE;
	
	hResult = poBenefitLink->get_PKId(&lPKId);
	if (lPKId == 0)
	{
		sErr = CStrFnc::LoadResString(scUpdBnfLnk);
		sErr += TEXT("\n\r");
		sErr += CStrFnc::LoadResString(icErrReqFld);
		sErr += TEXT("PKId");
		return (FlagExceptionMsg(IID_IAdmin, scUpdBnfLnk, sErr));
	}
	
	hResult = poBenefitLink->get_Employee(&pEmployee);
	if (hResult != NOERROR)
	{
		sErr = CStrFnc::LoadResString(scUpdBnfLnk);
		sErr += TEXT("\n\r");
		sErr += CStrFnc::LoadResString(icErrReqFld);
		sErr += TEXT("Employee.PKId");
		return (FlagExceptionMsg(IID_IAdmin, scUpdBnfLnk, sErr));
	}
	else
	{
		if (pEmployee)
		{
			hResult = pEmployee->get_PKId(&lEmployeePKId);
			hResult = pEmployee->Release();
			if (lEmployeePKId == 0)
			{
				sErr = CStrFnc::LoadResString(scUpdBnfLnk);
				sErr += TEXT("\n\r");
				sErr += CStrFnc::LoadResString(icErrReqFld);
				sErr += TEXT("Employee.PKId");
				return (FlagExceptionMsg(IID_IAdmin, scUpdBnfLnk, sErr));
			}
		}
		else
		{
			sErr = CStrFnc::LoadResString(scUpdBnfLnk);
			sErr += TEXT("\n\r");
			sErr += CStrFnc::LoadResString(icErrReqFld);
			sErr += TEXT("Employee.PKId");
			return (FlagExceptionMsg(IID_IAdmin, scUpdBnfLnk, sErr));
		}
	}
	
	hResult = poBenefitLink->get_Benefit(&pBenefit);
	if (hResult != NOERROR)
	{
		sErr = CStrFnc::LoadResString(scUpdBnfLnk);
		sErr += TEXT("\n\r");
		sErr += CStrFnc::LoadResString(icErrReqFld);
		sErr += TEXT("Benefit.PKId");
		return (FlagExceptionMsg(IID_IAdmin, scUpdBnfLnk, sErr));
	}
	else
	{
		if (pBenefit)
		{
			hResult = pBenefit->get_PKId(&lBenefitPKId);
			hResult = pBenefit->Release();
			if (lPKId == 0)
			{
				sErr = CStrFnc::LoadResString(scUpdBnfLnk);
				sErr += TEXT("\n\r");
				sErr += CStrFnc::LoadResString(icErrReqFld);
				sErr += TEXT("Benefit.PKId");
				return (FlagExceptionMsg(IID_IAdmin, scUpdBnfLnk, sErr));
			}
		}
		else
		{
			sErr = CStrFnc::LoadResString(scUpdBnfLnk);
			sErr += TEXT("\n\r");
			sErr += CStrFnc::LoadResString(icErrReqFld);
			sErr += TEXT("Employee.PKId");
			return (FlagExceptionMsg(IID_IAdmin, scUpdBnfLnk, sErr));
		}
	}

	hResult = poBenefitLink->get_CoverageId(&lCoverageId);
	if (lCoverageId == 0)
	{
		sErr = CStrFnc::LoadResString(scUpdBnfLnk);
		sErr += TEXT("\n\r");
		sErr += CStrFnc::LoadResString(icErrReqFld);
		sErr += TEXT("Benefit.CoverageId");
		return (FlagExceptionMsg(IID_IAdmin, scUpdBnfLnk, sErr));
	}
	
	hResult = poBenefitLink->get_TaxId(&lTaxId);
	if (lTaxId == 0)
	{
		sErr = CStrFnc::LoadResString(scUpdBnfLnk);
		sErr += TEXT("\n\r");
		sErr += CStrFnc::LoadResString(icErrReqFld);
		sErr += TEXT("Benefit.TaxId");
		return (FlagExceptionMsg(IID_IAdmin, scUpdBnfLnk, sErr));
	}
	if (m_bUseODBC)
	{
		if (ExecBooleanSQL(sSQL, pbSuccess) != NOERROR)
		{
			sErr = CStrFnc::LoadResString(scUpdBnf);
			FlagExceptionMsg(IID_IAdmin, scUpdBnfLnk, sErr);
			return (NOERROR);
		}
	}
	else
	{
		hResult = ::CoCreateInstance(CLSID_CList, NULL, THISCLSCTX, IID_IList, (void **)&ParameterList);
		
		::VariantInit(&vResult);
		vResult.vt = VT_I4;
		V_I4(&vResult) = lPKId;
		hResult = ParameterList->Add(TEXT("@PKId"),&vResult);
		
		::VariantInit(&vResult);
		vResult.vt = VT_I4;
		V_I4(&vResult) = lBenefitPKId;
		hResult = ParameterList->Add(TEXT("@BenefitId"),&vResult);
		
		::VariantInit(&vResult);
		vResult.vt = VT_I4;
		V_I4(&vResult) = lCoverageId;
		hResult = ParameterList->Add(TEXT("@CoverageId"),&vResult);
		
		::VariantInit(&vResult);
		vResult.vt = VT_I4;
		V_I4(&vResult) = lEmployeePKId;
		hResult = ParameterList->Add(TEXT("@EmployeeId"),&vResult);
		
		::VariantInit(&vResult);
		vResult.vt = VT_I4;
		V_I4(&vResult) = lTaxId;
		hResult = ParameterList->Add(TEXT("@TaxId"),&vResult);
		
		hResult = poBenefitLink->get_Amount(&cyAmount);
		::VariantInit(&vResult);
		vResult.vt = VT_CY;
		V_CY(&vResult) = cyAmount;
		hResult = ParameterList->Add(TEXT("@Amount"),&vResult);
		
		sSQL = CStrFnc::LoadResString(scUpd_EmplBnf);
		hResult = ParameterList->QueryInterface(IID_IDispatch, (void **)&pIDispatch);
		if (ExecBooleanMDB(sSQL, pIDispatch, pbSuccess) != NOERROR)
		{
			pIDispatch->Release();
			ParameterList->Release();
			sErr = CStrFnc::LoadResString(scUpdBnfLnk);
			FlagExceptionMsg(IID_IAdmin, scUpdBnfLnk, sErr);
			return (NOERROR);
		}
		pIDispatch->Release();
		ParameterList->Release();
	}
	
	poBenefitLink->Term();
	poBenefitLink->Release();
	*pbSuccess = TRUE;
	
	return (NOERROR);
}

//Update_Dependant
//Purpose	- Updates a specified dependant with given data
//Inputs	- (poDependant)			- dependant object containing information to be updated
//			- (pbSuccess)		- boolean output target
//Outputs	- (pbSuccess)		- boolean success value
HRESULT CAdmin::Update_Dependant(IDependant * poDependant, BOOL * pbSuccess)
{
	ASSERT(poDependant != NULL);
	ASSERT(pbSuccess != NULL);
	
	HRESULT hResult;
	_bstr_t sSQL;
	_bstr_t sErr;
	CStrFnc sTemp;
	VARIANT vResult;
	BSTR	bstrGet;
	long	lDepTypeId;
	BOOL	bGender;
	BOOL	bStatus;
	_bstr_t bstrTemp;
	DATE	TempDate;
	long	lDependantPKId;
	long	lEmployeePKId;
	IEmployee * pEmployee;
	IDispatch * pIDispatch;
	IList *		ParameterList;
	
	
	bstrGet = NULL;
	hResult		= NOERROR;
	*pbSuccess	= FALSE;
	
	hResult = poDependant->get_PKId(&lDependantPKId);
	if (lDependantPKId == 0)
	{
		sErr = CStrFnc::LoadResString(scUpdDep);
		sErr += TEXT("\n\r");
		sErr += CStrFnc::LoadResString(icErrReqFld);
		sErr += TEXT("PKId");
		return (FlagExceptionMsg(IID_IAdmin, scUpdDep, sErr));
	}
	
	hResult = poDependant->get_DepTypeId(&lDepTypeId);
	if (lDepTypeId == 0)
	{
		sErr = CStrFnc::LoadResString(scUpdDep);
		sErr += TEXT("\n\r");
		sErr += CStrFnc::LoadResString(icErrReqFld);
		sErr += TEXT("DepTypeId");
		return (FlagExceptionMsg(IID_IAdmin, scUpdDep, sErr));
	}
	
	hResult = poDependant->get_Employee(&pEmployee);
	if (hResult != NOERROR)
	{
		sErr = CStrFnc::LoadResString(scUpdDep);
		sErr += TEXT("\n\r");
		sErr += CStrFnc::LoadResString(icErrReqFld);
		sErr += TEXT("Employee.PKId");
		return (FlagExceptionMsg(IID_IAdmin, scUpdDep, sErr));
	}
	else
	{
		if (pEmployee)
		{
			hResult = pEmployee->get_PKId(&lEmployeePKId);
			hResult = pEmployee->Release();
			if (lEmployeePKId == 0)
			{
				sErr = CStrFnc::LoadResString(scUpdDep);
				sErr += TEXT("\n\r");
				sErr += CStrFnc::LoadResString(icErrReqFld);
				sErr += TEXT("Employee.PKId");
				return (FlagExceptionMsg(IID_IAdmin, scUpdDep, sErr));
			}
		}
		else
		{
			sErr = CStrFnc::LoadResString(scUpdBnfLnk);
			sErr += TEXT("\n\r");
			sErr += CStrFnc::LoadResString(icErrReqFld);
			sErr += TEXT("Employee.PKId");
			return (FlagExceptionMsg(IID_IAdmin, scUpdBnfLnk, sErr));
		}
	}
	
	NULLBSTR(bstrGet);
	hResult = poDependant->get_FirstName(&bstrGet);
	if (CStrFnc::IsEmpty(bstrGet))
	{
		sErr = CStrFnc::LoadResString(scUpdDep);
		sErr += TEXT("\n\r");
		sErr += CStrFnc::LoadResString(icErrReqFld);
		sErr += TEXT("FirstName");
		return (FlagExceptionMsg(IID_IAdmin, scUpdDep, sErr));
	}
	
	NULLBSTR(bstrGet);
	hResult = poDependant->get_LastName(&bstrGet);
	if (CStrFnc::IsEmpty(bstrGet))
	{
		sErr = CStrFnc::LoadResString(scUpdDep);
		sErr += TEXT("\n\r");
		sErr += CStrFnc::LoadResString(icErrReqFld);
		sErr += TEXT("LastName");
		return (FlagExceptionMsg(IID_IAdmin, scUpdDep, sErr));
	}

	hResult = poDependant->get_Gender(&bGender);
	if (hResult != NOERROR)
	{
		sErr = CStrFnc::LoadResString(scUpdDep);
		sErr += TEXT("\n\r");
		sErr += CStrFnc::LoadResString(icErrReqFld);
		sErr += TEXT("Gender");
		return (FlagExceptionMsg(IID_IAdmin, scUpdDep, sErr));
	}

	hResult = poDependant->get_Status(&bStatus);
	if (hResult != NOERROR)
	{
		sErr = CStrFnc::LoadResString(scUpdDep);
		sErr += TEXT("\n\r");
		sErr += CStrFnc::LoadResString(icErrReqFld);
		sErr += TEXT("Status");
		return (FlagExceptionMsg(IID_IAdmin, scUpdDep, sErr));
	}

	if (m_bUseODBC)
	{
		sSQL = CStrFnc::LoadResString(scExecute);
		sSQL += CStrFnc::LoadResString(scUpd_Dependant);
		sSQL += scSpace;
		sTemp.Format(TEXT("%ld"), lDependantPKId);
		sSQL += sTemp.GetBstr();
		sSQL += scSpaceCommaSpace;
		sTemp.Format(TEXT("%ld"), lDepTypeId);
		sSQL += sTemp.GetBstr();
		sSQL += scSpaceCommaSpace;
		sTemp.Format(TEXT("%ld"), lEmployeePKId);
		sSQL += sTemp.GetBstr();
		sSQL += scSpaceCommaSpace;
		
		hResult = poDependant->get_DateOfBirth(&TempDate);
		if (CStrFnc::IsValidDate(TempDate))
		{
			sSQL += scSQ;
			sSQL += CStrFnc::FormatDate(TempDate, scDateFormat);
			sSQL += scSQ;
			sSQL += scSpaceCommaSpace;
		}
		else
		{
			sSQL += scNull;
			sSQL += scSpaceCommaSpace;
		}
		
		NULLBSTR(bstrGet);
		hResult = poDependant->get_FirstName(&bstrGet);
		sSQL += scSQ;
		sSQL += bstrGet;
		sSQL += scSQ;
		sSQL += scSpaceCommaSpace;
		hResult = poDependant->get_Gender(&bGender);
		if (bGender)
		{
			sSQL += scZeroStr;
			sSQL += scSpaceCommaSpace;
		}
		else
		{
			sSQL += scOneStr;
			sSQL += scSpaceCommaSpace;
		}
		
		NULLBSTR(bstrGet);
		hResult = poDependant->get_LastName(&bstrGet);
		sSQL += scSQ;
		sSQL += bstrGet;
		sSQL += scSQ;
		sSQL += scSpaceCommaSpace;
		hResult = poDependant->get_Status(&bStatus);
		if (bStatus)
		{
			sSQL += scZeroStr;
			sSQL += scSpaceCommaSpace;
		}
		else
		{
			sSQL += scOneStr;
			sSQL += scSpaceCommaSpace;
		}
		if (ExecBooleanSQL(sSQL, pbSuccess) != NOERROR)
		{
			sErr = CStrFnc::LoadResString(scUpdDep);
			FlagExceptionMsg(IID_IAdmin, scUpdDep, sErr);
			return (NOERROR);
		}
	}
	else
	{
		hResult = ::CoCreateInstance(CLSID_CList, NULL, THISCLSCTX, IID_IList, (void **)&ParameterList);
		
		::VariantInit(&vResult);
		vResult.vt = VT_I4;
		V_I4(&vResult) = lDependantPKId;
		hResult = ParameterList->Add(TEXT("@PKId"),&vResult);
		
		::VariantInit(&vResult);
		vResult.vt = VT_I4;
		V_I4(&vResult) = lDepTypeId;
		hResult = ParameterList->Add(TEXT("@DepTypeId"),&vResult);
		
		::VariantInit(&vResult);
		vResult.vt = VT_I4;
		V_I4(&vResult) = lEmployeePKId;
		hResult = ParameterList->Add(TEXT("@EmployeeId"),&vResult);
		
		hResult = poDependant->get_DateOfBirth(&TempDate);
		::VariantInit(&vResult);
		vResult.vt = VT_DATE;
		V_DATE(&vResult) = TempDate;
		hResult = ParameterList->Add(TEXT("@DateOfBirth"),&vResult);
		
		NULLBSTR(bstrGet);
		hResult = poDependant->get_FirstName(&bstrGet);
		::VariantInit(&vResult);
		vResult.vt = VT_BSTR;
		V_BSTR(&vResult) = bstrGet;
		hResult = ParameterList->Add(TEXT("@FirstName"),&vResult);
		
		::VariantInit(&vResult);
		vResult.vt = VT_I2;
		V_I4(&vResult) = bGender;
		hResult = ParameterList->Add(TEXT("@Gender"),&vResult);
		
		NULLBSTR(bstrGet);
		hResult = poDependant->get_LastName(&bstrGet);
		::VariantInit(&vResult);
		vResult.vt = VT_BSTR;
		V_BSTR(&vResult) = bstrGet;
		hResult = ParameterList->Add(TEXT("@LastName"),&vResult);
		
		::VariantInit(&vResult);
		vResult.vt = VT_I2;
		V_I4(&vResult) = bStatus;
		hResult = ParameterList->Add(TEXT("@Status"),&vResult);
		
		sSQL = CStrFnc::LoadResString(scUpd_Dependant);
		hResult = ParameterList->QueryInterface(IID_IDispatch, (void **)&pIDispatch);
		if (ExecBooleanMDB(sSQL, pIDispatch, pbSuccess) != NOERROR)
		{
			pIDispatch->Release();
			ParameterList->Release();
			sErr = CStrFnc::LoadResString(scUpdDep);
			FlagExceptionMsg(IID_IAdmin, scUpdDep, sErr);
			return (NOERROR);
		}
		pIDispatch->Release();
		ParameterList->Release();
	}
	
	poDependant->Term();
	poDependant->Release();
	*pbSuccess = TRUE;
	return (NOERROR);
}

//Update_Employee
//Purpose	- Updates a specified employee with given data
//Inputs	- (poEmployee)		- employee object containing information to be updated
//			- (pbSuccess)		- boolean output target
//Outputs	- (pbSuccess)		- boolean success value
HRESULT CAdmin::Update_Employee(IEmployee * poEmployee, BOOL * pbSuccess)
{
	HRESULT hResult;
	_bstr_t sSQL;
	_bstr_t sErr;
	CStrFnc sTemp;
	VARIANT vResult;
	BSTR	bstrGet;
	long	lEmployeePKId;
	long	lBankPKId;
	_bstr_t bstrTemp;
	BOOL	bGender;
	DATE	TempDate;
	CURRENCY Salary;
	IBank *		oBank;
	IDispatch *	pIDispatch;
	IList *		ParameterList;
	
	hResult = NOERROR;
	bstrGet = NULL;
	*pbSuccess = FALSE;
	
	hResult = poEmployee->get_PKId(&lEmployeePKId);
	if (lEmployeePKId == 0)
	{
		sErr = CStrFnc::LoadResString(scUpdEmpl);
		sErr += TEXT("\n\r");
		sErr += CStrFnc::LoadResString(icErrReqFld);
		sErr += TEXT("PKId");
		return (FlagExceptionMsg(IID_IAdmin, scUpdEmpl, sErr));
	}
	
	hResult = poEmployee->get_Bank(&oBank);
	if (hResult != NOERROR)
	{
		sErr = CStrFnc::LoadResString(scUpdEmpl);
		sErr += TEXT("\n\r");
		sErr += CStrFnc::LoadResString(icErrReqFld);
		sErr += TEXT("Bank.PKId");
		return (FlagExceptionMsg(IID_IAdmin, scUpdEmpl, sErr));
	}
	else
	{
		if (oBank)
		{
			hResult = oBank->get_PKId(&lBankPKId);
			hResult = oBank->Release();
			if (lBankPKId == 0)
			{
				sErr = CStrFnc::LoadResString(scUpdEmpl);
				sErr += TEXT("\n\r");
				sErr += CStrFnc::LoadResString(icErrReqFld);
				sErr += TEXT("Bank.PKId");
				return (FlagExceptionMsg(IID_IAdmin, scUpdEmpl, sErr));
			}
		}
		else
		{
			sErr = CStrFnc::LoadResString(scUpdEmpl);
			sErr += TEXT("\n\r");
			sErr += CStrFnc::LoadResString(icErrReqFld);
			sErr += TEXT("Bank.PKId");
			return (FlagExceptionMsg(IID_IAdmin, scUpdEmpl, sErr));
		}
	}
	
	NULLBSTR(bstrGet);
	poEmployee->get_FirstName(&bstrGet);
	if (CStrFnc::IsEmpty(bstrGet))
	{
		sErr = CStrFnc::LoadResString(scUpdEmpl);
		sErr += TEXT("\n\r");
		sErr += CStrFnc::LoadResString(icErrReqFld);
		sErr += TEXT("FirstName");
		return (FlagExceptionMsg(IID_IAdmin, scUpdEmpl, sErr));
	}
	
	NULLBSTR(bstrGet);
	poEmployee->get_LastName(&bstrGet);
	if (CStrFnc::IsEmpty(bstrGet))
	{
		sErr = CStrFnc::LoadResString(scUpdEmpl);
		sErr += TEXT("\n\r");
		sErr += CStrFnc::LoadResString(icErrReqFld);
		sErr += TEXT("LastName");
		return (FlagExceptionMsg(IID_IAdmin, scUpdEmpl, sErr));
	}
	
	NULLBSTR(bstrGet);
	poEmployee->get_Position(&bstrGet);
	if (CStrFnc::IsEmpty(bstrGet))
	{
		sErr = CStrFnc::LoadResString(scUpdEmpl);
		sErr += TEXT("\n\r");
		sErr += CStrFnc::LoadResString(icErrReqFld);
		sErr += TEXT("Position");
		return (FlagExceptionMsg(IID_IAdmin, scUpdEmpl, sErr));
	}
	
	NULLBSTR(bstrGet);
	poEmployee->get_StartDate(&TempDate);
	if (!CStrFnc::IsValidDate(TempDate))
	{
		sErr = CStrFnc::LoadResString(scUpdEmpl);
		sErr += TEXT("\n\r");
		sErr += CStrFnc::LoadResString(icErrReqFld);
		sErr += TEXT("StartDate");
		return (FlagExceptionMsg(IID_IAdmin, scUpdEmpl, sErr));
	}
	
	NULLBSTR(bstrGet);
	poEmployee->get_UserName(&bstrGet);
	if (::SysStringLen(bstrGet) <= 0)
	{
		sErr = CStrFnc::LoadResString(scUpdEmpl);
		sErr += TEXT("\n\r");
		sErr += CStrFnc::LoadResString(icErrReqFld);
		sErr += TEXT("UserName");
		return (FlagExceptionMsg(IID_IAdmin, scUpdEmpl, sErr));
	}
	
	hResult = poEmployee->get_Gender(&bGender);
	if (hResult != NOERROR)
	{
		sErr = CStrFnc::LoadResString(scUpdEmpl);
		sErr += TEXT("\n\r");
		sErr += CStrFnc::LoadResString(icErrReqFld);
		sErr += TEXT("Gender");
		return (FlagExceptionMsg(IID_IAdmin, scUpdEmpl, sErr));
	}

	NULLBSTR(bstrGet);
	
	if (m_bUseODBC)
	{
		sSQL = CStrFnc::LoadResString(scExecute);
		sSQL += CStrFnc::LoadResString(scUpd_Employee);
		sSQL += scSpace;
		sTemp.Format(TEXT("%ld"), lEmployeePKId);
		sSQL += sTemp.GetBstr();
		sSQL += scSpaceCommaSpace;
		sSQL += scSpace;
		sTemp.Format(TEXT("%ld"), lBankPKId);
		sSQL += sTemp.GetBstr();
		sSQL += scSpaceCommaSpace;
		
		NULLBSTR(bstrGet);
		hResult = poEmployee->get_Address1(&bstrGet);
		sTemp = bstrGet;
		sTemp.Trim(sTemp.Left(60));
		sSQL += scSQ;
		sSQL += sTemp.GetBstr();
		sSQL += scSQ;
		sSQL += scSpaceCommaSpace;
		
		NULLBSTR(bstrGet);
		hResult = poEmployee->get_Address2(&bstrGet);
		sTemp = bstrGet;
		sTemp.Trim(sTemp.Left(60));
		sSQL += scSQ;
		sSQL += sTemp.GetBstr();
		sSQL += scSQ;
		sSQL += scSpaceCommaSpace;
		
		NULLBSTR(bstrGet);
		hResult = poEmployee->get_BankAccount(&bstrGet);
		sTemp = bstrGet;
		sTemp.Trim(sTemp.Left(60));
		sSQL += scSQ;
		sSQL += sTemp.GetBstr();
		sSQL += scSQ;
		sSQL += scSpaceCommaSpace;
		
		NULLBSTR(bstrGet);
		hResult = poEmployee->get_City(&bstrGet);
		sTemp = bstrGet;
		sTemp.Trim(sTemp.Left(60));
		sSQL += scSQ;
		sSQL += sTemp.GetBstr();
		sSQL += scSQ;
		sSQL += scSpaceCommaSpace;
		
		hResult = poEmployee->get_DateOfBirth(&TempDate);
		if (CStrFnc::IsValidDate(TempDate))
		{
			sSQL += scSQ;
			sSQL += CStrFnc::FormatDate(TempDate, scDateFormat);
			sSQL += scSQ;
			sSQL += scSpaceCommaSpace;
		}
		else
		{
			sSQL += scNull;
			sSQL += scSpaceCommaSpace;
		}
		
		hResult = poEmployee->get_EndDate(&TempDate);
		if (CStrFnc::IsValidDate(TempDate))
		{
			sSQL += scSQ;
			sSQL += CStrFnc::FormatDate(TempDate, scDateFormat);
			sSQL += scSQ;
			sSQL += scSpaceCommaSpace;
		}
		else
		{
			sSQL += scNull;
			sSQL += scSpaceCommaSpace;
		}
		
		NULLBSTR(bstrGet);
		hResult = poEmployee->get_FirstName(&bstrGet);
		sTemp = bstrGet;
		sTemp.Trim(sTemp.Left(35));
		sSQL += scSQ;
		sSQL += sTemp.GetBstr();
		sSQL += scSQ;
		sSQL += scSpaceCommaSpace;
		
		hResult = poEmployee->get_Gender(&bGender);
		if (bGender)
		{
			sSQL += scZeroStr;
			sSQL += scSpaceCommaSpace;
		}
		else
		{
			sSQL += scOneStr;
			sSQL += scSpaceCommaSpace;
		}
		
		NULLBSTR(bstrGet);
		hResult = poEmployee->get_LastName(&bstrGet);
		sTemp = bstrGet;
		sTemp.Trim(sTemp.Left(35));
		sSQL += scSQ;
		sSQL += sTemp.GetBstr();
		sSQL += scSQ;
		sSQL += scSpaceCommaSpace;
		
		NULLBSTR(bstrGet);
		hResult = poEmployee->get_Position(&bstrGet);
		sTemp = bstrGet;
		sTemp.Trim(sTemp.Left(35));
		sSQL += scSQ;
		sSQL += sTemp.GetBstr();
		sSQL += scSQ;
		sSQL += scSpaceCommaSpace;
		
		hResult = poEmployee->get_StartDate(&TempDate);
		sSQL += CStrFnc::FormatDate(TempDate, scDateFormat);
		sSQL += scSpaceCommaSpace;
		
		hResult = poEmployee->get_Salary(&Salary);
		::VariantInit(&vResult);
		vResult.vt = VT_CY;
		V_CY(&vResult) = Salary;
		
		{
			VARIANT v;
			::VariantInit(&v);
			v.vt = VT_BSTR;
			::VariantCopy(&v, &vResult);
			sSQL += V_BSTR(&v);
		}
		sSQL += scSpaceCommaSpace;
		
		NULLBSTR(bstrGet);
		hResult = poEmployee->get_UserName(&bstrGet);
		sTemp = bstrGet;
		sTemp.Trim(sTemp.Left(35));
		sSQL += scSQ;
		sSQL += sTemp.GetBstr();
		sSQL += scSQ;
		
		if (ExecBooleanSQL(sSQL, pbSuccess) != NOERROR)
		{
			sErr = CStrFnc::LoadResString(scUpdEmpl);
			FlagExceptionMsg(IID_IAdmin, scUpdEmpl, sErr);
			return (NOERROR);
		}
	}
	else
	{
		hResult = ::CoCreateInstance(CLSID_CList, NULL, THISCLSCTX, IID_IList, (void **)&ParameterList);
		
		::VariantInit(&vResult);
		vResult.vt = VT_I4;
		V_I4(&vResult) = lEmployeePKId;
		hResult = ParameterList->Add(TEXT("@PKId"),&vResult);
		
		::VariantInit(&vResult);
		vResult.vt = VT_I4;
		V_I4(&vResult) = lBankPKId;
		hResult = ParameterList->Add(TEXT("@BankID"),&vResult);
		
		NULLBSTR(bstrGet);
		hResult = poEmployee->get_Address1(&bstrGet);
		::VariantInit(&vResult);
		vResult.vt = VT_BSTR;
		V_BSTR(&vResult) = bstrGet;
		hResult = ParameterList->Add(TEXT("@Address1"),&vResult);
		
		NULLBSTR(bstrGet);
		hResult = poEmployee->get_Address2(&bstrGet);
		::VariantInit(&vResult);
		vResult.vt = VT_BSTR;
		V_BSTR(&vResult) = bstrGet;
		hResult = ParameterList->Add(TEXT("@Address2"),&vResult);
		
		NULLBSTR(bstrGet);
		hResult = poEmployee->get_BankAccount(&bstrGet);
		::VariantInit(&vResult);
		vResult.vt = VT_BSTR;
		V_BSTR(&vResult) = bstrGet;
		hResult = ParameterList->Add(TEXT("@BankAccount"),&vResult);
		
		NULLBSTR(bstrGet);
		hResult = poEmployee->get_City(&bstrGet);
		::VariantInit(&vResult);
		vResult.vt = VT_BSTR;
		V_BSTR(&vResult) = bstrGet;
		hResult = ParameterList->Add(TEXT("@City"),&vResult);
		
		hResult = poEmployee->get_DateOfBirth(&TempDate);
		::VariantInit(&vResult);
		vResult.vt = VT_DATE;
		V_DATE(&vResult) = TempDate;
		hResult = ParameterList->Add(TEXT("@DateOfBirth"),&vResult);
		
		hResult = poEmployee->get_EndDate(&TempDate);
		::VariantInit(&vResult);
		vResult.vt = VT_DATE;
		V_DATE(&vResult) = TempDate;
		hResult = ParameterList->Add(TEXT("@EndDate"),&vResult);
		
		NULLBSTR(bstrGet);
		hResult = poEmployee->get_FirstName(&bstrGet);
		::VariantInit(&vResult);
		vResult.vt = VT_BSTR;
		V_BSTR(&vResult) = bstrGet;
		hResult = ParameterList->Add(TEXT("@FirstName"),&vResult);
		
		hResult = poEmployee->get_Gender(&bGender);
		::VariantInit(&vResult);
		vResult.vt = VT_BOOL;
		V_BOOL(&vResult) = bGender;
		hResult = ParameterList->Add(TEXT("@Gender"),&vResult);
		
		NULLBSTR(bstrGet);
		hResult = poEmployee->get_LastName(&bstrGet);
		::VariantInit(&vResult);
		vResult.vt = VT_BSTR;
		V_BSTR(&vResult) = bstrGet;
		hResult = ParameterList->Add(TEXT("@LastName"),&vResult);
		
		NULLBSTR(bstrGet);
		hResult = poEmployee->get_Position(&bstrGet);
		::VariantInit(&vResult);
		vResult.vt = VT_BSTR;
		V_BSTR(&vResult) = bstrGet;
		hResult = ParameterList->Add(TEXT("@Position"),&vResult);
		
		hResult = poEmployee->get_Salary(&Salary);
		::VariantInit(&vResult);
		vResult.vt = VT_CY;
		V_CY(&vResult) = Salary;
		hResult = ParameterList->Add(TEXT("@Salary"),&vResult);
		
		hResult = poEmployee->get_StartDate(&TempDate);
		::VariantInit(&vResult);
		vResult.vt = VT_DATE;
		V_DATE(&vResult) = TempDate;
		hResult = ParameterList->Add(TEXT("@StartDate"),&vResult);
		
		NULLBSTR(bstrGet);
		hResult = poEmployee->get_UserName(&bstrGet);
		::VariantInit(&vResult);
		vResult.vt = VT_BSTR;
		V_BSTR(&vResult) = bstrGet;
		hResult = ParameterList->Add(TEXT("@UserName"),&vResult);
		
		sSQL = CStrFnc::LoadResString(scUpd_Employee);
		hResult = ParameterList->QueryInterface(IID_IDispatch, (void **)&pIDispatch);
		if (ExecBooleanMDB(sSQL, pIDispatch, pbSuccess) != NOERROR)
		{
			pIDispatch->Release();
			ParameterList->Release();
			sErr = CStrFnc::LoadResString(scUpdEmpl);
			FlagExceptionMsg(IID_IAdmin, scUpdEmpl, sErr);
			return (NOERROR);
		}
		pIDispatch->Release();
		ParameterList->Release();
	}
	
	poEmployee->Term();
	poEmployee->Release();
	*pbSuccess = TRUE;
	return (NOERROR);
}

//VariantXToBstr
//Purpose	- To convert any variant type to a bstr type
//Inputs	- (pvX)		- pointer to variant to be coerced
//Outputs	- (BSTR)	- newly converted BSTR value
BSTR CAdmin::VariantXToBstr(VARIANT * pvX)
{
	BSTR	bstrY;
	BSTR	bstrResult;
	HRESULT hResult;
	CStrFnc	sFormat;
	
	hResult = NOERROR;

	switch (pvX->vt)
	{
	case VT_BOOL:
		hResult = VarBstrFromBool(pvX->boolVal, LOCALE_NOUSEROVERRIDE, 0, &bstrY);
		break;
	case VT_I1:
		hResult = VarBstrFromI1(pvX->cVal, LOCALE_NOUSEROVERRIDE, 0, &bstrY);
		break;
	case VT_UI1:
		hResult = VarBstrFromUI1(pvX->bVal, LOCALE_NOUSEROVERRIDE, 0, &bstrY);
		break;
	case VT_I2:
		hResult = VarBstrFromI2(pvX->iVal, LOCALE_NOUSEROVERRIDE, 0, &bstrY);
		break;
	case VT_UI2:
		hResult = VarBstrFromUI2(pvX->uiVal, LOCALE_NOUSEROVERRIDE, 0, &bstrY);
		break;
	case VT_I4:
		hResult = VarBstrFromI4(pvX->lVal, LOCALE_NOUSEROVERRIDE, 0, &bstrY);
		break;
	case VT_UI4:
		hResult = VarBstrFromUI4(pvX->ulVal, LOCALE_NOUSEROVERRIDE, 0, &bstrY);
		break;
	case VT_DECIMAL:
		hResult = VarBstrFromDec(pvX->pdecVal, LOCALE_NOUSEROVERRIDE, 0, &bstrY);
		break;
	case VT_R4:
		hResult = VarBstrFromR4(pvX->fltVal, LOCALE_NOUSEROVERRIDE, 0, &bstrY);
		break;
	case VT_R8:
		hResult = VarBstrFromR8(pvX->dblVal, LOCALE_NOUSEROVERRIDE, 0, &bstrY);
		break;
	case VT_CY:
		hResult = VarBstrFromCy(pvX->cyVal, LOCALE_NOUSEROVERRIDE, 0, &bstrY);
		break;
	case VT_DATE:
		{
		sFormat = TEXT("%m%d%Y");
		bstrY = CStrFnc::FormatDate(pvX->date, sFormat.GetBstr());
		}
		break;
	case VT_DISPATCH:
		hResult = VarBstrFromDisp(pvX->pdispVal, LOCALE_NOUSEROVERRIDE, 0, &bstrY);
		break;
	case VT_BSTR:
		bstrY = ::SysAllocString(pvX->bstrVal);
		hResult = NOERROR;
		break;
	default:
		return (NULL);
	}
	
	if (hResult == NOERROR)
		bstrResult = ::SysAllocString(bstrY);
	else
		bstrResult = NULL;
	
	return (bstrResult);
}

//VariantXToVariantBstr
//Purpose	- To convert any variant type to a variant bstr type
//Inputs	- (pvX)		- pointer to variant to be coerced
//Outputs	- (VARIANT)	- newly converted VARIANT value
VARIANT CAdmin::VariantXToVariantBstr(VARIANT * pvX)
{
	BSTR bstrY;
	VARIANT v;
	VARIANT x;
	
	bstrY = this->VariantXToBstr(pvX);
	::VariantInit(&v);
	::VariantInit(&x);
	x.vt = VT_BSTR;
	x.bstrVal = bstrY;
	::VariantCopyInd(&v, &x);
	return (v);
}

//Constructor/Destructor
//
//Turn off warning 4355
#pragma warning (disable : 4355)
CAdmin::CAdmin(	LPUNKNOWN pUnkOuter, PFNDESTROYED pfnDestroy)
: m_SupportErrorInfo(this, IID_IAdmin)
#pragma warning (default : 4355)
{
	m_oid = g_oid++;
	m_cRef = 0;
	m_pUnkOuter = pUnkOuter;
	m_pfnDestroy = pfnDestroy;
	m_pITypeInfo = NULL;
	m_Database = NULL;
	m_DbEngine = NULL;
	m_bstrDataSource = NULL;
	m_iSrcType = 0;
	m_bUseODBC = FALSE;
	m_Banks = NULL;
	m_Benefits = NULL;
	m_Employees = NULL;
	m_pIRemoteBenefitsServer = NULL;
	m_pClassFactory2 = NULL;
	return;
}

CAdmin::~CAdmin(void)
{
	Term();

	if (m_pITypeInfo)
	{
		m_pITypeInfo->Release();
		m_pITypeInfo = NULL;
	}

	if (m_pClassFactory2)
	{
		m_pClassFactory2->Release();
		m_pClassFactory2 = NULL;
	}

	return;
}

/////////////////////////////////////////////
// Standard IUnknown method implementations
/////////////////////////////////////////////

//CAdmin::QueryInterface
//
//Parameters:
//riid			-	input interface reference id
//ppvCAdminIf	-	return CAdmin interface pointer
STDMETHODIMP CAdmin::QueryInterface(REFIID riid, void ** ppvCAdminIf)
{
	*ppvCAdminIf = NULL;
	
	if (riid == IID_IAdmin || riid == IID_IUnknown || riid == IID_IDispatch)
		*ppvCAdminIf = this;
	
	if (riid == IID_ISupportErrorInfo)
		*ppvCAdminIf = &m_SupportErrorInfo;
	
	if (*ppvCAdminIf)
	{
		((LPUNKNOWN)*ppvCAdminIf)->AddRef();
		return (NOERROR);
	}
	
	return (E_NOINTERFACE);
}

STDMETHODIMP_(ULONG) CAdmin::AddRef(void)
{
	++m_cRef;
	Audit(m_oid, 1, TEXT("CAdmin"), m_cRef);
	return (m_cRef);
}

STDMETHODIMP_(ULONG) CAdmin::Release(void)
{
	--m_cRef;
	Audit(m_oid, 2, TEXT("CAdmin"), m_cRef);

	if (m_cRef == 0)
	{
		if (m_pfnDestroy)
			(*m_pfnDestroy)();
		Audit(m_oid, 0, TEXT("CAdmin gbldec"), g_lObjectCount);
		delete this;
		return (0);
	}
	else
		return (m_cRef);
}

//////////////////////////////////////////////
// Standard IDispatch method implementations
//////////////////////////////////////////////

STDMETHODIMP CAdmin::GetTypeInfoCount(UINT *pctInfo)
{
	*pctInfo = 1;

	return (NOERROR);
}

STDMETHODIMP CAdmin::GetTypeInfo(UINT itInfo, LCID lcid, ITypeInfo **ppITypeInfo)
{
	HRESULT     hResult;
	ITypeLib   *pITypeLib;
	ITypeInfo **ppITI;
	
	hResult = NOERROR;
	ppITI = NULL;

	if (itInfo != 0)
		return (TYPE_E_ELEMENTNOTFOUND);
	
	if (ppITypeInfo == NULL)
		return (E_POINTER);
	
	*ppITypeInfo = NULL;
	
	switch (PRIMARYLANGID(lcid))
	{
	case LANG_NEUTRAL:
	case LANG_ENGLISH:
		ppITI = &m_pITypeInfo;
		break;
		
	default:
		return (DISP_E_UNKNOWNLCID);
	}
	
	if (*ppITI == NULL)
	{
#ifdef _BENDLL
		hResult = LoadRegTypeLib(LIBID_BenefitsCPP, 1, 0, PRIMARYLANGID(lcid), &pITypeLib);
#else
		hResult = LoadRegTypeLib(LIBID_BenefitsCPPRemote, 1, 0, PRIMARYLANGID(lcid), &pITypeLib);
#endif
		if (hResult != NOERROR)
		{
			switch (PRIMARYLANGID(lcid))
			{
			case LANG_NEUTRAL:
			case LANG_ENGLISH:
#ifdef _BENDLL
				hResult = LoadTypeLib(TEXT("bencpp.tlb"), &pITypeLib);
#else
				hResult = LoadTypeLib(TEXT("bencppr.tlb"), &pITypeLib);
#endif
				break;
			}
		}
		
		if (hResult != NOERROR)
			return (hResult);
		
		hResult = pITypeLib->GetTypeInfoOfGuid(IID_IAdmin, ppITI);

		m_pITypeInfo = *ppITI;
		m_pITypeInfo->AddRef();

		pITypeLib->Release();
		
		if (hResult != NOERROR)
			return (hResult);
	}
	
	(*ppITI)->AddRef();
	*ppITypeInfo = *ppITI;
	
	return (NOERROR);
}

STDMETHODIMP CAdmin::GetIDsOfNames(REFIID riid, OLECHAR **rgszNames,
									UINT cNames, LCID lcid, DISPID *rgDispID)
{
	HRESULT     hResult;
	ITypeInfo  *pTI;
	
	hResult = NOERROR;

	if (riid != IID_NULL)
		return (DISP_E_UNKNOWNINTERFACE);
	
	hResult = GetTypeInfo(0, lcid, &pTI);
	
	if (hResult == NOERROR)
	{
		hResult = DispGetIDsOfNames(pTI, rgszNames, cNames, rgDispID);
		pTI->Release();
	}
	
	return (hResult);
}

STDMETHODIMP CAdmin::Invoke(DISPID dispID, REFIID riid
							, LCID lcid, unsigned short wFlags, DISPPARAMS *pDispParams
							, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
	HRESULT     hResult;
	ITypeInfo  *pTI;
	LANGID      langID;
	
	hResult = NOERROR;
	langID	= PRIMARYLANGID(lcid);

	if (riid != IID_NULL)
		return (DISP_E_UNKNOWNINTERFACE);
	
	hResult = GetTypeInfo(0, lcid, &pTI);
	
	if (hResult != NOERROR)
		return (hResult);
	
	hResult = DispInvoke(this, pTI, dispID, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr);
	
	pTI->Release();

	return (hResult);
}

CAdminFactory::CAdminFactory(void)
{
	m_oid = g_oid++;
	m_cRef = 0L;
	return;
}

CAdminFactory::~CAdminFactory(void)
{
	return;
}

//CAdminFactory::QueryInterface
//Purpose     - Used to return an interface pointer from the requested interface id.
//Inputs      - (REFIID)  riid
//              (void **) ppvCAdminCFIf
STDMETHODIMP CAdminFactory::QueryInterface(REFIID riid, void ** ppvCAdminCFIf)
{
	*ppvCAdminCFIf = NULL;
	
	if (riid == IID_IUnknown ||
		riid == IID_IClassFactory)
		*ppvCAdminCFIf = this;
	
	if (*ppvCAdminCFIf)
	{
		((LPUNKNOWN)*ppvCAdminCFIf)->AddRef();
		return (NOERROR);
	}
	
	return (E_NOINTERFACE);
}

STDMETHODIMP_(ULONG) CAdminFactory::AddRef(void)
{
	++m_cRef;
	Audit(m_oid, 1, TEXT("CAdminF"), m_cRef);
	return (m_cRef);
}

STDMETHODIMP_(ULONG) CAdminFactory::Release(void)
{
	--m_cRef;
	Audit(m_oid, 2, TEXT("CAdminF"), m_cRef);
	if (m_cRef == 0)
	{
		Audit(m_oid, 0, TEXT("CAdminF gbldec"), g_lObjectCount);
		delete this;
		return (0);
	}
	else
		return (m_cRef);
}

//CAdminFactory::CreateInstance
//Purpose     - Creates a CAdmin object and returns an Interface pointer
//Inputs      - (LPUNKNOWN) pUnkOuter
//              (REFIID)    riid
//				(void **)   ppvCAdminIf
STDMETHODIMP CAdminFactory::CreateInstance(LPUNKNOWN pUnkOuter, REFIID riid, void ** ppvCAdminIf)
{
	PCAdmin            pCAdmin;
	HRESULT            hResult;
	
	hResult			= NOERROR;
	*ppvCAdminIf	= NULL;
	
	if (pUnkOuter && riid != IID_IUnknown)
		return (CLASS_E_NOAGGREGATION);
	
	pCAdmin = new CAdmin(pUnkOuter, ObjectDestroyed);
	
	if (pCAdmin == NULL)
		return (hResult);
	
	if (pCAdmin->CFInit())
		hResult = pCAdmin->QueryInterface(riid, ppvCAdminIf);
	
	if (hResult != NOERROR)
		delete pCAdmin;
	else
	{
		g_lObjectCount++;
		Audit(m_oid, 0, TEXT("CAdmin gblinc"), g_lObjectCount);
	}
	
	return (hResult);
}

//CAdminFactory::LockServer
//Purpose     - Increments/Decrements the lock count to prevent class factory server shutdown.
//Inputs      - (BOOL)    fLock
STDMETHODIMP CAdminFactory::LockServer(BOOL fLock)
{
	if (fLock)
		g_lLockedObjects++;
	else
		g_lLockedObjects--;
	
	return (NOERROR);
}

//CAdmin::CFInit
//
BOOL CAdmin::CFInit(void)
{
	LPUNKNOWN pIUnknown;
	
	pIUnknown = this;
	if (m_pUnkOuter)
		pIUnknown = m_pUnkOuter;
	
	return (TRUE);
}

