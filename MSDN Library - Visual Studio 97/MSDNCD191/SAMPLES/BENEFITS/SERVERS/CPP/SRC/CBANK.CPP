//cbank.cpp   - Copyright (C) 1996 Microsoft Corporation
//            - All rights reserved.
//Date        - 10/19/96
//Authors     - Ken Bergmann, Robert Coleridge
//Purpose     - Implementation File for the CBank object

#pragma warning (disable : 4800)	// turns off bool warning

#include "cbank.h"
#include "cienum.h"
#include "csaferay.h"

extern long g_oid;

//Globals used to determine when to unload a In-Proc or Local Server.
extern ULONG       g_lObjectCount;
extern ULONG       g_lLockedObjects;

STDMETHODIMP CBank::putref_Admin(IAdmin * poAdmin) 
{
	if (m_poAdmin)
	{
		m_poAdmin->Release();
		m_poAdmin = NULL;
	}

	m_poAdmin = (IAdmin *)poAdmin;

	if (m_poAdmin)
		m_poAdmin->AddRef();

	return (NOERROR);
}


STDMETHODIMP CBank::get_PKId(long FAR * plPKId) 
{
    *plPKId = m_lPKId;
	return (NOERROR);
}


STDMETHODIMP CBank::put_PKId(long lPKId) 
{
	m_lPKId = lPKId;
	return (NOERROR);
}


STDMETHODIMP CBank::get_Address1(BSTR FAR * pbstrAddress1) 
{
	if (m_bExpanded)
		::SysReAllocString(pbstrAddress1,m_bstrAddress1);
	else
	{
		if (Expand())
			::SysReAllocString(pbstrAddress1,m_bstrAddress1);
		else
		{
			Term();
			return (FlagException(IID_IBank, scBankExpand, icErrExpand));
		}
	}
	return (NOERROR);
}

STDMETHODIMP CBank::put_Address1(BSTR bstrAddress1) 
{
	::SysReAllocString(&m_bstrAddress1,bstrAddress1);
	return (NOERROR);
}

STDMETHODIMP CBank::get_Address2(BSTR FAR * pbstrAddress2) 
{
	if (m_bExpanded)
		::SysReAllocString(pbstrAddress2, m_bstrAddress2);
	else
	{
		if (Expand())
			::SysReAllocString(pbstrAddress2, m_bstrAddress2);
		else
		{
			Term();
			return (FlagException(IID_IBank, scBankExpand, icErrExpand));
		}
	}
	return (NOERROR);
}


STDMETHODIMP CBank::put_Address2(BSTR bstrAddress2) 
{
	::SysReAllocString(&m_bstrAddress2, bstrAddress2);
	return (NOERROR);
}


STDMETHODIMP CBank::get_City(BSTR FAR * pbstrCity) 
{
	if (m_bExpanded)
		::SysReAllocString(pbstrCity, m_bstrCity);
	else
	{
		if (Expand())
			::SysReAllocString(pbstrCity, m_bstrCity);
		else
		{
			Term();
			return (FlagException(IID_IBank, scBankExpand, icErrExpand));
		}
	}

	return (NOERROR);
}

STDMETHODIMP CBank::put_City(BSTR bstrCity) 
{
    ::SysReAllocString(&m_bstrCity, bstrCity);
	return (NOERROR);
}


STDMETHODIMP CBank::get_Code(BSTR FAR * pbstrCode) 
{
	if (m_bExpanded)
		::SysReAllocString(pbstrCode, m_bstrCode);
	else
	{
		if (Expand())
			::SysReAllocString(pbstrCode, m_bstrCode);
		else
		{
			Term();
			return (FlagException(IID_IBank, scBankExpand, icErrExpand));
		}
	}

	return (NOERROR);
}


STDMETHODIMP CBank::put_Code(BSTR bstrCode) 
{
    ::SysReAllocString(&m_bstrCode, bstrCode);
	return (NOERROR);
}


STDMETHODIMP CBank::get_Name(BSTR FAR * pbstrName) 
{
	if (m_bExpanded)
		::SysReAllocString(pbstrName, m_bstrName);
	else
	{
		if (Expand())
			::SysReAllocString(pbstrName, m_bstrName);
		else
		{
			Term();
			return (FlagException(IID_IBank, scBankExpand, icErrExpand));
		}
	}
	
	return (NOERROR);
}


STDMETHODIMP CBank::put_Name(BSTR bstrName) 
{
	::SysReAllocString(&m_bstrName, bstrName);
	return (NOERROR);
}


STDMETHODIMP CBank::get_Employees(IListEmployee * FAR * ppcEmployees) 
{
	if (m_bExpanded)
		*ppcEmployees = m_poEmployees;
	else
	{
		if (Expand())
			*ppcEmployees = m_poEmployees;
		else
		{
			Term();
			return (FlagException(IID_IBank, scBankExpand, icErrExpand));
		}
	}
	
	if (m_poEmployees)
		m_poEmployees->AddRef();

	return (NOERROR);
}


void CBank::Clear()
{
	if (m_poEmployees)
	{
		m_poEmployees->Release();
		m_poEmployees = NULL;
	}

	m_bExpanded = FALSE;

	EmptyBSTR(m_bstrAddress1);
	EmptyBSTR(m_bstrAddress2);
	EmptyBSTR(m_bstrCity);
	EmptyBSTR(m_bstrCode);
	EmptyBSTR(m_bstrName);
	return;
}


STDMETHODIMP CBank::Term() 
{
	Clear();
	return (NOERROR);
}

//GetBankData
//Purpose:
//Select the appropriate SQL query statement and retrieve the information 
//regarding the bank object, such as its name, addresses, code, etc.
BOOL CBank::GetBankData()
{
	IList * pColl;
	HRESULT hResult;
	BOOL	bUseODBC;
	VARIANT vValue;
	BOOL	bSuccess;
	_bstr_t sSQL;
	CStrFnc sTemp;

	hResult = NOERROR;
	bSuccess = FALSE;
	bUseODBC = FALSE;

	hResult = m_poAdmin->get_UseODBC(&bUseODBC);
	if (hResult != NOERROR)
		return (false);

	//Load the appropriate SQL string from the resources
	sSQL = CStrFnc::LoadResString(bUseODBC ? scGet_Bank_SQL : scGet_Bank_MDB);
	sTemp.Format(TEXT(" %ld"), m_lPKId); 
	sSQL += sTemp.GetBstr();

	hResult = ::CoCreateInstance(CLSID_CList, NULL, THISCLSCTX, IID_IList, (void **)&pColl);
	if (hResult != NOERROR)
		return (false);

	hResult = m_poAdmin->ExecFillList(sSQL, &pColl, &bSuccess);
	if (hResult != NOERROR)
	{
		pColl->Release();
		return (false);
	}

	if (bSuccess)
	{
		::VariantInit(&vValue); 
		vValue.vt = VT_BSTR;
		hResult = pColl->get_Item(TEXT("Address1"), &vValue);
		if (hResult == NOERROR)
			::SysReAllocString(&m_bstrAddress1, V_BSTR(&vValue));

		::VariantInit(&vValue); 
		vValue.vt = VT_BSTR;
		hResult = pColl->get_Item(TEXT("Address2"), &vValue);
		if (hResult == NOERROR)
			::SysReAllocString(&m_bstrAddress2, V_BSTR(&vValue));

		::VariantInit(&vValue); 
		vValue.vt = VT_BSTR;
		hResult = pColl->get_Item(TEXT("City"), &vValue);
		if (hResult == NOERROR)
			::SysReAllocString(&m_bstrCity, V_BSTR(&vValue));

		::VariantInit(&vValue); 
		vValue.vt = VT_BSTR;
		hResult = pColl->get_Item(TEXT("Code"), &vValue);
		if (hResult == NOERROR)
			::SysReAllocString(&m_bstrCode, V_BSTR(&vValue));

		::VariantInit(&vValue); 
		vValue.vt = VT_BSTR;
		hResult = pColl->get_Item(TEXT("Name"), &vValue);
		if (hResult == NOERROR)
			::SysReAllocString(&m_bstrName, V_BSTR(&vValue));

	}

	pColl->Release();

	return (TRUE);
}

//GetEmployeesData
//
//Purpose:
//Retrieve all the employees and their data.
//Assumes that List m_poEmployees is empty
BOOL CBank::GetEmployeesData()
{
	IListEmployee * pAdminLstEmployees;
	BOOL	bSuccess;
	HRESULT hResult;
	BOOL	bUseODBC;
	_bstr_t sSQL;
	CStrFnc sTemp;
	VARIANT	varTemp;
	long	lLB;
	long	lUB;
	long	lPKid;
	long	lCount;
	IEmployee* poEmp;

	hResult	 = NOERROR;
	bUseODBC = FALSE;
	lLB		 = 0;
	lUB		 = 0;
	poEmp	 = NULL;

	hResult = m_poAdmin->get_UseODBC(&bUseODBC);
	if (hResult != NOERROR)
		return (false);

	sSQL = CStrFnc::LoadResString(bUseODBC ? scLvw_EmployeeByBankId_SQL : scLvw_EmployeeByBankId_MDB);
	sTemp.Format(TEXT(" %ld"), m_lPKId);
	sSQL += sTemp.GetBstr();

	::VariantInit(&varTemp);
	varTemp.vt = VT_BSTR | VT_ARRAY;

	hResult = m_poAdmin->ExecFillArray(sSQL, &varTemp, false, &bSuccess);
	if (hResult != NOERROR)
		return (false);

	if (bSuccess)
	{
		CSafeArray SA(V_ARRAY(&varTemp));

		if (m_poEmployees)
		{
			m_poEmployees->Release();
			m_poEmployees = NULL;
		}
			
		hResult = ::CoCreateInstance(CLSID_CListEmployee, NULL, THISCLSCTX, IID_IListEmployee, (void **)&m_poEmployees);
		if (hResult != NOERROR)
		{
			if (m_poEmployees)
			{
				m_poEmployees->Release();
				m_poEmployees = NULL;
			}

			return (false);
		}

		lLB = SA.LBound(2);
		lUB = SA.UBound(2);

		hResult = m_poAdmin->get_Employees(&pAdminLstEmployees);
		if (hResult == NOERROR)
		{
			if (pAdminLstEmployees)
			{
				for (lCount = lLB; lCount < lUB; lCount++)
				{
					poEmp	= NULL;
					CStrFnc sWork(SA.bstrGet((long)icLvwPKId,lCount));
					lPKid	= sWork.Val();
					hResult = pAdminLstEmployees->get_Item(lPKid, &poEmp);
					if (hResult == NOERROR)
					{
						m_poEmployees->Add(poEmp);
						if (poEmp)
							poEmp->Release();
					}
				}
				pAdminLstEmployees->Release();
			}
		}
	}

	return (TRUE);
}


BOOL CBank::Expand()
{
	//Check to see if the Primary Key is assigned
    if (m_lPKId)
	{		
		Clear();					//reset all members
		if (m_poAdmin)
		{
			if (GetBankData())		//get bank's address, id, etc.
			{
				if (GetEmployeesData())
					m_bExpanded = TRUE;
				else
				{	
					Term();				//terminate
					m_bExpanded = FALSE;
					FlagException(IID_IBank, scBankExpand, icErrExpand);
				}		
			}
			else					//failed to fill List
			{	
				Term();				//terminate
				m_bExpanded = FALSE;
				FlagException(IID_IBank, scBankExpand, icErrExpand);
			}		
		}
		else
		{
			Term();					//terminate
			m_bExpanded = FALSE;
			FlagException(IID_IBank, scBankExpand, icErrExpand);
		}
	}
	else
	{
		m_bExpanded = TRUE;
	}

    return (m_bExpanded);
}

//Constructor/Destructor
//
//Turn off warning 4355
#pragma warning (disable : 4355)
CBank::CBank(LPUNKNOWN pUnkOuter, PFNDESTROYED pfnDestroy)
			: m_SupportErrorInfo(this, IID_IBank)
#pragma warning (default : 4355)
{
	m_oid		 = g_oid++;
	m_cRef		 = 0;
	m_pUnkOuter	 = pUnkOuter;
	m_pfnDestroy = pfnDestroy;
	m_pITypeInfo = NULL;
	m_bExpanded	 = FALSE;
	m_bstrAddress1 = NULL;
	m_bstrAddress2 = NULL;
	m_bstrCity	 = NULL;
	m_bstrCode	 = NULL;
	m_bstrName	 = NULL;
	m_poAdmin	 = NULL;
	m_poEmployees = NULL;
	m_lPKId		 = 0;

	return;
}


CBank::~CBank(void)
{
	Term();

	if (m_poAdmin)
	{
		m_poAdmin->Release();
		m_poAdmin = NULL;
	}

	if (m_pITypeInfo)
	{
		m_pITypeInfo->Release();
		m_pITypeInfo = NULL;
	}

	return;
}

/////////////////////////////////////////////
// Standard IUnknown method implementations
/////////////////////////////////////////////

//CBank::QueryInterface
//
//Parameters:
//riid	-	input interface reference id
//ppv	-	return CBank interface pointer
STDMETHODIMP CBank::QueryInterface(REFIID riid, void ** ppv)
{
	*ppv = NULL;

	if (riid == IID_IBank || riid == IID_IUnknown || riid == IID_IDispatch)
		*ppv = this;

	if (riid == IID_ISupportErrorInfo)
		*ppv = &m_SupportErrorInfo;

	if (*ppv)
	{
		((LPUNKNOWN)*ppv)->AddRef();
		return (NOERROR);
	}

	return (E_NOINTERFACE);
}

STDMETHODIMP_(ULONG) CBank::AddRef(void)
{
	++m_cRef;
	Audit(m_oid, 1, TEXT("CBank"), m_cRef);
	return (m_cRef);
}

STDMETHODIMP_(ULONG) CBank::Release(void)
{
	--m_cRef;
	Audit(m_oid, 2, TEXT("CBank"), m_cRef);

	if (m_cRef == 0)
	{
		if (m_pfnDestroy)
			(*m_pfnDestroy)();
		Audit(m_oid, 0, TEXT("CBank gbldec"), g_lObjectCount);
		delete this;
		return (0);
	}
	else
		return (m_cRef);
}

//////////////////////////////////////////////
// Standard IDispatch method implementations
//////////////////////////////////////////////

STDMETHODIMP CBank::GetTypeInfoCount(UINT *pctInfo)
{
	*pctInfo = 1;

	return (NOERROR);
}

STDMETHODIMP CBank::GetTypeInfo(UINT itInfo, LCID lcid, ITypeInfo **ppITypeInfo)
{
	HRESULT     hResult;
	ITypeLib   *pITypeLib;
	ITypeInfo **ppITI;

	hResult	= NOERROR;
	ppITI	= NULL;

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

		hResult = pITypeLib->GetTypeInfoOfGuid(IID_IBank, ppITI);

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

STDMETHODIMP CBank::GetIDsOfNames(REFIID riid, OLECHAR **rgszNames,
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

STDMETHODIMP CBank::Invoke(DISPID dispID,	REFIID riid, LCID lcid,
				 unsigned short wFlags, DISPPARAMS *pDispParams,
				 VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
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

CBankFactory::CBankFactory(void)
{
	m_oid = g_oid++;
	m_cRef = 0L;
	return;
}

CBankFactory::~CBankFactory(void)
{
	return;
}

//CBankFactory::QueryInterface
//Purpose     - Used to return an interface pointer from the requested interface id.
//Inputs      - (REFIID)  riid
//              (void **) ppvCBankIf
STDMETHODIMP CBankFactory::QueryInterface(REFIID riid, void ** ppvCBankIf)
{
	*ppvCBankIf = NULL;

	if (riid == IID_IUnknown ||
		riid == IID_IClassFactory)
		*ppvCBankIf = this;

	if (*ppvCBankIf)
	{
		((LPUNKNOWN)*ppvCBankIf)->AddRef();
		return (NOERROR);
	}

	return (E_NOINTERFACE);
}

STDMETHODIMP_(ULONG) CBankFactory::AddRef(void)
{
	++m_cRef;
	Audit(m_oid, 1, TEXT("CBankF"), m_cRef);
	return (m_cRef);
}

STDMETHODIMP_(ULONG) CBankFactory::Release(void)
{
	--m_cRef;
	Audit(m_oid, 2, TEXT("CBankF"), m_cRef);
	if (m_cRef == 0)
	{
		delete this;
		return (0);
	}
	else
		return (m_cRef);
}

//CBankFactory::CreateInstance
//Purpose     - Creates a CBank object and returns an Interface pointer
//Inputs      - (LPUNKNOWN) pUnkOuter
//              (REFIID)    riid
//				(void **)   ppvCBenefIf
STDMETHODIMP CBankFactory::CreateInstance(LPUNKNOWN pUnkOuter, REFIID riid, void ** ppvCBankIf)
{
	PCBank		pCBank;
	HRESULT		hResult;

	hResult		= NOERROR;
	*ppvCBankIf	= NULL;

	if (pUnkOuter && riid != IID_IUnknown)
		return (CLASS_E_NOAGGREGATION);

	pCBank = new CBank(pUnkOuter, ObjectDestroyed);

	if (pCBank == NULL)
		return (hResult);

	if (pCBank->CFInit())
		hResult = pCBank->QueryInterface(riid, ppvCBankIf);

	if (hResult != NOERROR)
		delete pCBank;
	else
	{
		g_lObjectCount++;
		Audit(m_oid, 0, TEXT("CBank gblinc"), g_lObjectCount);
	}

	return (hResult);
}

//CBenefitFactory::LockServer
//Purpose     - Increments/Decrements the lock count to prevent class factory server shutdown.
//Inputs      - (BOOL)    fLock
STDMETHODIMP CBankFactory::LockServer(BOOL fLock)
{
	if (fLock)
		g_lLockedObjects++;
	else
		g_lLockedObjects--;

	return (NOERROR);
}

//CBank::CFInit
//
BOOL CBank::CFInit(void)
{
	LPUNKNOWN pIUnknown = this;

	if (m_pUnkOuter)
		pIUnknown = m_pUnkOuter;

	return (TRUE);
}
