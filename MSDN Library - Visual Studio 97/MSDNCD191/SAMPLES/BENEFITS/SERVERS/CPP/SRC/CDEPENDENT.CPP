//cdependent.cpp  - Copyright (C) 1996 Microsoft Corporation
//            - All rights reserved.
//Date        - 10/19/96
//Authors     - Ken Bergmann, Robert Coleridge
//Purpose     - Implementation file for CDependant object

#pragma warning (disable : 4800)	// turns off bool warning

#include "cdependent.h"

extern long g_oid;

//Globals used to determine when to unload a In-Proc or Local Server.
extern ULONG       g_lObjectCount;
extern ULONG       g_lLockedObjects;

STDMETHODIMP CDependant::putref_Admin(IAdmin * poAdmin) 
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

STDMETHODIMP CDependant::get_PKId(long FAR* plPKId) 
{
	*plPKId = m_lPKId;
	return (NOERROR);
}

STDMETHODIMP CDependant::put_PKId(long lPKId) 
{
	m_lPKId = lPKId;
	return (NOERROR);
}

STDMETHODIMP CDependant::get_DepTypeId(long FAR* plDepTypeId) 
{
	if (m_bExpanded)
		*plDepTypeId = m_lDepTypeId;
	else
	{
		if (Expand())
			*plDepTypeId = m_lDepTypeId;
		else
		{
			Term();
			return (FlagException(IID_IDependant, scDepExpand, icErrExpand));
		}
	}
	return (NOERROR);
}

STDMETHODIMP CDependant::put_DepTypeId(long lDepTypeId) 
{
	m_lDepTypeId = lDepTypeId;
	return (NOERROR);
}

STDMETHODIMP CDependant::get_DateOfBirth(DATE * pvdtDateOfBirth) 
{
	if (m_bExpanded)
		*pvdtDateOfBirth = m_vDateOfBirth;
	else
	{
		if (Expand())
			*pvdtDateOfBirth = m_vDateOfBirth;
		else
		{
			Term();
			return (FlagException(IID_IDependant, scDepExpand, icErrExpand));
		}
	}
	return (NOERROR);
}

STDMETHODIMP CDependant::put_DateOfBirth(DATE pvdtDateOfBirth) 
{
	m_vDateOfBirth = pvdtDateOfBirth;
	return (NOERROR);
}

STDMETHODIMP CDependant::get_Employee(IEmployee * FAR* ppoEmployee) 
{
	if (m_bExpanded)
		*ppoEmployee = (IEmployee *)m_poEmployee;
	else
	{
		if (Expand())
			*ppoEmployee = (IEmployee *)m_poEmployee;
		else
		{
			Term();
			return (FlagException(IID_IDependant, scDepExpand, icErrExpand));
		}
	}

	if (m_poEmployee)
		m_poEmployee->AddRef();

	return (NOERROR);
}

STDMETHODIMP CDependant::putref_Employee(IEmployee * poEmployee) 
{
	if (m_poEmployee)
	{
		m_poEmployee->Release();
		m_poEmployee = NULL;
	}

	m_poEmployee = (IEmployee*)poEmployee;

	if (m_poEmployee)
		m_poEmployee->AddRef();

	return (NOERROR); 
}

STDMETHODIMP CDependant::get_FirstName(BSTR FAR * pbstrFirstName) 
{
	if (m_bExpanded)
		::SysReAllocString(pbstrFirstName,m_bstrFirstName);
	else
	{
		if (Expand())
			::SysReAllocString(pbstrFirstName,m_bstrFirstName);
		else
		{
			Term();
			return (FlagException(IID_IDependant, scDepExpand, icErrExpand));
		}
	}
	return (NOERROR);
}

STDMETHODIMP CDependant::put_FirstName(BSTR bstrFirstName) 
{
	::SysReAllocString(&m_bstrFirstName, bstrFirstName);	
	return (NOERROR);
}

STDMETHODIMP CDependant::get_Gender(BOOL FAR* pbGender) 
{
	if (m_bExpanded)
		*pbGender = m_iGender;
	else
	{
		if (Expand())
			*pbGender = m_iGender;
		else
		{
			Term();
			return (FlagException(IID_IDependant, scDepExpand, icErrExpand));
		}
	}
	return (NOERROR);
}

STDMETHODIMP CDependant::put_Gender(BOOL iGender) 
{
	m_iGender = iGender;
	return (NOERROR);
}

STDMETHODIMP CDependant::get_LastName(BSTR FAR* pbstrLastName) 
{
	if (m_bExpanded)
		::SysReAllocString(pbstrLastName,m_bstrLastName);
	else
	{
		if (Expand())
			::SysReAllocString(pbstrLastName,m_bstrLastName);
		else
		{
			Term();
			return (FlagException(IID_IDependant, scDepExpand, icErrExpand));
		}
	}
	return (NOERROR);
}

STDMETHODIMP CDependant::put_LastName(BSTR bstrLastName) 
{
	::SysReAllocString(&m_bstrLastName, bstrLastName);	
	return (NOERROR);
}

STDMETHODIMP CDependant::get_Status(BOOL FAR* piStatus)
{
	if (m_bExpanded)
		*piStatus = m_iStatus;
	else
	{
		if (Expand())
			*piStatus = m_iStatus;
		else
		{
			Term();
			return (FlagException(IID_IDependant, scDepExpand, icErrExpand));
		}
	}
	return (NOERROR);
}

STDMETHODIMP CDependant::put_Status(BOOL iStatus) 
{
	m_iStatus = iStatus;
	return (NOERROR);
}

STDMETHODIMP CDependant::Term() 
{
	Clear();
	return (NOERROR);
}

//CDependant::Clear
//Purpose - Preserve m_poAdmin reference. Remove other object references.
void CDependant::Clear()
{
	m_iGender = 0;
	m_iStatus = 0;
	m_bExpanded = FALSE;
	m_vDateOfBirth = 0;
	m_lDepTypeId = 0;

	if (m_poEmployee)
	{
		m_poEmployee->Release();
		m_poEmployee = NULL;
	}

	EmptyBSTR(m_bstrFirstName);
	EmptyBSTR(m_bstrLastName);
}

BOOL CDependant::Expand()
{
	if (m_lPKId)
	{
		Clear();
		if (m_poAdmin)
		{
			if (GetDependantData())
				m_bExpanded = TRUE;
			else
			{
				Term();
				m_bExpanded = FALSE;
				FlagException(IID_IDependant, scDepExpand, icErrExpand);
			}
		}
		else
		{
			Term();
			m_bExpanded = FALSE;
			FlagException(IID_IDependant, scDepExpand, icErrExpand);
		}
	}
	else
	{
		m_bExpanded = TRUE;
	}

	return (m_bExpanded);
}

BOOL CDependant::GetDependantData()
{
	IListEmployee * pAdminLstEmployees;
	IList * pColl;
	BOOL	bResult;
	BOOL	bSuccess;
	HRESULT hResult;
	BOOL	bUseODBC;
	_bstr_t sSQL;
	CStrFnc sTemp;
	VARIANT vValue;

	hResult	= NOERROR;
	bResult = FALSE;
	bSuccess = FALSE;
	bUseODBC = FALSE;

	hResult = m_poAdmin->get_UseODBC(&bUseODBC);
	if (hResult != NOERROR)
		return (false);

	sSQL = CStrFnc::LoadResString(bUseODBC ? scGet_Dependant_SQL : scGet_Dependant_MDB);

	sTemp.Format(TEXT(" %ld"), m_lPKId); 
	sSQL += sTemp.GetBstr();

	hResult = ::CoCreateInstance(CLSID_CList, NULL, THISCLSCTX, IID_IList, (void **)&pColl);
    if (hResult == NOERROR)
	{
		hResult = m_poAdmin->ExecFillList(sSQL, &pColl, &bSuccess);
		if (hResult == NOERROR)
		{
			if (bSuccess)
			{
				::VariantInit(&vValue);
				vValue.vt = VT_I4;
				hResult = pColl->get_Item(TEXT("DepTypeId"), &vValue);
				if (hResult == NOERROR)
					m_lDepTypeId = CStrFnc::VariantBstrToLong(vValue);

				::VariantInit(&vValue); 
				vValue.vt = VT_DATE;
				hResult = pColl->get_Item(TEXT("DateOfBirth"), &vValue);
				if (hResult == NOERROR)
					m_vDateOfBirth = CStrFnc::FormatDate(V_BSTR(&vValue));
				
				::VariantInit(&vValue); 
				vValue.vt = VT_BSTR;
				hResult = pColl->get_Item(TEXT("FirstName"), &vValue);
				if (hResult == NOERROR)
					::SysReAllocString(&m_bstrFirstName, V_BSTR(&vValue));

				::VariantInit(&vValue); 
				vValue.vt = VT_I2;
				hResult = pColl->get_Item(TEXT("Gender"), &vValue);
				if (hResult == NOERROR)
					m_iGender = CStrFnc::Boolean(V_BSTR(&vValue));

				::VariantInit(&vValue); 
				vValue.vt = VT_BSTR;
				hResult = pColl->get_Item(TEXT("LastName"), &vValue);
				if (hResult == NOERROR)
					::SysReAllocString(&m_bstrLastName, V_BSTR(&vValue));

				if (m_poEmployee)
				{
					m_poEmployee->Release();
					m_poEmployee = NULL;
				}

				::VariantInit(&vValue);
				hResult = pColl->get_Item(TEXT("EmployeeId"), &vValue);
				if (hResult == NOERROR)
				{
					hResult = m_poAdmin->get_Employees(&pAdminLstEmployees);
					if (hResult == NOERROR)
					{
						if (pAdminLstEmployees)
						{
							hResult = pAdminLstEmployees->get_Item(CStrFnc::VariantBstrToLong(vValue), &m_poEmployee);
							if (hResult == NOERROR)
								bResult = TRUE;
							else
							{
								if (m_poEmployee)
								{
									m_poEmployee->Release();
									m_poEmployee = NULL;
								}
							}
							pAdminLstEmployees->Release();
						}
					}
				}
			}
		}
		pColl->Release();;
	}
	return (bResult);
}


//Constructor/Destructor
//
//Turn off warning 4355
#pragma warning (disable : 4355)
CDependant::CDependant(	LPUNKNOWN pUnkOuter, PFNDESTROYED pfnDestroy)
				: m_SupportErrorInfo(this, IID_IDependant)
#pragma warning (default : 4355)
{
	m_oid = g_oid++;
	m_cRef = 0;
	m_pUnkOuter = pUnkOuter;
	m_pfnDestroy = pfnDestroy;
	m_poAdmin = NULL;
	m_poEmployee = NULL;
	m_lPKId = 0;
    m_lDepTypeId = 0;
	m_bExpanded = FALSE;
	m_iGender = 0;
	m_iStatus = 0;
	m_bstrFirstName = NULL;
	m_bstrLastName = NULL;
	m_vDateOfBirth = 0;
	m_lDepTypeId = 0;
	m_pITypeInfo = NULL;

	return;
}

CDependant::~CDependant(void)
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

////////////////////////////////////////////
// Standard IUnknown method implementations
////////////////////////////////////////////

//CBenefitLink::QueryInterface
//
//Parameters:
//riid	-	input interface reference id
//ppvCDependant	-	return CBenefitLink interface pointer
STDMETHODIMP CDependant::QueryInterface(REFIID riid, void ** ppvCDependant)
{
	*ppvCDependant = NULL;

	if (riid == IID_IDependant || riid == IID_IUnknown || riid == IID_IDispatch)
		*ppvCDependant = this;

	if (riid == IID_ISupportErrorInfo)
		*ppvCDependant = &m_SupportErrorInfo;

	if (*ppvCDependant)
	{
		((LPUNKNOWN)*ppvCDependant)->AddRef();
		return (NOERROR);
	}

	return (E_NOINTERFACE);
}

STDMETHODIMP_(ULONG) CDependant::AddRef(void)
{
	++m_cRef;
	Audit(m_oid, 1, TEXT("CDependant"), m_cRef);
	return (m_cRef);
}

STDMETHODIMP_(ULONG) CDependant::Release(void)
{
	--m_cRef;
	Audit(m_oid, 2, TEXT("CDependant"), m_cRef);

	if (m_cRef == 0)
	{
		if (m_pfnDestroy)
			(*m_pfnDestroy)();
		Audit(m_oid, 0, TEXT("CDependant gbldec"), g_lObjectCount);
		delete this;
		return (0);
	}
	else
		return (m_cRef);
}

//////////////////////////////////////////////
// Standard IDispatch method implementations
//////////////////////////////////////////////

STDMETHODIMP CDependant::GetTypeInfoCount(UINT *pctInfo)
{
	*pctInfo = 1;

	return (NOERROR);
}

STDMETHODIMP CDependant::GetTypeInfo(UINT itInfo, LCID lcid, ITypeInfo **ppITypeInfo)
{
	HRESULT     hResult;
	ITypeLib   *pITypeLib;
	ITypeInfo **ppITI;

	hResult = NOERROR;
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

		hResult = pITypeLib->GetTypeInfoOfGuid(IID_IDependant, ppITI);
		
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

STDMETHODIMP CDependant::GetIDsOfNames(REFIID riid, OLECHAR **rgszNames,
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

STDMETHODIMP CDependant::Invoke(DISPID dispID,	REFIID riid, LCID lcid,
				 unsigned short wFlags, DISPPARAMS *pDispParams,
				 VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
	HRESULT     hResult;
	ITypeInfo * pTI;
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

//CBank::CFInit
//
BOOL CDependant::CFInit(void)
{
	LPUNKNOWN pIUnknown;

	pIUnknown = this;

	if (m_pUnkOuter)
		pIUnknown = m_pUnkOuter;

	return (TRUE);
}

CDependantFactory::CDependantFactory(void)
{
	m_oid	= g_oid++;
	m_cRef	= 0L;
	return;
}

CDependantFactory::~CDependantFactory(void)
{
	return;
}

//CDependantFactory::QueryInterface
//Purpose     - Used to return an interface pointer from the requested interface id.
//Inputs      - (REFIID)  riid
//              (void **) ppvCDependant
STDMETHODIMP CDependantFactory::QueryInterface(REFIID riid, void ** ppvCDependantIf)
{
	*ppvCDependantIf = NULL;

	if (riid == IID_IUnknown ||
		riid == IID_IClassFactory)
		*ppvCDependantIf = this;

	if (*ppvCDependantIf)
	{
		((LPUNKNOWN)*ppvCDependantIf)->AddRef();
		return (NOERROR);
	}

	return (E_NOINTERFACE);
}

STDMETHODIMP_(ULONG) CDependantFactory::AddRef(void)
{
	++m_cRef;
	Audit(m_oid, 1, TEXT("CDependantF"), m_cRef);
	return (m_cRef);
}

STDMETHODIMP_(ULONG) CDependantFactory::Release(void)
{
	--m_cRef;
	Audit(m_oid, 2, TEXT("CDependantF"), m_cRef);
	if (m_cRef == 0)
	{
		delete this;
		return (0);
	}
	else
		return (m_cRef);
}

//CDependantFactory::CreateInstance
//Purpose     - Creates a CDependant object and returns an Interface pointer
//Inputs      - (LPUNKNOWN) pUnkOuter
//              (REFIID)    riid
//				(void **)   ppvCDependantIf
STDMETHODIMP CDependantFactory::CreateInstance(LPUNKNOWN pUnkOuter, REFIID riid, void ** ppvCDependantIf)
{
	PCDependant	pCDependant;
	HRESULT		hResult;

	*ppvCDependantIf = NULL;
	hResult = NOERROR;

	if (pUnkOuter && riid != IID_IUnknown)
		return (CLASS_E_NOAGGREGATION);

	pCDependant = new CDependant(pUnkOuter, ObjectDestroyed);

	if (pCDependant == NULL)
		return (hResult);

	if (pCDependant->CFInit())
		hResult = pCDependant->QueryInterface(riid, ppvCDependantIf);

	if (hResult != NOERROR)
		delete pCDependant;
	else
	{
		g_lObjectCount++;
		Audit(m_oid, 0, TEXT("CDependant gblinc"), g_lObjectCount);
	}

	return (hResult);
}

//CDependantFactory::LockServer
//Purpose     - Increments/Decrements the lock count to prevent class factory server shutdown.
//Inputs      - (BOOL)    fLock
STDMETHODIMP CDependantFactory::LockServer(BOOL fLock)
{
	if (fLock)
		g_lLockedObjects++;
	else
		g_lLockedObjects--;

	return (NOERROR);
}

