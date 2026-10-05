//cbenefitlink.cpp  - Copyright (C) 1996 Microsoft Corporation
//            - All rights reserved.
//Date        - 10/19/96
//Authors     - Ken Bergmann, Robert Coleridge
//Purpose     - Implementation file for the CBenefitLink object

#pragma warning (disable : 4800)	// turns off bool warning

#include "cbenefitlink.h"

extern long g_oid;

//Globals used to determine when to unload a In-Proc or Local Server.
extern ULONG       g_lObjectCount;
extern ULONG       g_lLockedObjects;

STDMETHODIMP CBenefitLink::get_Amount(CURRENCY __RPC_FAR * newVal) 
{
	if (newVal == NULL)
		return (E_POINTER);

	if (m_bExpanded)
		*newVal = m_Amount;
	else
	{
		if (Expand())
			*newVal = m_Amount;
		else
		{
			Term();
			return (FlagException(IID_IBenefitLink, scBnfLnkExpand, icErrExpand));
		}
	}
	return (NOERROR);
}

STDMETHODIMP CBenefitLink::putref_Admin(IAdmin * poAdmin) 
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

STDMETHODIMP CBenefitLink::get_PKId(long FAR * plPKId) 
{
	*plPKId = m_lPKId;
	return (NOERROR);
}

STDMETHODIMP CBenefitLink::put_PKId(long lPKId) 
{
	m_lPKId = lPKId;
	return (NOERROR);
}

STDMETHODIMP CBenefitLink::put_Amount(CURRENCY cyAmount)
{
	m_Amount = cyAmount;
	return (NOERROR);
}

STDMETHODIMP CBenefitLink::get_Benefit(IBenefit * FAR * poBenefit)
{
	if (m_bExpanded)
		*poBenefit = (IBenefit *)m_poBenefit;
	else
	{
		if (Expand())
			*poBenefit = (IBenefit *)m_poBenefit;
		else
		{
			Term();
			return (FlagException(IID_IBenefitLink, scBnfLnkExpand, icErrExpand));
		}
	}

	if (m_poBenefit)
		m_poBenefit->AddRef();

	return (NOERROR);
}

STDMETHODIMP CBenefitLink::putref_Benefit(IBenefit * poBenefit)
{
	if (m_poBenefit)
	{
		m_poBenefit->Release();
		m_poBenefit = NULL;
	}

	m_poBenefit = (IBenefit *)poBenefit;

	if (m_poBenefit)
		m_poBenefit->AddRef();

	return (NOERROR);
}

STDMETHODIMP CBenefitLink::get_CoverageId(long FAR * plCoverageId)
{
	if (m_bExpanded)
		*plCoverageId = m_CoverageId;
	else
	{
		if (Expand())
			*plCoverageId = m_CoverageId;
		else
		{
			Term();
			return (FlagException(IID_IBenefitLink, scBnfLnkExpand, icErrExpand));
		}
	}
	return (NOERROR);
}

STDMETHODIMP CBenefitLink::put_CoverageId(long lCoverageId)
{
	m_CoverageId = lCoverageId;
	return (NOERROR);
}

STDMETHODIMP CBenefitLink::get_Employee(IEmployee * FAR * poEmployee)
{
	if (m_bExpanded)
		*poEmployee = (IEmployee *)m_poEmployee;
	else
	{
		if (Expand())
			*poEmployee = (IEmployee *)m_poEmployee;
		else
		{
			Term();
			return (FlagException(IID_IBenefitLink, scBnfLnkExpand, icErrExpand));
		}
	}

	if (m_poEmployee)
		m_poEmployee->AddRef();

	return (NOERROR);
}

STDMETHODIMP CBenefitLink::putref_Employee(IEmployee * poEmployee)
{
	if (m_poEmployee)
	{
		m_poEmployee->Release();
		m_poEmployee = NULL;
	}

	m_poEmployee = (IEmployee *)poEmployee;

	if (m_poEmployee)
		m_poEmployee->AddRef();

	return (NOERROR);
}

STDMETHODIMP CBenefitLink::get_TaxId(long FAR * plTaxId)
{
	if (m_bExpanded)
		*plTaxId = m_TaxId;
	else
	{
		if (Expand())
			*plTaxId = m_TaxId;
		else
		{
			Term();
			return (FlagException(IID_IBenefitLink, scBnfLnkExpand, icErrExpand));
		}
	}
	return (NOERROR);
}

STDMETHODIMP CBenefitLink::put_TaxId(long lTaxId)
{
	m_TaxId = lTaxId;
	return (NOERROR);
}

STDMETHODIMP CBenefitLink::Term(THIS)
{
	Clear();
	return (NOERROR);
}

void CBenefitLink::Clear()
{
	m_CoverageId = 0;
	m_TaxId = 0;
	m_bExpanded = FALSE;
	m_Amount.int64 = 0;
	if (m_poEmployee)
	{
		m_poEmployee->Release();
		m_poEmployee = NULL;
	}

	if (m_poBenefit)
	{
		m_poBenefit->Release();
		m_poBenefit = NULL;
	}

	return;
}

BOOL CBenefitLink::Expand()
{
	IList *			pList;
	_bstr_t			sSQL;
	CStrFnc			sTemp;
	BOOL			bUseODBC;
	HRESULT			hResult;
	VARIANT			vItemData;
	IListEmployee * pAdminLstEmployees;
	IListBenefit *	pAdminLstBenefits;
	BOOL			bSuccess;

	hResult = NOERROR;

    if (m_lPKId)
	{		
		Clear();
		if (m_poAdmin)
		{
			hResult = ::CoCreateInstance(CLSID_CList, NULL, THISCLSCTX, IID_IList, (void **)&pList);
			if (hResult == NOERROR)
			{
				hResult = m_poAdmin->get_UseODBC(&bUseODBC);
				if (hResult != NOERROR)
				{
					pList->Release();
					return (hResult);
				}

				sSQL = CStrFnc::LoadResString(bUseODBC ? scGet_EmplBnf_SQL: scGet_EmplBnf_MDB);
				sSQL += TEXT(" ");

				sTemp.Format(TEXT("%ld"), m_lPKId);
				sSQL += sTemp.GetBstr();

				hResult = m_poAdmin->ExecFillList(sSQL, &pList, &bSuccess);
				if (hResult == NOERROR)
				{
					if (bSuccess)
					{
						::VariantInit(&vItemData);
						hResult = pList->get_Item(TEXT("CoverageId"), &vItemData);
						if (hResult == NOERROR)
							m_CoverageId = CStrFnc::VariantBstrToLong(vItemData);

						::VariantInit(&vItemData);
						hResult = pList->get_Item(TEXT("TaxId"), &vItemData);
						if (hResult == NOERROR)
							m_TaxId = CStrFnc::VariantBstrToLong(vItemData);

						::VariantInit(&vItemData);
						hResult = pList->get_Item(TEXT("Amount"), &vItemData);
						if (hResult == NOERROR)
							m_Amount = CStrFnc::BstrToCy(V_BSTR(&vItemData));

						if (m_poEmployee)
						{
							m_poEmployee->Release();
							m_poEmployee = NULL;
						}

						::VariantInit(&vItemData);
						hResult = pList->get_Item(TEXT("EmployeeId"), &vItemData);
						if (hResult == NOERROR)
						{
							hResult = m_poAdmin->get_Employees(&pAdminLstEmployees);
							if (hResult == NOERROR)
							{
								if (pAdminLstEmployees)
								{
									hResult = pAdminLstEmployees->get_Item(CStrFnc::VariantBstrToLong(vItemData), &m_poEmployee);
									if (hResult != NOERROR)
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

						if (m_poBenefit)
						{
							m_poBenefit->Release();
							m_poBenefit = NULL;
						}
						
						::VariantInit(&vItemData);
						hResult = pList->get_Item(TEXT("BenefitId"), &vItemData);
						if (hResult == NOERROR)
						{
							hResult = m_poAdmin->get_Benefits(&pAdminLstBenefits);
							if (hResult == NOERROR)
							{
								if (pAdminLstBenefits)
								{
									hResult = pAdminLstBenefits->get_Item(CStrFnc::VariantBstrToLong(vItemData), &m_poBenefit);
									if (hResult != NOERROR)
									{
										if (m_poBenefit)
										{
											m_poBenefit->Release();
											m_poBenefit = NULL;
										}
									}
									pAdminLstBenefits->Release();
								}
							}
						}
						
						m_bExpanded = TRUE;
					}
				}
				pList->Release();
			}
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
CBenefitLink::CBenefitLink(	LPUNKNOWN pUnkOuter, PFNDESTROYED pfnDestroy)
			: m_SupportErrorInfo(this, IID_IBenefitLink)
#pragma warning (default : 4355)
{
	m_oid = g_oid++;
	m_cRef = 0;
	m_pUnkOuter = pUnkOuter;
	m_pfnDestroy = pfnDestroy;
	m_CoverageId = 0;
	m_TaxId = 0;
	m_bExpanded = FALSE;
	m_Amount.int64 = 0;
	m_poEmployee = NULL;
	m_poBenefit = NULL;
	m_poAdmin = NULL;
	m_lPKId = 0;
	m_pITypeInfo = NULL;

	return;
}

CBenefitLink::~CBenefitLink(void)
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
//ppv	-	return CBenefitLink interface pointer
STDMETHODIMP CBenefitLink::QueryInterface(REFIID riid, void ** ppv)
{
	*ppv = NULL;

	if (riid == IID_IBenefitLink || riid == IID_IUnknown || riid == IID_IDispatch)
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

STDMETHODIMP_(ULONG) CBenefitLink::AddRef(void)
{
	++m_cRef;
	Audit(m_oid, 1, TEXT("CBenefitLink"), m_cRef);
	return (m_cRef);
}

STDMETHODIMP_(ULONG) CBenefitLink::Release(void)
{
	--m_cRef;
	Audit(m_oid, 2, TEXT("CBenefitLink"), m_cRef);

	if (m_cRef == 0L)
	{
		if (m_pfnDestroy)
			(*m_pfnDestroy)();
		Audit(m_oid, 0, TEXT("CBenefitLink gbldec"), g_lObjectCount);
		delete this;
		return (0);
	}
	else
		return (m_cRef);
}

//////////////////////////////////////////////
// Standard IDispatch method implementations
//////////////////////////////////////////////

STDMETHODIMP CBenefitLink::GetTypeInfoCount(UINT *pctInfo)
{
	*pctInfo = 1;

	return (NOERROR);
}

STDMETHODIMP CBenefitLink::GetTypeInfo(UINT itInfo, LCID lcid, ITypeInfo **ppITypeInfo)
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

		hResult = pITypeLib->GetTypeInfoOfGuid(IID_IBenefitLink, ppITI);
		
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

STDMETHODIMP CBenefitLink::GetIDsOfNames(REFIID riid, OLECHAR **rgszNames, 
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

STDMETHODIMP CBenefitLink::Invoke(DISPID dispID,	REFIID riid, LCID lcid,
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

CBenefitLinkFactory::CBenefitLinkFactory(void)
{
	m_oid = g_oid++;
	m_cRef = 0L;
	return;
}

CBenefitLinkFactory::~CBenefitLinkFactory(void)
{
	return;
}

//CBenefitLinkFactory::QueryInterface
//Purpose     - Used to return an interface pointer from the requested interface id.
//Inputs      - (REFIID)  riid
//              (void **) ppvCBenefitLinkIf
STDMETHODIMP CBenefitLinkFactory::QueryInterface(REFIID riid, void ** ppvCBenefitLinkIf)
{
	*ppvCBenefitLinkIf = NULL;

	if (riid == IID_IUnknown ||
		riid == IID_IClassFactory)
		*ppvCBenefitLinkIf = this;

	if (*ppvCBenefitLinkIf)
	{
		((LPUNKNOWN)*ppvCBenefitLinkIf)->AddRef();
		return (NOERROR);
	}

	return (E_NOINTERFACE);
}

STDMETHODIMP_(ULONG) CBenefitLinkFactory::AddRef(void)
{
	++m_cRef;
	Audit(m_oid, 1, TEXT("CBenefitLinkF"), m_cRef);
	return (m_cRef);
}

STDMETHODIMP_(ULONG) CBenefitLinkFactory::Release(void)
{
	--m_cRef;
	Audit(m_oid, 2, TEXT("CBenefitLinkF"), m_cRef);
	if (m_cRef == 0)
	{
		delete this;
		return (0);
	}
	else
		return (m_cRef);
}

//CBenefitLinkFactory::CreateInstance
//Purpose     - Creates a CBenefit object and returns an Interface pointer
//Inputs      - (LPUNKNOWN) pUnkOuter
//              (REFIID)    riid
//				(void **)   ppvCBenefIf
STDMETHODIMP CBenefitLinkFactory::CreateInstance(LPUNKNOWN pUnkOuter, REFIID riid, void ** ppvCBenefitLinkIf)
{
	PCBenefitLink     pCBenefLink;
	HRESULT           hResult;

	hResult = NOERROR;
	*ppvCBenefitLinkIf = NULL;

	if (pUnkOuter && riid != IID_IUnknown)
		return (CLASS_E_NOAGGREGATION);

	pCBenefLink = new CBenefitLink(pUnkOuter, ObjectDestroyed);

	if (pCBenefLink == NULL)
		return (hResult);

	if (pCBenefLink->CFInit())
		hResult = pCBenefLink->QueryInterface(riid, ppvCBenefitLinkIf);

	if (hResult != NOERROR)
		delete pCBenefLink;
	else
	{
		g_lObjectCount++;
		Audit(m_oid, 0, TEXT("CBenefitLink gblinc"), g_lObjectCount);
	}

	return (hResult);
}

//CBenefitLinkFactory::LockServer
//Purpose     - Increments/Decrements the lock count to prevent class factory server shutdown.
//Inputs      - (BOOL)    fLock
STDMETHODIMP CBenefitLinkFactory::LockServer(BOOL fLock)
{
	if (fLock)
		g_lLockedObjects++;
	else
		g_lLockedObjects--;

	return (NOERROR);
}

//CBank::CFInit
//
BOOL CBenefitLink::CFInit(void)
{
	LPUNKNOWN pIUnknown = this;

	if (m_pUnkOuter)
		pIUnknown = m_pUnkOuter;

	return (TRUE);
}
