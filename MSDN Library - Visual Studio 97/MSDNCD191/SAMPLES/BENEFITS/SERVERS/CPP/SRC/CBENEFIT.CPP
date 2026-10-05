//cbenefit.cpp  - Copyright (C) 1996 Microsoft Corporation
//            - All rights reserved.
//Date        - 10/19/96
//Authors     - Ken Bergmann, Robert Coleridge
//Purpose     - Implementation file for the CBenefit object.

#pragma warning (disable : 4800)	// turns off bool perf warning

#include "cbenefit.h"

extern long g_oid;

//Globals used to determine when to unload a In-Proc or Local Server.
extern ULONG       g_lObjectCount;
extern ULONG       g_lLockedObjects;

STDMETHODIMP CBenefit::putref_Admin(IAdmin * poAdmin)
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

STDMETHODIMP CBenefit::get_PKId(long FAR * plPKId) 
{
    *plPKId = m_lPKId;
	return (NOERROR);
}

STDMETHODIMP CBenefit::put_PKId(long lPKId) 
{
	m_lPKId = lPKId;
	return (NOERROR);
}

STDMETHODIMP CBenefit::get_BnfTaxId(long FAR * plBnfTaxId) 
{
	if (m_bExpanded)
		*plBnfTaxId = m_lBnfTaxId;
	else
	{
		if (Expand())
			*plBnfTaxId = m_lBnfTaxId;
		else
		{
			Term();
			return (FlagException(IID_IBenefit, scBnfExpand, icErrExpand));
		}
	}
	return (NOERROR);
}

STDMETHODIMP CBenefit::put_BnfTaxId(long lBnfTaxId) 
{
	m_lBnfTaxId = lBnfTaxId;
	return (NOERROR);
}

STDMETHODIMP CBenefit::get_BnfTypeId(long FAR * plBnfTypeId) 
{
	if (m_bExpanded)
		*plBnfTypeId = m_lBnfTypeId;
	else
	{
		if (Expand())
			*plBnfTypeId = m_lBnfTypeId;
		else
		{
			Term();
			return (FlagException(IID_IBenefit, scBnfExpand, icErrExpand));
		}
	}
	return (NOERROR);
}

STDMETHODIMP CBenefit::put_BnfTypeId(long lBnfTypeId) 
{
	m_lBnfTypeId = lBnfTypeId;
	return (NOERROR);
}

STDMETHODIMP CBenefit::get_Cost(long FAR * plCost) 
{
	if (m_bExpanded)
		*plCost = m_lCost;
	else
	{
		if (Expand())
			*plCost = m_lCost;
		else
		{
			Term();
			return (FlagException(IID_IBenefit, scBnfExpand, icErrExpand));
		}
	}
	return (NOERROR);
}

STDMETHODIMP CBenefit::put_Cost(long lCost) 
{
    m_lCost = lCost;
	return (NOERROR);
}

STDMETHODIMP CBenefit::get_Description(BSTR FAR * pbstrDescription) 
{
	if (m_bExpanded)
	{
		if (*pbstrDescription)
			::SysReAllocString(pbstrDescription, m_bstrDescription);
		else
			*pbstrDescription = ::SysAllocString(m_bstrDescription);
	}
	else
	{
		if (Expand())
		{
			if (*pbstrDescription)
				::SysReAllocString(pbstrDescription, m_bstrDescription);
			else
				*pbstrDescription = ::SysAllocString(m_bstrDescription);
		}
		else
		{
			Term();
			return (FlagException(IID_IBenefit, scBnfExpand, icErrExpand));
		}
	}
	return (NOERROR);
}

STDMETHODIMP CBenefit::put_Description(BSTR bstrDescription) 
{
    ::SysReAllocString(&m_bstrDescription, bstrDescription);
	return (NOERROR);
}


STDMETHODIMP CBenefit::get_Increment(short FAR * piIncrement) 
{
	if (m_bExpanded)
		*piIncrement = m_iIncrement;
	else
	{
		if (Expand())
			*piIncrement = m_iIncrement;
		else
		{
			Term();
			return (FlagException(IID_IBenefit, scBnfExpand, icErrExpand));
		}
	}
	return (NOERROR);
}

STDMETHODIMP CBenefit::put_Increment(short iIncrement) 
{
	m_iIncrement = iIncrement;
	return (NOERROR);
}

STDMETHODIMP CBenefit::get_Maximum(long FAR * plMaximum) 
{
	if (m_bExpanded)
		*plMaximum = m_lMaximum;
	else
	{
		if (Expand())
			*plMaximum = m_lMaximum;
		else
		{
			Term();
			return (FlagException(IID_IBenefit, scBnfExpand, icErrExpand));
		}
	}
	return (NOERROR);
}

STDMETHODIMP CBenefit::put_Maximum(long lMaximum) 
{
    m_lMaximum = lMaximum;
	return (NOERROR);
}

STDMETHODIMP CBenefit::get_Minimum(long FAR * plMinimum) 
{
	if (m_bExpanded)
		*plMinimum = m_lMinimum;
	else
	{
		if (Expand())
			*plMinimum = m_lMinimum;
		else
		{
			Term();
			return (FlagException(IID_IBenefit, scBnfExpand, icErrExpand));
		}
	}
	return (NOERROR);
}

STDMETHODIMP CBenefit::put_Minimum(long lMinimum) 
{
	m_lMinimum = lMinimum;
	return (NOERROR);
}

STDMETHODIMP CBenefit::get_Multiples(short FAR * piMultiples) 
{
	if (m_bExpanded)
		*piMultiples = m_iMultiples;
	else
	{
		if (Expand())
			*piMultiples = m_iMultiples;
		else
		{
			Term();
			return (FlagException(IID_IBenefit, scBnfExpand, icErrExpand));
		}
	}
	return (NOERROR);
}

STDMETHODIMP CBenefit::put_Multiples(short iMultiples) 
{
	m_iMultiples = iMultiples;
	return (NOERROR);
}

STDMETHODIMP CBenefit::Term() 
{
	Clear();
	return (NOERROR);
}

void CBenefit::Clear()
{
	m_bExpanded	= FALSE;
	m_iIncrement= 0;
	m_iMultiples= 0;
	m_lBnfTaxId	= 0;
	m_lBnfTypeId= 0;
	m_lCost		= 0;
	m_lMaximum	= 0;
	m_lMinimum	= 0;

	EmptyBSTR(m_bstrDescription);
}


//CBenefit::Expand
//Purpose     - Manages the retrieval of private data
//Outputs     - (BOOL) Success
BOOL CBenefit::Expand()
{
	IList * pColl;
	_bstr_t sSQL;
	CStrFnc sTemp;
	long	lValue;
	int		iValue;
	VARIANT vValue;
	BOOL	bUseODBC;
	BOOL	bSuccess;
	HRESULT	hResult;

	hResult		= NOERROR;
	lValue		= 0;
	iValue		= 0;
	bUseODBC	= FALSE;
	m_bExpanded = FALSE;

	if (m_lPKId)
	{ 
		Clear();
		if (m_poAdmin)
		{
			hResult = ::CoCreateInstance(CLSID_CList, NULL, THISCLSCTX, IID_IList, (void **)&pColl);
			if (hResult == NOERROR)
			{
				hResult = m_poAdmin->get_UseODBC(&bUseODBC);
				if (hResult != NOERROR)
				{
					pColl->Release();
					return (false);
				}
   
				//Load either ODBC or SQL query string from the resources
				sSQL = CStrFnc::LoadResString(bUseODBC ? scGet_Benefit_SQL: scGet_Benefit_MDB);
				sSQL += TEXT(" ");

				//add the primary key id to the end of the string
				sTemp.Format(TEXT("%ld"), m_lPKId);
				sSQL += sTemp.GetBstr();

				//Fill the List by executing the SQL statement
				hResult = m_poAdmin->ExecFillList(sSQL, &pColl, &bSuccess);
				if (hResult == NOERROR)
				{
					if (bSuccess)
					{
						//Get the benefits tax id
						::VariantInit(&vValue);
						hResult = pColl->get_Item(TEXT("BnfTaxId"), &vValue);
						m_lBnfTaxId = 0;
						if (hResult == NOERROR)
							m_lBnfTaxId = CStrFnc::VariantBstrToLong(vValue);

						//Get the benefits type id
						::VariantClear(&vValue);
						hResult = pColl->get_Item(TEXT("BnfTypeId"), &vValue);
						m_lBnfTypeId = 0;
						if (hResult == NOERROR)
							m_lBnfTypeId = CStrFnc::VariantBstrToLong(vValue);

						//Get the cost
						::VariantClear(&vValue);
						hResult = pColl->get_Item(TEXT("Cost"), &vValue);
						m_lCost = 0;
						if (hResult == NOERROR)
							m_lCost = CStrFnc::VariantBstrToLong(vValue);
						
						//Get the description
						::VariantClear(&vValue);
						hResult = pColl->get_Item(TEXT("Description"), &vValue);
						if (hResult == NOERROR)
							::SysReAllocString(&m_bstrDescription, V_BSTR(&vValue));
						
						//Get the increment value
						::VariantClear(&vValue);
						hResult = pColl->get_Item(TEXT("Increment"), &vValue);
						m_iIncrement = 0;
						if (hResult == NOERROR)
							m_iIncrement = CStrFnc::VariantBstrToLong(vValue);
		
						//Get the maximum value
						::VariantClear(&vValue);
						hResult = pColl->get_Item(TEXT("Maximum"), &vValue);
						m_lMaximum = 0;
						if (hResult == NOERROR)
							m_lMaximum = CStrFnc::VariantBstrToLong(vValue);
						
						//Get the minimum value
						::VariantClear(&vValue);
						hResult = pColl->get_Item(TEXT("Minimum"), &vValue);
						m_lMinimum = 0;
						if (hResult == NOERROR)
							m_lMinimum = CStrFnc::VariantBstrToLong(vValue);
						
						//Get the multiples
						::VariantClear(&vValue);
						hResult = pColl->get_Item(TEXT("Multiples"), &vValue);
						m_iMultiples = 0;
						if (hResult == NOERROR)
							m_iMultiples = CStrFnc::VariantBstrToLong(vValue);

						m_bExpanded = TRUE;
					}
				}
				pColl->Release();
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
CBenefit::CBenefit(	LPUNKNOWN pUnkOuter, PFNDESTROYED pfnDestroy)
				: m_SupportErrorInfo(this, IID_IBenefit)
#pragma warning (default : 4355)
{
	m_oid		= g_oid++;
	m_cRef		= 0;
	m_pUnkOuter = pUnkOuter;
	m_pfnDestroy= pfnDestroy;
	m_bExpanded = FALSE;
	m_iIncrement= 0;
	m_iMultiples= 0;
	m_lBnfTaxId = 0;
	m_lBnfTypeId= 0;
	m_lCost		= 0;
	m_lMaximum	= 0;
	m_lMinimum	= 0;
	m_bstrDescription = NULL;
	m_poAdmin	= NULL;
	m_lPKId		= 0;
	m_pITypeInfo= NULL;

	return;
}

CBenefit::~CBenefit(void)
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

BOOL CBenefit::CFInit(void)
{
	LPUNKNOWN pIUnknown;

	pIUnknown = this;

	if (m_pUnkOuter)
		pIUnknown = m_pUnkOuter;

	return (TRUE);
}
////////////////////////////////////////////
// Standard IUnknown method implementations
////////////////////////////////////////////

//CBenefit::QueryInterface
//
//Parameters:
//riid	-	input interface reference id
//ppv	-	return CBenefit interface pointer
STDMETHODIMP CBenefit::QueryInterface(REFIID riid, void ** ppv)
{
	*ppv = NULL;

	if (riid == IID_IBenefit || riid == IID_IUnknown || riid == IID_IDispatch)
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

STDMETHODIMP_(ULONG) CBenefit::AddRef(void)
{
	++m_cRef;
	Audit(m_oid, 1, TEXT("CBenefit"), m_cRef);
	return (m_cRef);
}

STDMETHODIMP_(ULONG) CBenefit::Release(void)
{
	--m_cRef;
	Audit(m_oid, 2, TEXT("CBenefit"), m_cRef);

	if (m_cRef == 0)
	{
		if (m_pfnDestroy)
			(*m_pfnDestroy)();
		Audit(m_oid, 0, TEXT("CBenefit gbldec"), g_lObjectCount);
		delete this;
		return (0);
	}
	else
		return (m_cRef);
}

//////////////////////////////////////////////
// Standard IDispatch method implementations
//////////////////////////////////////////////

STDMETHODIMP CBenefit::GetTypeInfoCount(UINT *pctInfo)
{
	*pctInfo = 1;

	return (NOERROR);
}

STDMETHODIMP CBenefit::GetTypeInfo(UINT itInfo, LCID lcid, ITypeInfo **ppITypeInfo)
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

		hResult = pITypeLib->GetTypeInfoOfGuid(IID_IBenefit, ppITI);
		
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

STDMETHODIMP CBenefit::GetIDsOfNames(REFIID riid, OLECHAR **rgszNames,
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

STDMETHODIMP CBenefit::Invoke(DISPID dispID,	REFIID riid, LCID lcid,
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

CBenefitFactory::CBenefitFactory(void)
{
	m_oid	= g_oid++;
	m_cRef	= 0L;
	return;
}

CBenefitFactory::~CBenefitFactory(void)
{
	return;
}

//CBenefitFactory::QueryInterface
//Purpose     - Used to return an interface pointer from the requested interface id.
//Inputs      - (REFIID)  riid
//              (void **) ppvCBenefIf
STDMETHODIMP CBenefitFactory::QueryInterface(REFIID riid, void ** ppvCBenefitIf)
{
	*ppvCBenefitIf = NULL;

	if (riid == IID_IUnknown ||
		riid == IID_IClassFactory)
		*ppvCBenefitIf = this;

	if (*ppvCBenefitIf)
	{
		((LPUNKNOWN)*ppvCBenefitIf)->AddRef();
		return (NOERROR);
	}

	return (E_NOINTERFACE);
}

STDMETHODIMP_(ULONG) CBenefitFactory::AddRef(void)
{
	++m_cRef;
	Audit(m_oid, 1, TEXT("CBenefitF"), m_cRef);
	return (m_cRef);
}

STDMETHODIMP_(ULONG) CBenefitFactory::Release(void)
{
	--m_cRef;
	Audit(m_oid, 2, TEXT("CBenefitF"), m_cRef);
	if (m_cRef == 0)
	{
		delete this;
		return (0);
	}
	else
		return (m_cRef);
}

//CBenefitFactory::CreateInstance
//Purpose     - Creates a CBenefit object and returns an Interface pointer
//Inputs      - (LPUNKNOWN) pUnkOuter
//              (REFIID)    riid
//				(void **)   ppvCBenefIf
STDMETHODIMP CBenefitFactory::CreateInstance(LPUNKNOWN pUnkOuter, REFIID riid, void ** ppvCBenefIf)
{
	PCBenefit	pCBenefit;
	HRESULT		hResult;

	*ppvCBenefIf = NULL;
	hResult		 = NOERROR;

	if (pUnkOuter && riid != IID_IUnknown)
		return (CLASS_E_NOAGGREGATION);

	pCBenefit = new CBenefit(pUnkOuter, ObjectDestroyed);

	if (pCBenefit == NULL)
		return (hResult);

	if (pCBenefit->CFInit())
		hResult = pCBenefit->QueryInterface(riid, ppvCBenefIf);

	if (hResult != NOERROR)
		delete pCBenefit;
	else
	{
		g_lObjectCount++;
		Audit(m_oid, 0, TEXT("CBenefit gblinc"), g_lObjectCount);
	}

	return (hResult);
}

//CBenefitFactory::LockServer
//Purpose     - Increments/Decrements the lock count to prevent class factory server shutdown.
//Inputs      - (BOOL)    fLock
STDMETHODIMP CBenefitFactory::LockServer(BOOL fLock)
{
	if (fLock)
		g_lLockedObjects++;
	else
		g_lLockedObjects--;

	return (NOERROR);
}

