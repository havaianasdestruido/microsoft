//clistbenefitlink.cpp  - Copyright (C) 1996 Microsoft Corporation
//            - All rights reserved.
//Date        - 10/19/96
//Authors     - Ken Bergmann, Robert Coleridge
//Purpose     - Implementation file for the CListBenefitLink object.

#pragma warning (disable : 4800)	// turns off bool warning

#include "clistbenefitlink.h"
#include "cienum.h"
#include "ckeyll.h"
#include "clist.h"
#include "cstrfunc.h"

extern long g_oid;

//Globals used to determine when to unload a In-Proc or Local Server.
extern ULONG       g_lObjectCount;
extern ULONG       g_lLockedObjects;

STDMETHODIMP CListBenefitLink::Add(IBenefitLink * pvItem)
{
	long lIndex;
	CKeyedLinkedListItem * pPrevItem;
	CKeyedLinkedListItem * pItem;
	BSTR sKey;
	VARIANT v;

	if (pvItem == NULL)
		return (E_POINTER);

	pvItem->get_PKId(&lIndex);
	sKey = CList::GenListKey(lIndex);

	lIndex = m_pLL->Find(sKey, &pPrevItem, &pItem);
	if (lIndex != LIST_ERROR)
		return (FlagException(IID_IListBenefitLink, icErrBadIndex, icErrBadIndex));

	::VariantInit(&v);
	v.vt = VT_DISPATCH;
	V_DISPATCH(&v) = pvItem;
	lIndex = m_pLL->Insert(sKey, &v);

	if (lIndex == LIST_ERROR)
		return (FlagException(IID_IListBenefitLink, icErrBadIndex, icErrBadIndex));

	return (NOERROR);
}

STDMETHODIMP CListBenefitLink::get_Count(long FAR* plCount)
{
	*plCount = m_pLL->GetCount();
	return (NOERROR);
}

STDMETHODIMP CListBenefitLink::put_Item(IBenefitLink * newVal)
{
	long lIndex;
	BSTR sKey;
	VARIANT v;
	
	newVal->get_PKId(&lIndex);
	sKey = CList::GenListKey(lIndex);

	::VariantInit(&v);
	v.vt = VT_DISPATCH;
	V_DISPATCH(&v) = newVal;
	m_pLL->Update(sKey, &v);

	return (NOERROR);
}

STDMETHODIMP CListBenefitLink::get_Item(long lKey, IBenefitLink ** pvItem)
{ 
	BSTR sKey;
	CKeyedLinkedListItem * pItem;
	long lIndex;
	
	sKey = CList::GenListKey(lKey);
	lIndex = m_pLL->Find(sKey, NULL, NULL);
	if (lIndex <= LIST_ERROR)
		return (FlagException(IID_IListBenefitLink, icErrBadIndex, icErrBadIndex));

	if (m_pLL)
	{
		if (lIndex >= m_pLL->GetCount())
			return (FlagException(IID_IListBenefitLink, icErrBadIndex, icErrBadIndex));
	}

	if (!m_pLL->Get(lIndex, &pItem))
		return (FlagException(IID_IListBenefitLink, icErrBadIndex, icErrBadIndex));

	*pvItem = (IBenefitLink *)V_DISPATCH(pItem->GetData());
	(*pvItem)->AddRef();

	return (NOERROR);
}

//CListBenefitLink::get_NewEnum
//Purpose:
//Returns an enumerator (IEnumVARIANT) for the items curently in the list. 
//The NewEnum property is restricted and so is invisible to users of an 
//automation controller's scripting language. Automation controllers that support
//a 'For Each' statement to iterate through the elements of a list will use
//the enumerator returned by NewEnum. The enumerator creates a snapshot of the
//the current state of the list.

STDMETHODIMP CListBenefitLink::get__NewEnum(IUnknown ** ppIEnumVariantIUnk)
{
	CIEnumVariant * pCIEnumVariant;
	HRESULT hResult;

	pCIEnumVariant	= NULL;
	hResult			= NOERROR;
	*ppIEnumVariantIUnk = NULL;

	hResult = CIEnumVariant::Create(m_pLL, m_pLL->GetCount(), &pCIEnumVariant);
	if (hResult != NOERROR)
	{
		FlagException(IID_IListBenefitLink, icErrBadIndex, icErrBadIndex);
		goto error;
	}

	hResult = pCIEnumVariant->QueryInterface(IID_IUnknown, (VOID **)ppIEnumVariantIUnk);    
	if (hResult != NOERROR) 
	{
		FlagException(IID_IListBenefitLink, icErrBadIndex, icErrBadIndex);
		goto error;
	}
	
	return (NOERROR); 

error:
	if (pCIEnumVariant)
		delete pCIEnumVariant;   
	return (hResult);
}

STDMETHODIMP CListBenefitLink::Remove(long lKey)
{
	BSTR sKey;
	long lIndex;

	sKey = CList::GenListKey(lKey);
	lIndex = m_pLL->Find(sKey, NULL, NULL);

	if (lIndex <= LIST_ERROR)
		return (FlagException(IID_IListBenefitLink, icErrBadIndex, icErrBadIndex));

	if (m_pLL)
	{
		if (lIndex >= m_pLL->GetCount())
			return (FlagException(IID_IListBenefitLink, icErrBadIndex, icErrBadIndex));
	}

	m_pLL->Delete(sKey);

    return (NOERROR);
}

//Turn off warning 4355
#pragma warning (disable : 4355)
CListBenefitLink::CListBenefitLink(	LPUNKNOWN pUnkOuter, PFNDESTROYED pfnDestroy)
				: m_SupportErrorInfo(this, IID_IListBenefitLink)
#pragma warning (default : 4355)
{
	m_oid = g_oid++;
	m_cRef = 0;
	m_pUnkOuter = pUnkOuter;
	m_pfnDestroy = pfnDestroy;

	m_pLL = new CKeyedLinkedList();

	m_pITypeInfo = NULL;
	return;
}

CListBenefitLink::~CListBenefitLink(void)
{
	if (m_pLL)
		delete m_pLL;

	if (m_pITypeInfo)
	{
		m_pITypeInfo->Release();
		m_pITypeInfo = NULL;
	}

	return;
}

//CListBenefitLink::CFInit
//Purpose:
//Performs any intiailization of a CBase that's prone to failure
//that we also use internally before exposing the object outside.
//Called by class factory
BOOL CListBenefitLink::CFInit(void)
{
	LPUNKNOWN pIUnknown;

	pIUnknown = this;
	if (m_pUnkOuter)
		pIUnknown = m_pUnkOuter;

	return (TRUE);
}

//CListBenefitLink::QueryInterface
//Parameters:
//riid		-	input interface reference id
//ppvCListBenefLinkIf	-	return CBank interface pointer

STDMETHODIMP CListBenefitLink::QueryInterface(REFIID riid, void ** ppvCListBenefLinkIf)
{
	*ppvCListBenefLinkIf = NULL;

	if (riid == IID_IListBenefitLink || riid == IID_IUnknown || riid == IID_IDispatch)
		*ppvCListBenefLinkIf = this;

	if (riid == IID_ISupportErrorInfo)
		*ppvCListBenefLinkIf = &m_SupportErrorInfo;

	if (*ppvCListBenefLinkIf)
	{
		((LPUNKNOWN)*ppvCListBenefLinkIf)->AddRef();
		return (NOERROR);
	}

	return (E_NOINTERFACE);
}

STDMETHODIMP_(ULONG) CListBenefitLink::AddRef(void)
{
	++m_cRef;
	Audit(m_oid, 1, TEXT("CListBenefitLink"), m_cRef);
	return (m_cRef);
}

STDMETHODIMP_(ULONG) CListBenefitLink::Release(void)
{
	--m_cRef;
	Audit(m_oid, 2, TEXT("CListBenefitLink"), m_cRef);
	if (m_cRef == 0)
	{
		if (m_pfnDestroy)
			(*m_pfnDestroy)();

		Audit(m_oid, 0, TEXT("CListBenefitLink gbldec"), g_lObjectCount);
		delete this;
		return (0);
	}
	else
		return (m_cRef);
}

STDMETHODIMP CListBenefitLink::GetTypeInfoCount(UINT *pctInfo)
{
	*pctInfo = 1;	// we only have 1 typelib to return info on
	return (NOERROR);
}

STDMETHODIMP CListBenefitLink::GetTypeInfo(UINT itInfo, LCID lcid, ITypeInfo **ppITypeInfo)
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

		hResult = pITypeLib->GetTypeInfoOfGuid(IID_IListBenefitLink, ppITI);
		
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

STDMETHODIMP CListBenefitLink::GetIDsOfNames(REFIID riid, OLECHAR **rgszNames, UINT cNames, LCID lcid, DISPID *rgDispID)
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

STDMETHODIMP CListBenefitLink::Invoke
(
	DISPID dispID, 
	REFIID riid,
    LCID lcid, 
	unsigned short wFlags, 
	DISPPARAMS *pDispParams,
    VARIANT *pVarResult, 
	EXCEPINFO *pExcepInfo, 
	UINT *puArgErr)
{
	HRESULT     hResult;
	ITypeInfo  *pTI;
	LANGID      langID;

	hResult	= NOERROR;
	langID = PRIMARYLANGID(lcid);

	if (riid != IID_NULL)
		return (DISP_E_UNKNOWNINTERFACE);

	hResult = GetTypeInfo(0, lcid, &pTI);

	if (hResult != NOERROR)
		return (hResult);

	hResult = DispInvoke(this, pTI, dispID, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr);

	pTI->Release();
	return (hResult);
}

CListBenefitLinkFactory::CListBenefitLinkFactory(void)
{
	m_oid = g_oid++;
	m_cRef = 0L;
	return;
}

CListBenefitLinkFactory::~CListBenefitLinkFactory(void)
{
	return;
}

//CListBenefitLinkFactory::QueryInterface
//Purpose:
//Used to return an interface pointer from the
//requested interface id.
//Parameters:
//riid					-	Input reference interface id.
//ppvCListBenefLinkCFIf	-	Returned interface pointer.
STDMETHODIMP CListBenefitLinkFactory::QueryInterface(REFIID riid, void ** ppvCListBenefLinkCFIf)
{
	*ppvCListBenefLinkCFIf = NULL;

	if (riid == IID_IUnknown || riid == IID_IClassFactory)
		*ppvCListBenefLinkCFIf = this;

	if (*ppvCListBenefLinkCFIf)
	{
		((LPUNKNOWN)*ppvCListBenefLinkCFIf)->AddRef();
		return (NOERROR);
	}

	return (E_NOINTERFACE);
}

//CListBenefitLinkFactory::AddRef
//Purpose:
//Increments class factory reference count.
STDMETHODIMP_(ULONG) CListBenefitLinkFactory::AddRef(void)
{
	++m_cRef;
	Audit(m_oid, 1, TEXT("CListBenefitLinkF"), m_cRef);
	return (m_cRef);
}

//CListBenefitLinkFactory::Release
//Purpose:
//Decrements class factory reference count. Deletes
//the object when reference count reaches zero.
STDMETHODIMP_(ULONG) CListBenefitLinkFactory::Release(void)
{
	--m_cRef;
	Audit(m_oid, 2, TEXT("CListBenefitLinkF"), m_cRef);
	if (m_cRef == 0)
	{
		delete this;
		return (0);
	}
	else
		return (m_cRef);
}

//CListBenefitLinkFactory::CreateInstance
//Purpose:
//Creates a CListBenefitLink object and returns an Interface pointer
//for the referenced interface id (riid).
//Parameters:
//pUnkOuter			-	Input outer unknown pointer or NULL.
//riid				-	Input reference interface id.
//ppvCListBenefLinkIf	-	Return pointer to a CListBenefitLink interface.
STDMETHODIMP CListBenefitLinkFactory::CreateInstance(LPUNKNOWN pUnkOuter, REFIID riid, void ** ppvCListBenefLinkIf)
{
	PCListBenefitLink   pCListBenefitLink;
	HRESULT             hResult;

	hResult = NOERROR;
	*ppvCListBenefLinkIf = NULL;

	if (pUnkOuter && riid != IID_IUnknown)
		return (CLASS_E_NOAGGREGATION);

	pCListBenefitLink = new CListBenefitLink(pUnkOuter, ObjectDestroyed);

	if (pCListBenefitLink == NULL)
		return (hResult);

	if (pCListBenefitLink->CFInit())
		hResult = pCListBenefitLink->QueryInterface(riid, ppvCListBenefLinkIf);

	if (hResult != NOERROR)
		delete pCListBenefitLink;
	else
	{
		g_lObjectCount++;
		Audit(m_oid, 0, TEXT("CListBenefitLink gblinc"), g_lObjectCount);
	}

	return (hResult);
}

//CListBenefitLinkFactory::LockServer
//Purpose:
//Increments/Decrements the lock count to prevent class
//factory server shutdown.
//Parameters:
//fLock	-	TRUE implies addref this server. FALSE
//			implies release this server.

STDMETHODIMP CListBenefitLinkFactory::LockServer(BOOL fLock)
{
	if (fLock)
		g_lLockedObjects++;
	else
		g_lLockedObjects--;

	return (NOERROR);
}
