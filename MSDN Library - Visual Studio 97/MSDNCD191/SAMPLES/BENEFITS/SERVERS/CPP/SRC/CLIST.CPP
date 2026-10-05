//clist.cpp  - Copyright (C) 1996 Microsoft Corporation
//            - All rights reserved.
//Date        - 10/19/96
//Authors     - Ken Bergmann, Robert Coleridge
//Purpose     - Implementation file for the Clist object.

#pragma warning (disable : 4800)	// turns off bool warning

#include "clist.h"
#include "cienum.h"
#include "ckeyll.h"
#include "cstrfunc.h"

extern long g_oid;

//Globals used to determine when to unload a In-Proc or Local Server.
extern ULONG       g_lObjectCount;
extern ULONG       g_lLockedObjects;

STDMETHODIMP CList::Add(BSTR bstrKey,VARIANT FAR* pvItem)
{
	long lIndex;
	CKeyedLinkedListItem * pPrevItem;
	CKeyedLinkedListItem * pItem;

	if (pvItem == NULL)
		return (E_POINTER);

	lIndex = m_pLL->Find(bstrKey, &pPrevItem, &pItem);
	if (lIndex != LIST_ERROR)
		return (FlagException(IID_IList, icErrBadIndex, icErrBadIndex));

	lIndex = m_pLL->Insert(bstrKey, pvItem);
	if (lIndex == LIST_ERROR)
		return (FlagException(IID_IList, icErrBadIndex, icErrBadIndex));

	return (NOERROR);
}

STDMETHODIMP CList::get_Count(long FAR* plCount)
{
	*plCount = m_pLL->GetCount();
	return (NOERROR);
}

STDMETHODIMP CList::put_Item(BSTR bstrKey, VARIANT newVal)
{
	m_pLL->Update(bstrKey, &newVal);
	return (NOERROR);
}

STDMETHODIMP CList::get_Item(BSTR bstrKey, VARIANT FAR* pvItem)
	{ 
	CKeyedLinkedListItem * pItem;
	long lIndex;
	
	lIndex = m_pLL->Find(bstrKey, NULL, NULL);
	if (lIndex <= LIST_ERROR)
		return (FlagException(IID_IList, icErrBadIndex, icErrBadIndex));

	if (m_pLL)
	{
		if (lIndex >= m_pLL->GetCount())
			return (FlagException(IID_IList, icErrBadIndex, icErrBadIndex));
	}

	if (!m_pLL->Get(lIndex, &pItem))
		return (FlagException(IID_IList, icErrBadIndex, icErrBadIndex));

	::VariantInit(pvItem);
	::VariantCopyInd(pvItem, pItem->GetData());

	return (NOERROR);
}

//CList::get_NewEnum
//Purpose:
//Returns an enumerator (IEnumVARIANT) for the items curently in the list. 
//The NewEnum property is restricted and so is invisible to users of an 
//automation controller's scripting language. Automation controllers that support
//a 'For Each' statement to iterate through the elements of a list will use
//the enumerator returned by NewEnum. The enumerator creates a snapshot of the
//the current state of the list.

STDMETHODIMP CList::get__NewEnum(IUnknown ** ppIEnumVariantIUnk)
{
	CIEnumVariant * pCIEnumVariant;
	HRESULT hResult;

	pCIEnumVariant		= NULL;
	hResult				= NOERROR;
	*ppIEnumVariantIUnk = NULL;

	hResult = CIEnumVariant::Create(m_pLL, m_pLL->GetCount(), &pCIEnumVariant);
	if (hResult != NOERROR)
	{
		FlagException(IID_IList, icErrBadIndex, icErrBadIndex);
		goto error;
	}

	hResult = pCIEnumVariant->QueryInterface(IID_IUnknown, (VOID **)ppIEnumVariantIUnk);    
	if (hResult != NOERROR)
	{
		FlagException(IID_IList, icErrBadIndex, icErrBadIndex);
		goto error;
	}
	
	return (NOERROR); 

error:
	if (pCIEnumVariant)
		delete pCIEnumVariant;   
	return (hResult);
}

STDMETHODIMP CList::Remove(BSTR bstrKey)
{
	long lIndex;

	lIndex = m_pLL->Find(bstrKey, NULL, NULL);

	if (lIndex <= LIST_ERROR)
		return (FlagException(IID_IList, icErrBadIndex, icErrBadIndex));

	if (m_pLL)
	{
		if (lIndex >= m_pLL->GetCount())
			return (FlagException(IID_IList, icErrBadIndex, icErrBadIndex));
	}

	m_pLL->Delete(bstrKey);

    return (NOERROR);
}

//Turn off warning 4355
#pragma warning (disable : 4355)
CList::CList(	LPUNKNOWN pUnkOuter, PFNDESTROYED pfnDestroy)
				: m_SupportErrorInfo(this, IID_IList)
#pragma warning (default : 4355)
{
	m_oid			= g_oid++;
	m_cRef			= 0;
	m_pUnkOuter		= pUnkOuter;
	m_pfnDestroy	= pfnDestroy;

	m_pLL = new CKeyedLinkedList();

	m_pITypeInfo = NULL;
	return;
}

CList::~CList(void)
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

//CList::CFInit
//Purpose:
//Performs any intiailization of a CBase that's prone to failure
//that we also use internally before exposing the object outside.
//Called by class factory
BOOL CList::CFInit(void)
{
	LPUNKNOWN pIUnknown;

	pIUnknown = this;
	if (m_pUnkOuter)
		pIUnknown = m_pUnkOuter;

	return (TRUE);
}

BSTR CList::GenListKey(long lIndex)
{
	CStrFnc sKey;
	sKey.Format(TEXT("ID=%ld"), lIndex);
	return (sKey.GetBstr());
}

//CList::QueryInterface
//Parameters:
//riid			-	input interface reference id
//ppvCListIf	-	return CList interface pointer
STDMETHODIMP CList::QueryInterface(REFIID riid, void ** ppvCListIf)
{
	*ppvCListIf = NULL;

	if (riid == IID_IList || riid == IID_IUnknown || riid == IID_IDispatch)
		*ppvCListIf = this;

	if (riid == IID_ISupportErrorInfo)
		*ppvCListIf = &m_SupportErrorInfo;

	if (*ppvCListIf)
	{
		((LPUNKNOWN)*ppvCListIf)->AddRef();
		return (NOERROR);
	}

	return (E_NOINTERFACE);
}

STDMETHODIMP_(ULONG) CList::AddRef(void)
{
	++m_cRef;
	Audit(m_oid, 1, TEXT("CList"), m_cRef);
	return (m_cRef);
}

STDMETHODIMP_(ULONG) CList::Release(void)
{
	--m_cRef;
	Audit(m_oid, 2, TEXT("CList"), m_cRef);
	if (m_cRef == 0)
	{
		if (m_pfnDestroy)
			(*m_pfnDestroy)();

		Audit(m_oid, 0, TEXT("CList gbldec"), g_lObjectCount);
		delete this;
		return (0);
	}
	else
		return (m_cRef);
}

STDMETHODIMP CList::GetTypeInfoCount(UINT *pctInfo)
{
	*pctInfo = 1;	// we only have 1 typelib to return info on
	return (NOERROR);
}

STDMETHODIMP CList::GetTypeInfo(UINT itInfo, LCID lcid, ITypeInfo **ppITypeInfo)
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
			return DISP_E_UNKNOWNLCID;
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

		hResult = pITypeLib->GetTypeInfoOfGuid(IID_IList, ppITI);

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

STDMETHODIMP CList::GetIDsOfNames(REFIID riid, OLECHAR **rgszNames, UINT cNames, LCID lcid, DISPID *rgDispID)
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

STDMETHODIMP CList::Invoke
(
	DISPID dispID, 
	REFIID riid,
    LCID lcid, 
	unsigned short wFlags, 
	DISPPARAMS *pDispParams,
    VARIANT *pVarResult, 
	EXCEPINFO *pExcepInfo, 
	UINT *puArgErr
){
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

CListFactory::CListFactory(void)
{
	m_oid = g_oid++;
	m_cRef = 0L;
	return;
}

CListFactory::~CListFactory(void)
{
	return;
}

//CListFactory::QueryInterface
//Purpose:
//Used to return an interface pointer from the
//requested interface id.
//Parameters:
//riid			-	Input reference interface id.
//ppvCListCFIf	-	Returned interface pointer.
STDMETHODIMP CListFactory::QueryInterface(REFIID riid, void ** ppvCListCFIf)
{
	*ppvCListCFIf = NULL;

	if (riid == IID_IUnknown ||	riid == IID_IClassFactory)
		*ppvCListCFIf = this;

	if (*ppvCListCFIf)
	{
		((LPUNKNOWN)*ppvCListCFIf)->AddRef();
		return (NOERROR);
	}

	return (E_NOINTERFACE);
}

//CListFactory::AddRef
//Purpose:Increments class factory reference count.
STDMETHODIMP_(ULONG) CListFactory::AddRef(void)
{
	++m_cRef;
	Audit(m_oid, 1, TEXT("CListF"), m_cRef);
	return (m_cRef);
}

//CListFactory::Release
//Purpose:
//Decrements class factory reference count. Deletes
//the object when reference count reaches zero.
STDMETHODIMP_(ULONG) CListFactory::Release(void)
{
	--m_cRef;
	Audit(m_oid, 2, TEXT("CListF"), m_cRef);
	if (m_cRef == 0)
	{
		delete this;
		return (0);
	}
	else
		return (m_cRef);
}

//CListFactory::CreateInstance
//Purpose:
//Creates a CList object and returns an Interface pointer
//for the referenced interface id (riid).
//Parameters:
//pUnkOuter		-	Input outer unknown pointer or NULL.
//riid			-	Input reference interface id.
//ppvCListIf	-	Return pointer to a CList interfce.
STDMETHODIMP CListFactory::CreateInstance(LPUNKNOWN pUnkOuter, REFIID riid, void ** ppvCListIf)
{
	PCList            pCList;
	HRESULT           hResult;

	hResult		= NOERROR;
	*ppvCListIf = NULL;

	if (pUnkOuter && riid != IID_IUnknown)
		return (CLASS_E_NOAGGREGATION);

	pCList = new CList(pUnkOuter, ObjectDestroyed);

	if (pCList == NULL)
		return (hResult);

	if (pCList->CFInit())
		hResult = pCList->QueryInterface(riid, ppvCListIf);

	if (hResult != NOERROR)
		delete pCList;
	else
	{
		g_lObjectCount++;
		Audit(m_oid, 0, TEXT("CList gblinc"), g_lObjectCount);
	}

	return (hResult);
}

//CListFactory::LockServer
//Purpose:
//Increments/Decrements the lock count to prevent class
//factory server shutdown.
//Parameters:
//fLock	-	TRUE implies addref this server. FALSE
//			implies release this server.
STDMETHODIMP CListFactory::LockServer(BOOL fLock)
{
	if (fLock)
		g_lLockedObjects++;
	else
		g_lLockedObjects--;

	return (NOERROR);
}
