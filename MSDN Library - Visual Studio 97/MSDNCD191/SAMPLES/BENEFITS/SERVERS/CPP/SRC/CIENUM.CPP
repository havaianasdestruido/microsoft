//CIEnum.cpp  - Copyright (C) 1996 Microsoft Corporation
//            - All rights reserved.
//Date        - 10/19/96
//Authors     - Ken Bergmann, Robert Coleridge
//Purpose     - Implementation file for IEnumVariant enumerator object

#pragma warning (disable : 4800)	// turns off bool warning

#include "support.h"
#include "cienum.h"

extern long g_oid;

//Create
//Purpose	- Creates an instance of the IEnumVARIANT enumerator object and initializes it.
//Inputs	- (CKeyedLinkedList *) pLL
//			  (ULONG) lElements 
//Outputs	- (CIEnumVariant **) ppCIEnumVariant
HRESULT CIEnumVariant::Create(CKeyedLinkedList * pLL, ULONG lElements, CIEnumVariant ** ppCIEnumVariant)
{   
    HRESULT					hResult;
    CIEnumVariant *			pCIEnumVariant;
	CKeyedLinkedListItem *	pItem;
                      
    hResult			= NOERROR;
    pCIEnumVariant	= NULL;
    *ppCIEnumVariant= NULL;
    
    pCIEnumVariant = new CIEnumVariant();
    if (pCIEnumVariant == NULL)
        goto error; 
        
    pCIEnumVariant->m_lRef = 0;
    
	if (pLL)
	{
		pItem = pLL->GetNext();
		pItem = pLL->GetNext(pItem);
		while (pItem)
		{
			pCIEnumVariant->m_pLL->Insert(pItem->GetKey(), pItem->GetData());
			pCIEnumVariant->m_lElements++;
			pItem = pLL->GetNext(pItem);
		}
	}

    *ppCIEnumVariant = pCIEnumVariant;
    return (NOERROR);

error:;
    if (pCIEnumVariant == NULL)
        return (E_OUTOFMEMORY);   
                              
    delete pCIEnumVariant;
    return (hResult);
}

//CIEnumVariant
//Purpose     - Constructor for CIEnumVariant object. Initializes members to NULL.
CIEnumVariant::CIEnumVariant()
{    
	m_oid		= g_oid++;
    m_pLL		= new CKeyedLinkedList;
	m_lLBound	= 0;
	m_lElements	= 0;
	m_lCurrent	= 0;
}

//~CIEnumVariant
//Purpose     - Destructor for CIEnumVariant object.
CIEnumVariant::~CIEnumVariant()
{                   
    if (m_pLL)
		delete m_pLL;
}

//QueryInterface
//Purpose     - Returns an interface pointer for the specified IID.
//Inputs      - (REFIID) riid				- interface reference id
//Outputs     - (void**) ppvCIEnumVariantIf	- CIEnumVariant interface pointer
STDMETHODIMP CIEnumVariant::QueryInterface(REFIID iid, void ** ppvCIEnumVariantIf) 
{   
    *ppvCIEnumVariantIf = NULL;
        
    if (iid == IID_IUnknown || iid == IID_IEnumVARIANT) 
        *ppvCIEnumVariantIf = this;     
    else
		return (E_NOINTERFACE); 

    AddRef();
    return (NOERROR);    
}

//AddRef
//Purpose     - Increments "this" objects reference count
STDMETHODIMP_(ULONG) CIEnumVariant::AddRef(void)
{
	++m_lRef;
	Audit(m_oid, 1, TEXT("CIEnum"), m_lRef);
    return (m_lRef);
}

//Release
//Purpose     - Decrements "this" objects reference count.
STDMETHODIMP_(ULONG) CIEnumVariant::Release(void)
{
	--m_lRef;
	Audit(m_oid, 2, TEXT("CIEnum"), m_lRef);
    if (m_lRef == 0)
	{
        delete this;
		return (0);
	}
	else
		return (m_lRef);
}

//Next
//Purpose     - Retrieves a list of the next "lElements" elements.
//Inputs	- (ULONG) lElements
//			- (VARIANT *)	pValue
//Outputs	- (ULONG)		plElementFetched  	
//			- (VARIANT *)	pValue
STDMETHODIMP CIEnumVariant::Next(ULONG lElements, VARIANT * pValue, ULONG * plElementFetched)
{ 
	ASSERT(pValue != NULL);

	HRESULT	hResult;
	ULONG	lIndex;
	ULONG	lItemCount;
	long	lPosInList;
	CKeyedLinkedListItem * pItem;

	hResult		= NOERROR;
	lIndex		= 0;
	lItemCount	= 0;
	lPosInList	= 0;
	pItem		= NULL;

	if (plElementFetched)
		*plElementFetched = 0;

	// Retrieve the next lElements elements.
	for (
		lPosInList = m_lCurrent, lItemCount = 0, lIndex = 0;
		lPosInList < (long)(m_lLBound + m_lElements) && lItemCount < lElements;
		lPosInList++, lItemCount++,	lIndex++
		)
	{
		::VariantInit(&pValue[lIndex]);
		m_pLL->Get(lPosInList, &pItem);
		::VariantCopyInd(&pValue[lIndex], pItem->GetData());
	}

	if (plElementFetched)
		*plElementFetched = lItemCount;

	m_lCurrent = lPosInList;

	if (lItemCount)
		return (NOERROR);
	else
		return (S_FALSE);
}

//Skip
//Purpose	- Skips the next lElements elements. Implements IEnumVARIANT::Skip.
//Inputs	- (ULONG) lElements 	
STDMETHODIMP CIEnumVariant::Skip(ULONG lElements)
{   
	m_lCurrent += lElements; 
	if (m_lCurrent > (long)(m_lLBound + m_lElements))
	{
		m_lCurrent = m_lLBound + m_lElements;
		return (E_FAIL);
	} 
	else
		return (NOERROR);
}

//Reset
//Purpose	- Resets the current element in the enumerator to the beginning. Implements IEnumVARIANT::Reset.
STDMETHODIMP CIEnumVariant::Reset()
{ 
	m_lCurrent = m_lLBound;
	return (NOERROR);
}

//Clone
//Purpose	- Creates a copy of the current enumeration state.
//Outputs	- (IEnumVARIANT **) ppenum
STDMETHODIMP CIEnumVariant::Clone(IEnumVARIANT ** ppenum)
{
	ASSERT(ppenum != NULL);

	HRESULT			hResult;
	CIEnumVariant *	penum;

	hResult = NOERROR;
	penum	= NULL;
	hResult	= 0;
	*ppenum	= NULL;

	hResult = CIEnumVariant::Create(m_pLL, m_lElements, &penum);
	if (hResult != NOERROR)
		goto error;        

	penum->AddRef();
	penum->m_lCurrent = m_lCurrent; 

	*ppenum = penum;        
	return (NOERROR);

error:
	if (penum)
		penum->Release();

	return (hResult);
}
