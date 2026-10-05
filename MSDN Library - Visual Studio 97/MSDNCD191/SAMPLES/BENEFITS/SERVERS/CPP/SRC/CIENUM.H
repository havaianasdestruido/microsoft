//CIEnum.h    - Copyright (C) 1996 Microsoft Corporation
//            - All rights reserved.
//Date        - 10/19/96
//Authors     - Ken Bergmann, Robert Coleridge
//Purpose     - Header file for IEnumVariant enumerator object

#ifndef __CIENUMVARIANT_H__
#define __CIENUMVARIANT_H__

#include <windows.h>
#include "ckeyll.h"

class CIEnumVariant : public IEnumVARIANT
{
private:
	//Private member variables
    ULONG	m_lRef;			// Reference count                                 
    ULONG	m_lElements;	// Number of elements in enumerator. 
    long	m_lLBound;		// Lower bound of index.
    long	m_lCurrent;		// Current index.

public:
	//Public member variables
    CKeyedLinkedList *	m_pLL;   // Link list that holds elements.
	long				m_oid;

	//Public member methods
    CIEnumVariant();
    ~CIEnumVariant();
    static HRESULT Create(CKeyedLinkedList * pLL, ULONG lElements, CIEnumVariant ** ppCIEnumVariant);

    // IUnknown methods 
    STDMETHODIMP_(ULONG) AddRef();
    STDMETHODIMP QueryInterface(REFIID riid, void ** ppvObj);
    STDMETHODIMP_(ULONG) Release();

    // IEnumVARIANT methods 
    STDMETHODIMP Clone(IEnumVARIANT ** ppIEnumVARIANT);
    STDMETHODIMP Next(ULONG cElements, VARIANT * pValue, ULONG * plElementFetched);
    STDMETHODIMP Reset();
    STDMETHODIMP Skip(ULONG lElements);
};

typedef CIEnumVariant *PCIEnumVariant;

#endif //#ifndef __CIENUMVARIANT_H__
