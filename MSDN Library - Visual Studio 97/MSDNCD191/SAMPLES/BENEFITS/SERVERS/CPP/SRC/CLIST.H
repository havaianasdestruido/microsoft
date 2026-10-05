//clist.cpp  - Copyright (C) 1996 Microsoft Corporation
//            - All rights reserved.
//Date        - 10/19/96
//Authors     - Ken Bergmann, Robert Coleridge
//Purpose     - 

#ifndef __CLIST_H__
#define __CLIST_H__

#include "support.h"
#include "ckeyll.h"

#ifdef _BENDLL
#include "benefits.h"
#else
#include "benefitsr.h"
#endif

class CList : public IList
{
protected:
	ULONG           m_cRef;             //Object reference count
	LPUNKNOWN       m_pUnkOuter;        //Controlling unknown
	PFNDESTROYED    m_pfnDestroy;       //To call on closure

	//Our interfaces
    CSupportErrorInfo m_SupportErrorInfo; // ISupportErrorInfo interface implementation 
	ITypeInfo      *m_pITypeInfo;      //Type information

	////////////////////////////////////////////////////////////////////////////
protected:
	// internal vars
	CKeyedLinkedList * m_pLL;		// Linked list that holds list items.

public:
	// helper functions
	static BSTR GenListKey(long lIndex);
	long m_oid;

	// Standard routines
	CList(LPUNKNOWN, PFNDESTROYED);
	~CList(void);
	BOOL CFInit(void);	// called by class factory

	//Non-delegating object IUnknown
	STDMETHODIMP         QueryInterface(REFIID, void **);
	STDMETHODIMP_(ULONG) AddRef(void);
	STDMETHODIMP_(ULONG) Release(void);

	//IDispatch members
	STDMETHODIMP GetTypeInfoCount(UINT *);
	STDMETHODIMP GetTypeInfo(UINT, LCID, ITypeInfo **);
	STDMETHODIMP GetIDsOfNames(REFIID, OLECHAR **, UINT, LCID, DISPID *);
	STDMETHODIMP Invoke(DISPID, REFIID, LCID, WORD, DISPPARAMS *, VARIANT *, EXCEPINFO *, UINT *);

	//CList functions
	STDMETHODIMP Add(BSTR bstrKey,VARIANT FAR* pvItem);
	STDMETHODIMP get_Count(long FAR* plCount);
	STDMETHODIMP get_Item(BSTR bstrKey, VARIANT FAR* pvItem);
	STDMETHODIMP put_Item(BSTR bstrKey, VARIANT newVal);
	STDMETHODIMP get__NewEnum(IUnknown * FAR* ppIEnumVariantIUnk);
	STDMETHODIMP Remove(BSTR bstrKey);
};

typedef CList *PCList;

class CListFactory : public IClassFactory
{
protected:
	ULONG		m_cRef;
	long		m_oid;

public:
	CListFactory(void);
	~CListFactory(void);

	//IUnknown members
	STDMETHODIMP         QueryInterface(REFIID, void **);
	STDMETHODIMP_(ULONG) AddRef(void);
	STDMETHODIMP_(ULONG) Release(void);

	//IClassFactory members
	STDMETHODIMP     CreateInstance(LPUNKNOWN, REFIID, void **);
	STDMETHODIMP     LockServer(BOOL);
};

typedef CListFactory *PCListFactory;

#endif //#ifndef __CLIST_H__
