//CListBank.h  - Copyright (C) 1996 Microsoft Corporation
//            - All rights reserved.
//Date        - 10/19/96
//Authors     - Ken Bergmann, Robert Coleridge
//Purpose     - Header file for the CListBank object.

#ifndef __CLISTBANK_H__
#define __CLISTBANK_H__

#include "support.h"
#include "ckeyll.h"


#ifdef _BENDLL
#include "benefits.h"
#else
#include "benefitsr.h"
#endif


class CListBank : public IListBank
{
private:
	// Private member variables
	ULONG           m_cRef;					//Object reference count
	LPUNKNOWN       m_pUnkOuter;			//Controlling unknown
	PFNDESTROYED    m_pfnDestroy;			//To call on closure

	ITypeInfo*			m_pITypeInfo;		//Type information
    CSupportErrorInfo	m_SupportErrorInfo; // ISupportErrorInfo interface implementation 
	CKeyedLinkedList *	m_pLL;				// Linked list that holds list items.

public:
	// Public member variables
	long m_oid;

	// Public member methods
	CListBank(LPUNKNOWN, PFNDESTROYED);
	~CListBank(void);
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

	//CListBank members
	STDMETHODIMP Add(IBank * pvItem);
	STDMETHODIMP Remove(long lIndex);

	//CListBank properties
	STDMETHODIMP get_Count(long FAR* plCount);
	STDMETHODIMP get_Item(long lIndex, IBank ** pvItem);
	STDMETHODIMP put_Item(IBank * newVal);
	STDMETHODIMP get__NewEnum(IUnknown * FAR* ppIEnumVariantIUnk);
};

typedef CListBank *PCListBank;

class CListBankFactory : public IClassFactory
{
 protected:
	ULONG	m_cRef;
	long	m_oid;

 public:
	CListBankFactory(void);
	~CListBankFactory(void);

	//IUnknown members
	STDMETHODIMP         QueryInterface(REFIID, void **);
	STDMETHODIMP_(ULONG) AddRef(void);
	STDMETHODIMP_(ULONG) Release(void);

	//IClassFactory members
	STDMETHODIMP     CreateInstance(LPUNKNOWN, REFIID, void **);
	STDMETHODIMP     LockServer(BOOL);
};

typedef CListBankFactory *PCListBankFactory;

#endif //#ifndef __CLISTBANK_H__
