//clistdenpenant.h  - Copyright (C) 1996 Microsoft Corporation
//            - All rights reserved.
//Date        - 10/19/96
//Authors     - Ken Bergmann, Robert Coleridge
//Purpose     - Header file for CListDependant object.

#ifndef __CLISTDEPENDANT_H__
#define __CLISTDEPENDANT_H__

#include "support.h"
#include "ckeyll.h"

#ifdef _BENDLL
#include "benefits.h"
#else
#include "benefitsr.h"
#endif

class CListDependant : public IListDependant
{
private:
	//Private member variables
	ULONG           m_cRef;             //Object reference count
	LPUNKNOWN       m_pUnkOuter;        //Controlling unknown
	PFNDESTROYED    m_pfnDestroy;       //To call on closure

	ITypeInfo *     m_pITypeInfo;      //Type information
	long m_oid;
	CKeyedLinkedList * m_pLL;		// Linked list that holds list items.
    CSupportErrorInfo m_SupportErrorInfo; // ISupportErrorInfo interface implementation 

public:
	//Public member variables
	CListDependant(LPUNKNOWN, PFNDESTROYED);
	~CListDependant(void);
	BOOL CFInit(void);	// called by class factory

	//IUnknown members
	STDMETHODIMP         QueryInterface(REFIID, void **);
	STDMETHODIMP_(ULONG) AddRef(void);
	STDMETHODIMP_(ULONG) Release(void);

	//IDispatch members
	STDMETHODIMP GetTypeInfoCount(UINT *);
	STDMETHODIMP GetTypeInfo(UINT, LCID, ITypeInfo **);
	STDMETHODIMP GetIDsOfNames(REFIID, OLECHAR **, UINT, LCID, DISPID *);
	STDMETHODIMP Invoke(DISPID, REFIID, LCID, WORD, DISPPARAMS *, VARIANT *, EXCEPINFO *, UINT *);

	//CListDependant members
	STDMETHODIMP Add(IDependant * pvItem);
	STDMETHODIMP Remove(long lIndex);

	//CListDependant members
	STDMETHODIMP get_Count(long FAR* plCount);
	STDMETHODIMP get_Item(long lIndex, IDependant ** pvItem);
	STDMETHODIMP put_Item(IDependant * newVal);
	STDMETHODIMP get__NewEnum(IUnknown * FAR* ppIEnumVariantIUnk);
};

typedef CListDependant *PCListDependant;

class CListDependantFactory : public IClassFactory
{
protected:
	ULONG		m_cRef;
	long		m_oid;

public:
	CListDependantFactory (void);
	~CListDependantFactory (void);

	//IUnknown members
	STDMETHODIMP         QueryInterface(REFIID, void **);
	STDMETHODIMP_(ULONG) AddRef(void);
	STDMETHODIMP_(ULONG) Release(void);

	//IClassFactory members
	STDMETHODIMP     CreateInstance(LPUNKNOWN, REFIID, void **);
	STDMETHODIMP     LockServer(BOOL);
};

typedef CListDependantFactory  *PCListDependantFactory ;

#endif //#ifndef __CLISTDEPENDANT_H__
