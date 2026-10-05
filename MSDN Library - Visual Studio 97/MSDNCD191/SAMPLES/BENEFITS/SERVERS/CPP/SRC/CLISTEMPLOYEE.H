//clistemployee.cpp  - Copyright (C) 1996 Microsoft Corporation
//            - All rights reserved.
//Date        - 10/19/96
//Authors     - Ken Bergmann, Robert Coleridge
//Purpose     - Header file for CListEmployee object.

#ifndef __CLISTEMPLOYEE_H__
#define __CLISTEMPLOYEE_H__

#include "support.h"
#include "ckeyll.h"

#ifdef _BENDLL
#include "benefits.h"
#else
#include "benefitsr.h"
#endif

class CListEmployee : public IListEmployee
{
private:
	ULONG           m_cRef;             //Object reference count
	LPUNKNOWN       m_pUnkOuter;        //Controlling unknown
	PFNDESTROYED    m_pfnDestroy;       //To call on closure

	ITypeInfo *		m_pITypeInfo;		//Type information
	CKeyedLinkedList * m_pLL;			// Linked list that holds list items.
	long			m_oid;
    CSupportErrorInfo m_SupportErrorInfo; // ISupportErrorInfo interface implementation 

public:
	//Private member methods
	CListEmployee(LPUNKNOWN, PFNDESTROYED);
	~CListEmployee(void);
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

	//CList members
	STDMETHODIMP Add(IEmployee * pvItem);
	STDMETHODIMP Remove(long lKey);

	//CList properties
	STDMETHODIMP get_Count(long FAR* plCount);
	STDMETHODIMP get_Item(long lKey, IEmployee ** pvItem);
	STDMETHODIMP put_Item(IEmployee * newVal);
	STDMETHODIMP get__NewEnum(IUnknown * FAR* ppIEnumVariantIUnk);
};

typedef CListEmployee *PCListEmployee;

class CListEmployeeFactory : public IClassFactory
{
protected:
	ULONG           m_cRef;
	long m_oid;

public:
	CListEmployeeFactory(void);
	~CListEmployeeFactory(void);

	//IUnknown members
	STDMETHODIMP         QueryInterface(REFIID, void **);
	STDMETHODIMP_(ULONG) AddRef(void);
	STDMETHODIMP_(ULONG) Release(void);

	//IClassFactory members
	STDMETHODIMP     CreateInstance(LPUNKNOWN, REFIID, void **);
	STDMETHODIMP     LockServer(BOOL);
};

typedef CListEmployeeFactory *PCListEmployeeFactory;

#endif //#ifndef __CLISTEMPLOYEE_H__
