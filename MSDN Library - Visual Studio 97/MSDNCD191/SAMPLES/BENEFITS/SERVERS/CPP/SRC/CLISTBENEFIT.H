//CListBenefit.h  - Copyright (C) 1996 Microsoft Corporation
//            - All rights reserved.
//Date        - 10/19/96
//Authors     - Ken Bergmann, Robert Coleridge
//Purpose     - Header file for the CListBenefit object.

#ifndef __CLISTBENEFIT_H__
#define __CLISTBENEFIT_H__

#include "support.h"
#include "ckeyll.h"

#ifdef _BENDLL
#include "benefits.h"
#else
#include "benefitsr.h"
#endif

// forward referencing

class CListBenefit : public IListBenefit
{
 private:
	//Private member variables
	ULONG           m_cRef;             //Object reference count
	LPUNKNOWN       m_pUnkOuter;        //Controlling unknown
	PFNDESTROYED    m_pfnDestroy;       //To call on closure

	ITypeInfo *			m_pITypeInfo;		//Type information
	CKeyedLinkedList *	m_pLL;		// Linked list that holds list items.
	long				m_oid;
    CSupportErrorInfo	m_SupportErrorInfo;	// ISupportErrorInfo interface implementation 

 public:
	//Public member variables
	CListBenefit(LPUNKNOWN, PFNDESTROYED);
	~CListBenefit(void);
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

	//CListBenefit members
	STDMETHODIMP Add(IBenefit * pvItem);
	STDMETHODIMP Remove(long lIndex);

	//CListBenefit properties
	STDMETHODIMP get_Count(long FAR* plCount);
	STDMETHODIMP get_Item(long lIndex, IBenefit ** pvItem);
	STDMETHODIMP put_Item(IBenefit * newVal);
	STDMETHODIMP get__NewEnum(IUnknown * FAR* ppIEnumVariantIUnk);
};

typedef CListBenefit *PCListBenefit;

class CListBenefitFactory : public IClassFactory
{
 protected:
	ULONG		m_cRef;
	long		m_oid;

 public:
	CListBenefitFactory(void);
	~CListBenefitFactory(void);

	//IUnknown members
	STDMETHODIMP         QueryInterface(REFIID, void **);
	STDMETHODIMP_(ULONG) AddRef(void);
	STDMETHODIMP_(ULONG) Release(void);

	//IClassFactory members
	STDMETHODIMP     CreateInstance(LPUNKNOWN, REFIID, void **);
	STDMETHODIMP     LockServer(BOOL);
};

typedef CListBenefitFactory *PCListBenefitFactory;

#endif //#ifndef __CLISTBENEFIT_H__
