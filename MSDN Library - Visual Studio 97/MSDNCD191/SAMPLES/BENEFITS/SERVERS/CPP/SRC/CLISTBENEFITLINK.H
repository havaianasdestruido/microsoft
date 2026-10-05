//clistbenefitlink.h  - Copyright (C) 1996 Microsoft Corporation
//            - All rights reserved.
//Date        - 10/19/96
//Authors     - Ken Bergmann, Robert Coleridge
//Purpose     - Header file for the CListBenefitLink object.

#ifndef __CLISTBENEFITLINK_H__
#define __CLISTBENEFITLINK_H__

#include "support.h"
#include "ckeyll.h"

#ifdef _BENDLL
#include "benefits.h"
#else
#include "benefitsr.h"
#endif

class CListBenefitLink : public IListBenefitLink
{
private:
	//Private member methods
	ULONG           m_cRef;             //Object reference count
	LPUNKNOWN       m_pUnkOuter;        //Controlling unknown
	PFNDESTROYED    m_pfnDestroy;       //To call on closure

	long			m_oid;
	ITypeInfo *     m_pITypeInfo;      //Type information
	CKeyedLinkedList * m_pLL;		// Linked list that holds list items.
    CSupportErrorInfo m_SupportErrorInfo; // ISupportErrorInfo interface implementation 

public:
	//Public member methods
	CListBenefitLink(LPUNKNOWN, PFNDESTROYED);
	~CListBenefitLink(void);
	BOOL CFInit(void);	// called by class factory

	//IUnknown methods
	STDMETHODIMP         QueryInterface(REFIID, void **);
	STDMETHODIMP_(ULONG) AddRef(void);
	STDMETHODIMP_(ULONG) Release(void);

	//IDispatch members
	STDMETHODIMP GetTypeInfoCount(UINT *);
	STDMETHODIMP GetTypeInfo(UINT, LCID, ITypeInfo **);
	STDMETHODIMP GetIDsOfNames(REFIID, OLECHAR **, UINT, LCID, DISPID *);
	STDMETHODIMP Invoke(DISPID, REFIID, LCID, WORD, DISPPARAMS *, VARIANT *, EXCEPINFO *, UINT *);

	//CListBenefitLink methods
	STDMETHODIMP Add(IBenefitLink * pvItem);
	STDMETHODIMP Remove(long lIndex);

	//CListBenefitLink properties
	STDMETHODIMP get_Count(long FAR* plCount);
	STDMETHODIMP get_Item(long lIndex, IBenefitLink ** pvItem);
	STDMETHODIMP put_Item(IBenefitLink * newVal);
	STDMETHODIMP get__NewEnum(IUnknown * FAR* ppIEnumVariantIUnk);
};

typedef CListBenefitLink *PCListBenefitLink;

class CListBenefitLinkFactory : public IClassFactory
{
protected:
	ULONG		m_cRef;
	long		m_oid;

public:
	CListBenefitLinkFactory(void);
	~CListBenefitLinkFactory(void);

	//IUnknown members
	STDMETHODIMP         QueryInterface(REFIID, void **);
	STDMETHODIMP_(ULONG) AddRef(void);
	STDMETHODIMP_(ULONG) Release(void);

	//IClassFactory members
	STDMETHODIMP     CreateInstance(LPUNKNOWN, REFIID, void **);
	STDMETHODIMP     LockServer(BOOL);
};

typedef CListBenefitLinkFactory *PCListBenefitLinkFactory;

#endif //#ifndef __CLISTBENEFITLINK_H__
