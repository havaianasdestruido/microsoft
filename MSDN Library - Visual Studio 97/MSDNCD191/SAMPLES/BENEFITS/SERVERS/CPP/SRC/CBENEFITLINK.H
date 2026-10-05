//cbenefitlink.h  - Copyright (C) 1996 Microsoft Corporation
//            - All rights reserved.
//Date        - 10/19/96
//Authors     - Ken Bergmann, Robert Coleridge
//Purpose     - Header file for the CBenefitLink object

#ifndef __CBENEFITLINK_H__
#define __CBENEFITLINK_H__

#include "support.h"

#ifdef _BENDLL
#include "benefits.h"
#else
#include "benefitsr.h"
#endif

class CBenefitLink : public IBenefitLink
{
private:
	// Private Member Variables
	ULONG           m_cRef;             //Object reference count
	LPUNKNOWN       m_pUnkOuter;        //Controlling unknown
	PFNDESTROYED    m_pfnDestroy;       //To call on closure

	IAdmin *		m_poAdmin;
	CURRENCY		m_Amount;
	IBenefit *		m_poBenefit;
	LONG			m_CoverageId;
	IEmployee *		m_poEmployee;
	BOOL			m_bExpanded;
	LONG			m_lPKId;
	LONG			m_TaxId;

    CSupportErrorInfo m_SupportErrorInfo; // ISupportErrorInfo interface implementation 
	ITypeInfo *		m_pITypeInfo;      //Type information

	// Private Member Methods
	void Clear();
	BOOL Expand();
	long m_oid;

public:
	// Public member methods
	CBenefitLink(LPUNKNOWN, PFNDESTROYED);
	~CBenefitLink(void);
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

	//IBenefitLink members
	STDMETHODIMP Term(THIS);

	//IBenefitLink properties
	STDMETHODIMP putref_Admin(IAdmin * poAdmin);
	STDMETHODIMP get_Amount(CURRENCY __RPC_FAR *newVal);
	STDMETHODIMP put_Amount(CURRENCY newVal);
	STDMETHODIMP get_Benefit(IBenefit * FAR* poBenefit);
	STDMETHODIMP putref_Benefit(IBenefit * poBenefit);
	STDMETHODIMP get_CoverageId(long FAR* plCoverageId);
	STDMETHODIMP put_CoverageId(long lCoverageId);
	STDMETHODIMP get_Employee(IEmployee * FAR* poEmployee);
	STDMETHODIMP putref_Employee(IEmployee * poEmployee);
	STDMETHODIMP get_PKId(long FAR* plPKId);
	STDMETHODIMP put_PKId(long lPKId);
	STDMETHODIMP get_TaxId(long FAR* plTaxId);
	STDMETHODIMP put_TaxId(long lTaxId);
};

typedef CBenefitLink *PCBenefitLink;

class CBenefitLinkFactory : public IClassFactory
{
private:
	//Private member variables
	ULONG		m_cRef;
	long		m_oid;

public:
	//Private member methods
	CBenefitLinkFactory(void);
	~CBenefitLinkFactory(void);

	//IUnknown members
	STDMETHODIMP         QueryInterface(REFIID, void **);
	STDMETHODIMP_(ULONG) AddRef(void);
	STDMETHODIMP_(ULONG) Release(void);

	//IClassFactory members
	STDMETHODIMP     CreateInstance(LPUNKNOWN, REFIID, void **);
	STDMETHODIMP     LockServer(BOOL);
};

typedef CBenefitLinkFactory *PCBenefitLinkFactory;

#endif //#ifndef __CBENEFITLINK_H__
