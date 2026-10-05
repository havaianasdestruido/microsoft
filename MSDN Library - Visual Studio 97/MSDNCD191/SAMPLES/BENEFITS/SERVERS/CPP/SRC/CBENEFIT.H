//cbenefit.h  - Copyright (C) 1996 Microsoft Corporation
//            - All rights reserved.
//Date        - 10/19/96
//Authors     - Ken Bergmann, Robert Coleridge
//Purpose     - Header file for the CBenefit object.

#ifndef __CBENEFIT_H__
#define __CBENEFIT_H__

#include "support.h"

#ifdef _BENDLL
#include "benefits.h"
#else
#include "benefitsr.h"
#endif

class CBenefit : public IBenefit
{
private:
	// Private Member Variables
	BSTR         m_bstrDataSource;
	BSTR		 m_bstrDescription;
	BOOL         m_bExpanded;
	ULONG        m_cRef;             //Object reference count
	int          m_iIncrement;
	int 		 m_iMultiples;
	long		 m_lBnfTaxId;
	long		 m_lBnfTypeId;
	long		 m_lCost;
	long		 m_lMaximum;
	long		 m_lMinimum;
	long		 m_lPKId;
	ITypeInfo *	 m_pITypeInfo;       //Type information
	LPUNKNOWN    m_pUnkOuter;        //Controlling unknown
	PFNDESTROYED m_pfnDestroy;       //To call on closure
	IAdmin *     m_poAdmin;	         //Pointer to the "parent" CAdmin object
	CSupportErrorInfo m_SupportErrorInfo; //ISupportErrorInfo interface implementation 

	// Private Member Methods
	void Clear();
	BOOL Expand();
	long m_oid;

public:
	//Public Member Methods
	CBenefit(LPUNKNOWN, PFNDESTROYED);
	~CBenefit(void);
	BOOL CFInit(void);

	//IUnknown Members
	STDMETHODIMP         QueryInterface(REFIID, void **);
	STDMETHODIMP_(ULONG) AddRef(void);
	STDMETHODIMP_(ULONG) Release(void);

	//IDispatch members
	STDMETHODIMP GetTypeInfoCount(UINT *);
	STDMETHODIMP GetTypeInfo(UINT, LCID, ITypeInfo **);
	STDMETHODIMP GetIDsOfNames(REFIID, OLECHAR **, UINT, LCID, DISPID *);
	STDMETHODIMP Invoke(DISPID, REFIID, LCID, WORD, DISPPARAMS *, VARIANT *, EXCEPINFO *, UINT *);

	//IBenefit members
	STDMETHODIMP Term(THIS);

	//IBenefit properties
	STDMETHOD(putref_Admin)(IAdmin * poAdmin);
	STDMETHODIMP get_BnfTaxId(long FAR* plBnfTaxId);
	STDMETHODIMP put_BnfTaxId(long lBnfTaxId);
	STDMETHODIMP get_BnfTypeId(long FAR* plBnfTypeId);
	STDMETHODIMP put_BnfTypeId(long lBnfTypeId);
	STDMETHODIMP get_Cost(long FAR* plCost);
	STDMETHODIMP put_Cost(long lCost);
	STDMETHODIMP get_Description(BSTR FAR* pbstrDescription);
	STDMETHODIMP put_Description(BSTR bstrDescription);
	STDMETHODIMP get_Increment(short FAR* piIncrement);
	STDMETHODIMP put_Increment(short iIncrement);
	STDMETHODIMP get_Maximum(long FAR* plMaximum);
	STDMETHODIMP put_Maximum(long lMaximum);
	STDMETHODIMP get_Minimum(long FAR* plMinimum);
	STDMETHODIMP put_Minimum(long lMinimum);
	STDMETHODIMP get_Multiples(short FAR* piMultiples);
	STDMETHODIMP put_Multiples(short iMultiples);
	STDMETHODIMP get_PKId(long FAR* plPKId);
	STDMETHODIMP put_PKId(long lPKId);
};

typedef CBenefit * PCBenefit;

class CBenefitFactory : public IClassFactory
{
 private:
	ULONG  m_cRef;
	long m_oid;

 public:
	CBenefitFactory(void);
	~CBenefitFactory(void);

	//IUnknown members
	STDMETHODIMP         QueryInterface(REFIID, void **);
	STDMETHODIMP_(ULONG) AddRef(void);
	STDMETHODIMP_(ULONG) Release(void);

	//IClassFactory members
	STDMETHODIMP     CreateInstance(LPUNKNOWN, REFIID, void **);
	STDMETHODIMP     LockServer(BOOL);
};

typedef CBenefitFactory *PCBenefitFactory;

#endif //#ifndef __CBENEFIT_H__
