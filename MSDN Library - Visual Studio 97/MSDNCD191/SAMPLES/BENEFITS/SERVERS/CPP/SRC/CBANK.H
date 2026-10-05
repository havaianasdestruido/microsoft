//cbank.h	  - Copyright (C) 1996 Microsoft Corporation
//            - All rights reserved.
//Date        - 10/19/96
//Authors     - Ken Bergmann, Robert Coleridge
//Purpose     - Header file for the CBank object.

#ifndef __CBANK_H__
#define __CBANK_H__

#include "support.h"

#ifdef _BENDLL
#include "benefits.h"
#else
#include "benefitsr.h"
#endif

class CBank : public IBank
{
 private:
	// Private Member Variables
	BOOL			  m_bExpanded;
	BSTR			  m_bstrAddress1;
	BSTR			  m_bstrAddress2;
	BSTR			  m_bstrCity;
	BSTR			  m_bstrCode;
	BSTR			  m_bstrName;
	ULONG             m_cRef;             //Object reference count
	long			  m_lPKId;			  //Primary key ID
	LPUNKNOWN         m_pUnkOuter;        //Controlling unknown
	PFNDESTROYED      m_pfnDestroy;       //To call on closure
	ITypeInfo*        m_pITypeInfo;       //Type information
	IAdmin *		  m_poAdmin;
	IListEmployee *   m_poEmployees;       //List of employees, can be empty
	CSupportErrorInfo m_SupportErrorInfo; // ISupportErrorInfo interface implementation 

	// Private Member Mmethods
	void Clear();
	BOOL Expand();
	BOOL GetBankData();
	BOOL GetEmployeesData();
	long m_oid;

 public:
	// Public Member Mmethods
	CBank(LPUNKNOWN, PFNDESTROYED);
	~CBank(void);
	BOOL CFInit(void);	// called by class factory

	// IUnknown members
	STDMETHODIMP         QueryInterface(REFIID, void **);
	STDMETHODIMP_(ULONG) AddRef(void);
	STDMETHODIMP_(ULONG) Release(void);

	//IDispatch members
	STDMETHODIMP GetTypeInfoCount(UINT *);
	STDMETHODIMP GetTypeInfo(UINT, LCID, ITypeInfo **);
	STDMETHODIMP GetIDsOfNames(REFIID, OLECHAR **, UINT, LCID, DISPID *);
	STDMETHODIMP Invoke(DISPID, REFIID, LCID, WORD, DISPPARAMS *, VARIANT *, EXCEPINFO *, UINT *);

	//IBank functions
	STDMETHODIMP Term();

	//IBank properties
	STDMETHODIMP putref_Admin(IAdmin * poAdmin);
	STDMETHODIMP get_Address1(BSTR FAR* pbstrAddress1);
	STDMETHODIMP put_Address1(BSTR bstrAddress1);
	STDMETHODIMP get_Address2(BSTR FAR* pbstrAddress2);
	STDMETHODIMP put_Address2(BSTR bstrAddress2);
	STDMETHODIMP get_Employees(IListEmployee * FAR* ppcEmployees);
	STDMETHODIMP get_City(BSTR FAR* pbstrCity);
	STDMETHODIMP put_City(BSTR bstrCity);
	STDMETHODIMP get_Code(BSTR FAR* pbstrCode);
	STDMETHODIMP put_Code(BSTR bstrCode);
	STDMETHODIMP get_Name(BSTR FAR* pbstrName);
	STDMETHODIMP put_Name(BSTR bstrName);
	STDMETHODIMP get_PKId(long FAR* plPKId);
	STDMETHODIMP put_PKId(long lPKId);
};

typedef CBank *PCBank;

//Class Factory Declaration

class CBankFactory : public IClassFactory
{
private:
	ULONG		m_cRef;
	long		m_oid;

public:
	CBankFactory(void);
	~CBankFactory(void);

	//IUnknown members
	STDMETHODIMP         QueryInterface(REFIID, void **);
	STDMETHODIMP_(ULONG) AddRef(void);
	STDMETHODIMP_(ULONG) Release(void);

	//IClassFactory members
	STDMETHODIMP     CreateInstance(LPUNKNOWN, REFIID, void **);
	STDMETHODIMP     LockServer(BOOL);
};

typedef CBankFactory *PCBankFactory;

#endif //#ifndef __CBANK_H__
