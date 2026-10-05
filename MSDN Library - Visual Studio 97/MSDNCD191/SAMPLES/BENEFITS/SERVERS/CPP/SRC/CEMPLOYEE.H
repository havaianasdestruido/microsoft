//cemployee.h - Copyright (C) 1996 Microsoft Corporation
//            - All rights reserved.
//Date        - 10/19/96
//Authors     - Ken Bergmann, Robert Coleridge
//Purpose     - Header file for CEmployee object

#ifndef __CEMPLOYEE_H__
#define __CEMPLOYEE_H__

#include "support.h"

#ifdef _BENDLL
#include "benefits.h"
#else
#include "benefitsr.h"
#endif

class CEmployee : public IEmployee
{
private:
	//Private member variables
	ULONG				m_cRef;             //Object reference count
	LPUNKNOWN			m_pUnkOuter;        //Controlling unknown
	PFNDESTROYED		m_pfnDestroy;       //To call on closure

	BOOL				m_bExpanded;
	BSTR				m_bstrAddress1;
	BSTR				m_bstrAddress2;
	BSTR				m_bstrBankAccount;
	BSTR				m_bstrCity;
	BSTR				m_bstrFirstName;
	BSTR				m_bstrLastName;
	BSTR				m_bstrPosition;
	BSTR				m_bstrUserName;
	CURRENCY			m_cySalary;
	long				m_lPKId;
	long				m_oid;
	BOOL				m_pbGender;
	ITypeInfo *			m_pITypeInfo;      //Type information
	IListBenefitLink *	m_plstBenefitLinks;
	IListDependant *	m_plstDependants;
	IAdmin *			m_poAdmin;
	IBank *			    m_poBank;
	CSupportErrorInfo	m_SupportErrorInfo; // ISupportErrorInfo interface implementation 
	DATE				m_vDateOfBirth;
	DATE				m_vdtEndDate;
	DATE				m_vdtStartDate;

	//Private member methods
	void Clear();
	BOOL Expand();
	BOOL GetEmployeeData();

public:
	//Public member methods
	CEmployee(LPUNKNOWN, PFNDESTROYED);
	~CEmployee(void);
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

	//IEmployee functions
	STDMETHODIMP Term(THIS);

	//IEmployee properties
    STDMETHODIMP putref_Admin(IAdmin * poAdmin);
	STDMETHODIMP get_Address1(BSTR FAR* pbstrAddress1);
	STDMETHODIMP put_Address1(BSTR bstrAddress1);
	STDMETHODIMP get_Address2(BSTR FAR* pbstrAddress2);
	STDMETHODIMP put_Address2(BSTR bstrAddress2);
	STDMETHODIMP get_Bank(IBank * FAR* ppoBank);
	STDMETHODIMP putref_Bank(IBank * poBank);
	STDMETHODIMP get_BankAccount(BSTR FAR* pbstrBankAccount);
	STDMETHODIMP put_BankAccount(BSTR bstrBankAccount);
	STDMETHODIMP get_BenefitLinks(IListBenefitLink * FAR* ppcLinkBenefits);
	STDMETHODIMP get_City(BSTR FAR* pbstrCity);
	STDMETHODIMP put_City(BSTR bstrCity);
	STDMETHODIMP get_DateOfBirth(DATE * pvdtDateOfBirth);
	STDMETHODIMP put_DateOfBirth(DATE pvdtDateOfBirth);
	STDMETHODIMP get_Dependants(IListDependant * FAR* ppcDependents);
	STDMETHODIMP get_EndDate(DATE * pvdtEndDate);
	STDMETHODIMP put_EndDate(DATE pvtEndDate);
	STDMETHODIMP get_FirstName(BSTR FAR* pbstrFirstName);
	STDMETHODIMP put_FirstName(BSTR bstrFirstName);
	STDMETHODIMP get_Gender(BOOL FAR* pbGender);
	STDMETHODIMP put_Gender(BOOL bGender);
	STDMETHODIMP get_LastName(BSTR FAR* pbstrLastName);
	STDMETHODIMP put_LastName(BSTR bstrLastName);
	STDMETHODIMP get_PKId(long FAR* plPKId);
	STDMETHODIMP put_PKId(long lPKId);
	STDMETHODIMP get_Position(BSTR FAR* pbstrPosition);
	STDMETHODIMP put_Position(BSTR bstrPosition);
	STDMETHODIMP get_Salary(CURRENCY * pcySalary);
	STDMETHODIMP put_Salary(CURRENCY pcySalary);
	STDMETHODIMP get_StartDate(DATE * pvdtStartDate);
	STDMETHODIMP put_StartDate(DATE pvtStartDate);
	STDMETHODIMP get_UserName(BSTR FAR* pbstrUserName);
	STDMETHODIMP put_UserName(BSTR bstrUserName);
};

typedef CEmployee *PCEmployee;

class CEmployeeFactory : public IClassFactory
{
private:
	ULONG		m_cRef;
	long		m_oid;

public:
	CEmployeeFactory(void);
	~CEmployeeFactory(void);

	//IUnknown members
	STDMETHODIMP         QueryInterface(REFIID, void **);
	STDMETHODIMP_(ULONG) AddRef(void);
	STDMETHODIMP_(ULONG) Release(void);

	//IClassFactory members
	STDMETHODIMP     CreateInstance(LPUNKNOWN, REFIID, void **);
	STDMETHODIMP     LockServer(BOOL);
};

typedef CEmployeeFactory *PCEmployeeFactory;

#endif //#ifndef __CEMPLOYEE_H__
