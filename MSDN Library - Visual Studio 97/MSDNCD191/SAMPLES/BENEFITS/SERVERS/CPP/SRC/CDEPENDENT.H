//CDependant.h- Copyright (C) 1996 Microsoft Corporation
//            - All rights reserved.
//Date        - 10/19/96
//Authors     - Ken Bergmann, Robert Coleridge
//Purpose     - Header file for CDependant object

#ifndef __CDEPENDANT_H__
#define __CDEPENDANT_H__

#include "support.h"

#ifdef _BENDLL
#include "benefits.h"
#else
#include "benefitsr.h"
#endif

class CDependant : public IDependant
{
private:
	//Private member variables
	ULONG        m_cRef;             //Object reference count
	LPUNKNOWN    m_pUnkOuter;        //Controlling unknown
	PFNDESTROYED m_pfnDestroy;       //To call on closure
	BSTR        m_bstrDataSource;
	BSTR		m_bstrFirstName;
	BSTR		m_bstrLastName;
	BOOL		m_bExpanded;
	int			m_iGender;          //0:male, 1:female, 
	int         m_iStatus;          //0:Dependant, 1:beneficiary, 
	long		m_lPKId;			//Primary Key ID
    long		m_lDepTypeId;		//Dependant Type ID 
	IAdmin*		m_poAdmin;
	IEmployee*	m_poEmployee;		//Reference to the employee
	DATE		m_vDateOfBirth;		//Date of birth  
    CSupportErrorInfo m_SupportErrorInfo; // ISupportErrorInfo interface implementation 
	ITypeInfo*	m_pITypeInfo;		//Type information
	long		m_oid;

	//Private member methods
	void Clear();
	BOOL Expand();
	BSTR GetSQLStatement();
	BOOL GetDependantData();

public:
	//Public member methods
	CDependant(LPUNKNOWN, PFNDESTROYED);
	~CDependant(void);
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

	//IDependant functions
	STDMETHODIMP Term(THIS);

	//IDependant properties
	STDMETHODIMP putref_Admin(IAdmin * poAdmin);
	STDMETHODIMP get_DateOfBirth(DATE * pvdtDateOfBirth);
	STDMETHODIMP put_DateOfBirth(DATE pvdtDateOfBirth);
	STDMETHODIMP get_DepTypeId(long FAR* plDepTypeId);
	STDMETHODIMP put_DepTypeId(long lDepTypeId);
	STDMETHODIMP get_Employee(IEmployee * FAR* ppoEmployee);
	STDMETHODIMP putref_Employee(IEmployee * poEmployee);
	STDMETHODIMP get_FirstName(BSTR FAR* pbstrFirstName);
	STDMETHODIMP put_FirstName(BSTR bstrFirstName);
	STDMETHODIMP get_Gender(BOOL FAR* pbGender);
	STDMETHODIMP put_Gender(BOOL bGender);
	STDMETHODIMP get_LastName(BSTR FAR* pbstrLastName);
	STDMETHODIMP put_LastName(BSTR bstrLastName);
	STDMETHODIMP get_PKId(long FAR* plPKId);
	STDMETHODIMP put_PKId(long lPKId);
	STDMETHODIMP get_Status(BOOL FAR* pbStatus);
	STDMETHODIMP put_Status(BOOL bStatus);
};

typedef CDependant *PCDependant;

class CDependantFactory : public IClassFactory
{
protected:
	ULONG		m_cRef;
	long		m_oid;

public:
	CDependantFactory(void);
	~CDependantFactory(void);

	//IUnknown members
	STDMETHODIMP         QueryInterface(REFIID, void **);
	STDMETHODIMP_(ULONG) AddRef(void);
	STDMETHODIMP_(ULONG) Release(void);

	//IClassFactory members
	STDMETHODIMP     CreateInstance(LPUNKNOWN, REFIID, void **);
	STDMETHODIMP     LockServer(BOOL);
};

typedef CDependantFactory *PCDependantFactory;

#endif //__CDEPENDANT_H__
