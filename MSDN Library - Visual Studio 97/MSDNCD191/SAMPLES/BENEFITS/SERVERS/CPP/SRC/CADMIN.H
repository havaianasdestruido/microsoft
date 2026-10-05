//CAdmin.h  - Copyright (C) 1996 Microsoft Corporation
//            - All rights reserved.
//Date        - 10/19/96
//Authors     - Ken Bergmann, Robert Coleridge
//Purpose     - Header File for CAdmin object

#ifndef __CADMIN_H__
#define __CADMIN_H__

#import "dao350.dll" named_guids no_namespace

//#include <malloc.h>
//#include <wtypes.h>
#include <comutil.h>
#include "support.h"
//#include "cstrfunc.h"

#ifdef _BENDLL
#include "benefits.h"
#else
#include "benefitsr.h"
#endif

class CAdmin : public IAdmin
{
 private:
	// Private Member Variables
	BSTR				m_bstrDataSource;
	BOOL				m_bUseODBC;
	IAdmin *			m_CAdmin;
	ULONG           	m_cRef;             	//Object reference count
	_DBEnginePtr		m_DbEngine;
	DatabasePtr			m_Database;
	short				m_iSrcType;
	long				m_lErrNo;
	long				m_oid;
	IDispatch *			m_pIRemoteBenefitsServer;
	ITypeInfo *		  	m_pITypeInfo;	   	//Type information
	PFNDESTROYED    	m_pfnDestroy;		//To call on closure
	LPUNKNOWN       	m_pUnkOuter;        	//Controlling unknown
	EXCEPINFO			m_RemoteServer_ExceptionInfo;
	CSupportErrorInfo 	m_SupportErrorInfo;	//ISupportErrorInfo interface implementation 
	BOOL				mInitializedAlready;
	BSTR				sErrDesc;

	LPCLASSFACTORY2     m_pClassFactory2;

	// Private Member Methods
 	BOOL		ValidateField(BSTR bstrField);

 public:
	// Public Member Variables
	IListBank *			m_Banks;
	IListBenefit *		m_Benefits;
	IListEmployee *		m_Employees;

	// Public Member Methods
	CAdmin(LPUNKNOWN, PFNDESTROYED);
	~CAdmin(void);

	BOOL	CFInit(void);	// called by class factory
	HRESULT Insert_BenefitLink(IBenefitLink * poBenefitLink, BOOL * pbSuccess);
	HRESULT Insert_Benefit(IBenefit * poBenefit, BOOL * pbSuccess);
	HRESULT Insert_Dependant(IDependant * poDependant, BOOL * pbSuccess);
	HRESULT Insert_Bank(IBank * poBank, BOOL * pbSuccess);
	HRESULT Insert_Employee(IEmployee * poEmployee, BOOL * pbSuccess);
	HRESULT Delete_Employee(IEmployee * poEmployee, BOOL * pbSuccess);
	HRESULT Delete_Bank(IBank * poBank, BOOL * pbSuccess);
	HRESULT Delete_Dependant(IDependant * poDependant, BOOL * pbSuccess);
	HRESULT Delete_Benefit(IBenefit * poBenefit, BOOL * pbSuccess);
	HRESULT Delete_BenefitLink(IBenefitLink * poBenefitLink, BOOL * pbSuccess);
	HRESULT	StartRemoteObject(BSTR bstrRemoteProgId, BOOL bUseODBC, BSTR DataSource, int iSourceType, int iRemoteSvr, int iLinkSvr);
	HRESULT Update_Employee(IEmployee * poEmployee, BOOL * pbSuccess);
	HRESULT Update_Bank(IBank * poBank, BOOL * pbSuccess);
	HRESULT Update_Dependant(IDependant * poDependant, BOOL * pbSuccess);
	HRESULT Update_Benefit(IBenefit * poBenefit, BOOL * pbSuccess);
	HRESULT Update_BenefitLink(IBenefitLink * poBenefitLink, BOOL * pbSuccess);
	BSTR	VariantXToBstr(VARIANT * pvX);
	VARIANT VariantXToVariantBstr(VARIANT * pvX);

	//IUnknown members
	STDMETHODIMP         QueryInterface(REFIID, void **);
	STDMETHODIMP_(ULONG) AddRef(void);
	STDMETHODIMP_(ULONG) Release(void);

	//IDispatch members
	STDMETHODIMP GetTypeInfoCount(UINT *);
	STDMETHODIMP GetTypeInfo(UINT, LCID, ITypeInfo **);
	STDMETHODIMP GetIDsOfNames(REFIID, OLECHAR **, UINT, LCID, DISPID *);
	STDMETHODIMP Invoke(DISPID, REFIID, LCID, WORD, DISPPARAMS *, VARIANT *, EXCEPINFO *, UINT *);

	//IAdmin functions
	STDMETHODIMP get_UseODBC(BOOL * pbUseODBC);
	STDMETHODIMP get_Banks(IListBank ** ppOut);
	STDMETHODIMP get_Benefits(IListBenefit ** ppOut);
	STDMETHODIMP get_Employees(IListEmployee ** ppOut);

	STDMETHODIMP Delete(VARIANT * pvEmployee, VARIANT * pvBank, VARIANT * pvDependant, VARIANT * pvBenefit, VARIANT * pvBenefitLink, BOOL * pbSuccess);
	STDMETHODIMP ExecBooleanMDB(BSTR sQry, IDispatch * lstParameters, BOOL * pbSuccess);
	STDMETHODIMP ExecBooleanSQL(BSTR sQry, BOOL * pbSuccess);
	STDMETHODIMP ExecFillArray(BSTR sQry, VARIANT * pSafearray, BOOL bAddColumnName, BOOL * pbSuccess);
	STDMETHODIMP ExecFillList(BSTR sQry, IList ** ppcList, BOOL * pbSuccess);
	STDMETHODIMP FillListBanks();
	STDMETHODIMP FillListBenefits();
	STDMETHODIMP FillListEmployees();
	STDMETHODIMP GetBenefitTaxTypes(IList * * ppcOutput, BOOL * pbSuccess);
	STDMETHODIMP GetBenefitTypes(IList * * ppcOutput, BOOL * pbSuccess);
	STDMETHODIMP GetDependantTypes(IList * * ppcOutput, BOOL * pbSuccess);
	STDMETHODIMP GetDomain(BSTR sSQL, IList ** ppOut, BOOL * pbSuccess);
	STDMETHODIMP GetEmployeeBenefitCoverageTypes(IList * * ppcOutput, BOOL * pbSuccess);
	STDMETHODIMP GetEmployeeBenefitTaxTypes(IList * * ppcOutput, BOOL * pbSuccess);
	STDMETHODIMP GetListBenefits(VARIANT * pvSARRAY);
	STDMETHODIMP GetListBanks(VARIANT * pvSARRAY);
	STDMETHODIMP GetListEmployees(VARIANT * pvSARRAY);
	STDMETHODIMP GetNumEmplsPerBenefit(IList * * ppcOutput, long lCoverageID, long lTaxID, VARIANT * bGender, BOOL * pbSuccess);
	STDMETHODIMP GetNumEmplsPerBenefitType(IList * * ppcOutput, long lCoverageID, long lTaxID, VARIANT * bGender, BOOL * pbSuccess);
	STDMETHODIMP IsUserValid(BSTR bstrUserId, int * piSuccess);
	STDMETHODIMP Init(BOOL bUseODBC, BSTR DataSource, int iSourceType, int iRemoteSvr, int iLinkSvr, BOOL * pbSuccess);
	STDMETHODIMP Insert(VARIANT * pvEmployee, VARIANT * pvBank, VARIANT * pvDependant, VARIANT * pvBenefit, VARIANT * pvBenefitLink, BOOL * pbSuccess);
	STDMETHODIMP Term();
	STDMETHODIMP Update(VARIANT * pvEmployee, VARIANT * pvBank, VARIANT * pvDependant, VARIANT * pvBenefit, VARIANT * pvBenefitLink, BOOL * pbSuccess);
};

typedef CAdmin *PCAdmin;

//class factory object
class CAdminFactory : public IClassFactory
{
private:
	ULONG           m_cRef;
	long			m_oid;

public:
	CAdminFactory(void);
	~CAdminFactory(void);

	//IUnknown members
	STDMETHODIMP         QueryInterface(REFIID, void **);
	STDMETHODIMP_(ULONG) AddRef(void);
	STDMETHODIMP_(ULONG) Release(void);

	//IClassFactory members
	STDMETHODIMP     CreateInstance(LPUNKNOWN, REFIID, void **);
	STDMETHODIMP     LockServer(BOOL);
};

typedef CAdminFactory *PCAdminFactory;

#endif //#ifndef __CADMIN_H__
