//cemployee.cpp  - Copyright (C) 1996 Microsoft Corporation
//            - All rights reserved.
//Date        - 10/19/96
//Authors     - Ken Bergmann, Robert Coleridge
//Purpose     - Implementation file for CEmployee object

#pragma warning (disable : 4800)	// turns off bool warning

#include "cemployee.h"
#include "cienum.h"
#include "csaferay.h"

extern long g_oid;

//Globals used to determine when to unload a In-Proc or Local Server.
extern ULONG       g_lObjectCount;
extern ULONG       g_lLockedObjects;

STDMETHODIMP CEmployee::get_Gender(BOOL FAR* pbGender) 
{
	if (m_bExpanded)
		*pbGender = m_pbGender;
	else
	{
		if (Expand())
			*pbGender = m_pbGender;
		else
		{
			Term();
			return (FlagException(IID_IEmployee, scEmplExpand, icErrExpand));
		}
	}
	return (NOERROR);
}

STDMETHODIMP CEmployee::put_Gender(BOOL iGender) 
{
	m_pbGender = iGender;
	return (NOERROR);
}

STDMETHODIMP CEmployee::putref_Admin(IAdmin * poAdmin)
{
	if (m_poAdmin)
	{
		m_poAdmin->Release();
		m_poAdmin = NULL;
	}

	m_poAdmin = (IAdmin *)poAdmin;
	
	if (m_poAdmin)
		m_poAdmin->AddRef();

	return (NOERROR);
}

STDMETHODIMP CEmployee::get_PKId(long FAR* plPKId)
{
	*plPKId = m_lPKId;
	return (NOERROR);
}

STDMETHODIMP CEmployee::put_PKId(long lPKId)
{
	m_lPKId = lPKId;
	return (NOERROR);
}

STDMETHODIMP CEmployee::get_Address1(BSTR FAR * pbstrAddress1)
{
	if (m_bExpanded)
		::SysReAllocString(pbstrAddress1,m_bstrAddress1);
	else
	{
		if (Expand())
			::SysReAllocString(pbstrAddress1,m_bstrAddress1);
		else
		{
			Term();
			return (FlagException(IID_IEmployee, scEmplExpand, icErrExpand));
		}
	}

	return (NOERROR);
}

STDMETHODIMP CEmployee::put_Address1(BSTR bstrAddress1)
{
	::SysReAllocString(&m_bstrAddress1, bstrAddress1);
	return (NOERROR);
}

STDMETHODIMP CEmployee::get_Address2(BSTR FAR* pbstrAddress2)
{
	if (m_bExpanded)
		::SysReAllocString(pbstrAddress2,m_bstrAddress2);
	else
	{
		if (Expand())
			::SysReAllocString(pbstrAddress2,m_bstrAddress2);
		else
		{
			Term();
			return (FlagException(IID_IEmployee, scEmplExpand, icErrExpand));
		}
	}

	return (NOERROR);
}

STDMETHODIMP CEmployee::put_Address2(BSTR bstrAddress2)
{
	::SysReAllocString(&m_bstrAddress2, bstrAddress2);
	return (NOERROR);
}

STDMETHODIMP CEmployee::get_City(BSTR FAR * pbstrCity)
{
	if (m_bExpanded)
		::SysReAllocString(pbstrCity,m_bstrCity);
	else
	{
		if (Expand())
			::SysReAllocString(pbstrCity,m_bstrCity);
		else
		{
			Term();
			return (FlagException(IID_IEmployee, scEmplExpand, icErrExpand));
		}
	}

	return (NOERROR);
}

STDMETHODIMP CEmployee::put_City(BSTR bstrCity)
{
	::SysReAllocString(&m_bstrCity, bstrCity);
	return (NOERROR);
}

STDMETHODIMP CEmployee::get_Bank(IBank * FAR* ppoBank)
{
	if (m_bExpanded)
		*ppoBank = (IBank *)m_poBank;
	else
	{
		if (Expand())
			*ppoBank = (IBank *)m_poBank;
		else
		{
			Term();
			return (FlagException(IID_IEmployee, scEmplExpand, icErrExpand));
		}
	}

	if (m_poBank)
		m_poBank->AddRef();

	return (NOERROR);
}

STDMETHODIMP CEmployee::putref_Bank(IBank * poBank)
{
	if (m_poBank)
	{
		m_poBank->Term();
		m_poBank->Release();
		m_poBank = NULL;
	}

	m_poBank = (IBank *)poBank;
	
	if (m_poBank)
		m_poBank->AddRef();

	return (NOERROR);
}

STDMETHODIMP CEmployee::get_BankAccount(BSTR FAR* pbstrBankAccount)
{
	if (m_bExpanded)
		::SysReAllocString(pbstrBankAccount,m_bstrBankAccount);
	else
	{
		if (Expand())
			::SysReAllocString(pbstrBankAccount,m_bstrBankAccount);
		else
		{
			Term();
			return (FlagException(IID_IEmployee, scEmplExpand, icErrExpand));
		}
	}

	return (NOERROR);
}

STDMETHODIMP CEmployee::put_BankAccount(BSTR bstrBankAccount)
{
	::SysReAllocString(&m_bstrBankAccount, bstrBankAccount);
	return (NOERROR);
}

STDMETHODIMP CEmployee::get_Position(BSTR FAR* pbstrPosition)
{
	if (m_bExpanded)
		::SysReAllocString(pbstrPosition,m_bstrPosition);
	else
	{
		if (Expand())
			::SysReAllocString(pbstrPosition,m_bstrPosition);
		else
		{
			Term();
			return (FlagException(IID_IEmployee, scEmplExpand, icErrExpand));
		}
	}

	return (NOERROR);
}

STDMETHODIMP CEmployee::put_Position(BSTR bstrPosition)
{

	::SysReAllocString(&m_bstrPosition, bstrPosition);
	return (NOERROR);
}

STDMETHODIMP CEmployee::get_Salary(CURRENCY * pcySalary)
{
	if (m_bExpanded)
		*pcySalary = m_cySalary;
	else
	{
		if (Expand())
			*pcySalary = m_cySalary;
		else
		{
			Term();
			return (FlagException(IID_IEmployee, scEmplExpand, icErrExpand));
		}
	}

	return (NOERROR);
}

STDMETHODIMP CEmployee::put_Salary(CURRENCY pcySalary)
{
	m_cySalary = pcySalary;
	return (NOERROR);
}

STDMETHODIMP CEmployee::get_StartDate(DATE * pvdtStartDate)
{
	if (m_bExpanded)
		*pvdtStartDate = m_vdtStartDate;
	else
	{
		if (Expand())
			*pvdtStartDate = m_vdtStartDate;
		else
		{
			Term();
			return (FlagException(IID_IEmployee, scEmplExpand, icErrExpand));
		}
	}

	return (NOERROR);
}

STDMETHODIMP CEmployee::put_StartDate(DATE pvtStartDate)
{
	m_vdtStartDate = pvtStartDate;
	return (NOERROR);
}

STDMETHODIMP CEmployee::get_EndDate(DATE * pvdtEndDate)
{
	if (m_bExpanded)
		*pvdtEndDate = m_vdtEndDate;
	else
	{
		if (Expand())
			*pvdtEndDate = m_vdtEndDate;
		else
		{
			Term();
			return (FlagException(IID_IEmployee, scEmplExpand, icErrExpand));
		}
	}

	return (NOERROR);
}

STDMETHODIMP CEmployee::put_EndDate(DATE pvtEndDate)
{
	m_vdtEndDate = pvtEndDate;
	return (NOERROR);
}

STDMETHODIMP CEmployee::get_UserName(BSTR FAR* pbstrUserName)
{
	if (m_bExpanded)
		::SysReAllocString(pbstrUserName,m_bstrUserName);
	else
	{
		if (Expand())
			::SysReAllocString(pbstrUserName,m_bstrUserName);
		else
		{
			Term();
			return (FlagException(IID_IEmployee, scEmplExpand, icErrExpand));
		}
	}

	return (NOERROR);
}

STDMETHODIMP CEmployee::put_UserName(BSTR bstrUserName)
{
	::SysReAllocString(&m_bstrUserName, bstrUserName);
	return (NOERROR);
}

STDMETHODIMP CEmployee::get_Dependants(IListDependant * FAR* ppcDependants)
{
	if (m_bExpanded)
		*ppcDependants = (IListDependant *)m_plstDependants;
	else
	{
		if (Expand())
			*ppcDependants = (IListDependant *)m_plstDependants;
		else
		{
			Term();
			return (FlagException(IID_IEmployee, scEmplExpand, icErrExpand));
		}
	}

	if (m_plstDependants)
		m_plstDependants->AddRef();

	return (NOERROR);
}

STDMETHODIMP CEmployee::get_BenefitLinks(IListBenefitLink * FAR* ppcLinkBenefits)
	{
	// pass back value
	if (m_bExpanded)
		*ppcLinkBenefits = (IListBenefitLink *)m_plstBenefitLinks;
	else
	{
		if (Expand())
			*ppcLinkBenefits = (IListBenefitLink *)m_plstBenefitLinks;
		else
		{
			Term();
			return (FlagException(IID_IEmployee, scEmplExpand, icErrExpand));
		}
	}

	if (m_plstBenefitLinks)
		m_plstBenefitLinks->AddRef();

	return (NOERROR);
}

STDMETHODIMP CEmployee::Term()
{
	Clear();
	return (NOERROR);
}

BOOL CEmployee::Expand()
{
    if (m_lPKId)
	{		
		Clear();
		if (m_poAdmin)
		{
			if (GetEmployeeData())
				m_bExpanded = TRUE;
			else
			{	
				Term();
				m_bExpanded = FALSE;
				FlagException(IID_IEmployee, scEmplExpand, icErrExpand);
			}		
		}
		else
		{
			Term();
			m_bExpanded = FALSE;
			FlagException(IID_IEmployee, scEmplExpand, icErrExpand);
		}
	}
	else
	{
		m_bExpanded = TRUE;
	}

    return (m_bExpanded);
}

BOOL CEmployee::GetEmployeeData()
{
	IList	  * pList;
	_bstr_t		sSQL;
	CStrFnc		sTemp;
	BOOL		bUseODBC;
	VARIANT		v;
	VARIANT		vItemData;
	BOOL		bSuccess;
	HRESULT		hResult = NOERROR;
	IListBank * pAdminLstBanks;

	CIEnumVariant * pForEach;
	IDependant * pDep;
	IBenefitLink * pBnfLnk;
	long	lCount;
	ULONG	uCount;
	long	lMin;
	long	lMax;
	long	i;

	VARIANT varTemp;
	
	lCount = 0;
	bUseODBC = FALSE;
	bSuccess = FALSE;
	hResult = NOERROR;

	hResult = m_poAdmin->get_UseODBC(&bUseODBC);
	if (hResult != NOERROR)
		return (false);

	if (bUseODBC)
		sSQL = CStrFnc::LoadResString(scGet_Employee_SQL);
	else
		sSQL = CStrFnc::LoadResString(scGet_Employee_MDB);

	sTemp.Format(TEXT(" %ld"), m_lPKId); 
	sSQL += sTemp.GetBstr();

	hResult = ::CoCreateInstance(CLSID_CList, NULL, THISCLSCTX, IID_IList, (void **)&pList);
	if (hResult != NOERROR)
		return (false);

	hResult = m_poAdmin->ExecFillList(sSQL, &pList, &bSuccess);
	if (hResult == NOERROR)
	{
		if (bSuccess)
		{
			::VariantInit(&vItemData);
			hResult = pList->get_Item(TEXT("FirstName"), &vItemData);
			if (hResult == NOERROR)
				::SysReAllocString(&m_bstrFirstName, V_BSTR(&vItemData));

			::VariantInit(&vItemData);
			hResult = pList->get_Item(TEXT("LastName"), &vItemData);
			if (hResult == NOERROR)
				::SysReAllocString(&m_bstrLastName, V_BSTR(&vItemData));

			::VariantInit(&vItemData);
			hResult = pList->get_Item(TEXT("Gender"), &vItemData);
			if (hResult == NOERROR)
				m_pbGender = CStrFnc::Boolean(V_BSTR(&vItemData));

			::VariantInit(&vItemData);
			hResult = pList->get_Item(TEXT("DateOfBirth"), &vItemData);
			if (hResult == NOERROR)
				m_vDateOfBirth = CStrFnc::FormatDate(V_BSTR(&vItemData));

			::VariantInit(&vItemData);
			hResult = pList->get_Item(TEXT("Address1"), &vItemData);
			if (hResult == NOERROR)
				::SysReAllocString(&m_bstrAddress1, V_BSTR(&vItemData));

			::VariantInit(&vItemData);
			hResult = pList->get_Item(TEXT("Address2"), &vItemData);
			if (hResult == NOERROR)
				::SysReAllocString(&m_bstrAddress2, V_BSTR(&vItemData));

			::VariantInit(&vItemData);
			hResult = pList->get_Item(TEXT("City"), &vItemData);
			if (hResult == NOERROR)
				::SysReAllocString(&m_bstrCity, V_BSTR(&vItemData));

			if (m_poBank)
			{
				m_poBank->Term();
				m_poBank->Release();
				m_poBank = NULL;
			}
			
			::VariantInit(&vItemData);
			hResult = pList->get_Item(TEXT("BankId"), &vItemData);
			if (hResult == NOERROR)
			{
				hResult = m_poAdmin->get_Banks(&pAdminLstBanks);
				if (hResult == NOERROR)
				{
					if (pAdminLstBanks)
					{
						hResult = pAdminLstBanks->get_Item(CStrFnc::VariantBstrToLong(vItemData), &m_poBank);
						if (hResult != NOERROR)
						{
							if (m_poBank)
							{
								m_poBank->Term();
								m_poBank->Release();
								m_poBank = NULL;
							}
						}
						pAdminLstBanks->Release();
					}
				}
			}

			::VariantInit(&vItemData);
			hResult = pList->get_Item(TEXT("BankAccount"), &vItemData);
			if (hResult == NOERROR)
				::SysReAllocString(&m_bstrBankAccount, V_BSTR(&vItemData));

			::VariantInit(&vItemData);
			hResult = pList->get_Item(TEXT("Position"), &vItemData);
			if (hResult == NOERROR)
				::SysReAllocString(&m_bstrPosition, V_BSTR(&vItemData));

			::VariantInit(&vItemData);
			hResult = pList->get_Item(TEXT("Salary"), &vItemData);
			if (hResult == NOERROR)
				m_cySalary = CStrFnc::BstrToCy(V_BSTR(&vItemData));

			::VariantInit(&vItemData);
			hResult = pList->get_Item(TEXT("StartDate"), &vItemData);
			if (hResult == NOERROR)
				m_vdtStartDate = CStrFnc::FormatDate(V_BSTR(&vItemData));

			::VariantInit(&vItemData);
			hResult = pList->get_Item(TEXT("EndDate"), &vItemData);
			if (hResult == NOERROR)
				m_vdtEndDate = CStrFnc::FormatDate(V_BSTR(&vItemData));

			::VariantInit(&vItemData);
			hResult = pList->get_Item(TEXT("UserName"), &vItemData);
			if (hResult == NOERROR)
				::SysReAllocString(&m_bstrUserName, V_BSTR(&vItemData));

			pList->Release();

			::VariantInit(&varTemp);

			if (m_plstDependants)
			{
				hResult = m_plstDependants->get__NewEnum((IUnknown **)&pForEach);
				::VariantInit(&varTemp);
				hResult = pForEach->Next(1, &varTemp, &uCount);
				while (hResult == NOERROR)
				{
					pDep = (IDependant *)V_DISPATCH(&varTemp);
					pDep->Term();
					pDep->Release();
					::VariantClear(&varTemp);
					hResult = pForEach->Next(1, &varTemp, &uCount);
				}
				m_plstDependants->Release();
				m_plstDependants = NULL;
			}

			hResult = ::CoCreateInstance(CLSID_CListDependant, NULL, THISCLSCTX, IID_IListDependant, (void **)&m_plstDependants);
			if (hResult != NOERROR)
			{
				if (m_plstDependants)
				{
					m_plstDependants->Release();
					m_plstDependants = NULL;
				}

				return (false);
			}

			if (bUseODBC)
				sSQL = CStrFnc::LoadResString(scLvw_DependantByEmployeeId_SQL);
			else
				sSQL = CStrFnc::LoadResString(scLvw_DependantByEmployeeId_MDB);

			sTemp.Empty();
			sTemp.Format(TEXT(" %ld"), m_lPKId); 
			sSQL += sTemp.GetBstr();

			::VariantInit(&v);
			v.vt = VT_BSTR | VT_ARRAY;
			
			hResult = m_poAdmin->ExecFillArray(sSQL, &v, false, &bSuccess);
			if (hResult == NOERROR)
			{
				if (bSuccess)
				{
					CSafeArray SA(V_ARRAY(&v));
					lMin = SA.LBound(2);
					lMax = SA.UBound(2);
					for (i = lMin; i < lMax; i++)
					{
						IDependant * pDep;
						pDep = NULL;
						hResult = ::CoCreateInstance(CLSID_CDependant, NULL, THISCLSCTX, IID_IDependant, (void **)&pDep);
						if (pDep)
						{
							pDep->putref_Admin(m_poAdmin);
							::VariantInit(&v);
							v.vt = VT_BSTR;
							V_BSTR(&v) = SA.bstrGet((long)icLvwPKId, (long)i);
							pDep->put_PKId(CStrFnc::VariantBstrToLong(v));
							m_plstDependants->Add(pDep);
							pDep->Release();
						}
					}
				}
			}

			::VariantInit(&varTemp);

			if (m_plstBenefitLinks)
			{
				hResult = m_plstBenefitLinks->get__NewEnum((IUnknown **)&pForEach);
				::VariantInit(&varTemp);
				hResult = pForEach->Next(1, &varTemp, &uCount);
				while (hResult == NOERROR)
				{
					pBnfLnk = (IBenefitLink *)V_DISPATCH(&varTemp);
					pBnfLnk->Term();
					pBnfLnk->Release();
					::VariantClear(&varTemp);
					hResult = pForEach->Next(1, &varTemp, &uCount);
				}
				m_plstBenefitLinks->Release();
				m_plstBenefitLinks = NULL;
			}

			hResult = ::CoCreateInstance(CLSID_CListBenefitLink, NULL, THISCLSCTX, IID_IListBenefitLink, (void **)&m_plstBenefitLinks);
			if (hResult != NOERROR)
				return (false);

			if (bUseODBC)
				sSQL = CStrFnc::LoadResString(scLvw_EmplBnfByEmployeeId_SQL);
			else
				sSQL = CStrFnc::LoadResString(scLvw_EmplBnfByEmployeeId_MDB);

			sTemp.Empty();
			sTemp.Format(TEXT(" %ld"), m_lPKId); 
			sSQL += sTemp.GetBstr();

			::VariantInit(&v);
			v.vt = VT_BSTR | VT_ARRAY;
			hResult = m_poAdmin->ExecFillArray(sSQL, &v, false, &bSuccess);
			if (hResult == NOERROR)
			{
				if (bSuccess)
				{
					CSafeArray SA(V_ARRAY(&v));
					lMin = SA.LBound(2);
					lMax = SA.UBound(2);
					for (i = lMin; i < lMax; i++)
					{
						IBenefitLink * pBL;
						pBL = NULL;
						hResult = ::CoCreateInstance(CLSID_CBenefitLink, NULL, THISCLSCTX, IID_IBenefitLink, (void **)&pBL);
						if (pBL)
						{
							pBL->putref_Admin(m_poAdmin);
							::VariantInit(&v);
							v.vt = VT_BSTR;
							V_BSTR(&v) = SA.bstrGet((long)icLvwPKId, (long)i);
							pBL->put_PKId(CStrFnc::VariantBstrToLong(v));
							::VariantInit(&v);
							m_plstBenefitLinks->Add(pBL);
							pBL->Release();
						}
					}
				}
			}
			return (TRUE);
		}
	}
	return (false);
}

STDMETHODIMP CEmployee::get_DateOfBirth(DATE * pvdtDateOfBirth)
{
	if (m_bExpanded)
		*pvdtDateOfBirth = m_vDateOfBirth;
	else
	{
		if (Expand())
			*pvdtDateOfBirth = m_vDateOfBirth;
		else
		{
			Term();
			return (FlagException(IID_IEmployee, scEmplExpand, icErrExpand));
		}
	}

	return (NOERROR);
}

STDMETHODIMP CEmployee::put_DateOfBirth(DATE pvdtDateOfBirth)
{
	m_vDateOfBirth = pvdtDateOfBirth;
	return (NOERROR);
}

STDMETHODIMP CEmployee::get_FirstName(BSTR FAR* pbstrFirstName)
{
	if (m_bExpanded)
		::SysReAllocString(pbstrFirstName, m_bstrFirstName);
	else
	{
		if (Expand())
			::SysReAllocString(pbstrFirstName, m_bstrFirstName);
		else
		{
			Term();
			return (FlagException(IID_IEmployee, scEmplExpand, icErrExpand));
		}
	}

	return (NOERROR);
}

STDMETHODIMP CEmployee::put_FirstName(BSTR bstrFirstName)
{
	::SysReAllocString(&m_bstrFirstName, bstrFirstName);
	return (NOERROR);
}

STDMETHODIMP CEmployee::get_LastName(BSTR FAR* pbstrLastName)
{
	if (m_bExpanded)
		::SysReAllocString(pbstrLastName, m_bstrLastName);
	else
	{
		if (Expand())
			::SysReAllocString(pbstrLastName, m_bstrLastName);
		else
		{
			Term();
			return (FlagException(IID_IEmployee, scEmplExpand, icErrExpand));
		}
	}

	return (NOERROR);
}

STDMETHODIMP CEmployee::put_LastName(BSTR bstrLastName)
{
	::SysReAllocString(&m_bstrLastName, bstrLastName);
	return (NOERROR);
}

void CEmployee::Clear()
{
	CIEnumVariant * pForEach;
	IDependant *	pDep;
	IBenefitLink *	pBnfLnk;
	long	lCount;
	HRESULT hResult;
	ULONG	uCount;
	VARIANT varTemp;
	
	lCount = 0;
	hResult = NOERROR;

	::VariantInit(&varTemp);

	if (m_plstBenefitLinks)
	{
		hResult = m_plstBenefitLinks->get__NewEnum((IUnknown **)&pForEach);
		::VariantInit(&varTemp);
		hResult = pForEach->Next(1, &varTemp, &uCount);
		while (hResult == NOERROR)
		{
			pBnfLnk = (IBenefitLink *)V_DISPATCH(&varTemp);
			pBnfLnk->Term();
			pBnfLnk->Release();
			::VariantClear(&varTemp);
			hResult = pForEach->Next(1, &varTemp, &uCount);
		}
		m_plstBenefitLinks->Release();
		m_plstBenefitLinks = NULL;
	}
	
	::VariantInit(&varTemp);

	if (m_plstDependants)
	{
		hResult = m_plstDependants->get__NewEnum((IUnknown **)&pForEach);
		::VariantInit(&varTemp);
		hResult = pForEach->Next(1, &varTemp, &uCount);
		while (hResult == NOERROR)
		{
			pDep = (IDependant *)V_DISPATCH(&varTemp);
			pDep->Term();
			pDep->Release();
			::VariantClear(&varTemp);
			hResult = pForEach->Next(1, &varTemp, &uCount);
		}
		m_plstDependants->Release();
		m_plstDependants = NULL;
	}

	if (m_poBank)
	{
		m_poBank->Term();
		m_poBank->Release();
		m_poBank = NULL;
	}

	m_pbGender = FALSE;
	m_vDateOfBirth = NULL;
	m_vdtStartDate = NULL;
	m_vdtEndDate = NULL;
	m_cySalary.int64 = 0;
	EmptyBSTR(m_bstrFirstName);
	EmptyBSTR(m_bstrLastName);
	EmptyBSTR(m_bstrAddress1);
	EmptyBSTR(m_bstrAddress2);
	EmptyBSTR(m_bstrCity);
	EmptyBSTR(m_bstrBankAccount);
	EmptyBSTR(m_bstrPosition);
	EmptyBSTR(m_bstrUserName);

	m_bExpanded = FALSE;
	return;
}


//Constructor/Destructor
//
//Turn off warning 4355
#pragma warning (disable : 4355)
CEmployee::CEmployee(	LPUNKNOWN pUnkOuter, PFNDESTROYED pfnDestroy)
				: m_SupportErrorInfo(this, IID_IEmployee)
#pragma warning (default : 4355)
{
	m_oid		= g_oid++;
	m_cRef		= 0;
	m_pUnkOuter = pUnkOuter;
	m_pfnDestroy= pfnDestroy;
	m_bExpanded = FALSE;
	m_poAdmin	= (IAdmin *)NULL;
	m_lPKId		= 0L;
	m_bstrFirstName = (BSTR)NULL;
	m_bstrLastName	= (BSTR)NULL;
	m_pbGender	= 0;
	m_vDateOfBirth = NULL;
	m_bstrAddress1 = (BSTR)NULL;
	m_bstrAddress2 = (BSTR)NULL;
	m_bstrCity	= (BSTR)NULL;
	m_poBank	= (IBank *)NULL;
	m_bstrBankAccount = (BSTR)NULL;
	m_bstrPosition	 = (BSTR)NULL;
	m_cySalary.int64 = 0;
	m_vdtStartDate	 = NULL;
	m_vdtEndDate	 = NULL;
	m_bstrUserName	 = (BSTR)NULL;
	m_plstDependants = (IListDependant *)NULL;
	m_plstBenefitLinks = (IListBenefitLink *)NULL;
	m_pITypeInfo = NULL;

	return;
}

CEmployee::~CEmployee(void)
{
	Term();

	if (m_poAdmin)
	{
		m_poAdmin->Release();
		m_poAdmin = NULL;
	}

	if (m_pITypeInfo)
	{
		m_pITypeInfo->Release();
		m_pITypeInfo = NULL;
	}
	return;
}


BOOL CEmployee::CFInit(void)
{
	LPUNKNOWN pIUnknown = this;

	if (m_pUnkOuter)
		pIUnknown = m_pUnkOuter;

	return (TRUE);
}

////////////////////////////////////////////
// Standard IUnknown method implementations
////////////////////////////////////////////

//CEmployee::QueryInterface
//
//Purpose     - Used to return an interface pointer from the requested interface id.
//Inputs      - (REFIID)  riid
//              (void **) ppvCEmployee
STDMETHODIMP CEmployee::QueryInterface(REFIID riid, void ** ppvCEmployee)
{
	*ppvCEmployee = NULL;

	if (riid == IID_IEmployee || riid == IID_IUnknown || riid == IID_IDispatch)
		*ppvCEmployee = this;

	if (riid == IID_ISupportErrorInfo)
		*ppvCEmployee = &m_SupportErrorInfo;

	if (*ppvCEmployee)
	{
		((LPUNKNOWN)*ppvCEmployee)->AddRef();
		return (NOERROR);
	}

	return (E_NOINTERFACE);
}

STDMETHODIMP_(ULONG) CEmployee::AddRef(void)
{
	++m_cRef;
	Audit(m_oid, 1, TEXT("CEmployee"), m_cRef);
	return (m_cRef);
}

STDMETHODIMP_(ULONG) CEmployee::Release(void)
{
	--m_cRef;
	Audit(m_oid, 2, TEXT("CEmployee"), m_cRef);

	if (m_cRef == 0)
	{
		if (m_pfnDestroy)
			(*m_pfnDestroy)();
		Audit(m_oid, 0, TEXT("CEmployee gbldec"), g_lObjectCount);
		delete this;
		return (0);
	}
	else
		return (m_cRef);
}

//////////////////////////////////////////////
// Standard IDispatch method implementations
//////////////////////////////////////////////

STDMETHODIMP CEmployee::GetTypeInfoCount(UINT *pctInfo)
{
	*pctInfo = 1;

	return (NOERROR);
}

STDMETHODIMP CEmployee::GetTypeInfo(UINT itInfo, LCID lcid, ITypeInfo **ppITypeInfo)
{
	HRESULT     hResult;
	ITypeLib   *pITypeLib;
	ITypeInfo **ppITI;

	hResult = NOERROR;
	ppITI	= NULL;

	if (itInfo != 0)
		return (TYPE_E_ELEMENTNOTFOUND);

	if (ppITypeInfo == NULL)
		return (E_POINTER);

	*ppITypeInfo = NULL;

	switch (PRIMARYLANGID(lcid))
	{
		case LANG_NEUTRAL:
		case LANG_ENGLISH:
			ppITI = &m_pITypeInfo;
			break;

		default:
			return (DISP_E_UNKNOWNLCID);
	}

	if (*ppITI == NULL)
	{
#ifdef _BENDLL
		hResult = LoadRegTypeLib(LIBID_BenefitsCPP, 1, 0, PRIMARYLANGID(lcid), &pITypeLib);
#else
		hResult = LoadRegTypeLib(LIBID_BenefitsCPPRemote, 1, 0, PRIMARYLANGID(lcid), &pITypeLib);
#endif
		if (hResult != NOERROR)
		{
			switch (PRIMARYLANGID(lcid))
			{
				case LANG_NEUTRAL:
				case LANG_ENGLISH:
#ifdef _BENDLL
					hResult = LoadTypeLib(TEXT("bencpp.tlb"), &pITypeLib);
#else
					hResult = LoadTypeLib(TEXT("bencppr.tlb"), &pITypeLib);
#endif
					break;
			}
		}

		if (hResult != NOERROR)
			return (hResult);

		hResult = pITypeLib->GetTypeInfoOfGuid(IID_IEmployee, ppITI);
		
		m_pITypeInfo = *ppITI;
		m_pITypeInfo->AddRef();
		
		pITypeLib->Release();

		if (hResult != NOERROR)
			return (hResult);
	}

	(*ppITI)->AddRef();
	*ppITypeInfo = *ppITI;

	return (NOERROR);
}

STDMETHODIMP CEmployee::GetIDsOfNames(REFIID riid, OLECHAR **rgszNames,
									  UINT cNames, LCID lcid, DISPID *rgDispID)
{
	HRESULT     hResult;
	ITypeInfo  *pTI;

	hResult = NOERROR;

	if (riid != IID_NULL)
		return (DISP_E_UNKNOWNINTERFACE);

	hResult = GetTypeInfo(0, lcid, &pTI);

	if (hResult == NOERROR)
	{
		hResult = DispGetIDsOfNames(pTI, rgszNames, cNames, rgDispID);
		pTI->Release();
	}

	return (hResult);
}

STDMETHODIMP CEmployee::Invoke (DISPID dispID,	REFIID riid, LCID lcid,
				 unsigned short wFlags, DISPPARAMS *pDispParams,
				 VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
	HRESULT     hResult;
	ITypeInfo  *pTI;
	LANGID      langID;

	hResult = NOERROR;
	langID	= PRIMARYLANGID(lcid);

	if (riid != IID_NULL)
		return (DISP_E_UNKNOWNINTERFACE);

	hResult = GetTypeInfo(0, lcid, &pTI);

	if (hResult != NOERROR)
		return (hResult);

	hResult = DispInvoke(this, pTI, dispID, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr);

	pTI->Release();
	return (hResult);
}

CEmployeeFactory::CEmployeeFactory(void)
{
	m_oid = g_oid++;
	m_cRef = 0L;
	return;
}

CEmployeeFactory::~CEmployeeFactory(void)
{
	return;
}

//CEmployeeFactory::QueryInterface
//Purpose     - Used to return an interface pointer from the requested interface id.
//Inputs      - (REFIID)  riid
//              (void **) ppvCEmployeefIf
STDMETHODIMP CEmployeeFactory::QueryInterface(REFIID riid, void ** ppvCEmployeeIf)
{
	*ppvCEmployeeIf = NULL;

	if (riid == IID_IUnknown ||
		riid == IID_IClassFactory)
		*ppvCEmployeeIf = this;

	if (*ppvCEmployeeIf)
	{
		((LPUNKNOWN)*ppvCEmployeeIf)->AddRef();
		return (NOERROR);
	}

	return (E_NOINTERFACE);
}

STDMETHODIMP_(ULONG) CEmployeeFactory::AddRef(void)
{
	++m_cRef;
	Audit(m_oid, 1, TEXT("CEmployeeF"), m_cRef);
	return (m_cRef);
}

STDMETHODIMP_(ULONG) CEmployeeFactory::Release(void)
{
	--m_cRef;
	Audit(m_oid, 2, TEXT("CEmployeeF"), m_cRef);
	if (m_cRef == 0)
	{
		delete this;
		return (0);
	}
	else
		return (m_cRef);
}

//CEmployeeFactory::CreateInstance
//Purpose     - Creates a CBenefit object and returns an Interface pointer
//Inputs      - (LPUNKNOWN) pUnkOuter
//              (REFIID)    riid
//				(void **)   ppvCBenefIf
STDMETHODIMP CEmployeeFactory::CreateInstance(LPUNKNOWN pUnkOuter, REFIID riid, void ** ppvCEmployeeIf)
{
	PCEmployee	pCEmployee;
	HRESULT		hResult;

	hResult		= NOERROR;
	*ppvCEmployeeIf	= NULL;

	if (pUnkOuter && riid != IID_IUnknown)
		return (CLASS_E_NOAGGREGATION);

	pCEmployee = new CEmployee(pUnkOuter, ObjectDestroyed);

	if (pCEmployee == NULL)
		return (hResult);

	if (pCEmployee->CFInit())
		hResult = pCEmployee->QueryInterface(riid, ppvCEmployeeIf);

	if (hResult != NOERROR)
		delete pCEmployee;
	else
	{
		g_lObjectCount++;
		Audit(m_oid, 0, TEXT("CEmployee gblinc"), g_lObjectCount);
	}

	return (hResult);
}

//CEmployeeFactory::LockServer
//Purpose     - Increments/Decrements the lock count to prevent class factory server shutdown.
//Inputs      - (BOOL)    fLock
STDMETHODIMP CEmployeeFactory::LockServer(BOOL fLock)
{
	if (fLock)
		g_lLockedObjects++;
	else
		g_lLockedObjects--;

	return (NOERROR);
}

