//DllFuncs.cpp  - Copyright (C) 1996 Microsoft Corporation
//            - All rights reserved.
//Date        - 10/19/96
//Authors     - Ken Bergmann, Robert Coleridge
//Purpose     - Implementation routines for the DLL.

#pragma warning (disable : 4800)	// turns off bool warning

#include <objbase.h>
#include <initguid.h>

#define INITGUIDS

#include "cadmin.h"
#include "cbank.h"
#include "cbenefit.h"
#include "cbenefitlink.h"
#include "cdependent.h"
#include "cemployee.h"
#include "clist.h"
#include "clistbank.h"
#include "clistbenefit.h"
#include "clistbenefitlink.h"
#include "clistdependant.h"
#include "clistemployee.h"

#include "benefits.h"

// leave under other includes to avoid conflicts with DAO350.tli includes
#include <stdio.h>
#include <stdlib.h>

ULONG		g_lObjectCount = 0;
ULONG		g_lLockedObjects = 0;
HINSTANCE	g_hInst = NULL;
long		g_oid = 1;

void ObjectDestroyed(void);

#ifdef __cplusplus
extern "C"
{
#endif
	
HRESULT APIENTRY DllGetClassObject(REFCLSID rclsid, REFIID riid, void ** ppv);
STDAPI DllCanUnloadNow(void);
STDAPI DllRegisterServer(void);
STDAPI DllUnregisterServer(void);

BOOL SetKeyAndValue(LPSTR pszKey, LPSTR pszSubkey, LPCSTR lpValueName, LPSTR pszValue);
BOOL RegisterServer();
BOOL UnregisterServer();
BOOL RegisterMyClass(LPSTR lpProgId, LPSTR lpLibId, CLSID ClsId, LPSTR lpDescription);
BOOL UnregisterMyClass(LPSTR lpProgId, CLSID ClsId);

//DllMain
//Purpose		- provides standard DLLMain entry point (see Win32 docs for details)
BOOL WINAPI DllMain(HANDLE hInstance, ULONG ulReason, void * pvReserved)
{
	g_hInst = hInstance;
	
	if (ulReason == DLL_PROCESS_DETACH)
		return (TRUE);
	else
	{
		if (ulReason != DLL_PROCESS_ATTACH)
			return (TRUE);
	}		
	return (TRUE);
}

//DllGetClassObject
//Purpose		- provides entry point for system to request object instantiation
//Inputs		- (rclsid)		- class id to get object from
//				- (riid)		- interface id to obtain
//				- (ppv)			- pointer to variable to hold returned object
//Outputs		- (HRESULT)		- standard COM/OLE result code
HRESULT APIENTRY DllGetClassObject(REFCLSID rclsid, REFIID riid, void ** ppv)
{
	HRESULT  hResult = NOERROR;
	
	if (rclsid == CLSID_CAdmin
		|| rclsid == CLSID_CBank
		|| rclsid == CLSID_CBenefit
		|| rclsid == CLSID_CBenefitLink
		|| rclsid == CLSID_CDependant
		|| rclsid == CLSID_CList
		|| rclsid == CLSID_CListBank
		|| rclsid == CLSID_CListBenefit
		|| rclsid == CLSID_CListBenefitLink
		|| rclsid == CLSID_CListDependant
		|| rclsid == CLSID_CListEmployee
		|| rclsid == CLSID_CEmployee)
	{
	}
	else
		return (E_FAIL);
	
	if (riid != IID_IUnknown && riid != IID_IClassFactory)
		return (E_NOINTERFACE);
	
	// replicate code to create new class factory
	if (rclsid == CLSID_CAdmin)
	{
		CAdminFactory *pClassFactory;
		
		pClassFactory = new CAdminFactory();
		
		if (pClassFactory == NULL)
			return (E_OUTOFMEMORY);
		
		hResult = pClassFactory->QueryInterface(riid, ppv);
		
		if (hResult != NOERROR)
			delete pClassFactory;
	}
	else
	{
		if (rclsid == CLSID_CBank)
		{
			CBankFactory *pClassFactory;
			
			pClassFactory = new CBankFactory();
			
			if (pClassFactory == NULL)
				return (E_OUTOFMEMORY);
			
			hResult = pClassFactory->QueryInterface(riid, ppv);
			
			if (hResult != NOERROR)
				delete pClassFactory;
		}
		else
		{
			if (rclsid == CLSID_CBenefit)
			{
				CBenefitFactory *pClassFactory;
				
				pClassFactory = new CBenefitFactory();
				
				if (pClassFactory == NULL)
					return (E_OUTOFMEMORY);
				
				hResult = pClassFactory->QueryInterface(riid, ppv);
				
				if (hResult != NOERROR)
					delete pClassFactory;
			}
			else
			{
				if (rclsid == CLSID_CDependant)
				{
					CDependantFactory *pClassFactory;
					
					pClassFactory = new CDependantFactory();
					
					if (pClassFactory == NULL)
						return (E_OUTOFMEMORY);
					
					hResult = pClassFactory->QueryInterface(riid, ppv);
					
					if (hResult != NOERROR)
						delete pClassFactory;
				}
				else
				{
					if (rclsid == CLSID_CEmployee)
					{
						CEmployeeFactory *pClassFactory;
						
						pClassFactory = new CEmployeeFactory();
						
						if (pClassFactory == NULL)
							return (E_OUTOFMEMORY);
						
						hResult = pClassFactory->QueryInterface(riid, ppv);
						
						if (hResult != NOERROR)
							delete pClassFactory;
					}
					else
					{
						if (rclsid == CLSID_CBenefitLink)
						{
							CBenefitLinkFactory *pClassFactory;
							
							pClassFactory = new CBenefitLinkFactory();
							
							if (pClassFactory == NULL)
								return (E_OUTOFMEMORY);
							
							hResult = pClassFactory->QueryInterface(riid, ppv);
							
							if (hResult != NOERROR)
								delete pClassFactory;
						}
						else
						{
							if (rclsid == CLSID_CList)
							{
								CListFactory *pClassFactory;
								
								pClassFactory = new CListFactory();
								
								if (pClassFactory == NULL)
									return (E_OUTOFMEMORY);
								
								hResult = pClassFactory->QueryInterface(riid, ppv);
								
								if (hResult != NOERROR)
									delete pClassFactory;
							}
							else
							{
								if (rclsid == CLSID_CListBank)
								{
									CListBankFactory *pClassFactory;
									
									pClassFactory = new CListBankFactory();
									
									if (pClassFactory == NULL)
										return (E_OUTOFMEMORY);
									
									hResult = pClassFactory->QueryInterface(riid, ppv);
									
									if (hResult != NOERROR)
										delete pClassFactory;
								}
								else
								{
									if (rclsid == CLSID_CListBenefit)
									{
										CListBenefitFactory *pClassFactory;
										
										pClassFactory = new CListBenefitFactory();
										
										if (pClassFactory == NULL)
											return (E_OUTOFMEMORY);
										
										hResult = pClassFactory->QueryInterface(riid, ppv);
										
										if (hResult != NOERROR)
											delete pClassFactory;
									}
									else
									{
										if (rclsid == CLSID_CListBenefitLink)
										{
											CListBenefitLinkFactory *pClassFactory;
											
											pClassFactory = new CListBenefitLinkFactory();
											
											if (pClassFactory == NULL)
												return (E_OUTOFMEMORY);
											
											hResult = pClassFactory->QueryInterface(riid, ppv);
											
											if (hResult != NOERROR)
												delete pClassFactory;
										}
										else
										{
											if (rclsid == CLSID_CListDependant)
											{
												CListDependantFactory *pClassFactory;
												
												pClassFactory = new CListDependantFactory();
												
												if (pClassFactory == NULL)
													return (E_OUTOFMEMORY);
												
												hResult = pClassFactory->QueryInterface(riid, ppv);
												
												if (hResult != NOERROR)
													delete pClassFactory;
											}
											else
											{
												if (rclsid == CLSID_CListEmployee)
												{
													CListEmployeeFactory *pClassFactory;
													
													pClassFactory = new CListEmployeeFactory();
													
													if (pClassFactory == NULL)
														return (E_OUTOFMEMORY);
													
													hResult = pClassFactory->QueryInterface(riid, ppv);
													
													if (hResult != NOERROR)
														delete pClassFactory;
												}
											}
										}
									}
								}
							}
						}
					}
				}
			}
		}
	}												
	return (hResult);
}

//DllCanUnloadNow
//Purpose		- provides entry point for system to query if dll can be unloaded
//Outputs		- true or false if can be unloaded
STDAPI DllCanUnloadNow(void)
{
	if (g_lObjectCount == 0l && g_lLockedObjects == 0L)
		return (S_OK);
	else
		return (S_FALSE);
}

//DllRegisterServer
//Purpose		- provides entry point for system to register server via RegSvr32 or Setup
//Outputs		- true or false if can be unloaded
STDAPI DllRegisterServer(void)
{
	if (!RegisterServer())
		return (SELFREG_E_CLASS);
	
	return (S_OK);
}

//DllUnregisterServer
//Purpose		- provides entry point for system to unregister server via RegSvr32 or Setup
//Outputs		- true or false if can be unloaded
STDAPI DllUnregisterServer(void)
{
	if (!UnregisterServer())
		return (SELFREG_E_CLASS);
	
	return (NOERROR);
}

/////////////////////////////////////////////////////////////////////////////
// Helper to work with the Registry

//SetKeyAndValue
//Purpose		- to create and set a specified registry key
//Inputs		- (pszKey)		- top level registry key name
//				- (pszSubkey)	- sub key to set
//				- (lpValueName)	- value name to set
//				- (pszValue)	- value to set sub key to
//Outputs		- (BOOL)		- true if successful, false if not
BOOL SetKeyAndValue(LPSTR pszKey, LPSTR pszSubkey, LPCSTR lpValueName, LPSTR pszValue)
{
	HKEY	hKey;
	CHAR	szKey[256];
	
	strcpy(szKey, pszKey);
	
	if (pszSubkey)
	{
		strcat(szKey, "\\");
		strcat(szKey, pszSubkey);
	}
	
	if (RegCreateKeyExA(
			HKEY_CLASSES_ROOT,
			szKey,
			0,
			NULL,
			REG_OPTION_NON_VOLATILE,
			KEY_ALL_ACCESS,
			NULL, 
			&hKey,
			NULL) != ERROR_SUCCESS)
		return (FALSE);
	
	if (pszValue)
		RegSetValueExA(hKey, lpValueName, 0, REG_SZ, (BYTE *)pszValue,  strlen(pszValue)+1);
	
	RegCloseKey(hKey);
	return (TRUE);
}

//RegisterServer
//Purpose		- to perform all registrations necessary to register server
//Outputs		- (BOOL)		- true if successful, false if not
BOOL RegisterServer()
{
	BOOL		bResult		= TRUE;
	HRESULT		hResult		= NOERROR;
	ITypeLib*	pITypeLib	= NULL;
	LPTSTR		pszIDLib;
	CHAR		szCLSID[_MAX_PATH];
	CHAR		szIDLib[_MAX_PATH];
	CHAR		szModule[_MAX_PATH+10];
	TCHAR		szOLEModule[_MAX_PATH + 10];
	
	StringFromGUID2(CLSID_CAdmin, (LPOLESTR)szCLSID, sizeof(szCLSID));
	StringFromCLSID(LIBID_BenefitsCPP, &pszIDLib);
	
	::WideCharToMultiByte(CP_ACP, 0, pszIDLib, -1, szIDLib, sizeof(szIDLib), NULL, NULL);

	if (!RegisterMyClass("BenefitsCPP.CAdmin", szIDLib, CLSID_CAdmin, "Benefits Admin Class (Ver 1.0)"))
		bResult = FALSE;
	
	if (!RegisterMyClass("BenefitsCPP.CBank", szIDLib, CLSID_CBank, "Benefits Bank Class (Ver 1.0)"))
		bResult = FALSE;
	
	if (!RegisterMyClass("BenefitsCPP.CBenefit", szIDLib, CLSID_CBenefit, "Benefits Benefit Class (Ver 1.0)"))
		bResult = FALSE;
	
	if (!RegisterMyClass("BenefitsCPP.CBenefitLink", szIDLib, CLSID_CBenefitLink, "Benefits Benefit Link Class (Ver 1.0)"))
		bResult = FALSE;
	
	if (!RegisterMyClass("BenefitsCPP.CDependant", szIDLib, CLSID_CDependant, "Benefits Dependant Class (Ver 1.0)"))
		bResult = FALSE;
	
	if (!RegisterMyClass("BenefitsCPP.CEmployee", szIDLib, CLSID_CEmployee, "Benefits Employee Class (Ver 1.0)"))
		bResult = FALSE;
	
	if (!RegisterMyClass("BenefitsCPP.CList", szIDLib, CLSID_CList, "Benefits List Class (Ver 1.0)"))
		bResult = FALSE;
	
	if (!RegisterMyClass("BenefitsCPP.CListBank", szIDLib, CLSID_CListBank, "Benefits List Bank Class (Ver 1.0)"))
		bResult = FALSE;
	
	if (!RegisterMyClass("BenefitsCPP.CListBenefit", szIDLib, CLSID_CListBenefit, "Benefits List Benefit Class (Ver 1.0)"))
		bResult = FALSE;
	
	if (!RegisterMyClass("BenefitsCPP.CListBenefitLink", szIDLib, CLSID_CListBenefitLink, "Benefits List Benefit Link Class (Ver 1.0)"))
		bResult = FALSE;
	
	if (!RegisterMyClass("BenefitsCPP.CListDependant", szIDLib, CLSID_CListDependant, "Benefits List Dependant Class (Ver 1.0)"))
		bResult = FALSE;
	
	if (!RegisterMyClass("BenefitsCPP.CListEmployee", szIDLib, CLSID_CListEmployee, "Benefits List Employee Class (Ver 1.0)"))
		bResult = FALSE;

	ZeroMemory(szModule, sizeof (szModule));
	GetModuleFileNameA(g_hInst, szModule, _MAX_PATH+10);

	::MultiByteToWideChar(CP_ACP, 0, szModule, -1, szOLEModule, _MAX_PATH + 10);
	
	hResult = ::LoadTypeLib((LPOLESTR)szOLEModule, &pITypeLib);
	if (hResult != NOERROR)
		return (false);
	
	hResult = ::RegisterTypeLib(pITypeLib, (LPOLESTR)szOLEModule, NULL);//HKCR\TypeLib entry
	if (hResult != NOERROR)
		bResult = FALSE;
	else
		pITypeLib->Release();
	
	return (bResult);
}

//RegisterMyClass
//Purpose		- helper function to assist in registering a specific registry key
//Inputs		- (lpProgId)		- prog id of key to register
//				- (lpLibId)			- lib id of key to register
//				- (ClsId)			- GUID of object to register
//				- (lpDescription)	- description of this object
//Outputs		- (BOOL)		- true if successful, false if not
BOOL RegisterMyClass(LPSTR lpProgId, LPSTR lpLibId, CLSID ClsId, LPSTR lpDescription)
{
	BOOL	bResult = TRUE;
	HRESULT	hResult = NOERROR;
	LPTSTR	pszOLEID;
	CHAR	szCLSID[_MAX_PATH+10];
	CHAR	szID[_MAX_PATH+10];
	CHAR	szModule[_MAX_PATH+10];
	
	StringFromCLSID(ClsId, &pszOLEID);
	hResult = ::WideCharToMultiByte(
					CP_ACP,
					0,
					pszOLEID,
					-1,
					szID,
					sizeof(szID), NULL, NULL);

	if (!SetKeyAndValue(lpProgId, NULL, NULL, lpProgId))
		bResult = FALSE;
	
	if (!SetKeyAndValue(lpProgId, "CLSID", NULL, szID))
		bResult = FALSE;

	strcpy(szCLSID, "CLSID\\");
	strcat(szCLSID, szID);
	if (!SetKeyAndValue(szCLSID, NULL, NULL, lpDescription))
		bResult = FALSE;
	if (!SetKeyAndValue(szCLSID, "TypeLib", NULL, lpLibId))
		bResult = FALSE;
	
	ZeroMemory(szModule, sizeof (szModule));
	GetModuleFileNameA(g_hInst, szModule, sizeof(szModule));
	
	if (!SetKeyAndValue(szCLSID, "InProcServer32", NULL, szModule))
		bResult = FALSE;   
	
	if (!SetKeyAndValue(szCLSID, "ProgID", NULL, lpProgId))
		bResult = FALSE;
	
	if (!SetKeyAndValue(szCLSID, "VersionIndependentProgID", NULL, lpProgId))
		bResult = FALSE;
	
	if (!SetKeyAndValue(szCLSID, "Programmable", NULL, ""))
		bResult = FALSE;
	
	return (bResult);
}

//UnregisterServer
//Purpose		- to unregister the entre library
//Outputs		- (BOOL)		- true if successful, false if not
BOOL UnregisterServer()
{
	BOOL  bResult = TRUE;
	
	HRESULT hResult = ::UnRegisterTypeLib(LIBID_BenefitsCPP, 1, 0, 0, SYS_WIN32);
	if (hResult != NOERROR)
		return (false);

	if (!UnregisterMyClass("BenefitsCPP.CAdmin", CLSID_CAdmin))
		bResult = FALSE;
	
	if (!UnregisterMyClass("BenefitsCPP.CBank", CLSID_CBank))
		bResult = FALSE;
	
	if (!UnregisterMyClass("BenefitsCPP.CBenefit", CLSID_CBenefit))
		bResult = FALSE;
	
	if (!UnregisterMyClass("BenefitsCPP.CBenefitLink", CLSID_CBenefitLink))
		bResult = FALSE;
	
	if (!UnregisterMyClass("BenefitsCPP.CDependant", CLSID_CDependant))
		bResult = FALSE;
	
	if (!UnregisterMyClass("BenefitsCPP.CEmployee", CLSID_CEmployee))
		bResult = FALSE;
	
	if (!UnregisterMyClass("BenefitsCPP.CList", CLSID_CList))
		bResult = FALSE;
	
	if (!UnregisterMyClass("BenefitsCPP.CListBank", CLSID_CListBank))
		bResult = FALSE;
	
	if (!UnregisterMyClass("BenefitsCPP.CListBenefit", CLSID_CListBenefit))
		bResult = FALSE;
	
	if (!UnregisterMyClass("BenefitsCPP.CListBenefitLink", CLSID_CListBenefitLink))
		bResult = FALSE;
	
	if (!UnregisterMyClass("BenefitsCPP.CListDependant", CLSID_CListDependant))
		bResult = FALSE;
	
	if (!UnregisterMyClass("BenefitsCPP.CListEmployee", CLSID_CListEmployee))
		bResult = FALSE;
	
	return (bResult);
}

//UnregisterMyClass
//Purpose		- to unregister a specified object
//Inputs		- (lpProgId)	- prog id of object to unregister
//				- (ClsId)		- class id of object
//Outputs		- (BOOL)		- true if successful, false if not
BOOL UnregisterMyClass(LPSTR lpProgId, CLSID ClsId)
{
	BOOL	bResult = TRUE;	
	HRESULT hResult;
	LPTSTR	pszOLEID;
	CHAR	szCLSID[_MAX_PATH];
	CHAR	szID[_MAX_PATH];
	CHAR	szTemp[_MAX_PATH];
	
	StringFromCLSID(ClsId, &pszOLEID);
	
	hResult = ::WideCharToMultiByte(
					CP_ACP,
					0,
					pszOLEID,
					-1,
					szID,
					sizeof(szID), NULL, NULL);

	strcpy(szCLSID, "CLSID\\");
	strcat(szCLSID, szID);
	
	sprintf(szTemp, "%s\\%s", szCLSID, "InProcServer32");
	if (RegDeleteKeyA(HKEY_CLASSES_ROOT, szTemp) != ERROR_SUCCESS)
		bResult = FALSE;
	
	sprintf(szTemp, "%s\\%s", szCLSID, "ProgId");
	if (RegDeleteKeyA(HKEY_CLASSES_ROOT, szTemp) != ERROR_SUCCESS)
		bResult = FALSE;
	
	sprintf(szTemp, "%s\\%s", szCLSID, "VersionIndependentProgID");
	if (RegDeleteKeyA(HKEY_CLASSES_ROOT, szTemp) != ERROR_SUCCESS)
		bResult = FALSE;
	
	sprintf(szTemp, "%s\\%s", szCLSID, "Programmable");
	if (RegDeleteKeyA(HKEY_CLASSES_ROOT, szTemp) != ERROR_SUCCESS)
		bResult = FALSE;
	
	sprintf(szTemp, "%s\\%s", szCLSID, "TypeLib");
	if (RegDeleteKeyA(HKEY_CLASSES_ROOT, szTemp) != ERROR_SUCCESS)
		bResult = FALSE;
	
	if (RegDeleteKeyA(HKEY_CLASSES_ROOT, szCLSID) != ERROR_SUCCESS)
		bResult = FALSE;
	
	sprintf(szTemp, "%s\\%s", lpProgId, "CLSID");
	if (RegDeleteKeyA(HKEY_CLASSES_ROOT, szTemp) != ERROR_SUCCESS)
		bResult = FALSE; 
	
	if (RegDeleteKeyA(HKEY_CLASSES_ROOT, lpProgId) != ERROR_SUCCESS)
		bResult = FALSE; 

	return (bResult);
}

#ifdef __cplusplus
}
#endif
	
//ObjectDestroyed
// Purpose:
//  Function for an object to call when it gets destroyed.  Since
//  we're in a DLL we only track the number of objects here, letting
//  DllCanUnloadNow take care of the rest.
void ObjectDestroyed(void)
{
	g_lObjectCount--;
	return;
}
