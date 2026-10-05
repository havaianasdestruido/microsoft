//benexe.cpp  - Copyright (C) 1996 Microsoft Corporation
//            - All rights reserved.
//Date        - 10/19/96
//Authors     - Greg Shumate
//Purpose     - Implementation routines for the Benefits Exe.

#pragma warning (disable : 4800)	// turns off bool warning

#ifdef UNICODE
#define TURNBACKON
#undef UNICODE
#endif

#include <objbase.h>
#include <initguid.h>

#define INITGUIDS
#include "benexe.h"
#include "wchar.h"
#include "tchar.h"

#include "cadmin.h"
#include "cbank.h"
#include "cbenefit.h"
#include "cbenefitlink.h"
#include "cdependent.h"
#include "cemployee.h"
#include "clist.h"
#include "clistbank.h"
#include "clistbenefit.h"
#include "clistbenefitLink.h"
#include "clistdependant.h"
#include "clistemployee.h"
#include "support.h"
#include <stdio.h>

//Make window handle global so other code can cause a shutdown
HWND        g_hWnd = NULL;
HINSTANCE	g_hInst = NULL;
ULONG       g_lObjectCount = 0;
ULONG       g_lLockedObjects = 0;

long g_oid = 1;

extern "C"
{
BOOL RegisterMyClass(LPSTR lpProgId, LPSTR lpLibId, CLSID ClsId, LPSTR lpDescription);
BOOL RegisterServer();
BOOL SetKeyAndValue(LPSTR pszKey, LPSTR pszSubkey, LPCSTR lpValueName, LPSTR pszValue);
BOOL UnregisterMyClass(LPSTR lpProgId, CLSID ClsId);
BOOL UnregisterServer();
}

//////////////////////////////////////////////////////////
// WinMain
//
// Purpose: Main entry point of application.

int PASCAL WinMain
(
	HINSTANCE	hInst,
	HINSTANCE	hInstPrev,
	LPSTR		pszCmdLine,
	int			nCmdShow
){
    MSG         msg;
    PAPP        pApp;

	g_hInst = hInst;
    pApp = new CApp(g_hInst, hInstPrev, pszCmdLine, nCmdShow);

    if (pApp == NULL)
        return (-1);

    if (pApp->Register() == FALSE)
	{
		if (pApp->Init())
	    {
			while (GetMessage(&msg, NULL, 0,0 ))
			{
	            TranslateMessage(&msg);
		        DispatchMessage(&msg);
		    }
		}
	}	

    delete pApp;
    return (msg.wParam);
}


// BenExeWndProc
// Purpose: Standard window class procedure.

LRESULT APIENTRY BenExeWndProc
(
	HWND	hWnd,
	UINT	iMsg,
	WPARAM	wParam,
	LPARAM	lParam
){
    PAPP	pApp;

    pApp = (PAPP)GetWindowLong(hWnd, BENEXEWL_STRUCTURE);

    switch (iMsg)
	{
        case WM_NCCREATE:
            pApp = (PAPP)(((LPCREATESTRUCT)lParam)->lpCreateParams);
            SetWindowLong(hWnd, BENEXEWL_STRUCTURE, (LONG)pApp);
            return (DefWindowProc(hWnd, iMsg, wParam, lParam));

        case WM_DESTROY:
            PostQuitMessage(0);
            break;

        default:
            return (DefWindowProc(hWnd, iMsg, wParam, lParam));
	}

    return (0L);
}

// ObjectDestroyed
// Purpose: Function for the Benefits object to call when it gets destroyed.
// We destroy the main window if the proper conditions are met
// for shutdown.


void ObjectDestroyed(void)
{
    g_lObjectCount--;

    //No more objects and no locks, shut the app down.
    if (0L == g_lObjectCount && 0L == g_lLockedObjects && IsWindow(g_hWnd))
		PostMessage(g_hWnd, WM_DESTROY, 0, 0L);

    return;
}

// CApp::CApp
// Parameters:
// hInst           HINSTANCE of the Application from WinMain
// hInstPrev       HINSTANCE of a previous instance from WinMain
// pszCmdLine      LPSTR of the command line.
// nCmdShow        UINT specifying how to show the app window,
//                 from WinMain.

CApp::CApp
(
	HINSTANCE	hInst,
	HINSTANCE	hInstPrev,
	LPSTR		pszCmdLine,
	UINT		nCmdShow
){
    //Initialize WinMain parameter holders.
    m_hInst      = hInst;
    m_hInstPrev  = hInstPrev;
    m_pszCmdLine = pszCmdLine;
    m_nCmdShow   = nCmdShow;

    m_hWnd = NULL;
	for (UINT i = 0; i < eLAST; i++)
	{
		m_dwRegCO[i] = 0;
		m_pIClassFactory[i] = NULL;
	}
    m_fInitialized = FALSE;

    return;
}


CApp::~CApp(void)
{

    if (m_fInitialized)
	{
		for (UINT i = 0; i < eLAST; i++)
		{
			//Opposite of CoRegisterClassObject; class factory ref is now 1
			if (m_dwRegCO[i] != 0L)
				CoRevokeClassObject(m_dwRegCO[i]);

			//The last Release, which frees the class factory.
			if (m_pIClassFactory[i])
			{
				m_pIClassFactory[i]->Release();
				m_pIClassFactory[i] = NULL;
			}
		}

        CoUninitialize();
	}
    return;
}


// CApp::Init
// Purpose: Initializes an CApp object by registering window classes,
// creating the main window, and doing anything else prone to
// failure.  If this function fails the caller should guarantee
// that the destructor is called.

BOOL CApp::Init(void)
{
    WNDCLASS    wc;

	m_fInitialized = TRUE;

    if (!m_hInstPrev)
	{
        wc.style          = CS_HREDRAW | CS_VREDRAW;
        wc.lpfnWndProc    = BenExeWndProc;
        wc.cbClsExtra     = 0;
        wc.cbWndExtra     = CBWNDEXTRA;
        wc.hInstance      = m_hInst;
        wc.hIcon          = NULL;
        wc.hCursor        = NULL;
        wc.hbrBackground  = (HBRUSH)(COLOR_WINDOW + 1);
        wc.lpszMenuName   = NULL;
        wc.lpszClassName  = TEXT("BenefitsCPPRemote");

        if (!RegisterClass(&wc))
            return (false);
	}

    m_hWnd = CreateWindow(TEXT("BenefitsCPPRemote"),
						TEXT("BenefitsCPPRemote"),
						WS_OVERLAPPEDWINDOW,
						35, 35, 350, 250,
						NULL, NULL, m_hInst, this);

    if (m_hWnd == NULL)
        return (false);

    g_hWnd = m_hWnd;

    if (!InitCF())
        return (false);

	// return true if successful
    return (TRUE);
}


//////////////////////////////////////////////////////////
// RegisterServer

BOOL CApp::Register(void)
{
    //Fail if we're run outside of CoGetClassObject
    if ((lstrcmpiA(m_pszCmdLine, "-RegServer") == 0) ||
        (lstrcmpiA(m_pszCmdLine, "/RegServer") == 0))
		return (RegisterServer());
	else
	if ((lstrcmpiA(m_pszCmdLine, "-UnregServer") == 0) ||
        (lstrcmpiA(m_pszCmdLine, "/UnregServer") == 0 ))
		return (UnregisterServer());
	else
        return (false);
}

	
BOOL RegisterServer()
{
	BOOL	bResult;
	CHAR	szModule[_MAX_PATH+10];
	LPOLESTR	pszIDLib;
	CHAR	szIDLib[_MAX_PATH];
	CHAR	szCLSID[_MAX_PATH];
	HRESULT	hResult;
	ITypeLib* pITypeLib;
  
	hResult		= NOERROR;
	bResult		= TRUE;
	pITypeLib	= NULL;

	//get clsid for the server 
	StringFromGUID2(CLSID_CAdmin, (LPOLESTR)szCLSID, sizeof(szCLSID));
	
	//Get clsid for the type library
	StringFromCLSID(LIBID_BenefitsCPPRemote, &pszIDLib);
	
	hResult = ::WideCharToMultiByte(
						CP_ACP,
						0,
						pszIDLib,
						-1, //::SysStringLen(pszOLEID), 
						szIDLib,
						sizeof(szIDLib), NULL, NULL);

 
	if (!RegisterMyClass("BenefitsCPPRemote.CAdmin", szIDLib, CLSID_CAdmin, "Benefits Admin Class (Ver 1.0)"))
		bResult = FALSE;

	if (!RegisterMyClass("BenefitsCPPRemote.CBank", szIDLib, CLSID_CBank, "Benefits Bank Class (Ver 1.0)"))
		bResult = FALSE;

	if (!RegisterMyClass("BenefitsCPPRemote.CBenefit", szIDLib, CLSID_CBenefit, "Benefits Benefit Class (Ver 1.0)"))
		bResult = FALSE;

	if (!RegisterMyClass("BenefitsCPPRemote.CBenefitLink", szIDLib, CLSID_CBenefitLink, "Benefits Benefit Link Class (Ver 1.0)"))
		bResult = FALSE;

	if (!RegisterMyClass("BenefitsCPPRemote.CDependant", szIDLib, CLSID_CDependant, "Benefits Dependant Class (Ver 1.0)"))
		bResult = FALSE;

	if (!RegisterMyClass("BenefitsCPPRemote.CEmployee", szIDLib, CLSID_CEmployee, "Benefits Employee Class (Ver 1.0)"))
		bResult = FALSE;

	if (!RegisterMyClass("BenefitsCPPRemote.CList", szIDLib, CLSID_CList, "Benefits List Class (Ver 1.0)"))
		bResult = FALSE;

	if (!RegisterMyClass("BenefitsCPPRemote.CListBank", szIDLib, CLSID_CListBank, "Benefits List Bank Class (Ver 1.0)"))
		bResult = FALSE;

	if (!RegisterMyClass("BenefitsCPPRemote.CListBenefit", szIDLib, CLSID_CListBenefit, "Benefits List Benefit Class (Ver 1.0)"))
		bResult = FALSE;

	if (!RegisterMyClass("BenefitsCPPRemote.CListBenefitLink", szIDLib, CLSID_CListBenefitLink, "Benefits List Benefit Link Class (Ver 1.0)"))
		bResult = FALSE;

	if (!RegisterMyClass("BenefitsCPPRemote.CListDependant", szIDLib, CLSID_CListDependant, "Benefits List Dependant Class (Ver 1.0)"))
		bResult = FALSE;

	if (!RegisterMyClass("BenefitsCPPRemote.CListEmployee", szIDLib, CLSID_CListEmployee, "Benefits List Employee Class (Ver 1.0)"))
		bResult = FALSE;

	// get path to this module in ANSI form, convert to UNICODE for LoadTypeLib
	ZeroMemory(szModule, sizeof (szModule));
	GetModuleFileNameA(g_hInst, szModule, _MAX_PATH+10);
	
	//Register type library
	OLECHAR szOLEModule[_MAX_PATH + 10];
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

// RegisterMyClass
// Purpose: Add Registry information for an object.

BOOL RegisterMyClass(LPSTR lpProgId, LPSTR lpLibId, CLSID ClsId, LPSTR lpDescription)
{
	BOOL	bResult;
	CHAR	szModule[_MAX_PATH+10];
	CHAR	szID[_MAX_PATH+10];
	CHAR	szCLSID[_MAX_PATH+10];
	LPOLESTR	pszOLEID;
   
	bResult = TRUE;

    //get clsid for the server 
	StringFromCLSID(ClsId, &pszOLEID);
	::WideCharToMultiByte(CP_ACP, 0, pszOLEID, -1, szID, sizeof(szID), NULL, NULL);

	//Create ProgID entries
	if (!SetKeyAndValue(lpProgId, NULL, NULL, lpProgId))
		bResult = FALSE;

	if (!SetKeyAndValue(lpProgId, "CLSID", NULL, szID))
		bResult = FALSE;

	//Create entries under CLSID
	strcpy(szCLSID, "CLSID\\");
	strcat(szCLSID, szID);

	if (!SetKeyAndValue(szCLSID, NULL, NULL, lpDescription))
		bResult = FALSE;

	if (!SetKeyAndValue(szCLSID, "TypeLib", NULL, lpLibId))
		bResult = FALSE;
	
	// get path to this module 
	ZeroMemory(szModule, sizeof (szModule));
	GetModuleFileNameA(g_hInst, szModule, sizeof(szModule));
	
	if (!SetKeyAndValue(szCLSID, "LocalServer32", NULL, szModule))
		bResult = FALSE;   

	if (!SetKeyAndValue(szCLSID, "ProgID", NULL, lpProgId))
		bResult = FALSE;

	if (!SetKeyAndValue(szCLSID, "VersionIndependentProgID", NULL, lpProgId))
		bResult = FALSE;

	if (!SetKeyAndValue(szCLSID, "Programmable", NULL, ""))
		bResult = FALSE;
	
	return (bResult);
}

// UnregisterServer
// Purpose: Delete our entries for the server from the Registry.

BOOL UnregisterServer()
{
	BOOL  bResult;

	bResult = TRUE;

	//Unregister type library
	HRESULT hResult = ::UnRegisterTypeLib(LIBID_BenefitsCPPRemote, 1, 0, 0, SYS_WIN32);
	if (hResult != NOERROR)
		return (false);

	if (!UnregisterMyClass("BenefitsCPPRemote.CAdmin", CLSID_CAdmin))
		bResult = FALSE;

	if (!UnregisterMyClass("BenefitsCPPRemote.CBank", CLSID_CBank))
		bResult = FALSE;

	if (!UnregisterMyClass("BenefitsCPPRemote.CBenefit", CLSID_CBenefit))
		bResult = FALSE;

	if (!UnregisterMyClass("BenefitsCPPRemote.CBenefitLink", CLSID_CBenefitLink))
		bResult = FALSE;

	if (!UnregisterMyClass("BenefitsCPPRemote.CDependant", CLSID_CDependant))
		bResult = FALSE;

	if (!UnregisterMyClass("BenefitsCPPRemote.CEmployee", CLSID_CEmployee))
		bResult = FALSE;

	if (!UnregisterMyClass("BenefitsCPPRemote.CList", CLSID_CList))
		bResult = FALSE;

	if (!UnregisterMyClass("BenefitsCPPRemote.CListBank", CLSID_CListBank))
		bResult = FALSE;

	if (!UnregisterMyClass("BenefitsCPPRemote.CListBenefit", CLSID_CListBenefit))
		bResult = FALSE;

	if (!UnregisterMyClass("BenefitsCPPRemote.CListBenefitLink", CLSID_CListBenefitLink))
		bResult = FALSE;

	if (!UnregisterMyClass("BenefitsCPPRemote.CListDependant", CLSID_CListDependant))
		bResult = FALSE;

	if (!UnregisterMyClass("BenefitsCPPRemote.CListEmployee", CLSID_CListEmployee))
		bResult = FALSE;
   
	if (bResult)
		MessageBox(NULL, TEXT("UnRegisterServer succeeded."), TEXT("BenefitsCPPRemote"), MB_OK);
	else
		MessageBox(NULL, TEXT("UnRegisterServer failed."), TEXT("BenefitsCPPRemote"), MB_OK);

	return (bResult);
}

// UnregisterMyClass
// Purpose: Remove registry for an object

BOOL UnregisterMyClass(LPSTR lpProgId, CLSID ClsId)
{
	BOOL	bResult;	
	CHAR	szID[_MAX_PATH];
	CHAR	szCLSID[_MAX_PATH];
	CHAR	szTemp[_MAX_PATH];
	LPOLESTR	pszOLEID;
	
	bResult = TRUE;	

	//get clsid for the server 
	StringFromCLSID(ClsId, &pszOLEID);
	HRESULT hResult = ::WideCharToMultiByte( CP_ACP, 0, pszOLEID, -1, szID,	sizeof(szID), NULL, NULL);

	//Create some base key strings for CLSID section.
	strcpy(szCLSID, "CLSID\\");
	strcat(szCLSID, szID);
	
	//delete "LocalServer32" key
	sprintf(szTemp, "%s\\%s", szCLSID, "LocalServer32");
	if (RegDeleteKeyA(HKEY_CLASSES_ROOT, szTemp) != ERROR_SUCCESS)
		bResult = FALSE;
	
	//delete ProgID key
	sprintf(szTemp, "%s\\%s", szCLSID, "ProgId");
	if (RegDeleteKeyA(HKEY_CLASSES_ROOT, szTemp) != ERROR_SUCCESS)
		bResult = FALSE;
	
	//delete VersionIndependentProgID key
	sprintf(szTemp, "%s\\%s", szCLSID, "VersionIndependentProgID");
	if (RegDeleteKeyA(HKEY_CLASSES_ROOT, szTemp) != ERROR_SUCCESS)
		bResult = FALSE;
	
	//delete Programmable key
	sprintf(szTemp, "%s\\%s", szCLSID, "Programmable");
	if (RegDeleteKeyA(HKEY_CLASSES_ROOT, szTemp) != ERROR_SUCCESS)
		bResult = FALSE;
	
	//delete "TypeLib" key
	sprintf(szTemp, "%s\\%s", szCLSID, "TypeLib");
	if (RegDeleteKeyA(HKEY_CLASSES_ROOT, szTemp) != ERROR_SUCCESS)
		bResult = FALSE;
	
	//delete ClassID of our object
	if (RegDeleteKeyA(HKEY_CLASSES_ROOT, szCLSID) != ERROR_SUCCESS)
		bResult = FALSE;
	
	sprintf(szTemp, "%s\\%s", lpProgId, "CLSID");
	if (RegDeleteKeyA(HKEY_CLASSES_ROOT, szTemp) != ERROR_SUCCESS)
		bResult = FALSE; 
	
    //Unregister ProgIDs
	if (RegDeleteKeyA(HKEY_CLASSES_ROOT, lpProgId) != ERROR_SUCCESS)
		bResult = FALSE; 

	return (bResult);
}


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
	
	if (RegCreateKeyExA( HKEY_CLASSES_ROOT, szKey, 0, NULL, 
		REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, NULL, 
		&hKey, NULL) != ERROR_SUCCESS)
		return (false);
	
	if (pszValue)
		RegSetValueExA(hKey, lpValueName, 0, REG_SZ, (BYTE *)pszValue,  strlen(pszValue)+1);
	
	RegCloseKey(hKey);
	return (TRUE);
}

// CApp::InitCF
// Purpose: Register all the Class Factories


BOOL CApp::InitCF(void)
	{
    // Create our class factory and register them for this application
    // using CoRegisterClassObject. We are able to service more than
    // one object at a time so we use REGCLS_MULTIPLEUSE.
	HRESULT hResult;

	hResult = NOERROR;

    m_pIClassFactory[eCAdmin] = new CAdminFactory();
    if (m_pIClassFactory[eCAdmin] == NULL)
        return (false);

    //Since we hold on to this, we should AddRef it.
    m_pIClassFactory[eCAdmin]->AddRef();

    hResult = CoRegisterClassObject(CLSID_CAdmin,
							 m_pIClassFactory[eCAdmin],
							 CLSCTX_LOCAL_SERVER,
							 REGCLS_MULTIPLEUSE,
							 &m_dwRegCO[eCAdmin]);
    if (hResult != NOERROR)
        return (false);

   
	m_pIClassFactory[eCBank] = new CBankFactory();
    if (m_pIClassFactory[eCBank] == NULL)
        return (false);

    //Since we hold on to this, we should AddRef it.
    m_pIClassFactory[eCBank]->AddRef();

    hResult = CoRegisterClassObject(CLSID_CBank,
							 m_pIClassFactory[eCBank],
							 CLSCTX_LOCAL_SERVER,
							 REGCLS_MULTIPLEUSE,
							 &m_dwRegCO[eCBank]);
    if (hResult != NOERROR)
        return (false);


	m_pIClassFactory[eCEmployee] = new CEmployeeFactory();
    if (m_pIClassFactory[eCEmployee] == NULL)
        return (false);

    //Since we hold on to this, we should AddRef it.
    m_pIClassFactory[eCEmployee]->AddRef();

    hResult = CoRegisterClassObject(CLSID_CEmployee,
							 m_pIClassFactory[eCEmployee],
							 CLSCTX_LOCAL_SERVER,
							 REGCLS_MULTIPLEUSE,
							 &m_dwRegCO[eCEmployee]);
    if (hResult != NOERROR)
        return (false);


	m_pIClassFactory[eCBenefit] = new CBenefitFactory();
    if (m_pIClassFactory[eCBenefit] == NULL)
        return (false);

    //Since we hold on to this, we should AddRef it.
    m_pIClassFactory[eCBenefit]->AddRef();

    hResult = CoRegisterClassObject(CLSID_CBenefit,
							 m_pIClassFactory[eCBenefit],
							 CLSCTX_LOCAL_SERVER,
							 REGCLS_MULTIPLEUSE,
							 &m_dwRegCO[eCBenefit]);
    if (hResult != NOERROR)
        return (false);

	m_pIClassFactory[eCDependant] = new CDependantFactory();
    if (m_pIClassFactory[eCDependant] == NULL)
        return (false);

    //Since we hold on to this, we should AddRef it.
    m_pIClassFactory[eCDependant]->AddRef();

    hResult = CoRegisterClassObject(CLSID_CDependant,
							 m_pIClassFactory[eCDependant],
							 CLSCTX_LOCAL_SERVER,
							 REGCLS_MULTIPLEUSE,
							 &m_dwRegCO[eCDependant]);
    if (hResult != NOERROR)
        return (false);


	m_pIClassFactory[eCBenefitLink] = new CBenefitLinkFactory();
    if (m_pIClassFactory[eCBenefitLink] == NULL)
        return (false);

    //Since we hold on to this, we should AddRef it.
    m_pIClassFactory[eCBenefitLink]->AddRef();

    hResult = CoRegisterClassObject(CLSID_CBenefitLink,
							 m_pIClassFactory[eCBenefitLink],
							 CLSCTX_LOCAL_SERVER,
							 REGCLS_MULTIPLEUSE,
							 &m_dwRegCO[eCBenefitLink]);
    if (hResult != NOERROR)
        return (false);


  	m_pIClassFactory[eCList] = new CListFactory();
    if (m_pIClassFactory[eCList] == NULL)
        return (false);

    //Since we hold on to this, we should AddRef it.
    m_pIClassFactory[eCList]->AddRef();

    hResult = CoRegisterClassObject(CLSID_CList,
							 m_pIClassFactory[eCList],
							 CLSCTX_LOCAL_SERVER,
							 REGCLS_MULTIPLEUSE,
							 &m_dwRegCO[eCList]);
    if (hResult != NOERROR)
        return (false);


  	m_pIClassFactory[eCListBank] = new CListBankFactory();
    if (m_pIClassFactory[eCListBank] == NULL)
        return (false);

    //Since we hold on to this, we should AddRef it.
    m_pIClassFactory[eCListBank]->AddRef();

    hResult = CoRegisterClassObject(CLSID_CListBank,
							 m_pIClassFactory[eCListBank],
							 CLSCTX_LOCAL_SERVER,
							 REGCLS_MULTIPLEUSE,
							 &m_dwRegCO[eCListBank]);
    if (hResult != NOERROR)
        return (false);

  	m_pIClassFactory[eCListEmployee] = new CListEmployeeFactory();
    if (m_pIClassFactory[eCListEmployee] == NULL)
        return (false);

    //Since we hold on to this, we should AddRef it.
    m_pIClassFactory[eCListEmployee]->AddRef();

    hResult = CoRegisterClassObject(CLSID_CListEmployee,
							 m_pIClassFactory[eCListEmployee],
							 CLSCTX_LOCAL_SERVER,
							 REGCLS_MULTIPLEUSE,
							 &m_dwRegCO[eCListEmployee]);
    if (hResult != NOERROR)
        return (false);


  	m_pIClassFactory[eCListBenefit] = new CListBenefitFactory();
    if (m_pIClassFactory[eCListBenefit] == NULL)
        return (false);

    //Since we hold on to this, we should AddRef it.
    m_pIClassFactory[eCListBenefit]->AddRef();

    hResult = CoRegisterClassObject(CLSID_CListBenefit,
							 m_pIClassFactory[eCListBenefit],
							 CLSCTX_LOCAL_SERVER,
							 REGCLS_MULTIPLEUSE,
							 &m_dwRegCO[eCListBenefit]);
    if (hResult != NOERROR)
        return (false);

  	m_pIClassFactory[eCListDependant] = new CListDependantFactory();
    if (m_pIClassFactory[eCListDependant] == NULL)
        return (false);

    //Since we hold on to this, we should AddRef it.
    m_pIClassFactory[eCListDependant]->AddRef();

    hResult = CoRegisterClassObject(CLSID_CListDependant,
							 m_pIClassFactory[eCListDependant],
							 CLSCTX_LOCAL_SERVER,
							 REGCLS_MULTIPLEUSE,
							 &m_dwRegCO[eCListDependant]);
    if (hResult != NOERROR)
        return (false);

  	m_pIClassFactory[eCListBenefitLink] = new CListBenefitLinkFactory();
    if (m_pIClassFactory[eCListBenefitLink] == NULL)
        return (false);

    //Since we hold on to this, we should AddRef it.
    m_pIClassFactory[eCListBenefitLink]->AddRef();

    hResult = CoRegisterClassObject(CLSID_CListBenefitLink,
							 m_pIClassFactory[eCListBenefitLink],
							 CLSCTX_LOCAL_SERVER,
							 REGCLS_MULTIPLEUSE,
							 &m_dwRegCO[eCListBenefitLink]);
    if (hResult != NOERROR)
        return (false);

	return (TRUE);
};

#ifdef TURNBACKON
#define UNICODE
#undef TURNBACKON
#endif
