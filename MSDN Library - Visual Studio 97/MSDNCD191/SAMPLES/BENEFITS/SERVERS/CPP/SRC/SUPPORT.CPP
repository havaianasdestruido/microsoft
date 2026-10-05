//support.cpp  - Copyright (C) 1996 Microsoft Corporation
//            - All rights reserved.
//Date        - 10/19/96
//Authors     - Ken Bergmann, Robert Coleridge
//Purpose     - Implementation file for FlagException & FlagExceptionMsg

#pragma warning (disable : 4800)	// turns off bool warning

#include <stdio.h>
#include <windows.h>
#include <tchar.h>

#include "support.h"
#include "cstrfunc.h"

void Audit(long oid, int iType, BSTR lpMsg, int iCount);

extern HINSTANCE	g_hInst;
extern ULONG		g_lObjectCount;
extern ULONG		g_lLockedObjects;
extern long			g_oid;

// FlagException
//
// Thread-safe exception handling means that we call
// CreateErrorInfo which gives us an ICreateErrorInfo pointer
// that we then use to set the error information (basically
// to set the fields of an EXCEPINFO structure.  We then
// call SetErrorInfo to attach this error to the current
// thread.  ITypeInfo::Invoke will look for this when it
// returns from whatever function was invokes by calling
// GetErrorInfo.
//
// Exception
//
// Purpose:
//  Raises an exception for Invoke from within ITypeInfo::Invoke
//  using the CreateErrorInfo API and the ICreateErrorInfo interface.
//  This is part of CStack since CStack derives from IDispatch
//  as well.
//
// Parameters:
//  guid		GUID of excepting thread
//  uSource		IDS id of source
//  uException	IDS id of exception

HRESULT FlagException(REFGUID guid, UINT uSource, UINT uException)
{
	HRESULT	hResult = NOERROR;

    _bstr_t bstrTemp = CStrFnc::LoadResString(uException);
	hResult = FlagExceptionMsg(guid, uSource, bstrTemp);

	return (hResult);
}

//
// FlagExceptionMsg
//
HRESULT FlagExceptionMsg(REFGUID guid, UINT uSource, BSTR bstrMsg)
{
	HRESULT				hResult = NOERROR;
	ICreateErrorInfo	*pICreateErr;
	IErrorInfo			*pIErr;

	if (CreateErrorInfo(&pICreateErr) != NOERROR)
		return (DISP_E_EXCEPTION);

	pICreateErr->SetHelpFile(NULL);
	pICreateErr->SetHelpContext(0L);
	pICreateErr->SetGUID(guid);

    _bstr_t bstrTemp = CStrFnc::LoadResString(uSource);
    pICreateErr->SetSource(bstrTemp);
    pICreateErr->SetDescription(bstrMsg);

	hResult = pICreateErr->QueryInterface(IID_IErrorInfo, (void **)&pIErr);
	if (hResult == NOERROR)
	{
		SetErrorInfo(0, pIErr);
		pIErr->Release();
	}

	pICreateErr->Release();

	return (DISP_E_EXCEPTION);
}

//
// CSupportErrorInfo::CSupportErrorInfo
//

CSupportErrorInfo::CSupportErrorInfo(IUnknown * punkObject, REFIID riid)
{
    m_punkObject = punkObject; 
    m_iid = riid;
}         

//
// CSupportErrorInfo::QueryInterface
//

STDMETHODIMP CSupportErrorInfo::QueryInterface(REFIID iid, void FAR* FAR* ppv) 
{   
    return (m_punkObject->QueryInterface(iid, ppv)); 
}

//
// CSupportErrorInfo::AddRef
//

STDMETHODIMP_(ULONG) CSupportErrorInfo::AddRef(void)
{
    return (m_punkObject->AddRef());
}

//
// CSupportErrorInfo::Release
///

STDMETHODIMP_(ULONG) CSupportErrorInfo::Release(void)
{
    return (m_punkObject->Release());
} 

//
// CSupportErrorInfo::InterfaceSupportsErrorInfo
//
// Purpose:
// Indicates whether an interface supports Error Information.
//

STDMETHODIMP CSupportErrorInfo::InterfaceSupportsErrorInfo(REFIID riid)  
{
    return ((riid == m_iid) ? NOERROR : (S_FALSE));
}

void Audit(long oid, int iType, BSTR lpMsg, int iCount)
{
#ifdef ENABLEAUDITING
	static HANDLE hFile = 0;
	static DWORD dwOpenAttributes = CREATE_ALWAYS;
	static long LineNo = 1;

	DWORD dwBytesWritten;
	char cValue[200];
	char cOut[100];

	sprintf(cOut, "%8ld, oid:%5ld ", LineNo++, oid);
	strcpy(cValue, cOut);

	::WideCharToMultiByte(CP_ACP,	0, lpMsg, -1, cOut, 50, NULL, NULL);
	strcat(cValue, cOut);
	switch (iType)
	{
	case 0: // no special msg
		break;
	case 1: // Addref
		strcat(cValue, " AddRef ");
		break;
	case 2: // release
		strcat(cValue, " Release ");
		break;
	}

	sprintf(cOut, " Cnt:%d, ObjCnt:%ld, LckObj:%ld\r\n", iCount, g_lObjectCount, g_lLockedObjects);
	strcat(cValue, cOut);

	hFile = ::CreateFile(
		TEXT("c:\\audit.log"),
		GENERIC_WRITE,
		0,
		NULL,
		dwOpenAttributes,
		FILE_ATTRIBUTE_NORMAL,
		NULL);
	if (hFile != INVALID_HANDLE_VALUE)
	{
		SetFilePointer(hFile, 0, NULL, FILE_END);
		WriteFile(hFile, cValue, strlen(cValue),&dwBytesWritten, NULL);
		CloseHandle(hFile);

		dwOpenAttributes = OPEN_EXISTING;
	}
	return;
#endif
}
