//support.h  - Copyright (C) 1996 Microsoft Corporation
//            - All rights reserved.
//Date        - 10/19/96
//Authors     - Ken Bergmann, Robert Coleridge
//Purpose     - Header file for FlagException & FlagExceptionMsg

#ifndef __SUPPORT_H__
#define __SUPPORT_H__

#ifdef _BENDLL
#define THISCLSCTX CLSCTX_INPROC_SERVER
#else
#define THISCLSCTX CLSCTX_LOCAL_SERVER
#endif

#include <windows.h>
#include <malloc.h>
#include <wtypes.h>
#include "resource.h"
#include "cstrfunc.h"

extern void Audit(long oid, int iType, BSTR lpMsg, int iCount);

#define scBlank				TEXT("")
#define scDateFormat		TEXT("%m%d%Y")
#define scNull				TEXT("NULL")
#define scOneStr			TEXT(" 1 ")
#define scSQ				TEXT("'")
#define scSpace				TEXT(" ")
#define scSpaceCommaSpace	TEXT(" , ")
#define icZero				0
#define scZeroStr			TEXT(" 0 ")

//ListViewIndex
#define icLvwDesc			0
#define icLvwPKId			1
#define icLvwCode			2

//function prototypes
HRESULT FlagException(REFGUID guid, UINT uSource, UINT uException);
HRESULT FlagExceptionMsg(REFGUID guid, UINT uSource, BSTR bstrMsg);
void ObjectDestroyed(void);


// assert macro
#ifdef _DEBUG
#include <assert.h>
#define ASSERT(x) assert(x)
#else
#define ASSERT(x)
#endif

//Type for an object-destroyed callback
typedef void (*PFNDESTROYED)(void);

// macro to assist with optional parameters
#define IsOptionalMissing(v) (((v)->vt == VT_EMPTY) || ((v)->vt == VT_ERROR))

//EmptyBSTR frees string then sets string to empty string
#define EmptyBSTR(s)			\
{								\
	if (NULL != s)				\
		::SysFreeString(s);		\
	s = ::SysAllocString(TEXT("")); \
}

//NULLBSTR frees string then sets string to NULL
#define NULLBSTR(s)				\
{								\
	if (NULL != s)				\
		::SysFreeString(s);		\
	s = NULL; 					\
}


// ISupportErrorInfo interface implementation
interface CSupportErrorInfo : public ISupportErrorInfo 
{
 public:
    // IUnknown methods 
    STDMETHOD(QueryInterface)(REFIID riid, void ** ppvObj);
    STDMETHOD_(ULONG, AddRef)(void);
    STDMETHOD_(ULONG, Release)(void); 
    
    // ISupportErrorInfo method
    STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid);  
     
    CSupportErrorInfo(IUnknown * punkObject, REFIID riid); 

 private:
    LPUNKNOWN	m_punkObject;  // IUnknown of Object that implements this interface 
    GUID		m_iid;
};     

typedef CSupportErrorInfo *PCSupportErrorInfo;

#endif //#ifndef __SUPPORT_H__
