#if !defined(SPINMFC_H__073F325C_30EB_11D0_B84A_00AA00C0803B__INCLUDED_)
#define SPINMFC_H__073F325C_30EB_11D0_B84A_00AA00C0803B__INCLUDED_

// SpinMFC.h : main header file for SPINMFC.DLL

#if !defined( __AFXCTL_H__ )
	#error include 'afxctl.h' before including this file
#endif

#include "resource.h"       // main symbols

/////////////////////////////////////////////////////////////////////////////
// CSpinMFCApp : See SpinMFC.cpp for implementation.

class CSpinMFCApp : public COleControlModule
{
public:
	BOOL InitInstance();
	int ExitInstance();
};

extern const GUID CDECL _tlid;
extern const WORD _wVerMajor;
extern const WORD _wVerMinor;

//{{AFX_INSERT_LOCATION}}
// Microsoft Developer Studio will insert additional declarations immediately before the previous line.

#endif // !defined(SPINMFC_H__073F325C_30EB_11D0_B84A_00AA00C0803B__INCLUDED)
