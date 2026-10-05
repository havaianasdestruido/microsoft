#if !defined(SPINMFCPPG_H__073F3266_30EB_11D0_B84A_00AA00C0803B__INCLUDED_)
#define SPINMFCPPG_H__073F3266_30EB_11D0_B84A_00AA00C0803B__INCLUDED_

// SpinMFCPpg.h : Declaration of the CSpinMFCPropPage property page class.

////////////////////////////////////////////////////////////////////////////
// CSpinMFCPropPage : See SpinMFCPpg.cpp.cpp for implementation.

class CSpinMFCPropPage : public COlePropertyPage
{
	DECLARE_DYNCREATE(CSpinMFCPropPage)
	DECLARE_OLECREATE_EX(CSpinMFCPropPage)

// Constructor
public:
	CSpinMFCPropPage();

// Dialog Data
	//{{AFX_DATA(CSpinMFCPropPage)
	enum { IDD = IDD_PROPPAGE_SPINMFC };
	double	m_Value;
	double	m_Max;
	double	m_Min;
	double	m_Step;
	BOOL    m_bCurrency;
	//}}AFX_DATA

// Implementation
protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

// Message maps
protected:
	//{{AFX_MSG(CSpinMFCPropPage)
		// NOTE - ClassWizard will add and remove member functions here.
		//    DO NOT EDIT what you see in these blocks of generated code !
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()

};

//{{AFX_INSERT_LOCATION}}
// Microsoft Developer Studio will insert additional declarations immediately before the previous line.

#endif // !defined(SPINMFCPPG_H__073F3266_30EB_11D0_B84A_00AA00C0803B__INCLUDED)
