#if !defined(SPINMFCCTL_H__073F3264_30EB_11D0_B84A_00AA00C0803B__INCLUDED_)
#define SPINMFCCTL_H__073F3264_30EB_11D0_B84A_00AA00C0803B__INCLUDED_

// SpinMFCCtl.h : Declaration of the CSpinMFCCtrl ActiveX Control class.
#include "Currency.h"
#include "UpDnButton.h"

/////////////////////////////////////////////////////////////////////////////
// CSpinMFCCtrl : See SpinMFCCtl.cpp for implementation.

class CSpinMFCCtrl : public COleControl
{
	DECLARE_DYNCREATE(CSpinMFCCtrl)
protected:
    CCurrency    m_Edit;
    CUpDnButton  m_UpBtn,
                 m_DnBtn;
// Constructor
public:
	CSpinMFCCtrl();
    void ButtonClicked(){FireOnChange();};

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CSpinMFCCtrl)
	public:
	virtual void OnDraw(CDC* pdc, const CRect& rcBounds, const CRect& rcInvalid);
	virtual void DoPropExchange(CPropExchange* pPX);
	virtual void OnResetState();
	//}}AFX_VIRTUAL

// Implementation
protected:
	~CSpinMFCCtrl();
    void UpdateSizes(int& w, int& h);

	DECLARE_OLECREATE_EX(CSpinMFCCtrl)    // Class factory and guid
	DECLARE_OLETYPELIB(CSpinMFCCtrl)      // GetTypeInfo
	DECLARE_PROPPAGEIDS(CSpinMFCCtrl)     // Property page IDs
	DECLARE_OLECTLTYPE(CSpinMFCCtrl)		// Type name and misc status

// Message maps
	//{{AFX_MSG(CSpinMFCCtrl)
	afx_msg int OnCreate(LPCREATESTRUCT lpCreateStruct);
	afx_msg void OnSize(UINT nType, int cx, int cy);
	afx_msg void OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags);
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()

// Dispatch maps
	//{{AFX_DISPATCH(CSpinMFCCtrl)
	afx_msg double GetValue();
	afx_msg void SetValue(double newValue);
	afx_msg double GetIncrement();
	afx_msg void SetIncrement(double newValue);
	afx_msg double GetMaximum();
	afx_msg void SetMaximum(double newValue);
	afx_msg double GetMinimum();
	afx_msg void SetMinimum(double newValue);
	afx_msg BOOL GetCurrency();
	afx_msg void SetCurrency(BOOL bNewValue);
	//}}AFX_DISPATCH
	DECLARE_DISPATCH_MAP()

	afx_msg void AboutBox();

// Event maps
	//{{AFX_EVENT(CSpinMFCCtrl)
	void FireOnChange()
		{FireEvent(eventidOnChange,EVENT_PARAM(VTS_NONE));}
	//}}AFX_EVENT
	DECLARE_EVENT_MAP()

// Dispatch and event IDs
public:
	enum {
	//{{AFX_DISP_ID(CSpinMFCCtrl)
	dispidValue = 1L,
	dispidIncrement = 2L,
	dispidMaximum = 3L,
	dispidMinimum = 4L,
	dispidCurrency = 5L,
	eventidOnChange = 1L,
	//}}AFX_DISP_ID
	};
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Developer Studio will insert additional declarations immediately before the previous line.

#endif // !defined(SPINMFCCTL_H__073F3264_30EB_11D0_B84A_00AA00C0803B__INCLUDED)
