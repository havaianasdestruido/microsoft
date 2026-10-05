#ifndef __UPDNBUTTON_H__
#define __UPDNBUTTON_H__

#include "Currency.h"

// UpDnButton.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CUpDnButton window

class CUpDnButton : public CButton
{
protected:
    BOOL        m_bUp;       //type of this button: Up or Down
    CCurrency*  m_pCurrency;

// Construction
public:
	CUpDnButton();

// Attributes
public:
    void SetDown(){m_bUp = FALSE;};
    void SetCurrency(CCurrency* pCurrency){m_pCurrency = pCurrency;};
// Operations
public:

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CUpDnButton)
	//}}AFX_VIRTUAL

// Implementation
public:
	virtual ~CUpDnButton();

	// Generated message map functions
protected:
	//{{AFX_MSG(CUpDnButton)
	afx_msg void OnLButtonDown(UINT nFlags, CPoint point);
	afx_msg void OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags);
	//}}AFX_MSG

	DECLARE_MESSAGE_MAP()
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Developer Studio will insert additional declarations immediately before the previous line.

#endif // __UPDNBUTTON_H__
