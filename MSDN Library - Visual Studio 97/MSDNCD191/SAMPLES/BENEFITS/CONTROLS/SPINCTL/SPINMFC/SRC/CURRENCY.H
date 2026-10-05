#ifndef __CURRENCY_H__
#define __CURRENCY_H__

// Currency.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CCurrency window

class CCurrency : public CEdit
{
// Construction
public:
	CCurrency();
    void Increase();
    void Decrease();
    double& GetValue();
	void SetValue(double newValue);
	double& GetMaximum();
	void SetMaximum(double newValue);
	double& GetMinimum();
	void SetMinimum(double newValue);
	double& GetIncrement();
	void SetIncrement(double newValue);
	BOOL& GetCurrency();
	void SetCurrency(BOOL newValue);
	void Done();
protected:
	BOOL SetNewValue(double newValue);
	CString GetConvert(COleCurrency* pCurrency);
	CString DoubleToCurrency(double);
	double CurrencyToDouble();
// Attributes
	double	m_Value,
			m_Step,
			m_Max,
			m_Min,
            m_ValueOnSetFocus; //helper member to check if m_Value was 
                               //changed when focus leaves the control
	BOOL    m_bCurrency;
     
// Operations
public:
    void ShowValue();

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CCurrency)
	protected:
	virtual BOOL PreCreateWindow(CREATESTRUCT& cs);
	//}}AFX_VIRTUAL

// Implementation
public:
	virtual ~CCurrency();

	// Generated message map functions
protected:
	//{{AFX_MSG(CCurrency)
	afx_msg void OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags);
	afx_msg void OnKillfocus();
	afx_msg void OnSetfocus();
	//}}AFX_MSG

	DECLARE_MESSAGE_MAP()
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Developer Studio will insert additional declarations immediately before the previous line.

#endif // __CURRENCY_H__
