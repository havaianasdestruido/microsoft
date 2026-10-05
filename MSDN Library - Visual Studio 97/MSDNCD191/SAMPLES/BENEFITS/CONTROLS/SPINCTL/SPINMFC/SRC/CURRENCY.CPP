// Currency.cpp : implementation file
//

#include "stdafx.h"
#include "SpinMFC.h"
#include "Currency.h"
#include "SpinMFCCtl.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CCurrency

CCurrency::CCurrency()
{
 	m_Value = 0.0;
    m_Max = 1000000.0;
    m_Min = 0.0;
    m_Step = 1.0;
	m_bCurrency = FALSE;
}

CCurrency::~CCurrency()
{
}

BEGIN_MESSAGE_MAP(CCurrency, CEdit)
	//{{AFX_MSG_MAP(CCurrency)
	ON_WM_KEYDOWN()
	ON_CONTROL_REFLECT(EN_KILLFOCUS, OnKillfocus)
	ON_CONTROL_REFLECT(EN_SETFOCUS, OnSetfocus)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CCurrency message handlers

BOOL CCurrency::PreCreateWindow(CREATESTRUCT& cs) 
{
	// TODO: Add your specialized code here and/or call the base class
	cs.style = WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL; 
	//cs.dwExStyle = WS_EX_STATICEDGE; //to make 3D look
	return CEdit::PreCreateWindow(cs);
}

void CCurrency::OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags) 
{
    CEdit::OnKeyDown(nChar, nRepCnt, nFlags);  

    if(nChar == VK_UP)
        Increase();
    if(nChar == VK_DOWN)
        Decrease();
	if(nChar == VK_RETURN)
	{
		//Get the current value typed by the user 
		UpdateData();
		Done();
	}
}

void CCurrency::Increase()
{
	ASSERT(m_hWnd != NULL);
	UpdateData();
	Done();       //transfer data from the edit control into m_Value  	

    m_Value += m_Step;
    if(m_Value > m_Max)
       m_Value = m_Max;
	if(m_Value < m_Min)
		m_Value = m_Min;
    ShowValue();
}

void CCurrency::Decrease()
{
	ASSERT(m_hWnd != NULL);
	UpdateData();
	Done();     

    m_Value -= m_Step;
    if(m_Value < m_Min)
        m_Value = m_Min;
    ShowValue();
}

void CCurrency::ShowValue()
{
	ASSERT(m_hWnd != NULL);
	CString s;

	if(m_bCurrency)
		s = DoubleToCurrency(m_Value);
	else
		s.Format(TEXT("%.2f"), m_Value);
	
	SetWindowText(s);   
	UpdateData(FALSE);
}

CString CCurrency::DoubleToCurrency(double num)
{
	COleCurrency Currency;
	long whole = long(m_Value);
	long fract = long((m_Value - whole)*10000);

	Currency.SetCurrency(whole, fract);
	CString s = GetConvert(&Currency);
	return s;
}

double CCurrency::CurrencyToDouble()
{
	ASSERT(m_hWnd != NULL);
	double dValue = m_Value;
	CString s;
    GetWindowText(s);

    COleCurrency temp;
    temp.SetCurrency(0,0);
    if(temp.ParseCurrency(s,0, MAKELCID(MAKELANGID(LOCALE_SYSTEM_DEFAULT, LOCALE_USER_DEFAULT),
                                SORT_DEFAULT)))
    {

		CURRENCY cur = CURRENCY(temp);
		LONGLONG num = cur.int64;
	
		dValue = ((double)num) / 10000;
    }
	else
         MessageBeep(MB_OK); //Invalid value, complain
    
	return dValue;
}

CString CCurrency::GetConvert(COleCurrency* pCurrency)
{
    CString sIn = pCurrency->Format();
    CString sOut;

	GetCurrencyFormat(
        MAKELCID(MAKELANGID(LOCALE_SYSTEM_DEFAULT, LOCALE_USER_DEFAULT),
                                SORT_DEFAULT), //locale
                      0, //flags
                      sIn, //input string
                      NULL,//formating info 
                      sOut.GetBuffer(100),
	                  100);
    sOut.ReleaseBuffer();
	                  
    return sOut;   
}

void CCurrency::OnKillfocus() 
{
    //Fire event only if the value has been changed since the last SetFocus
	UpdateData();
	
	Done(); //validate entered data
    
	if(m_Value != m_ValueOnSetFocus)
        ((CSpinMFCCtrl*)GetParent())->ButtonClicked();
}

void CCurrency::OnSetfocus() 
{
    //Save the current value to FireOnChange event when the focus goes away
    //and if the value has been changed
    m_ValueOnSetFocus = m_Value;
}

double& CCurrency::GetValue()
{
	if(m_hWnd)
		Done();//validate entered data

	return m_Value;
}


double& CCurrency::GetIncrement()
{
	//return the address of the m_dStep
	return m_Step;
}


double& CCurrency::GetMaximum()
{
	//return the address of the m_dMaximum
	return m_Max;
}

double& CCurrency::GetMinimum()
{
	//return the address of the m_dMin
	return m_Min;
}

void CCurrency::SetValue(double newValue)
{
	if(!SetNewValue(newValue))
		ShowValue();
}

void CCurrency::SetIncrement(double newValue)
{
	m_Step = newValue;
}

void CCurrency::SetMaximum(double newValue)
{
	m_Max = newValue;
}

void CCurrency::SetMinimum(double newValue)
{
	m_Min = newValue;
}

BOOL& CCurrency::GetCurrency()
{
	return m_bCurrency;
}

void CCurrency::SetCurrency(BOOL newValue)
{
	m_bCurrency = newValue;
	ShowValue();
}

BOOL CCurrency::SetNewValue(double newValue)
{
	BOOL bResult = TRUE;

	if(newValue == 0.0)
	{
		m_Value = newValue;
	}
	else if(newValue > m_Max)
	{
		m_Value = m_Max;
		bResult = FALSE;
	}
	else if(newValue < m_Min)
	{
		m_Value = m_Min;
		bResult = FALSE;
	}
	else
		m_Value = newValue;
	return bResult;
}

void CCurrency::Done()
{
	double temp = CurrencyToDouble();
		
	//Try to assign it to the member variable m_Value
	if(!SetNewValue(temp))
		MessageBeep(MB_OK); //The value was out of range
	
	ShowValue();
}
