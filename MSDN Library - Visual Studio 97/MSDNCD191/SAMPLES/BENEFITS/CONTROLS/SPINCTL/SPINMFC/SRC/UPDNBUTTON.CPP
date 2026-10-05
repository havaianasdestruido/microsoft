// UpDnButton.cpp : implementation file
//

#include "stdafx.h"
#include "SpinMFC.h"
#include "UpDnButton.h"
#include "SpinMFCCtl.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CUpDnButton

CUpDnButton::CUpDnButton()
{
    m_bUp = TRUE;
    m_pCurrency = NULL;
}

CUpDnButton::~CUpDnButton()
{
}


BEGIN_MESSAGE_MAP(CUpDnButton, CButton)
	//{{AFX_MSG_MAP(CUpDnButton)
	ON_WM_LBUTTONDOWN()
	ON_WM_KEYDOWN()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CUpDnButton message handlers

/////////////////////////////////////////////////////////////////////////////
// The button was clicked. Process the event.
void CUpDnButton::OnLButtonDown(UINT nFlags, CPoint point) 
{
	if(m_pCurrency)
        m_bUp == TRUE ? m_pCurrency->Increase() : m_pCurrency->Decrease();
	
    ((CSpinMFCCtrl*)GetParent())->ButtonClicked();
    CButton::OnLButtonDown(nFlags, point);
}

/////////////////////////////////////////////////////////////////////////////
// Process Up and Down arrow keys
void CUpDnButton::OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags) 
{
	if((nChar == VK_UP) && (m_pCurrency))
        m_pCurrency->Increase();
    if((nChar == VK_DOWN) && (m_pCurrency))
        m_pCurrency->Decrease();
    
    CButton::OnKeyDown(nChar, nRepCnt, nFlags);
}
