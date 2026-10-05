// SpinMFCPpg.cpp : Implementation of the CSpinMFCPropPage property page class.

#include "stdafx.h"
#include "SpinMFC.h"
#include "SpinMFCPpg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif


IMPLEMENT_DYNCREATE(CSpinMFCPropPage, COlePropertyPage)


/////////////////////////////////////////////////////////////////////////////
// Message map

BEGIN_MESSAGE_MAP(CSpinMFCPropPage, COlePropertyPage)
	//{{AFX_MSG_MAP(CSpinMFCPropPage)
	// NOTE - ClassWizard will add and remove message map entries
	//    DO NOT EDIT what you see in these blocks of generated code !
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()


/////////////////////////////////////////////////////////////////////////////
// Initialize class factory and guid

IMPLEMENT_OLECREATE_EX(CSpinMFCPropPage, "SPINMFC.SpinMFCPropPage.1",
	0x73f3257, 0x30eb, 0x11d0, 0xb8, 0x4a, 0, 0xaa, 0, 0xc0, 0x80, 0x3b)


/////////////////////////////////////////////////////////////////////////////
// CSpinMFCPropPage::CSpinMFCPropPageFactory::UpdateRegistry -
// Adds or removes system registry entries for CSpinMFCPropPage

BOOL CSpinMFCPropPage::CSpinMFCPropPageFactory::UpdateRegistry(BOOL bRegister)
{
	if (bRegister)
		return AfxOleRegisterPropertyPageClass(AfxGetInstanceHandle(),
			m_clsid, IDS_SPINMFC_PPG);
	else
		return AfxOleUnregisterClass(m_clsid, NULL);
}


/////////////////////////////////////////////////////////////////////////////
// CSpinMFCPropPage::CSpinMFCPropPage - Constructor

CSpinMFCPropPage::CSpinMFCPropPage() :
	COlePropertyPage(IDD, IDS_SPINMFC_PPG_CAPTION)
{
	//{{AFX_DATA_INIT(CSpinMFCPropPage)
	m_Value = 0.0;
	m_Max = 0.0;
	m_Min = 0.0;
	m_Step = 0.0;
	m_bCurrency = FALSE;
	//}}AFX_DATA_INIT
}


/////////////////////////////////////////////////////////////////////////////
// CSpinMFCPropPage::DoDataExchange - Moves data between page and properties

void CSpinMFCPropPage::DoDataExchange(CDataExchange* pDX)
{
	//{{AFX_DATA_MAP(CSpinMFCPropPage)
	DDP_Text(pDX, IDC_EDIT_VALUE, m_Value, _T("Value"));
	DDX_Text(pDX, IDC_EDIT_VALUE, m_Value);
	DDV_MinMaxDouble(pDX, m_Value, -100000.0, 1000000.0);
	DDP_Text(pDX, IDC_EDIT_MAX, m_Max, _T("Maximum"));
	DDX_Text(pDX, IDC_EDIT_MAX, m_Max);
	DDV_MinMaxDouble(pDX, m_Max, -1000000., 1000000.0);
	DDP_Text(pDX, IDC_EDIT_MIN, m_Min, _T("Minimum"));
	DDX_Text(pDX, IDC_EDIT_MIN, m_Min);
	DDV_MinMaxDouble(pDX, m_Min, -1000000., 1000000.0);
	DDP_Text(pDX, IDC_EDIT_INCREMENT, m_Step, _T("Increment"));
	DDX_Text(pDX, IDC_EDIT_INCREMENT, m_Step);
	DDV_MinMaxDouble(pDX, m_Step, -1000000.0, 1000000.0);
	DDP_Check(pDX, IDC_CHECK_CURRENCY, m_bCurrency, _T("Currency"));	
	DDX_Check(pDX, IDC_CHECK_CURRENCY, m_bCurrency);
	//}}AFX_DATA_MAP		
	DDP_PostProcessing(pDX);
}

/////////////////////////////////////////////////////////////////////////////
// CSpinMFCPropPage message handlers
