// SpinMFCCtl.cpp : Implementation of the CSpinMFCCtrl ActiveX Control class.

#include "stdafx.h"
#include "SpinMFC.h"
#include "SpinMFCCtl.h"
#include "SpinMFCPpg.h"


#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif


IMPLEMENT_DYNCREATE(CSpinMFCCtrl, COleControl)


/////////////////////////////////////////////////////////////////////////////
// Message map

BEGIN_MESSAGE_MAP(CSpinMFCCtrl, COleControl)
	//{{AFX_MSG_MAP(CSpinMFCCtrl)
	ON_WM_CREATE()
	ON_WM_SIZE()
	ON_WM_KEYDOWN()
	//}}AFX_MSG_MAP
	ON_OLEVERB(AFX_IDS_VERB_EDIT, OnEdit)
	ON_OLEVERB(AFX_IDS_VERB_PROPERTIES, OnProperties)
END_MESSAGE_MAP()


/////////////////////////////////////////////////////////////////////////////
// Dispatch map

BEGIN_DISPATCH_MAP(CSpinMFCCtrl, COleControl)
	//{{AFX_DISPATCH_MAP(CSpinMFCCtrl)
	DISP_PROPERTY_EX(CSpinMFCCtrl, "Value", GetValue, SetValue, VT_R8)
	DISP_PROPERTY_EX(CSpinMFCCtrl, "Increment", GetIncrement, SetIncrement, VT_R8)
	DISP_PROPERTY_EX(CSpinMFCCtrl, "Maximum", GetMaximum, SetMaximum, VT_R8)
	DISP_PROPERTY_EX(CSpinMFCCtrl, "Minimum", GetMinimum, SetMinimum, VT_R8)
	DISP_PROPERTY_EX(CSpinMFCCtrl, "Currency", GetCurrency, SetCurrency, VT_BOOL)
	//}}AFX_DISPATCH_MAP
	DISP_FUNCTION_ID(CSpinMFCCtrl, "AboutBox", DISPID_ABOUTBOX, AboutBox, VT_EMPTY, VTS_NONE)
END_DISPATCH_MAP()


/////////////////////////////////////////////////////////////////////////////
// Event map

BEGIN_EVENT_MAP(CSpinMFCCtrl, COleControl)
	//{{AFX_EVENT_MAP(CSpinMFCCtrl)
	EVENT_CUSTOM("onChange", FireOnChange, VTS_NONE)
	//}}AFX_EVENT_MAP
END_EVENT_MAP()


/////////////////////////////////////////////////////////////////////////////
// Property pages

// TODO: Add more property pages as needed.  Remember to increase the count!
BEGIN_PROPPAGEIDS(CSpinMFCCtrl, 1)
	PROPPAGEID(CSpinMFCPropPage::guid)
END_PROPPAGEIDS(CSpinMFCCtrl)


/////////////////////////////////////////////////////////////////////////////
// Initialize class factory and guid

IMPLEMENT_OLECREATE_EX(CSpinMFCCtrl, "SPINMFC.SpinMFCCtrl.1",
	0x73f3256, 0x30eb, 0x11d0, 0xb8, 0x4a, 0, 0xaa, 0, 0xc0, 0x80, 0x3b)


/////////////////////////////////////////////////////////////////////////////
// Type library ID and version

IMPLEMENT_OLETYPELIB(CSpinMFCCtrl, _tlid, _wVerMajor, _wVerMinor)


/////////////////////////////////////////////////////////////////////////////
// Interface IDs

const IID BASED_CODE IID_DSpinMFC =
		{ 0x73f3254, 0x30eb, 0x11d0, { 0xb8, 0x4a, 0, 0xaa, 0, 0xc0, 0x80, 0x3b } };
const IID BASED_CODE IID_DSpinMFCEvents =
		{ 0x73f3255, 0x30eb, 0x11d0, { 0xb8, 0x4a, 0, 0xaa, 0, 0xc0, 0x80, 0x3b } };


/////////////////////////////////////////////////////////////////////////////
// Control type information

static const DWORD BASED_CODE _dwSpinMFCOleMisc =
	OLEMISC_ACTIVATEWHENVISIBLE |
	OLEMISC_SETCLIENTSITEFIRST |
	OLEMISC_INSIDEOUT |
	OLEMISC_CANTLINKINSIDE |
	OLEMISC_RECOMPOSEONRESIZE;

IMPLEMENT_OLECTLTYPE(CSpinMFCCtrl, IDS_SPINMFC, _dwSpinMFCOleMisc)


/////////////////////////////////////////////////////////////////////////////
// CSpinMFCCtrl::CSpinMFCCtrlFactory::UpdateRegistry -
// Adds or removes system registry entries for CSpinMFCCtrl

BOOL CSpinMFCCtrl::CSpinMFCCtrlFactory::UpdateRegistry(BOOL bRegister)
{
	// TODO: Verify that your control follows apartment-model threading rules.
	// Refer to MFC TechNote 64 for more information.
	// If your control does not conform to the apartment-model rules, then
	// you must modify the code below, changing the 6th parameter from
	// afxRegInsertable | afxRegApartmentThreading to afxRegInsertable.

	if (bRegister)
		return AfxOleRegisterControlClass(
			AfxGetInstanceHandle(),
			m_clsid,
			m_lpszProgID,
			IDS_SPINMFC,
			IDB_SPINMFC,
			afxRegInsertable | afxRegApartmentThreading,
			_dwSpinMFCOleMisc,
			_tlid,
			_wVerMajor,
			_wVerMinor);
	else
		return AfxOleUnregisterClass(m_clsid, m_lpszProgID);
}


/////////////////////////////////////////////////////////////////////////////
// CSpinMFCCtrl::CSpinMFCCtrl - Constructor

CSpinMFCCtrl::CSpinMFCCtrl()
{
	InitializeIIDs(&IID_DSpinMFC, &IID_DSpinMFCEvents);

    m_DnBtn.SetDown();

	// Set default size of the control
    CDC dc;
    dc.CreateCompatibleDC(NULL);
    TEXTMETRIC tm;
    dc.GetTextMetrics(&tm);
    SetInitialSize(tm.tmAveCharWidth*8, tm.tmHeight+4);
}


/////////////////////////////////////////////////////////////////////////////
// CSpinMFCCtrl::~CSpinMFCCtrl - Destructor

CSpinMFCCtrl::~CSpinMFCCtrl()
{
	// TODO: Cleanup your control's instance data here.
}


/////////////////////////////////////////////////////////////////////////////
// CSpinMFCCtrl::OnDraw - Drawing function

void CSpinMFCCtrl::OnDraw(
			CDC* pdc, const CRect& rcBounds, const CRect& rcInvalid)
{
    CBrush brAmbientBack(TranslateColor(AmbientBackColor()));
	pdc->FillRect(rcBounds, &brAmbientBack);
	
	m_Edit.ShowValue();    

	// Draw the 3D frame around the control.  For consistency with the
	// other conrols on the Sample Web Page we use the width of the 
	// pen 4 and 2 instead of 2 and 1 (dark and light gray pens).
	// The alternative is to set the flag of the edit control 
	// (cs.dwExStyle = WS_EX_STATICEDGE;) in PrecreateWindow of the
	// CCurrency class.
	CPen penDark;  
	CPen penLight;
	
    if( penLight.CreatePen( PS_SOLID, 2, RGB(172, 172, 172)) )
    {
        // Select it into the device context
        // Save the old pen at the same time
        CPen* pOldPen = pdc->SelectObject(&penLight);

        // Draw with the pen
        pdc->MoveTo(rcBounds.left, rcBounds.bottom);
        pdc->LineTo(rcBounds.right, rcBounds.bottom);
        pdc->LineTo(rcBounds.right, rcBounds.top);	

		if(penDark.CreatePen(PS_SOLID, 4, RGB(0, 0, 0)))
        {
			// Select dark pen 
			pdc->SelectObject(&penDark);
	        pdc->MoveTo(rcBounds.right-3, rcBounds.top);
			pdc->LineTo(rcBounds.left, rcBounds.top);
			pdc->LineTo(rcBounds.left, rcBounds.bottom);
		}
		// Restore the old pen 
		pdc->SelectObject(pOldPen);
    } 
}


/////////////////////////////////////////////////////////////////////////////
// CSpinMFCCtrl::DoPropExchange - Persistence support

void CSpinMFCCtrl::DoPropExchange(CPropExchange* pPX)
{
	ExchangeVersion(pPX, MAKELONG(_wVerMinor, _wVerMajor));
	COleControl::DoPropExchange(pPX);

	// TODO: Call PX_ functions for each persistent custom property.
	PX_Double(pPX, _T("Value"), m_Edit.GetValue(), 0.0);
	PX_Double(pPX, _T("Increment"), m_Edit.GetIncrement(), 1.0);
	PX_Double(pPX, _T("Maximum"), m_Edit.GetMaximum(), 10000.0);
	PX_Double(pPX, _T("Minimum"), m_Edit.GetMinimum(), 0.0);
	PX_Bool(pPX, _T("Currency"), m_Edit.GetCurrency(), FALSE);
}


/////////////////////////////////////////////////////////////////////////////
// CSpinMFCCtrl::OnResetState - Reset control to default state

void CSpinMFCCtrl::OnResetState()
{
	COleControl::OnResetState();  // Resets defaults found in DoPropExchange

	// TODO: Reset any other control state here.
}


/////////////////////////////////////////////////////////////////////////////
// CSpinMFCCtrl::AboutBox - Display an "About" box to the user

void CSpinMFCCtrl::AboutBox()
{
	CDialog dlgAbout(IDD_ABOUTBOX_SPINMFC);
	dlgAbout.DoModal();
}


/////////////////////////////////////////////////////////////////////////////
// CSpinMFCCtrl message handlers

int CSpinMFCCtrl::OnCreate(LPCREATESTRUCT lpCreateStruct) 
{
	if (COleControl::OnCreate(lpCreateStruct) == -1)
		return -1;
	  
	// Create CCurrency object. The size will be set later
    CRect rc(0, 0, 0, 0);
    
    if(!m_Edit.Create(WS_CHILD | WS_VISIBLE, rc, this, 100))
       return -1;

	// Set the default system font. 
	HFONT hFont = (HFONT)::GetStockObject(ANSI_VAR_FONT);
	if(hFont)
		m_Edit.SendMessage(WM_SETFONT, (WPARAM)hFont,
								   MAKELPARAM(TRUE, 0));
    // Create Up button
    if(!m_UpBtn.Create(TEXT(""), WS_CHILD | WS_VISIBLE , rc, this, 101))
       return -1;
    m_UpBtn.SetCurrency(&m_Edit);
    HICON hIcon = AfxGetApp()->LoadIcon(IDI_ICON_UP);
    if(hIcon)
    {
        m_UpBtn.ModifyStyle(0, BS_ICON);
        m_UpBtn.SetIcon(hIcon);
    }

    // Create Down button
    if(!m_DnBtn.Create(TEXT(""), WS_CHILD | WS_VISIBLE , rc, this, 102))
       return -1;
    hIcon = AfxGetApp()->LoadIcon(IDI_ICON_DOWN);
    if(hIcon)
    {
        m_DnBtn.ModifyStyle(0, BS_ICON);
        m_DnBtn.SetIcon(hIcon);
    }
    m_DnBtn.SetCurrency(&m_Edit);

	return 0;
}

void CSpinMFCCtrl::OnSize(UINT nType, int cx, int cy) 
{
	UpdateSizes(cx, cy);
    COleControl::OnSize(nType, cx, cy);
}

void CSpinMFCCtrl::UpdateSizes(int& w, int& h)
{
    int btnH = h/2-2;
    int half = h/2;   //middle of the rectangle 

	int btnW = ::GetSystemMetrics(SM_CXHTHUMB); //width of the button
	m_UpBtn.MoveWindow(w-btnW-1, half-btnH, btnW, btnH);
    m_DnBtn.MoveWindow(w-btnW-1, half, btnW, btnH);
    m_Edit.MoveWindow(2, half-btnH, w-btnW-2, h-4); 
}

void CSpinMFCCtrl::OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags) 
{
	if(nChar == VK_UP)
        m_Edit.Increase();
    if(nChar == VK_DOWN)
        m_Edit.Decrease();	
	if(nChar == VK_RETURN)
		m_Edit.Done();

	COleControl::OnKeyDown(nChar, nRepCnt, nFlags);
}


double CSpinMFCCtrl::GetValue() 
{
	return m_Edit.GetValue();
}

void CSpinMFCCtrl::SetValue(double newValue) 
{
	m_Edit.SetValue(newValue);
	SetModifiedFlag();
}

double CSpinMFCCtrl::GetIncrement() 
{
	return m_Edit.GetIncrement();
}

void CSpinMFCCtrl::SetIncrement(double newValue) 
{
	m_Edit.SetIncrement(newValue);
	SetModifiedFlag();
}

double CSpinMFCCtrl::GetMaximum() 
{
	return m_Edit.GetMaximum();
}

void CSpinMFCCtrl::SetMaximum(double newValue) 
{
	m_Edit.SetMaximum(newValue);
	SetModifiedFlag();
}

double CSpinMFCCtrl::GetMinimum() 
{
	return m_Edit.GetMinimum();
}

void CSpinMFCCtrl::SetMinimum(double newValue) 
{
	m_Edit.SetMinimum(newValue);
	SetModifiedFlag();
}


BOOL CSpinMFCCtrl::GetCurrency() 
{
	return m_Edit.GetCurrency();
}

void CSpinMFCCtrl::SetCurrency(BOOL bNewValue) 
{
	m_Edit.SetCurrency(bNewValue);
	SetModifiedFlag();
}

