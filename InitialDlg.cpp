// InitialDlg1.cpp : implementation file
//

#include "stdafx.h"
#include "Sword.h"
#include "InitialDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CInitialDlg dialog


CInitialDlg::CInitialDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CInitialDlg::IDD, pParent)
{
	//{{AFX_DATA_INIT(CInitialDlg)
	m_strInit = _T("");
	//}}AFX_DATA_INIT
	
}


void CInitialDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CInitialDlg)
	DDX_Text(pDX, IDC_INIT_MSG, m_strInit);
	//}}AFX_DATA_MAP
}

void CInitialDlg::SetString(CString &strinit)
{
	m_strInit.Format("%s",strinit);
	UpdateData(FALSE);
	UpdateWindow();
}


BEGIN_MESSAGE_MAP(CInitialDlg, CDialog)
	//{{AFX_MSG_MAP(CInitialDlg)
	ON_WM_CTLCOLOR()
	ON_WM_TIMER()
	ON_WM_DESTROY()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CInitialDlg message handlers

HBRUSH CInitialDlg::OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor) 
{
	HBRUSH hbr = CDialog::OnCtlColor(pDC, pWnd, nCtlColor);

	//CDC* pDC

	pDC->SetTextColor(RGB(255,255,255));
	pDC->SetBkMode(TRANSPARENT);     
	hbr=(HBRUSH)GetStockObject(NULL_BRUSH);

	Invalidate();
    UpdateWindow();

	// TODO: Return a different brush if the default is not desired
	return hbr;
}

void CInitialDlg::OnTimer(UINT nIDEvent) 
{
	// TODO: Add your message handler code here and/or call default
	CDialog::OnTimer(nIDEvent);
}

void CInitialDlg::OnDestroy() 
{
	CDialog::OnDestroy();
	
	// TODO: Add your message handler code here
	
}
