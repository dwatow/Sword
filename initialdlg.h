#if !defined(AFX_CINITIALDLG_H__7C15A07B_969B_4373_A1CB_DA20841E9B23__INCLUDED_)
#define AFX_CINITIALDLG_H__7C15A07B_969B_4373_A1CB_DA20841E9B23__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// CInitialDlg.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CInitialDlg dialog

class CInitialDlg : public CDialog
{
// Construction
public:
	void SetString(CString& strinit);
	CInitialDlg(CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(CInitialDlg)
	enum { IDD = IDD_INITIAL };
	CString	m_strInit;
		// NOTE: the ClassWizard will add data members here
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CInitialDlg)
	public:
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(CInitialDlg)
	afx_msg HBRUSH OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor);
	afx_msg void OnTimer(UINT nIDEvent);
	afx_msg void OnDestroy();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_CINITIALDLG_H__7C15A07B_969B_4373_A1CB_DA20841E9B23__INCLUDED_)
