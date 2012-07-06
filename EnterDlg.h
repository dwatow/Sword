#if !defined(AFX_ENTERDLG_H__070C2C79_BB06_4B3F_8168_73B3DA1CD74F__INCLUDED_)
#define AFX_ENTERDLG_H__070C2C79_BB06_4B3F_8168_73B3DA1CD74F__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// EnterDlg.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CEnterDlg dialog

class CEnterDlg : public CDialog
{
// Construction
public:
	CEnterDlg(CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(CEnterDlg)
	enum { IDD = IDD_DIALOG_ENTER };
		// NOTE: the ClassWizard will add data members here
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CEnterDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(CEnterDlg)
		// NOTE: the ClassWizard will add member functions here
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_ENTERDLG_H__070C2C79_BB06_4B3F_8168_73B3DA1CD74F__INCLUDED_)
