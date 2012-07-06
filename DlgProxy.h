// DlgProxy.h : header file
//

#if !defined(AFX_DLGPROXY_H__793FA08D_549A_459E_A9D5_9D773A77B8CF__INCLUDED_)
#define AFX_DLGPROXY_H__793FA08D_549A_459E_A9D5_9D773A77B8CF__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

class CSwordDlg;

/////////////////////////////////////////////////////////////////////////////
// CSwordDlgAutoProxy command target

class CSwordDlgAutoProxy : public CCmdTarget
{
	DECLARE_DYNCREATE(CSwordDlgAutoProxy)

	CSwordDlgAutoProxy();           // protected constructor used by dynamic creation

// Attributes
public:
	CSwordDlg* m_pDialog;

// Operations
public:

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CSwordDlgAutoProxy)
	public:
	virtual void OnFinalRelease();
	//}}AFX_VIRTUAL

// Implementation
protected:
	virtual ~CSwordDlgAutoProxy();

	// Generated message map functions
	//{{AFX_MSG(CSwordDlgAutoProxy)
		// NOTE - the ClassWizard will add and remove member functions here.
	//}}AFX_MSG

	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(CSwordDlgAutoProxy)

	// Generated OLE dispatch map functions
	//{{AFX_DISPATCH(CSwordDlgAutoProxy)
		// NOTE - the ClassWizard will add and remove member functions here.
	//}}AFX_DISPATCH
	DECLARE_DISPATCH_MAP()
	DECLARE_INTERFACE_MAP()
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_DLGPROXY_H__793FA08D_549A_459E_A9D5_9D773A77B8CF__INCLUDED_)
