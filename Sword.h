// Sword.h : main header file for the SWORD application
//

#if !defined(AFX_SWORD_H__4F9D0BDB_EF98_447C_AC03_53A0AD92D55C__INCLUDED_)
#define AFX_SWORD_H__4F9D0BDB_EF98_447C_AC03_53A0AD92D55C__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#ifndef __AFXWIN_H__
	#error include 'stdafx.h' before including this file for PCH
#endif
#include "xlef.h"

#include "resource.h"		// main symbols

/////////////////////////////////////////////////////////////////////////////
// CSwordApp:
// See Sword.cpp for the implementation of this class
//

class CSwordApp : public CWinApp
{
public:
	CSwordApp();

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CSwordApp)
	public:
	virtual BOOL InitInstance();
	//}}AFX_VIRTUAL

// Implementation

	//{{AFX_MSG(CSwordApp)
		// NOTE - the ClassWizard will add and remove member functions here.
		//    DO NOT EDIT what you see in these blocks of generated code !
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};


/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_SWORD_H__4F9D0BDB_EF98_447C_AC03_53A0AD92D55C__INCLUDED_)
