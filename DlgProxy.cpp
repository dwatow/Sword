// DlgProxy.cpp : implementation file
//

#include "stdafx.h"
#include "Sword.h"
#include "DlgProxy.h"
#include "SwordDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CSwordDlgAutoProxy

IMPLEMENT_DYNCREATE(CSwordDlgAutoProxy, CCmdTarget)

CSwordDlgAutoProxy::CSwordDlgAutoProxy()
{
	EnableAutomation();
	
	// To keep the application running as long as an automation 
	//	object is active, the constructor calls AfxOleLockApp.
	AfxOleLockApp();

	// Get access to the dialog through the application's
	//  main window pointer.  Set the proxy's internal pointer
	//  to point to the dialog, and set the dialog's back pointer to
	//  this proxy.
	ASSERT (AfxGetApp()->m_pMainWnd != NULL);
	ASSERT_VALID (AfxGetApp()->m_pMainWnd);
	ASSERT_KINDOF(CSwordDlg, AfxGetApp()->m_pMainWnd);
	m_pDialog = (CSwordDlg*) AfxGetApp()->m_pMainWnd;
	m_pDialog->m_pAutoProxy = this;
}

CSwordDlgAutoProxy::~CSwordDlgAutoProxy()
{
	// To terminate the application when all objects created with
	// 	with automation, the destructor calls AfxOleUnlockApp.
	//  Among other things, this will destroy the main dialog
	if (m_pDialog != NULL)
		m_pDialog->m_pAutoProxy = NULL;
	AfxOleUnlockApp();
}

void CSwordDlgAutoProxy::OnFinalRelease()
{
	// When the last reference for an automation object is released
	// OnFinalRelease is called.  The base class will automatically
	// deletes the object.  Add additional cleanup required for your
	// object before calling the base class.

	CCmdTarget::OnFinalRelease();
}

BEGIN_MESSAGE_MAP(CSwordDlgAutoProxy, CCmdTarget)
	//{{AFX_MSG_MAP(CSwordDlgAutoProxy)
		// NOTE - the ClassWizard will add and remove mapping macros here.
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

BEGIN_DISPATCH_MAP(CSwordDlgAutoProxy, CCmdTarget)
	//{{AFX_DISPATCH_MAP(CSwordDlgAutoProxy)
		// NOTE - the ClassWizard will add and remove mapping macros here.
	//}}AFX_DISPATCH_MAP
END_DISPATCH_MAP()

// Note: we add support for IID_ISword to support typesafe binding
//  from VBA.  This IID must match the GUID that is attached to the 
//  dispinterface in the .ODL file.

// {6B491210-B660-4777-8021-845A45DEBD27}
static const IID IID_ISword =
{ 0x6b491210, 0xb660, 0x4777, { 0x80, 0x21, 0x84, 0x5a, 0x45, 0xde, 0xbd, 0x27 } };

BEGIN_INTERFACE_MAP(CSwordDlgAutoProxy, CCmdTarget)
	INTERFACE_PART(CSwordDlgAutoProxy, IID_ISword, Dispatch)
END_INTERFACE_MAP()

// The IMPLEMENT_OLECREATE2 macro is defined in StdAfx.h of this project
// {D4C31140-19C3-482A-B6C9-FA70F610CEC7}
IMPLEMENT_OLECREATE2(CSwordDlgAutoProxy, "Sword.Application", 0xd4c31140, 0x19c3, 0x482a, 0xb6, 0xc9, 0xfa, 0x70, 0xf6, 0x10, 0xce, 0xc7)

/////////////////////////////////////////////////////////////////////////////
// CSwordDlgAutoProxy message handlers
