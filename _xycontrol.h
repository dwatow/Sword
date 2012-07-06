#if !defined(AFX__XYCONTROL_H__50447410_1C9B_4B52_96B1_2013BB29D7DD__INCLUDED_)
#define AFX__XYCONTROL_H__50447410_1C9B_4B52_96B1_2013BB29D7DD__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// Machine generated IDispatch wrapper class(es) created by Microsoft Visual C++

// NOTE: Do not modify the contents of this file.  If this class is regenerated by
//  Microsoft Visual C++, your modifications will be overwritten.

/////////////////////////////////////////////////////////////////////////////
// C_xyControl wrapper class

class C_xyControl : public CWnd
{
protected:
	DECLARE_DYNCREATE(C_xyControl)
public:
	CLSID const& GetClsid()
	{
		static CLSID const clsid
			= { 0xef009364, 0xcdd3, 0x43f0, { 0xbe, 0x96, 0xe6, 0xa9, 0x20, 0xb2, 0x9d, 0x60 } };
		return clsid;
	}
	virtual BOOL Create(LPCTSTR lpszClassName,
		LPCTSTR lpszWindowName, DWORD dwStyle,
		const RECT& rect,
		CWnd* pParentWnd, UINT nID,
		CCreateContext* pContext = NULL)
	{ return CreateControl(GetClsid(), lpszWindowName, dwStyle, rect, pParentWnd, nID); }

    BOOL Create(LPCTSTR lpszWindowName, DWORD dwStyle,
		const RECT& rect, CWnd* pParentWnd, UINT nID,
		CFile* pPersist = NULL, BOOL bStorage = FALSE,
		BSTR bstrLicKey = NULL)
	{ return CreateControl(GetClsid(), lpszWindowName, dwStyle, rect, pParentWnd, nID,
		pPersist, bStorage, bstrLicKey); }

// Attributes
public:

// Operations
public:
	void BateiSet();
	void KokutaiSet();
	void SetRefProbe(LPDISPATCH* newValue);
	void SetXYGraphData();
	void SetRefCa(LPDISPATCH* newValue);
	void SetXYData(short ll, float xx, float yy);
	void ClearData();
	void SetVisible(long* lIndex);
	void SetVisibleAll(BOOL* bFlag);
	void AddXYGraphData(long* lIndex);
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX__XYCONTROL_H__50447410_1C9B_4B52_96B1_2013BB29D7DD__INCLUDED_)
