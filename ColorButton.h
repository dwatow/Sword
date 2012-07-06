#if !defined(AFX_COLORBUTTON_H__9EEFB785_A774_4BA5_B9D8_9553C8069A15__INCLUDED_)
#define AFX_COLORBUTTON_H__9EEFB785_A774_4BA5_B9D8_9553C8069A15__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// ColorButton.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CColorButton window

class CColorButton : public CButton
{
// Construction
public:
	CColorButton();

// Attributes
public:

// Operations
public:

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CColorButton)
	public:
	virtual void OnFinalRelease();

	//}}AFX_VIRTUAL

// Implementation
public:
	BOOL Attach(const UINT nID, CWnd* pParent, 
		const COLORREF BGColor = RGB(192, 192, 192), 	// gray button
		const COLORREF FGColor = RGB(1, 1, 1), 			// black text 
		const COLORREF DisabledColor = RGB(0, 128, 0),	// dark gray disabled text
		const UINT nBevel = 2
	);
	//void SetBGColor(COLORREF color = RGB(192, 192, 192), BOOL bRedraw=FALSE);
	void SetBGColor(COLORREF color = RGB(255, 0, 0));
	virtual ~CColorButton();

	// Generated message map functions
protected:
	void DrawLine(CDC *DC, long left, long top, long right, long bottom, COLORREF color);
	void DrawLine(CDC *DC, CRect EndPoints, COLORREF color);
	UINT GetBevel();

	void DrawFrame(CDC *DC, CRect R, int Inset);
	COLORREF GetBGColor();
	void DrawFilledRect(CDC *DC, CRect R, COLORREF color);
	//{{AFX_MSG(CColorButton)
	virtual void DrawItem(LPDRAWITEMSTRUCT lpDIS);
	//}}AFX_MSG

	DECLARE_MESSAGE_MAP()
	// Generated OLE dispatch map functions
	//{{AFX_DISPATCH(CColorButton)
		// NOTE - the ClassWizard will add and remove member functions here.
	//}}AFX_DISPATCH
	DECLARE_DISPATCH_MAP()
	DECLARE_INTERFACE_MAP()
private:
	COLORREF m_bg, m_disabled;
	UINT m_bevel;
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_COLORBUTTON_H__9EEFB785_A774_4BA5_B9D8_9553C8069A15__INCLUDED_)
