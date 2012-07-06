//----------------------------------------------------
// colorBtn.h : changed by r. wuppinger
//----------------------------------------------------

#ifndef __COLORBTN_H__
#define __COLORBTN_H__

/////////////////////////////////////////////////////////////////////////////
// colorBtn.h : header file for the CColorButton class
//
// Written by Bob Ryan (ryan@cyberzone.net)
// Copyright (c) 1998.
//
// This code may be freely distributable in any application.  If
// you make any changes to the source, please let me know so that
// I might make a better version of the class.
//
// This file is provided "as is" with no expressed or implied warranty.
// The author accepts no liability for any damage/loss of business that
// this product may cause.
//
/////////////////////////////////////////////////////////////////////////////

const COLORREF CLOUDBLUE = RGB(128, 184, 223);
const COLORREF WHITE = RGB(255, 255, 255);
const COLORREF BLACK = RGB(1, 1, 1);
const COLORREF DKGRAY = RGB(128, 128, 128);
const COLORREF LTGRAY = RGB(192, 192, 192);
const COLORREF YELLOW = RGB(255, 255, 0);
const COLORREF DKYELLOW = RGB(128, 128, 0);
const COLORREF RED = RGB(255, 0, 0);
const COLORREF DKRED = RGB(128, 0, 0);
const COLORREF BLUE = RGB(0, 0, 255);
const COLORREF DKBLUE = RGB(0, 0, 128);
const COLORREF CYAN = RGB(0, 255, 255);
const COLORREF DKCYAN = RGB(0, 128, 128);
const COLORREF GREEN = RGB(0, 255, 0);
const COLORREF DKGREEN = RGB(0, 128, 0);
const COLORREF MAGENTA = RGB(255, 0, 255);
const COLORREF DKMAGENTA = RGB(128, 0, 128);


#define CB_BG_DEFAULT		 LTGRAY
#define CB_FG_DEFAULT		 BLACK				// black text 
#define CB_SID_DEFAULT	 DKGRAY 			// dark gray disabled text


class CColorButton : public CButton
{
DECLARE_DYNAMIC(CColorButton)
public:
	CColorButton(); 
	virtual ~CColorButton(); 

	BOOL Attach(const UINT nID, CWnd* pParent, 
		const COLORREF BGColor = CB_BG_DEFAULT, 	// gray button
		const COLORREF FGColor = CB_FG_DEFAULT, 			// black text 
		const COLORREF DisabledColor = CB_SID_DEFAULT,	// dark gray disabled text
		const UINT nBevel = 2
	);

	void SetFGColor      ( COLORREF color = CB_FG_DEFAULT, BOOL bRedraw=FALSE) { m_fg = color;      if(bRedraw)  InvalidateRect(NULL);} 
	void SetBGColor      ( COLORREF color = CB_BG_DEFAULT, BOOL bRedraw=FALSE) { m_bg = color;      if(bRedraw)  InvalidateRect(NULL);}
    void SetDisabledColor(COLORREF color = CB_SID_DEFAULT, BOOL bRedraw=FALSE) { m_disabled= color; if(bRedraw)  InvalidateRect(NULL);}

	void SetColor( COLORREF colFG = CB_FG_DEFAULT,	COLORREF colBG= CB_BG_DEFAULT,	COLORREF colDIS =
CB_SID_DEFAULT, BOOL bRedraw=TRUE) 
			{ SetFGColor( colFG);
	      SetBGColor( colBG);
	    SetDisabledColor(colDIS);
				if(bRedraw) InvalidateRect(NULL);
			}

protected:
	virtual void DrawItem(LPDRAWITEMSTRUCT lpDIS);
	void DrawFrame(CDC *DC, CRect R, int Inset);
	void DrawFilledRect(CDC *DC, CRect R, COLORREF color);
	void DrawLine(CDC *DC, CRect EndPoints, COLORREF color);
	void DrawLine(CDC *DC, long left, long top, long right, long bottom, COLORREF color);
	void DrawButtonText(CDC *DC, CRect R, const char *Buf, COLORREF TextColor);

	COLORREF GetFGColor() { return m_fg; }	
	COLORREF GetBGColor() { return m_bg; }
	COLORREF GetDisabledColor() { return m_disabled; }
	UINT GetBevel() { return m_bevel; }

private:
	COLORREF m_fg, m_bg, m_disabled;
	UINT m_bevel;

};
#endif 
