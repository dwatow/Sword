// ColorButton.cpp : implementation file
//

#include "stdafx.h"
#include "Sword.h"
#include "ColorButton.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CColorButton

CColorButton::CColorButton()
{
	EnableAutomation();
}

CColorButton::~CColorButton()
{
}

void CColorButton::OnFinalRelease()
{
	// When the last reference for an automation object is released
	// OnFinalRelease is called.  The base class will automatically
	// deletes the object.  Add additional cleanup required for your
	// object before calling the base class.

	CButton::OnFinalRelease();
}


BEGIN_MESSAGE_MAP(CColorButton, CButton)
	//{{AFX_MSG_MAP(CColorButton)
		// NOTE - the ClassWizard will add and remove mapping macros here.
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

BEGIN_DISPATCH_MAP(CColorButton, CButton)
	//{{AFX_DISPATCH_MAP(CColorButton)
		// NOTE - the ClassWizard will add and remove mapping macros here.
	//}}AFX_DISPATCH_MAP
END_DISPATCH_MAP()

// Note: we add support for IID_IColorButton to support typesafe binding
//  from VBA.  This IID must match the GUID that is attached to the 
//  dispinterface in the .ODL file.

// {3BE4D653-2B15-42EF-B59C-301F2B2BE79D}
static const IID IID_IColorButton =
{ 0x3be4d653, 0x2b15, 0x42ef, { 0xb5, 0x9c, 0x30, 0x1f, 0x2b, 0x2b, 0xe7, 0x9d } };

BEGIN_INTERFACE_MAP(CColorButton, CButton)
	INTERFACE_PART(CColorButton, IID_IColorButton, Dispatch)
END_INTERFACE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CColorButton message handlers

//void CColorButton::SetBGColor(COLORREF color, BOOL bRedraw)
void CColorButton::SetBGColor(COLORREF color)
{
 m_bg = color;      
 //if(bRedraw)  
	 InvalidateRect(NULL);
}

void CColorButton::DrawItem(LPDRAWITEMSTRUCT lpDIS) 
{
CDC* pDC = CDC::FromHandle(lpDIS->hDC);

	UINT state = lpDIS->itemState; 
	CRect focusRect, btnRect;
	focusRect.CopyRect(&lpDIS->rcItem); 
	btnRect.CopyRect(&lpDIS->rcItem); 

	//
	// Set the focus rectangle to just past the border decoration
	//
	focusRect.left += 4;
	focusRect.right -= 4;
	focusRect.top += 4;
	focusRect.bottom -= 4;
	  
	//
	// Retrieve the button's caption
	//
	const int bufSize = 512;
	TCHAR buffer[bufSize];
	GetWindowText(buffer, bufSize);
	

	//
	// Draw and label the button using draw methods 
	//
	DrawFilledRect(pDC, btnRect, GetBGColor()); 
	DrawFrame(pDC, btnRect, GetBevel());
//	DrawButtonText(pDC, btnRect, buffer, GetFGColor());


	//
	// Now, depending upon the state, redraw the button (down image) if it is selected,
	// place a focus rectangle on it, or redisplay the caption if it is disabled
	//
	if (state & ODS_FOCUS) {
		DrawFocusRect(lpDIS->hDC, (LPRECT)&focusRect);
		if (state & ODS_SELECTED){ 
			DrawFilledRect(pDC, btnRect, GetBGColor()); 
			DrawFrame(pDC, btnRect, -1);


// ----> Changes!	// changes by RW: 

				// move the button text if it is pressed...
				CRect rectPressedBtnText = btnRect;
				
				// to the right and bottom...
				rectPressedBtnText.OffsetRect( 1, 1 );

				// ... and now paint it!
//			DrawButtonText(pDC, rectPressedBtnText, buffer, GetFGColor());

			// DrawButtonText(pDC, btnRect, buffer, GetFGColor());
			DrawFocusRect(lpDIS->hDC, (LPRECT)&focusRect);
		}
	}
	else if (state & ODS_DISABLED) {
		//COLORREF disabledColor = bg ^ 0xFFFFFF; // contrasting color
//		DrawButtonText(pDC, btnRect, buffer, GetDisabledColor());
	}
}

void CColorButton::DrawFilledRect(CDC *DC, CRect R, COLORREF color)
{
	CBrush B;
	B.CreateSolidBrush(color);
	DC->FillRect(R, &B);
}

COLORREF CColorButton::GetBGColor()
{
  return m_bg; 
}

void CColorButton::DrawFrame(CDC *DC, CRect R, int Inset)
{
	COLORREF dark, light, tlColor, brColor;
	int i, m, width;
	width = (Inset < 0)? -Inset : Inset;
	
	for (i = 0; i < width; i += 1) {
		m = 255 / (i + 2);
		dark = PALETTERGB(m, m, m);
		m = 192 + (63 / (i + 1));
		light = PALETTERGB(m, m, m);
		  
		if ( width == 1 ) {
			light = RGB(255, 255, 255);
			dark = RGB(128, 128, 128);
		}
		
		if ( Inset < 0 ) {
			tlColor = dark;
			brColor = light;
		}
		else {
			tlColor = light;
			brColor = dark;
		}
		
		DrawLine(DC, R.left, R.top, R.right, R.top, tlColor);					     
	  // Across top
		DrawLine(DC, R.left, R.top, R.left, R.bottom, tlColor); 				     
	  // Down left
	  
		if ( (Inset < 0) && (i == width - 1) && (width > 1) ) {
			DrawLine(DC, R.left + 1, R.bottom - 1, R.right, R.bottom - 1, RGB(1, 1, 1));// Across bottom
			DrawLine(DC, R.right - 1, R.top + 1, R.right - 1, R.bottom, RGB(1, 1, 1));	// Down right
		}
		else {
			DrawLine(DC, R.left + 1, R.bottom - 1, R.right, R.bottom - 1, brColor); 	// Across bottom
			DrawLine(DC, R.right - 1, R.top + 1, R.right - 1, R.bottom, brColor);		// Down right
		}
		InflateRect(R, -1, -1);
	}
}


UINT CColorButton::GetBevel()
{
    return m_bevel;
}

void CColorButton::DrawLine(CDC *DC, CRect EndPoints, COLORREF color)
{
	CPen newPen;
	newPen.CreatePen(PS_SOLID, 1, color);
	CPen *oldPen = DC->SelectObject(&newPen);
	DC->MoveTo(EndPoints.left, EndPoints.top);
	DC->LineTo(EndPoints.right, EndPoints.bottom);
	DC->SelectObject(oldPen);
	newPen.DeleteObject();
}

void CColorButton::DrawLine(CDC *DC, long left, long top, long right, long bottom, COLORREF color)
{
	CPen newPen;
	newPen.CreatePen(PS_SOLID, 1, color);
	CPen *oldPen = DC->SelectObject(&newPen);
	DC->MoveTo(left, top);
	DC->LineTo(right, bottom);
	DC->SelectObject(oldPen);
	newPen.DeleteObject();
}

BOOL CColorButton::Attach(const UINT nID, CWnd *pParent, const COLORREF BGColor, const COLORREF FGColor, const COLORREF DisabledColor, const UINT nBevel)
{
	if (!SubclassDlgItem(nID, pParent))
		return FALSE;

	//m_fg = FGColor;
	m_bg = BGColor; 
	m_disabled = RGB(128, 128, 128);
	m_bevel = nBevel;

	return TRUE;
}
