//MeasureFlow.cpp
#include "stdafx.h"
#include "MeasureFlow.h"

////////////////////////////////////////////////////////////////
//畫圖和顏色

//畫圈
void MsrFlow::DrawACircle(CPoint point) 
{
/*
	在此畫出來的圈是以160*160大小
	以中心點為參考點畫出來的
	原函數是以左上角為參考點，在此修正誤差，將決定好的參考點
	再往左往上各算一個半徑的距離，畫出來

*/
    //設定圈圈顏色=與背景相反顏色
	COLORREF CircleColor;
    CircleColor = RGB(255,255,255)-BackGroundColor;
	//設定畫圈目標物位址
	CWnd* pWndGrid = GetDlgItem(IDC_COLOR_PATTERN);
	CDC* pDC = pWndGrid->GetDC();
	//設定畫筆
	CPen aPen;
	aPen.CreatePen(PS_INSIDEFRAME,5,CircleColor);
	//設定暫存畫筆空間
	CPen* pOldPen = pDC ->SelectObject(&aPen);
	//對圓修正誤差，從左上調到中心（開方形可繪區塊）
	CPoint StartDrawPoint(-80,-80);
	CPoint EndDrawPoint(80,80);
	CRect* pRect = new CRect(point+StartDrawPoint,point+EndDrawPoint);
	//畫下去（畫下圓形，起點0,0，終點0,0）
	CPoint Start(0,0);
	CPoint End(0,0);
	pDC->Arc(pRect,Start,End);
	//畫筆換回預設值
	pDC->SelectObject(pOldPen);
}

void CSwordDlg::DrawMsrLabel(CPoint Point)
{
	CString SX;//Small x
	CString SY;//Small y
	CString Lv;//Small Lv
    CString str;//輸出字串

	//設定點
	CPoint ShiftPoint(0,165);
    CPoint RectPoint(200,200);
    CPoint StartDrawPoint = Point+ShiftPoint;//移到圈圈邊，寫字區塊開始點
    CPoint EndDrawPoint = RectPoint+StartDrawPoint;//寫字區塊結束點
	CRect rect(ShiftPoint,EndDrawPoint);//設定文字區塊

	CDC* pDC = new CClientDC(this);
    SX.Format("%f",m_IProbe.GetSx());
	SY.Format("%f",m_IProbe.GetSy());
	Lv.Format("%f",m_IProbe.GetLv());

    str = "x ="+SX+"\n"+"y ="+SY+"\n"+"Lv="+Lv;
	pDC->DrawText(str, &rect, DT_LEFT | DT_VCENTER);
	UpdateData(FALSE);
}


