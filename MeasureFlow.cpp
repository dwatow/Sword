//MeasureFlow.cpp
#include "stdafx.h"
#include "MeasureFlow.h"

////////////////////////////////////////////////////////////////
//�e�ϩM�C��

//�e��
void MsrFlow::DrawACircle(CPoint point) 
{
/*
	�b���e�X�Ӫ���O�H160*160�j�p
	�H�����I���Ѧ��I�e�X�Ӫ�
	���ƬO�H���W�����Ѧ��I�A�b���ץ��~�t�A�N�M�w�n���Ѧ��I
	�A�������W�U��@�ӥb�|���Z���A�e�X��

*/
    //�]�w����C��=�P�I���ۤ��C��
	COLORREF CircleColor;
    CircleColor = RGB(255,255,255)-BackGroundColor;
	//�]�w�e��ؼЪ���}
	CWnd* pWndGrid = GetDlgItem(IDC_COLOR_PATTERN);
	CDC* pDC = pWndGrid->GetDC();
	//�]�w�e��
	CPen aPen;
	aPen.CreatePen(PS_INSIDEFRAME,5,CircleColor);
	//�]�w�Ȧs�e���Ŷ�
	CPen* pOldPen = pDC ->SelectObject(&aPen);
	//���ץ��~�t�A�q���W�ը줤�ߡ]�}��Υiø�϶��^
	CPoint StartDrawPoint(-80,-80);
	CPoint EndDrawPoint(80,80);
	CRect* pRect = new CRect(point+StartDrawPoint,point+EndDrawPoint);
	//�e�U�h�]�e�U��ΡA�_�I0,0�A���I0,0�^
	CPoint Start(0,0);
	CPoint End(0,0);
	pDC->Arc(pRect,Start,End);
	//�e�����^�w�]��
	pDC->SelectObject(pOldPen);
}

void CSwordDlg::DrawMsrLabel(CPoint Point)
{
	CString SX;//Small x
	CString SY;//Small y
	CString Lv;//Small Lv
    CString str;//��X�r��

	//�]�w�I
	CPoint ShiftPoint(0,165);
    CPoint RectPoint(200,200);
    CPoint StartDrawPoint = Point+ShiftPoint;//��������A�g�r�϶��}�l�I
    CPoint EndDrawPoint = RectPoint+StartDrawPoint;//�g�r�϶������I
	CRect rect(ShiftPoint,EndDrawPoint);//�]�w��r�϶�

	CDC* pDC = new CClientDC(this);
    SX.Format("%f",m_IProbe.GetSx());
	SY.Format("%f",m_IProbe.GetSy());
	Lv.Format("%f",m_IProbe.GetLv());

    str = "x ="+SX+"\n"+"y ="+SY+"\n"+"Lv="+Lv;
	pDC->DrawText(str, &rect, DT_LEFT | DT_VCENTER);
	UpdateData(FALSE);
}


