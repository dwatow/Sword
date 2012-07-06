///////////////////////////////////////////////////////////////////////////////
/*
Color
void CSwordDlg::OnRed()     ����
void CSwordDlg::OnGreen()   ���
void CSwordDlg::OnBlue()    �Ŧ�
void CSwordDlg::OnWhite()   �զ�
void CSwordDlg::OnBlack()   �¦�
void CSwordDlg::OnUsercolor() �ۭq�C�����d

flk
void CSwordDlg::OnFlkSpix()		SubPixel
void CSwordDlg::OnFlk2l()		2Line INV
void CSwordDlg::OnFlk2l1()		2Line+1 INV
void CSwordDlg::OnFlkVs()		V-Stripe
void CSwordDlg::OnFlkCol()		Column
*/
//Pattern �����C��
#include "stdafx.h"
#include "Sword.h"
#include "SwordDlg.h"

void CSwordDlg::OnRed()     {    SetBackGroundColor(BK_Red);     }
void CSwordDlg::OnGreen()   {    SetBackGroundColor(BK_Green);   }
void CSwordDlg::OnBlue()    {    SetBackGroundColor(BK_Blue);    }
void CSwordDlg::OnWhite()   {    SetBackGroundColor(BK_White);   }
void CSwordDlg::OnBlack()   {    SetBackGroundColor(BK_Dark);    }

void CSwordDlg::OnUsercolor() 
{
    CColorDialog Clrdlg(GetBackGroundColor(), CC_FULLOPEN);
    if (Clrdlg.DoModal() == IDOK)
    {
        SetBackGroundColor(BK_Other,Clrdlg.GetColor());
        m_Brush.DeleteObject(); 
        m_Brush.CreateSolidBrush(GetBackGroundColor());
        Invalidate(TRUE);
    }
}

void CSwordDlg::OnFlkSpix() 
{
    SetBackGroundColor(BK_Other, 0x017F7F7F);
    CDC* pDC = new CClientDC(this);

    static   DWORD  hexBits[]   =   {
        0x00017F01,0x0080017F,
        0x0080017F,0x00017F01};

    CBitmap bm;
    bm.CreateBitmap(2,2,32,1, &hexBits);

    m_Brush.DeleteObject();
    m_Brush.CreatePatternBrush(&bm);

    pDC->Rectangle(CRect(0, 0, ScrmH, ScrmV));

    Invalidate(TRUE);
    UpdateWindow();

    DrawACircle(Set9Point(), CircleRadius);//�e��
	//CString str;
	//str.Format("�q��Fliker���A�ФŲ��ʶq��");
	//pDC->TextOut((int)Set9Point().x-CircleRadius, (int)Set9Point().y-CircleRadius, str);
    delete pDC;
}

void CSwordDlg::OnFlk2l() 
{
    SetBackGroundColor(BK_Other, 0x017F7F7F);
    CDC* pDC = new CClientDC(this);

    static DWORD hexBits[]   =   {
        0x00017F01,0x007F017F,
        0x00017F01,0x007F017F,
        0x007F017F,0x00017F01,
        0x007F017F,0x00017F01};

    CBitmap bm;
    bm.CreateBitmap(2,4,32,1, &hexBits);

    m_Brush.DeleteObject();
    m_Brush.CreatePatternBrush(&bm);

    pDC->Rectangle(CRect(0, 0, ScrmH, ScrmV));

    Invalidate(TRUE);    
    UpdateWindow();

    DrawACircle(Set9Point(), CircleRadius);//�e��
    delete pDC;
}

void CSwordDlg::OnFlk2l1() 
{
    SetBackGroundColor(BK_Other, 0x017F7F7F);
    CDC* pDC = new CClientDC(this);

    static DWORD hexBits[] = {
        0x00017F01,0x007F017F,
        0x00017F01,0x007F017F,
        0x007F017F,0x00017F01};

    CBitmap bm;
    bm.CreateBitmap(2,3,32,1, &hexBits);

    m_Brush.DeleteObject();
    m_Brush.CreatePatternBrush(&bm);

    pDC->Rectangle(CRect(0, 0, ScrmH, ScrmV));

    Invalidate(TRUE);    
    UpdateWindow();

    DrawACircle(Set9Point(), CircleRadius);//�e��
    delete pDC;
}

void CSwordDlg::OnFlkVs() 
{
    SetBackGroundColor(BK_Other, 0x017F7F7F);
    CDC* pDC = new CClientDC(this);

    static DWORD hexBits[] = {0x00017F01,0x007F017F};

    CBitmap bm;
    bm.CreateBitmap(2,1,32,1, &hexBits);

    m_Brush.DeleteObject();
    m_Brush.CreatePatternBrush(&bm);

    pDC->Rectangle(CRect(0, 0, ScrmH, ScrmV));

    Invalidate(TRUE);    
    UpdateWindow();

    DrawACircle(Set9Point(), CircleRadius);//�e��
    delete pDC;
}

void CSwordDlg::OnFlkCol() 
{
    SetBackGroundColor(BK_Other, 0x017F7F7F);
    CDC* pDC = new CClientDC(this);

    static DWORD hexBits[] = {
		0x00000000,0x00007F00};

    CBitmap bm;
    bm.CreateBitmap(2,1,32,1, &hexBits);

    m_Brush.DeleteObject();
    m_Brush.CreatePatternBrush(&bm);

    pDC->Rectangle(CRect(0, 0, ScrmH, ScrmV));

    Invalidate(TRUE);
    UpdateWindow();

    DrawACircle(Set9Point(), CircleRadius);//�e��
    delete pDC;
}
