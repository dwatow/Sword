#include "stdafx.h"
#include "Sword.h"
#include "SwordDlg.h"


//量測用的勾勾選項
void CSwordDlg::OnCheck9cvr49(){    UpdateData(TRUE);    }
void CSwordDlg::OnCheck9cvr1() {    UpdateData(TRUE);    }
void CSwordDlg::OnCheck49cvr9(){    UpdateData(TRUE);    }
void CSwordDlg::OnCheck49cvr1(){    UpdateData(TRUE);    }

//預覽用的勾勾選項
void CSwordDlg::OnCheckLv()    {    UpdateData(TRUE);    }
void CSwordDlg::OnCheckSx()    {    UpdateData(TRUE);    }
void CSwordDlg::OnCheckSy()    {    UpdateData(TRUE);    }
void CSwordDlg::OnCheckT()     {    UpdateData(TRUE);    }
void CSwordDlg::OnCHECKDuv()   {    UpdateData(TRUE);    }
void CSwordDlg::OnCheckDu()    {    UpdateData(TRUE);    }
void CSwordDlg::OnCheckDv()    {    UpdateData(TRUE);    }
void CSwordDlg::OnCheckX()     {    UpdateData(TRUE);    }
void CSwordDlg::OnCheckY()     {    UpdateData(TRUE);    }
void CSwordDlg::OnCheckZ()     {    UpdateData(TRUE);    }
void CSwordDlg::OnCheck1cvr9() {    UpdateData(TRUE);    }
void CSwordDlg::OnCheck1cvr49(){    UpdateData(TRUE);    }
//void CSwordDlg::OnCheckFlk()   {    UpdateData(TRUE);    }

//OQC量測Item
void CSwordDlg::OnCheckOqc5nits() {	UpdateData(TRUE); }

void CSwordDlg::OnCheckOqcw()     {	UpdateData(TRUE); }
void CSwordDlg::OnCheckOqcr()     {	UpdateData(TRUE); }
void CSwordDlg::OnCheckOqcg()     {	UpdateData(TRUE); }
void CSwordDlg::OnCheckOqcb()     {	UpdateData(TRUE); }
void CSwordDlg::OnCheckOqcd()     {	UpdateData(TRUE); }

void CSwordDlg::OnCheckOqcw9()
{	
	UpdateData(TRUE); 
	pFE -> EnableWindow(m_bool_oqc_w9); 
	pFromEdge  -> EnableWindow(m_bool_oqc_w9);
}
void CSwordDlg::OnCheckOqcw5()	  
{	
	UpdateData(TRUE); 
	p5FE -> EnableWindow(m_bool_oqc_w5 | m_bool_oqc_d5 | m_bool_oqc_r5 | m_bool_oqc_g5 | m_bool_oqc_b5); 
	p5FromEdge  -> EnableWindow(m_bool_oqc_w5 | m_bool_oqc_d5 | m_bool_oqc_r5 | m_bool_oqc_g5 | m_bool_oqc_b5);
}
void CSwordDlg::OnCheckOqcd5()
{	
	UpdateData(TRUE); 
	p5FE -> EnableWindow(m_bool_oqc_w5 | m_bool_oqc_d5 | m_bool_oqc_r5 | m_bool_oqc_g5 | m_bool_oqc_b5); 
	p5FromEdge  -> EnableWindow(m_bool_oqc_w5 | m_bool_oqc_d5 | m_bool_oqc_r5 | m_bool_oqc_g5 | m_bool_oqc_b5);
}
void CSwordDlg::OnCheckOqcr5()
{	
	UpdateData(TRUE); 
	p5FE -> EnableWindow(m_bool_oqc_w5 | m_bool_oqc_d5 | m_bool_oqc_r5 | m_bool_oqc_g5 | m_bool_oqc_b5); 
	p5FromEdge  -> EnableWindow(m_bool_oqc_w5 | m_bool_oqc_d5 | m_bool_oqc_r5 | m_bool_oqc_g5 | m_bool_oqc_b5);
}
void CSwordDlg::OnCheckOqcg5()
{	
	UpdateData(TRUE); 
	p5FE -> EnableWindow(m_bool_oqc_w5 | m_bool_oqc_d5 | m_bool_oqc_r5 | m_bool_oqc_g5 | m_bool_oqc_b5); 
	p5FromEdge  -> EnableWindow(m_bool_oqc_w5 | m_bool_oqc_d5 | m_bool_oqc_r5 | m_bool_oqc_g5 | m_bool_oqc_b5);
}
void CSwordDlg::OnCheckOqcb5()
{	
	UpdateData(TRUE); 
	p5FE -> EnableWindow(m_bool_oqc_w5 | m_bool_oqc_r5 | m_bool_oqc_g5 | m_bool_oqc_b5 | m_bool_oqc_d5); 
	p5FromEdge  -> EnableWindow(m_bool_oqc_w5 | m_bool_oqc_d5 | m_bool_oqc_r5 | m_bool_oqc_g5 | m_bool_oqc_b5);
}

void CSwordDlg::OnCheckOqcw49()	  {	UpdateData(TRUE); }
void CSwordDlg::OnCheckOqcd25()   {	UpdateData(TRUE); }

//搭便車
void CSwordDlg::OnChangeFromEdge(){ UpdateData(TRUE); }  //離邊的幾分之幾
    