//-----------------------------------------------------------------------------------------------
//  SEC�n�D�q�����t�A13
//        13�I�ϧΤ� form Excel�ɮ榡
//-----------------------------------------------------------------------------------------------
#include "stdafx.h"
#include <comdef.h>
#include "Sword.h"
#include "SwordDlg.h"
//#include "excel.h"
/*
#define SelectBox(x,y)              ZeroMemory(buf,sizeof(buf));sprintf(buf,"%c%d",x,y);range=objSheet.GetRange(COleVariant(buf),COleVariant(buf));
#define SetBoxDate(Format, Data)    ZeroMemory(buf,sizeof(buf));sprintf(buf,Format,Data);range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf));
*/

void CSwordDlg::OnSave13p() 
{
    // TODO: Add your control notification handler code here
    COleVariant VOptional((long)DISP_E_PARAMNOTFOUND,VT_ERROR);
    _Application objApp;
    Workbooks objBooks;
    _Workbook objBook; 
    Worksheets objSheets;
    _Worksheet objSheet,objSheetT;
    Range range,col,row;//,oCell;//,range2,range3;
    Interior cell;
    Font font;
    COleException e;

    char buf[200];  //�Ȧs���r��
    char buf1[200];
    char buf2[200];
    int i,j;
    
        ZeroMemory(buf1,sizeof(buf1));
        ZeroMemory(buf2,sizeof(buf2)); 

//Step 1.�sExcel���ε{��
    if(!objApp.CreateDispatch("Excel.Application",&e))
    {//����
        CString str;
        str.Format("Excel CreateDispatch() failed w/err 0x%08lx", e.m_sc),
        AfxMessageBox(str, MB_SETFOREGROUND);
        return;
    }
    
//Step 2.����Excel�ɮ׮榡
    objBooks = objApp.GetWorkbooks();
    objBook = objBooks.Add(VOptional);
    //objBook.AttachDispatch(objBooks.Add(_variant_t("C:\\test.xls"))); //�}�Ҥ@�Ӥw�s�b���ɮ�

//Step 3.�]�wSheet1
//��ʳ]�w
    objSheets = objBook.GetWorksheets();

    objSheet = objSheets.GetItem(COleVariant((short)1)); //�qsheet1�}�l
    objSheet.SetName("Report"); //�]�wsheet�W��

//Step 4.�}�l�]�w���e

//-----------------------------------------------------------------------------------------------
//           ���r�񧹡I�U���O��J��ơI�зǳư}�C�I
//-----------------------------------------------------------------------------------------------

//��J���
    for(i=0;i<5;i++){//value
    for(j=0;j<5;j++){//color
        ZeroMemory(buf,sizeof(buf));
        sprintf(buf,"%c%d",'B'+j,2+i);
        range = objSheet.GetRange(COleVariant(buf),COleVariant(buf));

        ZeroMemory(buf,sizeof(buf));
                                                   //  MsrThirteenValue[13][5]
//              if((i==0)&&(j==0))    sprintf(buf,"%3.4f",MsrThirteenValue[0][0]);
//         else if((i==0)&&(j==2))    sprintf(buf,"%3.4f",MsrThirteenValue[1][0]);
//         else if((i==0)&&(j==4))    sprintf(buf,"%3.4f",MsrThirteenValue[2][0]);
// 
//         else if((i==1)&&(j==1))    sprintf(buf,"%3.4f",MsrThirteenValue[3][0]);
//         else if((i==1)&&(j==3))    sprintf(buf,"%3.4f",MsrThirteenValue[4][0]);
// 
//         else if((i==2)&&(j==0))    sprintf(buf,"%3.4f",MsrThirteenValue[5][0]);
//         else if((i==2)&&(j==2))    sprintf(buf,"%3.4f",MsrThirteenValue[6][0]);
//         else if((i==2)&&(j==4))    sprintf(buf,"%3.4f",MsrThirteenValue[7][0]);
// 
//         else if((i==3)&&(j==1))    sprintf(buf,"%3.4f",MsrThirteenValue[8][0]);
//         else if((i==3)&&(j==3))    sprintf(buf,"%3.4f",MsrThirteenValue[9][0]);
// 
//         else if((i==4)&&(j==0))    sprintf(buf,"%3.4f",MsrThirteenValue[10][0]);
//         else if((i==4)&&(j==2))    sprintf(buf,"%3.4f",MsrThirteenValue[11][0]);
//         else if((i==4)&&(j==4))    sprintf(buf,"%3.4f",MsrThirteenValue[12][0]);
//         else                       sprintf(buf," ");

    range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf));
    }}

    range = objSheet.GetRange(COleVariant("D9"),COleVariant("D9"));
    range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t("=MAX(B2:F6)/AVERAGE(B2:F6)"));

//-----------------------------------------------------------------------------------------------
//           �榡�����I�U�����C��M�ؽu
//-----------------------------------------------------------------------------------------------
    range = objSheet.GetRange(COleVariant("B2"),COleVariant("F6"));
    range.BorderAround(_variant_t((long)1),3,1,_variant_t((long)RGB(0,0,0)));//�ʮؽu

    range = objSheet.GetRange(COleVariant("B7"),COleVariant("F11"));
    range.BorderAround(_variant_t((long)1),3,1,_variant_t((long)RGB(0,0,0)));//�ʮؽu
//-----------------------------------------------------------------------------------------------
//-----------------------------------------------------------------------------------------------
    objApp.SetVisible(TRUE);    //���Excel��
    objApp.SetUserControl(TRUE);

    range.ReleaseDispatch();
    objSheet.ReleaseDispatch();
    objSheets.ReleaseDispatch();
    objBook.ReleaseDispatch();
    objBooks.ReleaseDispatch();
    objApp.ReleaseDispatch();    

        SetBackGroundColor(); //�w�]�I���C��
        SetMaxWindow(0);
        ShowWinNormal();//�Y�p�n��ܪ��F��
        HideWinNormal();//�Y�p�n�������F�� 
        this->SetDlgItemInt(IDC_MAX,1);
}