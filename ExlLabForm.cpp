//-----------------------------------------------------------------------------------------------
//�ۤv����CA-210�Ϊ����
//    ... Excel�ɮ榡
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
//����Ϊ��]�L�ؽu�^
void CSwordDlg::OnAgingfrom() 
{
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

//RGB�����I
        for(i=0;i<3;i++){//value
        for(j=0;j<121;j++){//color
                ZeroMemory(buf,sizeof(buf));
            if(j == 0)
            {
                switch(i)
                {
                    case 0: strcpy(buf,"Lv"); break;
                    case 1: strcpy(buf,"x"); break;
                    case 2: strcpy(buf,"y"); break;
                    default: strcpy(buf,""); break;
                }            
                range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf));
            }
            else
            {
                sprintf(buf,"%c%d",'A'+i,1+j);
                range = objSheet.GetRange(COleVariant(buf),COleVariant(buf));
                
                ZeroMemory(buf,sizeof(buf));
    //                           Lab[n][value]
                sprintf(buf,"%3.2f",Lab[j][i]); //L

                range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf));
            }
        }}

//-----------------------------------------------------------------------------------------------
//           �榡�����I�U�����C��M�ؽu
//-----------------------------------------------------------------------------------------------


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

