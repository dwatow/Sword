//-----------------------------------------------------------------------------------------------
//  SEC要求量測的暗態13
//        13點圖形化 form Excel檔格式
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

    char buf[200];  //暫存的字串
    char buf1[200];
    char buf2[200];
    int i,j;
    
        ZeroMemory(buf1,sizeof(buf1));
        ZeroMemory(buf2,sizeof(buf2)); 

//Step 1.叫Excel應用程式
    if(!objApp.CreateDispatch("Excel.Application",&e))
    {//失敗
        CString str;
        str.Format("Excel CreateDispatch() failed w/err 0x%08lx", e.m_sc),
        AfxMessageBox(str, MB_SETFOREGROUND);
        return;
    }
    
//Step 2.產生Excel檔案格式
    objBooks = objApp.GetWorkbooks();
    objBook = objBooks.Add(VOptional);
    //objBook.AttachDispatch(objBooks.Add(_variant_t("C:\\test.xls"))); //開啟一個已存在的檔案

//Step 3.設定Sheet1
//手動設定
    objSheets = objBook.GetWorksheets();

    objSheet = objSheets.GetItem(COleVariant((short)1)); //從sheet1開始
    objSheet.SetName("Report"); //設定sheet名稱

//Step 4.開始設定內容

//-----------------------------------------------------------------------------------------------
//           表格字填完！下面是填入資料！請準備陣列！
//-----------------------------------------------------------------------------------------------

//填入資料
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
//           格式做完！下面做顏色和框線
//-----------------------------------------------------------------------------------------------
    range = objSheet.GetRange(COleVariant("B2"),COleVariant("F6"));
    range.BorderAround(_variant_t((long)1),3,1,_variant_t((long)RGB(0,0,0)));//粗框線

    range = objSheet.GetRange(COleVariant("B7"),COleVariant("F11"));
    range.BorderAround(_variant_t((long)1),3,1,_variant_t((long)RGB(0,0,0)));//粗框線
//-----------------------------------------------------------------------------------------------
//-----------------------------------------------------------------------------------------------
    objApp.SetVisible(TRUE);    //顯示Excel檔
    objApp.SetUserControl(TRUE);

    range.ReleaseDispatch();
    objSheet.ReleaseDispatch();
    objSheets.ReleaseDispatch();
    objBook.ReleaseDispatch();
    objBooks.ReleaseDispatch();
    objApp.ReleaseDispatch();    

        SetBackGroundColor(); //預設背景顏色
        SetMaxWindow(0);
        ShowWinNormal();//縮小要顯示的東西
        HideWinNormal();//縮小要消失的東西 
        this->SetDlgItemInt(IDC_MAX,1);
}