//-----------------------------------------------------------------------------------------------
//自己實驗CA-210用的表格
//    ... Excel檔格式
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
//實驗用表格（無框線）
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

//RGB中心點
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
//           格式做完！下面做顏色和框線
//-----------------------------------------------------------------------------------------------


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

