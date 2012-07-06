//-----------------------------------------------------------------------------------------------
//文展給的
//    2nd_Source_LCM_optics data_modify Excel檔格式
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

void CSwordDlg::OnSave1() 
{
    COleVariant VOptional((long)DISP_E_PARAMNOTFOUND,VT_ERROR);
    _Application objApp;
    Workbooks objBooks;
    _Workbook objBook; 
    Worksheets objSheets;
    _Worksheet objSheet;//,objSheetT;
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
    objSheet.SetName("32''6K"); //設定sheet名稱

//自動設定
    /*找到新增Sheet的語法吧！不然也只能新增三組*/
/*    
int Sheet;
short sBuf;
for(Sheet=0;Sheet<4;Sheet++)
{
    sBuf=(short)Sheet+1; 
    objSheet = objSheets.Add(3,4,5,1);
    objSheet = objSheets.GetItem(COleVariant(sBuf)); //從sheet1開始

    ZeroMemory(buf,sizeof(buf));
    switch(Sheet)
    {
    case 0: strcpy(buf,"32''6K"); break;
    case 1: strcpy(buf,"37''6K"); break;
    case 2: strcpy(buf,"37''6.5K"); break;
    case 3: strcpy(buf,"40''5K"); break;
    case 4: strcpy(buf,"55''6K"); break;
    default: strcpy(buf,"出錯囉！"); break;
    }
    objSheet.SetName(buf); //設定sheet名稱

*/
//Step 4.開始設定內容

    //設定MODULE NO.
    range = objSheet.GetRange(COleVariant("A2"),COleVariant("A4"));//選取A1:E1
    range.SetMergeCells(_variant_t(true));//合併儲存格
    range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t("MODULE\nNO.")); //內容

    //設定LCM ID
    range = objSheet.GetRange(COleVariant("B2"),COleVariant("B4"));//選取A2:E4
    range.SetMergeCells(_variant_t(true));//合併儲存格
    range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t("LCM ID")); //內容
    
    //設定LED Light Bar ID
    range = objSheet.GetRange(COleVariant("C2"),COleVariant("E4"));//選取A2:E4
    range.SetMergeCells(_variant_t(true));//合併儲存格
    range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t("LED Light Bar ID")); //內容

    //設定measure point
    range = objSheet.GetRange(COleVariant("F2"),COleVariant("F4"));//選取A2:E4
    range.SetMergeCells(_variant_t(true));//合併儲存格
    range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t("measure point")); //內容

    //設定Light bar optics
    range = objSheet.GetRange(COleVariant("G2"),COleVariant("I2"));
    range.SetMergeCells(_variant_t(true));//合併儲存格
    range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t("Light bar optics")); //內容
    
        range = objSheet.GetRange(COleVariant("G3"),COleVariant("G4"));
        range.SetMergeCells(_variant_t(true));//合併儲存格
        range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t("Lv")); //內容

        range = objSheet.GetRange(COleVariant("H3"),COleVariant("H4"));
        range.SetMergeCells(_variant_t(true));//合併儲存格
        range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t("x")); //內容

        range = objSheet.GetRange(COleVariant("I3"),COleVariant("I4"));
        range.SetMergeCells(_variant_t(true));//合併儲存格
        range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t("y")); //內容

    //設定Lumens / S-LED
    range = objSheet.GetRange(COleVariant("J2"),COleVariant("L2"));
    range.SetMergeCells(_variant_t(true));//合併儲存格
    range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t("Lumens / S-LED")); //內容

        range = objSheet.GetRange(COleVariant("J3"),COleVariant("J4"));
        range.SetMergeCells(_variant_t(true));//合併儲存格
        range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t("Avg. Lv")); //內容

        range = objSheet.GetRange(COleVariant("K3"),COleVariant("K4"));
        range.SetMergeCells(_variant_t(true));//合併儲存格
        range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t("Avg. x")); //內容

        range = objSheet.GetRange(COleVariant("L3"),COleVariant("L4"));
        range.SetMergeCells(_variant_t(true));//合併儲存格
        range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t("Avg. y")); //內容
    
    //設定Correction factor
    range = objSheet.GetRange(COleVariant("M2"),COleVariant("O2"));
    range.SetMergeCells(_variant_t(true));//合併儲存格
    range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t("Correction factor")); //內容

        range = objSheet.GetRange(COleVariant("M3"),COleVariant("M4"));
        range.SetMergeCells(_variant_t(true));//合併儲存格
        range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t("Lv")); //內容

        range = objSheet.GetRange(COleVariant("N3"),COleVariant("N4"));
        range.SetMergeCells(_variant_t(true));//合併儲存格
        range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t("x")); //內容

        range = objSheet.GetRange(COleVariant("O3"),COleVariant("O4"));
        range.SetMergeCells(_variant_t(true));//合併儲存格
        range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t("y")); //內容

    //設定WHITE
    range = objSheet.GetRange(COleVariant("P2"),COleVariant("AB2"));
    range.SetMergeCells(_variant_t(true));//合併儲存格
    range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t("WHITE")); //內容

        for(i=0;i<9;i++)
        {
            ZeroMemory(buf,sizeof(buf)); sprintf(buf,"%c3",'P'+i); 
            range = objSheet.GetRange(COleVariant(buf),COleVariant(buf));

                ZeroMemory(buf,sizeof(buf)); sprintf(buf,"%d",i+1);      
                range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf)); //內容
            
            ZeroMemory(buf,sizeof(buf)); sprintf(buf,"%c4",'P'+i);             
            range = objSheet.GetRange(COleVariant(buf),COleVariant(buf));

                ZeroMemory(buf,sizeof(buf));
                switch(i)
                {
                case 0: strcpy(buf,"左上"); break;
                case 1: strcpy(buf,"中上"); break;
                case 2: strcpy(buf,"右上"); break;
                case 3: strcpy(buf,"左中"); break;
                case 4: strcpy(buf,"中"); break;
                case 5: strcpy(buf,"右中"); break;
                case 6: strcpy(buf,"左下"); break;
                case 7: strcpy(buf,"中下"); break;
                case 8: strcpy(buf,"右下"); break;
                default: strcpy(buf,"出錯了！"); break;
                }

                range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf)); //內容
        }
        range = objSheet.GetRange(COleVariant("Y3"),COleVariant("Y4"));
        range.SetMergeCells(_variant_t(true));//合併儲存格
        range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t("Center x")); //內容

        range = objSheet.GetRange(COleVariant("Z3"),COleVariant("Z4"));
        range.SetMergeCells(_variant_t(true));//合併儲存格
        range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t("Center y")); //內容

        range = objSheet.GetRange(COleVariant("AA3"),COleVariant("AA3"));
        range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t("Brightness")); //內容

        range = objSheet.GetRange(COleVariant("AA4"),COleVariant("AA4"));
        range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t("(AVG)")); //內容

        range = objSheet.GetRange(COleVariant("AB3"),COleVariant("AB4"));
        range.SetMergeCells(_variant_t(true));//合併儲存格
        range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t("Uniformity")); //內容

    range = objSheet.GetRange(COleVariant("AC2"),COleVariant("AE3"));
    range.SetMergeCells(_variant_t(true));//合併儲存格
    range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t("RED")); //內容

    range = objSheet.GetRange(COleVariant("AF2"),COleVariant("AH3"));
    range.SetMergeCells(_variant_t(true));//合併儲存格
    range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t("GREEN")); //內容

    range = objSheet.GetRange(COleVariant("AI2"),COleVariant("AK3"));
    range.SetMergeCells(_variant_t(true));//合併儲存格
    range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t("BLUE")); //內容

    for(i=0;i<3;i++){
    for(j=0;j<3;j++){

        ZeroMemory(buf,sizeof(buf)); 
        sprintf(buf,"%c%c4",'A','C'+j+(i*3)); 
        range = objSheet.GetRange(COleVariant(buf),COleVariant(buf));        
        
        ZeroMemory(buf,sizeof(buf));
        switch(j)
        {
            case 0: strcpy(buf,"Lv"); break;
            case 1: strcpy(buf,"x"); break;
            case 2: strcpy(buf,"y"); break;
            default: strcpy(buf,"出錯了！"); break;
        }
        range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf)); //內容
    }}

    range = objSheet.GetRange(COleVariant("A5"),COleVariant("A12"));
    range.SetMergeCells(_variant_t(true));//合併儲存格
    range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t("1")); //內容
    range = objSheet.GetRange(COleVariant("B5"),COleVariant("B12"));
    range.SetMergeCells(_variant_t(true));//合併儲存格
    
        range = objSheet.GetRange(COleVariant("C5"),COleVariant("C8"));
        range.SetMergeCells(_variant_t(true));//合併儲存格
        range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t("1st")); //內容

        range = objSheet.GetRange(COleVariant("C9"),COleVariant("C12"));
        range.SetMergeCells(_variant_t(true));//合併儲存格
        range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t("2nd")); //內容

    range = objSheet.GetRange(COleVariant("A13"),COleVariant("A20"));
    range.SetMergeCells(_variant_t(true));//合併儲存格
    range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t("2")); //內容
    range = objSheet.GetRange(COleVariant("B13"),COleVariant("B20"));
    range.SetMergeCells(_variant_t(true));//合併儲存格

        range = objSheet.GetRange(COleVariant("C13"),COleVariant("C16"));
        range.SetMergeCells(_variant_t(true));//合併儲存格
        range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t("1st")); //內容

        range = objSheet.GetRange(COleVariant("C17"),COleVariant("C20"));
        range.SetMergeCells(_variant_t(true));//合併儲存格
        range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t("2nd")); //內容

    range = objSheet.GetRange(COleVariant("A21"),COleVariant("A28"));
    range.SetMergeCells(_variant_t(true));//合併儲存格
    range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t("3")); //內容
    range = objSheet.GetRange(COleVariant("B21"),COleVariant("B28"));
    range.SetMergeCells(_variant_t(true));//合併儲存格

        range = objSheet.GetRange(COleVariant("C21"),COleVariant("C24"));
        range.SetMergeCells(_variant_t(true));//合併儲存格
        range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t("1st")); //內容

        range = objSheet.GetRange(COleVariant("C25"),COleVariant("C28"));
        range.SetMergeCells(_variant_t(true));//合併儲存格
        range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t("2nd")); //內容

    for(i=0;i<12;i++){
    for(j=0;j< 2;j++){
        ZeroMemory(buf,sizeof(buf)); 
        sprintf(buf,"F%d",j+5+(i*2)); 
        range = objSheet.GetRange(COleVariant(buf),COleVariant(buf));        
        
        ZeroMemory(buf,sizeof(buf));
        sprintf(buf,"%d",j+1); 
        range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf)); //內容
    
    }}
//-----------------------------------------------------------------------------------------------
//           表格字填完！下面是填入資料！請準備陣列！
//-----------------------------------------------------------------------------------------------

for(i = 0; i < 13; i++)     
{
    if(i<11)
    {
        ZeroMemory(buf1,sizeof(buf1));     
        ZeroMemory(buf2,sizeof(buf2));
        sprintf(buf1,"%c30",'P'+i);      
        sprintf(buf2,"%c33",'P'+i); 
    }
    else//i >= 11
    {
        ZeroMemory(buf1,sizeof(buf1));     
        ZeroMemory(buf2,sizeof(buf2));
        sprintf(buf1,"A%c30",'A'+i-11);      
        sprintf(buf2,"A%c33",'A'+i-11); 
    }
    range = objSheet.GetRange(COleVariant(buf1),COleVariant(buf2));
    range.SetMergeCells(_variant_t(true));//合併儲存格
    
    if(i<9)//九點亮度值
    {
        ZeroMemory(buf,sizeof(buf));
        sprintf(buf,"%3.2f",MsrNineOValue[3][i][2]);
        range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf)); //內容
    }
    else if(i==9)//中心點x
    {
        ZeroMemory(buf,sizeof(buf));
        sprintf(buf,"%1.4f",MsrNineOValue[3][4][0]);
        range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf)); //內容
    }
    else if(i==10)//中心點y
    {
        ZeroMemory(buf,sizeof(buf));
        sprintf(buf,"%1.4f",MsrNineOValue[3][4][1]);
        range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf)); //內容
    }
    else if(i==11)//平均亮度
    {
        range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t("=AVERAGE(P30:X30)")); //內容
        range.SetNumberFormat(COleVariant("0.00"));
    }
    else if(i==12)//均齊度
    {
        range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t("=MIN(P30:X30)/MAX(P30:X30)*100")); //內容
        range.SetNumberFormat(COleVariant("0"));
    }
}    

for(i = 0; i < 3; i++) //WRGB...變數
{
    for(j=0;j<3;j++)//Lv,x,y變數
    {
        ZeroMemory(buf1,sizeof(buf1));     sprintf(buf1,"A%c30",'C'+(i*3)+j); 
        ZeroMemory(buf2,sizeof(buf2));     sprintf(buf2,"A%c33",'C'+(i*3)+j); 
        range = objSheet.GetRange(COleVariant(buf1),COleVariant(buf2));
        range.SetMergeCells(_variant_t(true));//合併儲存格

        if(j==0)//Lv
        {
            ZeroMemory(buf,sizeof(buf));
            sprintf(buf,"%3.2f",MsrCenter[i][j]);
            range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf)); //內容    
        }
        else//x,y
        {
            ZeroMemory(buf,sizeof(buf));
            sprintf(buf,"%1.4f",MsrCenter[i][j]);
            range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf)); //內容    
        }
    }
}

//-----------------------------------------------------------------------------------------------
//           格式做完！下面做顏色和框線
//-----------------------------------------------------------------------------------------------




//-----------------------------------------------------------------------------------------------
//-----------------------------------------------------------------------------------------------
/*    objSheet.SaveAs(
        _variant_t("D:\\jobs.xls"), 
        vFileFormat,_variant_t(""),
        _variant_t(""), 
        _variant_t(false),
        _variant_t(false), 
        vSaveAsAccessMode, 
        vSaveConflictResolution, 
        _variant_t(false)
        );*/
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