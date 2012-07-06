//-----------------------------------------------------------------------------------------------
//  Michael要求的
//         量26寸的 form Excel檔格式
//-----------------------------------------------------------------------------------------------
#include "stdafx.h"
//#include <comdef.h>
#include "Sword.h"
#include "SwordDlg.h"
//#include "excel.h"
//#include "xlef.h"
// #define SelectBox(x,y)              ZeroMemory(buf,sizeof(buf));sprintf(buf,"%c%d",x,y);range=objSheet.GetRange(COleVariant(buf),COleVariant(buf))
// #define SetBoxDate(Format, Data)    ZeroMemory(buf,sizeof(buf));sprintf(buf,Format,Data);range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf))


void CSwordDlg::OnTempfor26() 
{
//Step 1.叫Excel應用程式
	xlsFile t26Form;
    
//Step 2.產生Excel檔案格式
    //做路徑出來
    char t26FormPath[MAX_PATH];
    //dwCurDirPathLen = GetCurrentDirectory(MAX_PATH,szCurrentDirectory);//抓目前所在的目錄（路徑）
    strcpy(t26FormPath, szCurrentDirectory);
    strcat(t26FormPath, "\\t26FORM.xls");

//Step 3.設定Sheet1
	t26Form.Open(t26FormPath).SetSheetName(1,"Report");

//Step 4.開始設定內容

//-----------------------------------------------------------------------------------------------
//           表格字填完！下面是填入資料！請準備陣列！
//-----------------------------------------------------------------------------------------------

//填入資料
    int i,j;
    for(i=0;i<7;i++){
    for(j=0;j<7;j++){
		t26Form.SelectCell('D'+i,5+j).SetCell("%3.2f", MsrFortyNineOValue[0][i+j*7][0]);
    }}

    for(i=0;i<9;i++)
    {
		t26Form.SelectCell('E',15+i).SetCell("%3.2f", MsrNineOValue[0][i][0]);
		t26Form.SelectCell('F',15+i).SetCell("%1.4f", MsrNineOValue[0][i][1]);
		t26Form.SelectCell('G',15+i).SetCell("%1.4f", MsrNineOValue[0][i][2]);
		t26Form.SelectCell('H',15+i).SetCell("%5.0f", MsrNineOValue[0][i][3]);
		t26Form.SelectCell('I',15+i).SetCell("%1.4f", MsrNineOValue[0][i][4]);
    }
//-----------------------------------------------------------------------------------------------
//           格式做完！下面做顏色和框線
//-----------------------------------------------------------------------------------------------
	t26Form.SetVisible(TRUE);
//     objApp.SetUserControl(TRUE);

        SetBackGroundColor(); //預設背景顏色
        SetMaxWindow(0);
        ShowWinNormal();//縮小要顯示的東西
        HideWinNormal();//縮小要消失的東西 
        this->SetDlgItemInt(IDC_MAX,1);        
}