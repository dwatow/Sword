//-----------------------------------------------------------------------------------------------
//  SEC要求量測29點
//        存於Sheet 13點Sheet2 25點 Form Excel檔格式
//
//-----------------------------------------------------------------------------------------------
#include "stdafx.h"
#include "Sword.h"
#include "SwordDlg.h"


const CString CSwordDlg::FromSelector()
{
	CString str;
	CString tempstr;

	pOQCFormSel->GetWindowText(tempstr);
	str.Format("%s\\%s",szCurrentDirectory, tempstr);

    return str;
}

void CSwordDlg::OnButtonOqc() 
{
	switch (pOQCFormSel->GetCurSel())
	{
	case 0:
		OnSave2();//SEC Form
		break;

	case 1:
		{
			xlsFile WRGBD5Form;
			//char OQCFormPath[MAX_PATH];
			//strcpy(OQCFormPath, FromSelector());

			WRGBD5Form.New().SetSheetName(1, "亮度");
			WRGBD5Form.SetVisible(TRUE);

			WRGBD5Form.SelectCell("A1").SetCell("五點的值");


			char X1 = 'A';
			int  X2 = 4;
			WRGBD5Form.SelectCell(X1   , X2 -1).SetCell("W");
			WRGBD5Form.SelectCell(X1   , X2   ).SetCell(OQC_Date[0].w5[0].Lv);
			WRGBD5Form.SelectCell(X1 +2, X2   ).SetCell(OQC_Date[0].w5[1].Lv);
			WRGBD5Form.SelectCell(X1 +1, X2 +1).SetCell(OQC_Date[0].w5[2].Lv);
			WRGBD5Form.SelectCell(X1   , X2 +2).SetCell(OQC_Date[0].w5[3].Lv);
			WRGBD5Form.SelectCell(X1 +2, X2 +2).SetCell(OQC_Date[0].w5[4].Lv);
			
			X1 = 'E';
			X2 = 4;
			WRGBD5Form.SelectCell(X1   , X2 -1).SetCell("D");
			WRGBD5Form.SelectCell(X1   , X2   ).SetCell(OQC_Date[0].d5[0].Lv);
			WRGBD5Form.SelectCell(X1 +2, X2   ).SetCell(OQC_Date[0].d5[0].Lv);
			WRGBD5Form.SelectCell(X1 +1, X2 +1).SetCell(OQC_Date[0].d5[0].Lv);
			WRGBD5Form.SelectCell(X1   , X2 +2).SetCell(OQC_Date[0].d5[0].Lv);
			WRGBD5Form.SelectCell(X1 +2, X2 +2).SetCell(OQC_Date[0].d5[0].Lv);

			X1 = 'A';
			X2 = 9;
			WRGBD5Form.SelectCell(X1   , X2 -1).SetCell("R");
			WRGBD5Form.SelectCell(X1   , X2   ).SetCell(OQC_Date[0].r5[0].Lv);
			WRGBD5Form.SelectCell(X1 +2, X2   ).SetCell(OQC_Date[0].r5[0].Lv);
			WRGBD5Form.SelectCell(X1 +1, X2 +1).SetCell(OQC_Date[0].r5[0].Lv);
			WRGBD5Form.SelectCell(X1   , X2 +2).SetCell(OQC_Date[0].r5[0].Lv);
			WRGBD5Form.SelectCell(X1 +2, X2 +2).SetCell(OQC_Date[0].r5[0].Lv);

			X1 = 'E';
			X2 = 9;
			WRGBD5Form.SelectCell(X1   , X2 -1).SetCell("G");
			WRGBD5Form.SelectCell(X1   , X2   ).SetCell(OQC_Date[0].g5[0].Lv);
			WRGBD5Form.SelectCell(X1 +2, X2   ).SetCell(OQC_Date[0].g5[0].Lv);
			WRGBD5Form.SelectCell(X1 +1, X2 +1).SetCell(OQC_Date[0].g5[0].Lv);
			WRGBD5Form.SelectCell(X1   , X2 +2).SetCell(OQC_Date[0].g5[0].Lv);
			WRGBD5Form.SelectCell(X1 +2, X2 +2).SetCell(OQC_Date[0].g5[0].Lv);

			X1 = 'I';
			X2 = 9;
			WRGBD5Form.SelectCell(X1   , X2 -1).SetCell("B");
			WRGBD5Form.SelectCell(X1   , X2   ).SetCell(OQC_Date[0].b5[0].Lv);
			WRGBD5Form.SelectCell(X1 +2, X2   ).SetCell(OQC_Date[0].b5[0].Lv);
			WRGBD5Form.SelectCell(X1 +1, X2 +1).SetCell(OQC_Date[0].b5[0].Lv);
			WRGBD5Form.SelectCell(X1   , X2 +2).SetCell(OQC_Date[0].b5[0].Lv);
			WRGBD5Form.SelectCell(X1 +2, X2 +2).SetCell(OQC_Date[0].b5[0].Lv);


			
			WRGBD5Form.SetSheetName(2, "色度x");
			
			WRGBD5Form.SelectCell("A1").SetCell("五點的值");
			
			
			X1 = 'A';
			X2 = 4;
			WRGBD5Form.SelectCell(X1   , X2 -1).SetCell("W");
			WRGBD5Form.SelectCell(X1   , X2   ).SetCell(OQC_Date[0].w5[0].x);
			WRGBD5Form.SelectCell(X1 +2, X2   ).SetCell(OQC_Date[0].w5[1].x);
			WRGBD5Form.SelectCell(X1 +1, X2 +1).SetCell(OQC_Date[0].w5[2].x);
			WRGBD5Form.SelectCell(X1   , X2 +2).SetCell(OQC_Date[0].w5[3].x);
			WRGBD5Form.SelectCell(X1 +2, X2 +2).SetCell(OQC_Date[0].w5[4].x);
			
			X1 = 'E';
			X2 = 4;
			WRGBD5Form.SelectCell(X1   , X2 -1).SetCell("D");
			WRGBD5Form.SelectCell(X1   , X2   ).SetCell(OQC_Date[0].d5[0].x);
			WRGBD5Form.SelectCell(X1 +2, X2   ).SetCell(OQC_Date[0].d5[1].x);
			WRGBD5Form.SelectCell(X1 +1, X2 +1).SetCell(OQC_Date[0].d5[2].x);
			WRGBD5Form.SelectCell(X1   , X2 +2).SetCell(OQC_Date[0].d5[3].x);
			WRGBD5Form.SelectCell(X1 +2, X2 +2).SetCell(OQC_Date[0].d5[4].x);
			
			X1 = 'A';
			X2 = 9;
			WRGBD5Form.SelectCell(X1   , X2 -1).SetCell("R");
			WRGBD5Form.SelectCell(X1   , X2   ).SetCell(OQC_Date[0].r5[0].x);
			WRGBD5Form.SelectCell(X1 +2, X2   ).SetCell(OQC_Date[0].r5[1].x);
			WRGBD5Form.SelectCell(X1 +1, X2 +1).SetCell(OQC_Date[0].r5[2].x);
			WRGBD5Form.SelectCell(X1   , X2 +2).SetCell(OQC_Date[0].r5[3].x);
			WRGBD5Form.SelectCell(X1 +2, X2 +2).SetCell(OQC_Date[0].r5[4].x);
			
			X1 = 'E';
			X2 = 9;
			WRGBD5Form.SelectCell(X1   , X2 -1).SetCell("G");
			WRGBD5Form.SelectCell(X1   , X2   ).SetCell(OQC_Date[0].g5[0].x);
			WRGBD5Form.SelectCell(X1 +2, X2   ).SetCell(OQC_Date[0].g5[1].x);
			WRGBD5Form.SelectCell(X1 +1, X2 +1).SetCell(OQC_Date[0].g5[2].x);
			WRGBD5Form.SelectCell(X1   , X2 +2).SetCell(OQC_Date[0].g5[3].x);
			WRGBD5Form.SelectCell(X1 +2, X2 +2).SetCell(OQC_Date[0].g5[4].x);
			
			X1 = 'I';
			X2 = 9;
			WRGBD5Form.SelectCell(X1   , X2 -1).SetCell("B");
			WRGBD5Form.SelectCell(X1   , X2   ).SetCell(OQC_Date[0].b5[0].x);
			WRGBD5Form.SelectCell(X1 +2, X2   ).SetCell(OQC_Date[0].b5[1].x);
			WRGBD5Form.SelectCell(X1 +1, X2 +1).SetCell(OQC_Date[0].b5[2].x);
			WRGBD5Form.SelectCell(X1   , X2 +2).SetCell(OQC_Date[0].b5[3].x);
			WRGBD5Form.SelectCell(X1 +2, X2 +2).SetCell(OQC_Date[0].b5[4].x);

			WRGBD5Form.SetSheetName(3, "色度y");
			
			WRGBD5Form.SelectCell("A1").SetCell("五點的值");
			
			
			X1 = 'A';
			X2 = 4;
			WRGBD5Form.SelectCell(X1   , X2 -1).SetCell("W");
			WRGBD5Form.SelectCell(X1   , X2   ).SetCell(OQC_Date[0].w5[0].y);
			WRGBD5Form.SelectCell(X1 +2, X2   ).SetCell(OQC_Date[0].w5[1].y);
			WRGBD5Form.SelectCell(X1 +1, X2 +1).SetCell(OQC_Date[0].w5[2].y);
			WRGBD5Form.SelectCell(X1   , X2 +2).SetCell(OQC_Date[0].w5[3].y);
			WRGBD5Form.SelectCell(X1 +2, X2 +2).SetCell(OQC_Date[0].w5[4].y);
			
			X1 = 'E';
			X2 = 4;
			WRGBD5Form.SelectCell(X1   , X2 -1).SetCell("D");
			WRGBD5Form.SelectCell(X1   , X2   ).SetCell(OQC_Date[0].d5[0].y);
			WRGBD5Form.SelectCell(X1 +2, X2   ).SetCell(OQC_Date[0].d5[1].y);
			WRGBD5Form.SelectCell(X1 +1, X2 +1).SetCell(OQC_Date[0].d5[2].y);
			WRGBD5Form.SelectCell(X1   , X2 +2).SetCell(OQC_Date[0].d5[3].y);
			WRGBD5Form.SelectCell(X1 +2, X2 +2).SetCell(OQC_Date[0].d5[4].y);
			
			X1 = 'A';
			X2 = 9;
			WRGBD5Form.SelectCell(X1   , X2 -1).SetCell("R");
			WRGBD5Form.SelectCell(X1   , X2   ).SetCell(OQC_Date[0].r5[0].y);
			WRGBD5Form.SelectCell(X1 +2, X2   ).SetCell(OQC_Date[0].r5[1].y);
			WRGBD5Form.SelectCell(X1 +1, X2 +1).SetCell(OQC_Date[0].r5[2].y);
			WRGBD5Form.SelectCell(X1   , X2 +2).SetCell(OQC_Date[0].r5[3].y);
			WRGBD5Form.SelectCell(X1 +2, X2 +2).SetCell(OQC_Date[0].r5[4].y);
			
			X1 = 'E';
			X2 = 9;
			WRGBD5Form.SelectCell(X1   , X2 -1).SetCell("G");
			WRGBD5Form.SelectCell(X1   , X2   ).SetCell(OQC_Date[0].g5[0].y);
			WRGBD5Form.SelectCell(X1 +2, X2   ).SetCell(OQC_Date[0].g5[1].y);
			WRGBD5Form.SelectCell(X1 +1, X2 +1).SetCell(OQC_Date[0].g5[2].y);
			WRGBD5Form.SelectCell(X1   , X2 +2).SetCell(OQC_Date[0].g5[3].y);
			WRGBD5Form.SelectCell(X1 +2, X2 +2).SetCell(OQC_Date[0].g5[4].y);
			
			X1 = 'I';
			X2 = 9;
			WRGBD5Form.SelectCell(X1   , X2 -1).SetCell("B");
			WRGBD5Form.SelectCell(X1   , X2   ).SetCell(OQC_Date[0].b5[0].y);
			WRGBD5Form.SelectCell(X1 +2, X2   ).SetCell(OQC_Date[0].b5[1].y);
			WRGBD5Form.SelectCell(X1 +1, X2 +1).SetCell(OQC_Date[0].b5[2].y);
			WRGBD5Form.SelectCell(X1   , X2 +2).SetCell(OQC_Date[0].b5[3].y);
			WRGBD5Form.SelectCell(X1 +2, X2 +2).SetCell(OQC_Date[0].b5[4].y);

			SetBackGroundColor(); //預設背景顏色
			SetMaxWindow(0);
			ShowWinNormal();//縮小要顯示的東西
			HideWinNormal();//縮小要消失的東西 
			break;
		}

	default:
		{
			xlsFile OQCForm;		//Step 1.叫Excel應用程式
									//Step 2.產生Excel檔案格式
			//做路徑出來
			char OQCFormPath[MAX_PATH];
			//dwCurDirPathLen = GetCurrentDirectory(MAX_PATH,szCurrentDirectory);//抓目前所在的目錄（路徑）
			strcpy(OQCFormPath, FromSelector());
	
		//Step 3.設定Sheet1
			OQCForm.Open(OQCFormPath).SetSheetName(1,"Report");
	
		//Step 4.開始設定內容
		//-----------------------------------------------------------------------------------------------
		//           表格字填完！下面是填入資料！請準備陣列！
		//-----------------------------------------------------------------------------------------------
			OQCForm.SetVisible(TRUE);
		//填入資料
			int i;
			int Tai;
	//		CString x;
	
			for(Tai=0;Tai<10;++Tai)
			{
				//BarCode
				OQCForm.SelectCell('B',Tai+5).SetCell(OQC_Date[Tai].BarCode);
				//W9(Lv)
				for(i=0;i<9;++i)
					OQCForm.SelectCell('C'+i,Tai+5).SetCell("%3.2f", OQC_Date[Tai].w9[i].Lv);
				//W1(x,y)
				OQCForm.SelectCell('L',Tai+5).SetCell("%1.4f", OQC_Date[Tai].w.x);
				OQCForm.SelectCell('M',Tai+5).SetCell("%1.4f", OQC_Date[Tai].w.y);
				//R1(Lv,x,y)
				OQCForm.SelectCell('N',Tai+5).SetCell("%3.2f", OQC_Date[Tai].r.Lv);
				OQCForm.SelectCell('O',Tai+5).SetCell("%1.4f", OQC_Date[Tai].r.x);
				OQCForm.SelectCell('P',Tai+5).SetCell("%1.4f", OQC_Date[Tai].r.y);
				//G1(Lv,x,y)
				OQCForm.SelectCell('Q',Tai+5).SetCell("%3.2f", OQC_Date[Tai].g.Lv);
				OQCForm.SelectCell('R',Tai+5).SetCell("%1.4f", OQC_Date[Tai].g.x);
				OQCForm.SelectCell('S',Tai+5).SetCell("%1.4f", OQC_Date[Tai].g.y);
				//B1(Lv,x,y)
				OQCForm.SelectCell('T',Tai+5).SetCell("%3.2f", OQC_Date[Tai].b.Lv);
				OQCForm.SelectCell('U',Tai+5).SetCell("%1.4f", OQC_Date[Tai].b.x);
				OQCForm.SelectCell('V',Tai+5).SetCell("%1.4f", OQC_Date[Tai].b.y);
				//D1(Lv,x,y)
				OQCForm.SelectCell('W',Tai+5).SetCell("%3.2f", OQC_Date[Tai].d.Lv);
				OQCForm.SelectCell('X',Tai+5).SetCell("%1.4f", OQC_Date[Tai].d.x);
				OQCForm.SelectCell('Y',Tai+5).SetCell("%1.4f", OQC_Date[Tai].d.y);
				//D13(Lv)
	// 			OQCForm.SelectCell('Z',Tai+5).SetCell("%3.2f", OQC_Date[Tai].d13[0].Lv);
	// 			OQCForm.SelectCell('A','A',Tai+5).SetCell("%3.2f", OQC_Date[Tai].d13[1].Lv);
	// 			OQCForm.SelectCell('A','B',Tai+5).SetCell("%3.2f", OQC_Date[Tai].d13[2].Lv);
	// 			OQCForm.SelectCell('A','C',Tai+5).SetCell("%3.2f", OQC_Date[Tai].d13[3].Lv);
	// 			OQCForm.SelectCell('A','D',Tai+5).SetCell("%3.2f", OQC_Date[Tai].d13[4].Lv);
	// 			OQCForm.SelectCell('A','E',Tai+5).SetCell("%3.2f", OQC_Date[Tai].d13[5].Lv);
	// 			OQCForm.SelectCell('A','F',Tai+5).SetCell("%3.2f", OQC_Date[Tai].d13[6].Lv);
	// 			OQCForm.SelectCell('A','G',Tai+5).SetCell("%3.2f", OQC_Date[Tai].d13[7].Lv);
	// 			OQCForm.SelectCell('A','H',Tai+5).SetCell("%3.2f", OQC_Date[Tai].d13[8].Lv);
	// 			OQCForm.SelectCell('A','I',Tai+5).SetCell("%3.2f", OQC_Date[Tai].d13[9].Lv);
	// 			OQCForm.SelectCell('A','J',Tai+5).SetCell("%3.2f", OQC_Date[Tai].d13[10].Lv);
	// 			OQCForm.SelectCell('A','K',Tai+5).SetCell("%3.2f", OQC_Date[Tai].d13[11].Lv);
	// 			OQCForm.SelectCell('A','L',Tai+5).SetCell("%3.2f", OQC_Date[Tai].d13[12].Lv);
				//5nits(Lv)
	 			OQCForm.SelectCell('Z',Tai+5).SetCell("%3.2f", OQC_Date[Tai].d5nits9[0].Lv);
				for(i=0;i<8;++i)
					OQCForm.SelectCell('A','A'+i,Tai+5).SetCell("%3.2f", OQC_Date[Tai].d5nits9[i+1].Lv);
				//5nits(x,y)
				OQCForm.SelectCell('A','J',Tai+5).SetCell("%1.4f", OQC_Date[Tai].d5nits1.x);
				OQCForm.SelectCell('A','K',Tai+5).SetCell("%1.4f", OQC_Date[Tai].d5nits1.y);
			}
	
			for(Tai=0;Tai<10;++Tai)
			{
				//SPD Board編號
				OQCForm.SelectCell('A','N',Tai+5).SetCell(OQC_Date[Tai].DriverDevice);
				if(OQC_Date[Tai].Current) //電流
					OQCForm.SelectCell('A','O',Tai+5).SetCell(OQC_Date[Tai].Current);
				if(OQC_Date[Tai].ChannelNO) //CH
					OQCForm.SelectCell('A','P',Tai+5).SetCell(OQC_Date[Tai].ChannelNO);
				if(OQC_Date[Tai].WorkNum)
					OQCForm.SelectCell('A','Q',Tai+5).SetCell(OQC_Date[Tai].WorkNum);
			}
	
		//Step 3.設定Sheet2
			OQCForm.SetSheetName(2,"黑色25點");
	
	
			for(Tai=0;Tai<10;++Tai)
			{
				//Bar Code
				OQCForm.SelectCell('B',Tai+5).SetCell(OQC_Date[Tai].BarCode);
				//D1(Lv,x,y)
				OQCForm.SelectCell('C',Tai+5).SetCell("%3.2f", OQC_Date[Tai].d.Lv);
				OQCForm.SelectCell('D',Tai+5).SetCell("%1.4f", OQC_Date[Tai].d.x);
				OQCForm.SelectCell('E',Tai+5).SetCell("%1.4f", OQC_Date[Tai].d.y);
				//D25(Lv)
				for(i=0;i<21;++i)
					OQCForm.SelectCell('F'+i,Tai+5).SetCell("%3.2f", OQC_Date[Tai].d25[i].Lv);
				for(i=0;i<4;++i)
				{
				//	CString temp;
				//	temp.Format("OQC_Date[%d].d25[%d].Lv = %f\nOQC_Date[%d].d25[%d].Lv = %f", Tai, i+21,OQC_Date[Tai].d25[i+21].Lv, Tai, i+21,OQC_Date[Tai].d25[i].Lv);
				//	MessageBox(temp);
					OQCForm.SelectCell('A','A'+i,Tai+5).SetCell("%3.2f", OQC_Date[Tai].d25[i+21].Lv);
				}
			}
	
			//Step 4.設定Sheet3
			OQCForm.SetSheetName(3,"白色5點");
			
			
			for(Tai=0;Tai<10;++Tai)
			{
				//Bar Code
				OQCForm.SelectCell('B',Tai+5).SetCell(OQC_Date[Tai].BarCode);
				//W5(Lv,x,y)
				OQCForm.SelectCell('C',Tai+5).SetCell("%3.2f", OQC_Date[Tai].w5[0].Lv);
				OQCForm.SelectCell('D',Tai+5).SetCell("%1.4f", OQC_Date[Tai].w5[0].x);
				OQCForm.SelectCell('E',Tai+5).SetCell("%1.4f", OQC_Date[Tai].w5[0].y);
				//W5(Lv)
				OQCForm.SelectCell('F',Tai+5).SetCell("%3.2f", OQC_Date[Tai].w5[1].Lv);
				OQCForm.SelectCell('G',Tai+5).SetCell("%1.4f", OQC_Date[Tai].w5[1].x);
				OQCForm.SelectCell('H',Tai+5).SetCell("%1.4f", OQC_Date[Tai].w5[1].y);
				
				OQCForm.SelectCell('I',Tai+5).SetCell("%3.2f", OQC_Date[Tai].w5[2].Lv);
				OQCForm.SelectCell('J',Tai+5).SetCell("%1.4f", OQC_Date[Tai].w5[2].x);
				OQCForm.SelectCell('K',Tai+5).SetCell("%1.4f", OQC_Date[Tai].w5[2].y);
				
				OQCForm.SelectCell('L',Tai+5).SetCell("%3.2f", OQC_Date[Tai].w5[3].Lv);
				OQCForm.SelectCell('M',Tai+5).SetCell("%1.4f", OQC_Date[Tai].w5[3].x);
				OQCForm.SelectCell('N',Tai+5).SetCell("%1.4f", OQC_Date[Tai].w5[3].y);
				
				OQCForm.SelectCell('O',Tai+5).SetCell("%3.2f", OQC_Date[Tai].w5[4].Lv);
				OQCForm.SelectCell('P',Tai+5).SetCell("%1.4f", OQC_Date[Tai].w5[4].x);
				OQCForm.SelectCell('Q',Tai+5).SetCell("%1.4f", OQC_Date[Tai].w5[4].y);
			}
			SetBackGroundColor(); //預設背景顏色
			SetMaxWindow(0);
			ShowWinNormal();//縮小要顯示的東西
			HideWinNormal();//縮小要消失的東西 
			break;
		}
	}

	CString temp;
	temp.Format("表格轉出成功!!!");
	MessageBox(temp);

}


