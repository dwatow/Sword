//-----------------------------------------------------------------------------------------------
//靜鷥給的
//    SEC form Excel檔格式
//-----------------------------------------------------------------------------------------------
#include "stdafx.h"
//#include "xlef.h"
#include "Sword.h"
#include "SwordDlg.h"

void CSwordDlg::OnSave2() 
{
	int i;
	if(pOQCTaiSel->GetCurSel() != 10)//就是馬上輸出
		SECForm();
	else							//就是量完再轉檔
		for (i = 0; i<m_Tai; i++)
			SECForm(OQC_Date[i]);


}
//-----------------------------------------------------------------------------------------------
//-----------------------------------------------------------------------------------------------

void CSwordDlg::SECForm(MsrDate &QDate)
{
	int i,j;
	xlsFile SECForm;
	SECForm.New().SetSheetName(1,"Report");
	SECForm.SetVisible(TRUE);


	for(i=1;i<47;i++)
	{
		if(	(i>= 5)&&(i < 8) || 
			(i>= 9)&&(i <13) || 
			(i>=15)&&(i <18) || 
			(i==19)||(i==21) ||
			(i>=23)&&(i <25) || 
			(i>=26)&&(i <28) ||
			(i>=30)&&(i <38) ||
			(i>=39)&&(i <46) )
			SECForm.SelectCell('A', i).SetCellHeight((float)13.50);
		else if(i==14)
			SECForm.SelectCell('A', i).SetCellHeight((float)12.75);
		else if(i==18)
			SECForm.SelectCell('A', i).SetCellHeight((float)16.50);
		else
			SECForm.SelectCell('A', i).SetCellHeight((float)14.25);
	}

    for(j=0;j<10;j++)
    {
        if(j==1)//B
			SECForm.SelectCell('A'+j, 1).SetCellWidth((float)24.63);
        else 
			SECForm.SelectCell('A'+j, 1).SetCellWidth((float)8.38);
    }
	
	
	for(i=2;i<47;i++)
	{

		switch(i)
		{
			case  2: SECForm.SelectCell('B', i).SetCellWidth((float)24.63).SetCellBorder().SetCell("Module Serial Number"); break;
			case  3: SECForm.SelectCell('B', i).SetCellWidth((float)24.63).SetCellBorder().SetCell("LED-LGP Gap "); break;
			case  4: SECForm.SelectCell('B', i).SetCellWidth((float)24.63).SetCellBorder().SetCell("LED Spec. "); break;
			case  5: SECForm.SelectCell('B', i).SetCellWidth((float)24.63).SetCellBorder().SetCell("Sheet Spec. "); break;
			case  9: SECForm.SelectCell('B', i).SetCellWidth((float)24.63).SetCellBorder().SetCell("Center Point "); break;
			case 14: SECForm.SelectCell('B', i).SetCellWidth((float)24.63).SetCellBorder().SetCell("49 Point Brightness "); break;
			case 21: SECForm.SelectCell('B', i).SetCellWidth((float)24.63).SetCellBorder().SetCell("49 Point Avg, "); break;
			case 22: SECForm.SelectCell('B', i).SetCellWidth((float)24.63).SetCellBorder().SetCell("9 Point Avg, "); break;
			case 23: SECForm.SelectCell('B', i).SetCellWidth((float)24.63).SetCellBorder().SetCell("Low Frequency Uniformity\n(5 nits) "); break;
			case 26: SECForm.SelectCell('B', i).SetCellWidth((float)24.63).SetCellBorder().SetCell("Dark Corner 9 Point "); break;
			case 29: SECForm.SelectCell('B', i).SetCellWidth((float)24.63).SetCellBorder().SetCell("Color coordinate uni."); break;
   
			case 30: SECForm.SelectCell('B', i).SetCellWidth((float)24.63).SetCellBorder().SetCell("上左"); break;
			case 31: SECForm.SelectCell('B', i).SetCellWidth((float)24.63).SetCellBorder().SetCell("上中"); break;
			case 32: SECForm.SelectCell('B', i).SetCellWidth((float)24.63).SetCellBorder().SetCell("上右"); break;
			case 33: SECForm.SelectCell('B', i).SetCellWidth((float)24.63).SetCellBorder().SetCell("中左"); break;
			case 34: SECForm.SelectCell('B', i).SetCellWidth((float)24.63).SetCellBorder().SetCell("中中"); break;
			case 35: SECForm.SelectCell('B', i).SetCellWidth((float)24.63).SetCellBorder().SetCell("中右"); break;
			case 36: SECForm.SelectCell('B', i).SetCellWidth((float)24.63).SetCellBorder().SetCell("下左"); break;
			case 37: SECForm.SelectCell('B', i).SetCellWidth((float)24.63).SetCellBorder().SetCell("下中"); break;
			case 38: SECForm.SelectCell('B', i).SetCellWidth((float)24.63).SetCellBorder().SetCell("下右"); break;
   
			case 39: SECForm.SelectCell('B', i).SetCellWidth((float)24.63).SetCellBorder().SetCell("Cosmetic Picture(F/W)"); break;
			default: SECForm.SelectCell('B', i).SetCellWidth((float)24.63).SetCellBorder().SetCell(""); break;
		}
	}
	for(i=2;i<47;i++)
	{
		switch(i)
		{
		case  5:  SECForm.SelectCell('B', i, 'B', i+3).SetMergeCells();   break;
		case  9:  SECForm.SelectCell('B', i, 'B', i+4).SetMergeCells();   break;
		case 14:  SECForm.SelectCell('B', i, 'B', i+6).SetMergeCells();   break;
		case 23:  
		case 26:  SECForm.SelectCell('B', i, 'B', i+2).SetMergeCells();   break;
		case 39:  SECForm.SelectCell('B', i, 'B', i+7).SetMergeCells();   break;
		}
    }

	//設定上排
	SECForm.SelectCell("C2", "I2").SetMergeCells().SetCellBorder(1,3,1);	
	SECForm.SelectCell("B2", "I2").SetCellBorder(1, 3, 1);
    //3列
	SECForm.SelectCell("C3", "E3").SetMergeCells().SetCellBorder().SetCell("上 :  /  /");
	SECForm.SelectCell("F3", "G3").SetMergeCells().SetCellBorder().SetCell("LGP Version");
	SECForm.SelectCell("H3", "I3").SetMergeCells().SetCellBorder();
	SECForm.SelectCell("B3", "I3").SetCellBorder(1,3,1);
    
	//4列
	SECForm.SelectCell("C4", "E4").SetMergeCells().SetCellBorder().SetCell("00 EA /SLED");
	SECForm.SelectCell("F4", "G4").SetMergeCells().SetCellBorder().SetCell("Current");
	SECForm.SelectCell("H4", "I4").SetMergeCells().SetCellBorder().SetCell("mA");
	SECForm.SelectCell("B4", "I4").SetCellBorder(1,3,1);

    //Sheet Spec.列
	SECForm.SelectCell("C5", "I5").SetMergeCells().SetCellBorder();
	SECForm.SelectCell("C6", "I6").SetMergeCells().SetCellBorder();
	SECForm.SelectCell("C8", "I8").SetMergeCells().SetCellBorder();

	SECForm.SelectCell("B5", "I8").SetCellBorder(1,3,1);

	//Center Point列
    for(i=0;i<4;i++){
	for(j=0;j<5;j++){
		if(j == 0)
		{
			switch(i)
			{
				//case 0:  strcpy(buf,"");  break;
			case 1:  SECForm.SelectCell('C'+i, 9+j).SetCellBorder().SetCell("L"); break;
			case 2:  SECForm.SelectCell('C'+i, 9+j).SetCellBorder().SetCell("x"); break;
			case 3:  SECForm.SelectCell('C'+i, 9+j).SetCellBorder().SetCell("y"); break;
			}
		}
		if(i==0)
		{
			switch(j)
			{
				//case 0:  SECForm.SelectCell('C'+i, 9+i).SetCell("");      break;
			case 1:  SECForm.SelectCell('C'+i, 9+j).SetCellBorder().SetCell("White"); break;
			case 2:  SECForm.SelectCell('C'+i, 9+j).SetCellBorder().SetCell("Red");   break;
			case 3:  SECForm.SelectCell('C'+i, 9+j).SetCellBorder().SetCell("Green"); break;
			case 4:  SECForm.SelectCell('C'+i, 9+j).SetCellBorder().SetCell("Blue");  break;
			}
		}
		SECForm.SelectCell('C'+i, 9+j).SetCellBorder();
    }}
	SECForm.SelectCell("G9", "I13").SetMergeCells().SetCellBorder(1,3,1);
	SECForm.SelectCell("B9", "I13").SetCellBorder(1, 3, 1);

	//49 Point Brightness
	for(i=0;i<7;i++){
    for(j=0;j<7;j++){
		SECForm.SelectCell('C'+i, 14+j).SetCellBorder();
        if((i%2==1)&&(j%2==1))
			SECForm.SelectCell('C'+i, 14+j).SetCellColor(6);
    }}
	SECForm.SelectCell("B14", "I20").SetCellBorder(1, 3, 1);

	//49 Point Avg,
	SECForm.SelectCell("C21", "E21").SetMergeCells().SetCellBorder().SetCell("=AVERAGE(C14:I20)");
	//9 Point Avg,
	SECForm.SelectCell("C22", "E22").SetMergeCells().SetCellBorder().SetCell("=AVERAGE(D15,F15,H15,D17,F17,H17,D19,F19,H19)");

	SECForm.SelectCell("F21", "G22").SetMergeCells().SetCellBorder().SetCell("Uniformity");
	SECForm.SelectCell("H21", "I22").SetMergeCells().SetCellBorder().SetCell("=MIN(D15,F15,H15,D17,F17,H17,D19,F19,H19)/MAX(D15,F15,H15,D17,F17,H17,D19,F19,H19)");
	SECForm.SelectCell("B21", "I22").SetCellBorder(1,3,1);

	//Low Frequency Uniformity (5 nits)
	for(i=0;i<7;i++){
    for(j=0;j<16;j++){
		SECForm.SelectCell('C'+i, 23+j).SetCellBorder();
	}}

	SECForm.SelectCell("F23").SetCell("=ABS(C23-D23)/D23");
	SECForm.SelectCell("F24").SetCell("=ABS(C24-D24)/D23");
	SECForm.SelectCell("F25").SetCell("=ABS(C25-D25)/D23");
	SECForm.SelectCell("G23", "G25").SetMergeCells();

	SECForm.SelectCell("H23").SetCell("=ABS(E23-D23)/D23");
	SECForm.SelectCell("H24").SetCell("=ABS(E24-D24)/D24");
	SECForm.SelectCell("H25").SetCell("=ABS(E25-D25)/D25");
	SECForm.SelectCell("F26", "I28").SetMergeCells().SetCellColor(16);

	SECForm.SelectCell("B23", "I25").SetCellBorder(1, 3, 1);

	//Dark Corner 9 Point 
	SECForm.SelectCell("B26", "I28").SetCellBorder(1, 3, 1);

	//Color coordinate unit.
	SECForm.SelectCell("C29").SetCell("L");
	SECForm.SelectCell("D29").SetCell("x");
	SECForm.SelectCell("E29").SetCell("y");
	SECForm.SelectCell("F29").SetCell("ΔT");
	SECForm.SelectCell("G29").SetCell("Δuv");
	SECForm.SelectCell("H29", "I38").SetMergeCells().SetCellColor(16).SetCellBorder(1, 3, 1);

	//Cosmetic Picture(F/W)
	//Cosmetic Picture(F/W)
	SECForm.SelectCell("B21", "I22").SetCellBorder().SetCellBorder(1, 3, 1);
	SECForm.SelectCell("C23", "E28").SetCellBorder();
	SECForm.SelectCell("C39", "I46").SetMergeCells();//.SetCellBorder(1, 3, 1);
	SECForm.SelectCell("B39", "I46").SetCellBorder(1, 3, 1);

	SECForm.SelectCell("B29", "G29").SetCellBorder(1, 3, 1);
	SECForm.SelectCell("B2", "B46").SetCellBorder(1, 3, 1);


//-----------------------------------------------------------------------------------------------
//           表格字填完！下面是填入資料！請準備陣列！
//-----------------------------------------------------------------------------------------------
	SECForm.SetVisible(TRUE);

    SECForm.SelectCell("D10").SetCell("%3.2f", QDate.w.Lv); //[color][L]
    SECForm.SelectCell("E10").SetCell("%1.4f", QDate.w.x); //[color][x]
	SECForm.SelectCell("F10").SetCell("%1.4f", QDate.w.y); //[color][y]

    SECForm.SelectCell("D11").SetCell("%3.2f", QDate.r.Lv); //[color][L]
    SECForm.SelectCell("E11").SetCell("%1.4f", QDate.r.x); //[color][x]
	SECForm.SelectCell("F11").SetCell("%1.4f", QDate.r.y); //[color][y]

    SECForm.SelectCell("D12").SetCell("%3.2f", QDate.g.Lv); //[color][L]
    SECForm.SelectCell("E12").SetCell("%1.4f", QDate.g.x); //[color][x]
	SECForm.SelectCell("F12").SetCell("%1.4f", QDate.g.y); //[color][y]

	SECForm.SelectCell("D13").SetCell("%3.2f", QDate.b.Lv); //[color][L]
    SECForm.SelectCell("E13").SetCell("%1.4f", QDate.b.x); //[color][x]
	SECForm.SelectCell("F13").SetCell("%1.4f", QDate.b.y); //[color][y]

	
	//w,49點亮度值
	for(i=0;i<7;i++){//橫的
    for(j=0;j<7;j++){//直的
		SECForm.SelectCell('C'+i, 14+j).SetCell("%3.2f",QDate.w49[i+j*7].Lv);
    }}
	
	//白色 9點全部值
	for(i=0;i<5;i++){//value
    for(j=0;j<9;j++){//n
		//Color 012'3'4 = 'W'RGB,other		// MsrNineOValue[Color][n][value]
        	 if(i == 0)	    SECForm.SelectCell('C'+i, 30+j).SetCell("%3.2f",QDate.w9[j].Lv); //[W][1-9][L]
		else if(i == 1)     SECForm.SelectCell('C'+i, 30+j).SetCell("%1.4f",QDate.w9[j].x); //[W][1-9][T]
		else if(i == 2)     SECForm.SelectCell('C'+i, 30+j).SetCell("%1.4f",QDate.w9[j].y); //[W][1-9][T]
		else if(i == 3)     SECForm.SelectCell('C'+i, 30+j).SetCell("%4d"  ,QDate.w9[j].T); //[W][1-9][T]
		else if(i == 4)     SECForm.SelectCell('C'+i, 30+j).SetCell("%1.4f",QDate.w9[j].Duv); //[W][1-9][T]
	
// 		CString temp;
// 		temp.Format("SECForm()寫入Excel\nSECForm.SelectCell(\'C\'+i, 30+j).SetCell(\"%%4d\",QDate.w9[j].T);\n%d", 
// 			QDate.w9[j].T);
// 		MessageBox(temp);


	}}
	//黑色 9點全部值（四角漏光）
	SECForm.SelectCell("C26").SetCell("%3.2f",QDate.d25[0].Lv);
	SECForm.SelectCell("D26").SetCell("%3.2f",QDate.d25[5].Lv);
	SECForm.SelectCell("E26").SetCell("%3.2f",QDate.d25[6].Lv);

	SECForm.SelectCell("C27").SetCell("%3.2f",QDate.d25[11].Lv);
	SECForm.SelectCell("D27").SetCell("%3.2f",QDate.d25[12].Lv);
	SECForm.SelectCell("E27").SetCell("%3.2f",QDate.d25[13].Lv);
	
	SECForm.SelectCell("C28").SetCell("%3.2f",QDate.d25[14].Lv);
	SECForm.SelectCell("D28").SetCell("%3.2f",QDate.d25[19].Lv);
	SECForm.SelectCell("E28").SetCell("%3.2f",QDate.d25[20].Lv);

// 	for(i=0;i<3;i++){//value
//     for(j=0;j<3;j++){//n
		//SECForm.SelectCell('C'+j,26+i).SetCell("%3.2f",QDate.d25[i*3+j].Lv);
// 	}}
	
	//5nits 
	for(i=0;i<3;i++){//value
    for(j=0;j<3;j++){//n
		SECForm.SelectCell('C'+j,23+i).SetCell("%3.2f",QDate.d5nits9[i*3+j].Lv);
    }}
	//-----------------------------------------------------------------------------------------------
	//           格式做完！下面做顏色和框線
	//-----------------------------------------------------------------------------------------------
	//全部一起//字11
	SECForm.SelectCell("A1", "IV65536").SetCellHeight((short)14.25).SetFont("Verdana").SetFontSize((short)11);

	SetBackGroundColor(); //預設背景顏色
    SetMaxWindow(0);
    ShowWinNormal();//縮小要顯示的東西
    HideWinNormal();//縮小要消失的東西
    this->SetDlgItemInt(IDC_MAX,1);
}

void CSwordDlg::SECForm()
{
	int i,j;
	xlsFile SECForm;
	SECForm.New().SetSheetName(1,"Report");
	SECForm.SetVisible(TRUE);


	for(i=1;i<47;i++)
	{
		if(	(i>= 5)&&(i < 8) || 
			(i>= 9)&&(i <13) || 
			(i>=15)&&(i <18) || 
			(i==19)||(i==21) ||
			(i>=23)&&(i <25) || 
			(i>=26)&&(i <28) ||
			(i>=30)&&(i <38) ||
			(i>=39)&&(i <46) )
			SECForm.SelectCell('A', i).SetCellHeight((float)13.50);
		else if(i==14)
			SECForm.SelectCell('A', i).SetCellHeight((float)12.75);
		else if(i==18)
			SECForm.SelectCell('A', i).SetCellHeight((float)16.50);
		else
			SECForm.SelectCell('A', i).SetCellHeight((float)14.25);
	}

    for(j=0;j<10;j++)
    {
        if(j==1)//B
			SECForm.SelectCell('A'+j, 1).SetCellWidth((float)24.63);
        else 
			SECForm.SelectCell('A'+j, 1).SetCellWidth((float)8.38);
    }
	
	
	for(i=2;i<47;i++)
	{

		switch(i)
		{
			case  2: SECForm.SelectCell('B', i).SetCellWidth((float)24.63).SetCellBorder().SetCell("Module Serial Number"); break;
			case  3: SECForm.SelectCell('B', i).SetCellWidth((float)24.63).SetCellBorder().SetCell("LED-LGP Gap "); break;
			case  4: SECForm.SelectCell('B', i).SetCellWidth((float)24.63).SetCellBorder().SetCell("LED Spec. "); break;
			case  5: SECForm.SelectCell('B', i).SetCellWidth((float)24.63).SetCellBorder().SetCell("Sheet Spec. "); break;
			case  9: SECForm.SelectCell('B', i).SetCellWidth((float)24.63).SetCellBorder().SetCell("Center Point "); break;
			case 14: SECForm.SelectCell('B', i).SetCellWidth((float)24.63).SetCellBorder().SetCell("49 Point Brightness "); break;
			case 21: SECForm.SelectCell('B', i).SetCellWidth((float)24.63).SetCellBorder().SetCell("49 Point Avg, "); break;
			case 22: SECForm.SelectCell('B', i).SetCellWidth((float)24.63).SetCellBorder().SetCell("9 Point Avg, "); break;
			case 23: SECForm.SelectCell('B', i).SetCellWidth((float)24.63).SetCellBorder().SetCell("Low Frequency Uniformity\n(5 nits) "); break;
			case 26: SECForm.SelectCell('B', i).SetCellWidth((float)24.63).SetCellBorder().SetCell("Dark Corner 9 Point "); break;
			case 29: SECForm.SelectCell('B', i).SetCellWidth((float)24.63).SetCellBorder().SetCell("Color coordinate uni."); break;
   
			case 30: SECForm.SelectCell('B', i).SetCellWidth((float)24.63).SetCellBorder().SetCell("上左"); break;
			case 31: SECForm.SelectCell('B', i).SetCellWidth((float)24.63).SetCellBorder().SetCell("上中"); break;
			case 32: SECForm.SelectCell('B', i).SetCellWidth((float)24.63).SetCellBorder().SetCell("上右"); break;
			case 33: SECForm.SelectCell('B', i).SetCellWidth((float)24.63).SetCellBorder().SetCell("中左"); break;
			case 34: SECForm.SelectCell('B', i).SetCellWidth((float)24.63).SetCellBorder().SetCell("中中"); break;
			case 35: SECForm.SelectCell('B', i).SetCellWidth((float)24.63).SetCellBorder().SetCell("中右"); break;
			case 36: SECForm.SelectCell('B', i).SetCellWidth((float)24.63).SetCellBorder().SetCell("下左"); break;
			case 37: SECForm.SelectCell('B', i).SetCellWidth((float)24.63).SetCellBorder().SetCell("下中"); break;
			case 38: SECForm.SelectCell('B', i).SetCellWidth((float)24.63).SetCellBorder().SetCell("下右"); break;
   
			case 39: SECForm.SelectCell('B', i).SetCellWidth((float)24.63).SetCellBorder().SetCell("Cosmetic Picture(F/W)"); break;
			default: SECForm.SelectCell('B', i).SetCellWidth((float)24.63).SetCellBorder().SetCell(""); break;
		}
	}
	for(i=2;i<47;i++)
	{
		switch(i)
		{
		case  5:  SECForm.SelectCell('B', i, 'B', i+3).SetMergeCells();   break;
		case  9:  SECForm.SelectCell('B', i, 'B', i+4).SetMergeCells();   break;
		case 14:  SECForm.SelectCell('B', i, 'B', i+6).SetMergeCells();   break;
		case 23:  
		case 26:  SECForm.SelectCell('B', i, 'B', i+2).SetMergeCells();   break;
		case 39:  SECForm.SelectCell('B', i, 'B', i+7).SetMergeCells();   break;
		}
    }

	//設定上排
	SECForm.SelectCell("C2", "I2").SetMergeCells().SetCellBorder(1,3,1);	
	SECForm.SelectCell("B2", "I2").SetCellBorder(1, 3, 1);
    //3列
	SECForm.SelectCell("C3", "E3").SetMergeCells().SetCellBorder().SetCell("上 :  /  /");
	SECForm.SelectCell("F3", "G3").SetMergeCells().SetCellBorder().SetCell("LGP Version");
	SECForm.SelectCell("H3", "I3").SetMergeCells().SetCellBorder();
	SECForm.SelectCell("B3", "I3").SetCellBorder(1,3,1);
    
	//4列
	SECForm.SelectCell("C4", "E4").SetMergeCells().SetCellBorder().SetCell("00 EA /SLED");
	SECForm.SelectCell("F4", "G4").SetMergeCells().SetCellBorder().SetCell("Current");
	SECForm.SelectCell("H4", "I4").SetMergeCells().SetCellBorder().SetCell("mA");
	SECForm.SelectCell("B4", "I4").SetCellBorder(1,3,1);

    //Sheet Spec.列
	SECForm.SelectCell("C5", "I5").SetMergeCells().SetCellBorder();
	SECForm.SelectCell("C6", "I6").SetMergeCells().SetCellBorder();
	SECForm.SelectCell("C8", "I8").SetMergeCells().SetCellBorder();

	SECForm.SelectCell("B5", "I8").SetCellBorder(1,3,1);

	//Center Point列
    for(i=0;i<4;i++){
	for(j=0;j<5;j++){
		if(j == 0)
		{
			switch(i)
			{
				//case 0:  strcpy(buf,"");  break;
			case 1:  SECForm.SelectCell('C'+i, 9+j).SetCellBorder().SetCell("L"); break;
			case 2:  SECForm.SelectCell('C'+i, 9+j).SetCellBorder().SetCell("x"); break;
			case 3:  SECForm.SelectCell('C'+i, 9+j).SetCellBorder().SetCell("y"); break;
			}
		}
		if(i==0)
		{
			switch(j)
			{
				//case 0:  SECForm.SelectCell('C'+i, 9+i).SetCell("");      break;
			case 1:  SECForm.SelectCell('C'+i, 9+j).SetCellBorder().SetCell("White"); break;
			case 2:  SECForm.SelectCell('C'+i, 9+j).SetCellBorder().SetCell("Red");   break;
			case 3:  SECForm.SelectCell('C'+i, 9+j).SetCellBorder().SetCell("Green"); break;
			case 4:  SECForm.SelectCell('C'+i, 9+j).SetCellBorder().SetCell("Blue");  break;
			}
		}
		SECForm.SelectCell('C'+i, 9+j).SetCellBorder();
    }}
	SECForm.SelectCell("G9", "I13").SetMergeCells().SetCellBorder(1,3,1);
	SECForm.SelectCell("B9", "I13").SetCellBorder(1, 3, 1);

	//49 Point Brightness
	for(i=0;i<7;i++){
    for(j=0;j<7;j++){
		SECForm.SelectCell('C'+i, 14+j).SetCellBorder();
        if((i%2==1)&&(j%2==1))
			SECForm.SelectCell('C'+i, 14+j).SetCellColor(6);
    }}
	SECForm.SelectCell("B14", "I20").SetCellBorder(1, 3, 1);

	//49 Point Avg,
	SECForm.SelectCell("C21", "E21").SetMergeCells().SetCellBorder().SetCell("=AVERAGE(C14:I20)");
	//9 Point Avg,
	SECForm.SelectCell("C22", "E22").SetMergeCells().SetCellBorder().SetCell("=AVERAGE(D15,F15,H15,D17,F17,H17,D19,F19,H19)");

	SECForm.SelectCell("F21", "G22").SetMergeCells().SetCellBorder().SetCell("Uniformity");
	SECForm.SelectCell("H21", "I22").SetMergeCells().SetCellBorder().SetCell("=MIN(D15,F15,H15,D17,F17,H17,D19,F19,H19)/MAX(D15,F15,H15,D17,F17,H17,D19,F19,H19)");
	SECForm.SelectCell("B21", "I22").SetCellBorder(1,3,1);

	//Low Frequency Uniformity (5 nits)
	for(i=0;i<7;i++){
    for(j=0;j<16;j++){
		SECForm.SelectCell('C'+i, 23+j).SetCellBorder();
	}}

	SECForm.SelectCell("F23").SetCell("=ABS(C23-D23)/D23");
	SECForm.SelectCell("F24").SetCell("=ABS(C24-D24)/D23");
	SECForm.SelectCell("F25").SetCell("=ABS(C25-D25)/D23");
	SECForm.SelectCell("G23", "G25").SetMergeCells();

	SECForm.SelectCell("H23").SetCell("=ABS(E23-D23)/D23");
	SECForm.SelectCell("H24").SetCell("=ABS(E24-D24)/D24");
	SECForm.SelectCell("H25").SetCell("=ABS(E25-D25)/D25");
	SECForm.SelectCell("F26", "I28").SetMergeCells().SetCellColor(16);

	SECForm.SelectCell("B23", "I25").SetCellBorder(1, 3, 1);

	//Dark Corner 9 Point 
	SECForm.SelectCell("B26", "I28").SetCellBorder(1, 3, 1);

	//Color coordinate unit.
	SECForm.SelectCell("C29").SetCell("L");
	SECForm.SelectCell("D29").SetCell("x");
	SECForm.SelectCell("E29").SetCell("y");
	SECForm.SelectCell("F29").SetCell("ΔT");
	SECForm.SelectCell("G29").SetCell("Δuv");
	SECForm.SelectCell("H29", "I38").SetMergeCells().SetCellColor(16).SetCellBorder(1, 3, 1);

	//Cosmetic Picture(F/W)
	//Cosmetic Picture(F/W)
	SECForm.SelectCell("B21", "I22").SetCellBorder().SetCellBorder(1, 3, 1);
	SECForm.SelectCell("C23", "E28").SetCellBorder();
	SECForm.SelectCell("C39", "I46").SetMergeCells();//.SetCellBorder(1, 3, 1);
	SECForm.SelectCell("B39", "I46").SetCellBorder(1, 3, 1);

	SECForm.SelectCell("B29", "G29").SetCellBorder(1, 3, 1);
	SECForm.SelectCell("B2", "B46").SetCellBorder(1, 3, 1);


//-----------------------------------------------------------------------------------------------
//           表格字填完！下面是填入資料！請準備陣列！
//-----------------------------------------------------------------------------------------------
	SECForm.SetVisible(TRUE);

    SECForm.SelectCell("D10").SetCell("%3.2f", MsrCenter[0][0]); //[color][L]
    SECForm.SelectCell("E10").SetCell("%1.4f", MsrCenter[0][1]); //[color][x]
	SECForm.SelectCell("F10").SetCell("%1.4f", MsrCenter[0][2]); //[color][y]

    SECForm.SelectCell("D11").SetCell("%3.2f", MsrCenter[1][0]); //[color][L]
    SECForm.SelectCell("E11").SetCell("%1.4f", MsrCenter[1][1]); //[color][x]
	SECForm.SelectCell("F11").SetCell("%1.4f", MsrCenter[1][2]); //[color][y]

    SECForm.SelectCell("D12").SetCell("%3.2f", MsrCenter[2][0]); //[color][L]
    SECForm.SelectCell("E12").SetCell("%1.4f", MsrCenter[2][1]); //[color][x]
	SECForm.SelectCell("F12").SetCell("%1.4f", MsrCenter[2][2]); //[color][y]

	SECForm.SelectCell("D13").SetCell("%3.2f", MsrCenter[3][0]); //[color][L]
    SECForm.SelectCell("E13").SetCell("%1.4f", MsrCenter[3][1]); //[color][x]
	SECForm.SelectCell("F13").SetCell("%1.4f", MsrCenter[3][2]); //[color][y]

	
	//w,49點亮度值
	for(i=0;i<7;i++){//橫的
    for(j=0;j<7;j++){//直的
		SECForm.SelectCell('C'+i, 14+j).SetCell("%3.2f",MsrFortyNineOValue[0][i+j*7][0]);
    }}
	//白色 9點全部值
	for(i=0;i<5;i++){//value
    for(j=0;j<9;j++){//n
		//Color 012'3'4 = 'W'RGB,other		// MsrNineOValue[Color][n][value]
        	 if(i == 0)	    SECForm.SelectCell('C'+i, 30+j).SetCell("%3.2f", MsrNineOValue[0][j][0]); //[W][1-9][L]
		else if(i == 1)     SECForm.SelectCell('C'+i, 30+j).SetCell("%1.4f", MsrNineOValue[0][j][1]); //[W][1-9][T]
		else if(i == 2)     SECForm.SelectCell('C'+i, 30+j).SetCell("%1.4f", MsrNineOValue[0][j][2]); //[W][1-9][T]
		else if(i == 3)     SECForm.SelectCell('C'+i, 30+j).SetCell("%4.0f", MsrNineOValue[0][j][3]); //[W][1-9][T]
		else if(i == 4)     SECForm.SelectCell('C'+i, 30+j).SetCell("%1.4f", MsrNineOValue[0][j][4]); //[W][1-9][T]
	}}
	//黑色 9點全部值（四角漏光）
	for(i=0;i<3;i++){//value
    for(j=0;j<3;j++){//n
		SECForm.SelectCell('C'+j,26+i).SetCell("%3.2f",MsrNineOValue[4][i*3+j][0]);
	}}
	
	//5nits 
	for(i=0;i<3;i++){//value
    for(j=0;j<3;j++){//n
		SECForm.SelectCell('C'+j,23+i).SetCell("%3.2f", FiveNits[i*3+j][0]);
    }}
	//-----------------------------------------------------------------------------------------------
	//           格式做完！下面做顏色和框線
	//-----------------------------------------------------------------------------------------------
	//全部一起//字11
	SECForm.SelectCell("A1", "IV65536").SetCellHeight((short)14.25).SetFont("Verdana").SetFontSize((short)11);

	SetBackGroundColor(); //預設背景顏色
    SetMaxWindow(0);
    ShowWinNormal();//縮小要顯示的東西
    HideWinNormal();//縮小要消失的東西
    this->SetDlgItemInt(IDC_MAX,1);
}