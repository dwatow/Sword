#include "stdafx.h"
//#include <comdef.h>
#include "Sword.h"
#include "SwordDlg.h"
//#include "excel.h"

/*
void CSwordDlg::Msr9P()	            九點量測
void CSwordDlg::Msr5Nits9P()		5nits量測
void CSwordDlg::Msr1P()             中心點量測
void CSwordDlg::Msr5P()             5點量測
void CSwordDlg::MsrFine5Nits()      找5nits的量測
void CSwordDlg::Msr49P()			49點量測
void CSwordDlg::Msr13P()			13點量測
void CSwordDlg::Msr25P()			25點量測
void CSwordDlg::Msr29P()			29點量測
void CSwordDlg::MsrFLIC()			FLK量測
void CSwordDlg::MsrLab()			實驗量測用
*/

void CSwordDlg::Msr9P()
{
    //重畫視窗    
    Invalidate();
    UpdateWindow();

	m_ICa.SetDisplayMode(0);//設定為x,y,Lv顯示模式

    int i;
    for(Few = 0;Few < 9;++Few)
    {
		CurrentPoint = (GetBackGroundColor() == RGB(0,0,0)) ? SetD9Point(Few) : Set9Point(Few,m_floatFromEdge);//定位圈圈的位置
        DrawColorLabel(SetCtlPointQuadrant(CurrentPoint),CircleRadius);						//控制色塊

        Sleep(500);//量測動作停止或暫停（等使用者拿起量筒）
        if (MerOperation)
			MsrAI(AutoSelectMode(GetBackGroundColor()),TRUE);		//偵測.動畫.畫圈
		else
			MsrAct(TRUE);

        Invalidate();
        UpdateWindow();
        
        if(PassStopColor())//STOP
        {
            break;
        }
        else if(PassBackColor())//BACK
        {
            //顯示預覽數據
			//CString str;	str.Format("Few=%d",Few);	MessageBox(str);
			for(i=0;i<Few;++i)
            {
				CurrentPoint = (GetBackGroundColor() == RGB(0,0,0)) ? SetD9Point(i) : Set9Point(i,m_floatFromEdge);
                DrawMsrLabel(CurrentPoint, CircleRadius, &MsrNineOValue[ColorLabel][i][0]);//寫字
            }
			//str.Format("Few=%d",Few);	MessageBox(str);
			Few -= 2;
			if(Few < -1)
				break;
			else
				continue;
        }
        else
        {        
            Sleep(100);
            m_ICa.Measure(0);
			ColorLabel = SelectColorLabel(GetBackGroundColor());//用顏色標籤選擇儲存的記憶體

            //存值到記憶體
            //MsrNineOValue[   Color  ][ n ][value]
            MsrNineOValue[ColorLabel][Few][0] = static_cast<double>(m_IProbe.GetLv());

            MsrNineOValue[ColorLabel][Few][1] = static_cast<double>(m_IProbe.GetSx());
            MsrNineOValue[ColorLabel][Few][2] = static_cast<double>(m_IProbe.GetSy());

            MsrNineOValue[ColorLabel][Few][3] = static_cast<long>  (m_IProbe.GetT());
            MsrNineOValue[ColorLabel][Few][4] = static_cast<double>(m_IProbe.GetDuv());

            MsrNineOValue[ColorLabel][Few][5] = static_cast<double>(m_IProbe.GetUd());
            MsrNineOValue[ColorLabel][Few][6] = static_cast<double>(m_IProbe.GetVd());

            MsrNineOValue[ColorLabel][Few][7] = static_cast<double>(m_IProbe.GetX());
            MsrNineOValue[ColorLabel][Few][8] = static_cast<double>(m_IProbe.GetY());
            MsrNineOValue[ColorLabel][Few][9] = static_cast<double>(m_IProbe.GetZ());

            //顯示預覽數據
            for(i=0;i<=Few;++i)
            {
				CurrentPoint = (GetBackGroundColor() == RGB(0,0,0)) ? SetD9Point(i) : Set9Point(i,m_floatFromEdge);
				DrawMsrLabel(CurrentPoint, CircleRadius, &MsrNineOValue[ColorLabel][i][0]);//寫字
            }
            DrawMsrLabel(CurrentPoint, CircleRadius, &MsrNineOValue[ColorLabel][Few][0]);//寫字

        }
    }//for(Few = 0;Few<9;++Few)
    for(i=0;i<10;++i)//value
    {
                     //中心點[白色][所有值] = 九點[白色][中心點][所有值] 
		if(m_bool_9cvr1)    MsrCenter[0][i] = MsrNineOValue[0][4][i];//填入中心點量測值
        if(m_bool_9cvr49)//填入49點量測值
        {
             // 49點[白色][點數][所有值] = 九點[白色][點數][所有值]
            MsrFortyNineOValue[0][ 8][i] = MsrNineOValue[0][0][i];
            MsrFortyNineOValue[0][10][i] = MsrNineOValue[0][1][i];
            MsrFortyNineOValue[0][12][i] = MsrNineOValue[0][2][i];

            MsrFortyNineOValue[0][22][i] = MsrNineOValue[0][3][i];
            MsrFortyNineOValue[0][24][i] = MsrNineOValue[0][4][i];
            MsrFortyNineOValue[0][26][i] = MsrNineOValue[0][5][i];

            MsrFortyNineOValue[0][36][i] = MsrNineOValue[0][6][i];
            MsrFortyNineOValue[0][38][i] = MsrNineOValue[0][7][i];
            MsrFortyNineOValue[0][40][i] = MsrNineOValue[0][8][i];
        }
    }
}


void CSwordDlg::Msr5P()
{
    //重畫視窗    
    Invalidate();
    UpdateWindow();
	
	m_ICa.SetDisplayMode(0);//設定為x,y,Lv顯示模式
	
    int i;
    for(Few = 0;Few < 5;++Few)
    {
		CurrentPoint = Set5Point(Few, m_float5FromEdge);//定位圈圈的位置
        DrawColorLabel(SetCtlPointQuadrant(CurrentPoint),CircleRadius);						//控制色塊
		
        Sleep(500);//量測動作停止或暫停（等使用者拿起量筒）
        if (MerOperation)
			MsrAI(AutoSelectMode(GetBackGroundColor()),TRUE);		//偵測.動畫.畫圈
		else
			MsrAct(TRUE);
		
        Invalidate();
        UpdateWindow();
        
        if(PassStopColor())//STOP
        {
            break;
        }
        else if(PassBackColor())//BACK
        {
            //顯示預覽數據
			//CString str;	str.Format("Few=%d",Few);	MessageBox(str);
			for(i=0;i<Few;++i)
            {
				CurrentPoint = Set5Point(i);
                DrawMsrLabel(CurrentPoint, CircleRadius, &MsrFiveOValue[ColorLabel][i][0]);//寫字
            }
			//str.Format("Few=%d",Few);	MessageBox(str);
			Few -= 2;
			if(Few < -1)
				break;
			else
				continue;
        }
        else
        {        
            Sleep(100);
            m_ICa.Measure(0);
			ColorLabel = SelectColorLabel(GetBackGroundColor());//用顏色標籤選擇儲存的記憶體
			
            //存值到記憶體
            //MsrNineOValue[   Color  ][ n ][value]
			MsrFiveOValue[ColorLabel][Few][0] = static_cast<double>(m_IProbe.GetLv());
			
			MsrFiveOValue[ColorLabel][Few][1] = static_cast<double>(m_IProbe.GetSx());
			MsrFiveOValue[ColorLabel][Few][2] = static_cast<double>(m_IProbe.GetSy());
			
			MsrFiveOValue[ColorLabel][Few][3] = static_cast<long>  (m_IProbe.GetT());
			MsrFiveOValue[ColorLabel][Few][4] = static_cast<double>(m_IProbe.GetDuv());

			MsrFiveOValue[ColorLabel][Few][5] = static_cast<double>(m_IProbe.GetUd());
			MsrFiveOValue[ColorLabel][Few][6] = static_cast<double>(m_IProbe.GetVd());
			
			MsrFiveOValue[ColorLabel][Few][7] = static_cast<double>(m_IProbe.GetX());
			MsrFiveOValue[ColorLabel][Few][8] = static_cast<double>(m_IProbe.GetY());
			MsrFiveOValue[ColorLabel][Few][9] = static_cast<double>(m_IProbe.GetZ());


            //顯示預覽數據
            for(i=0;i<=Few;++i)
            {
				CurrentPoint = Set5Point(i);
				DrawMsrLabel(CurrentPoint, CircleRadius, &MsrFiveOValue[ColorLabel][i][0]);//寫字
            }
            DrawMsrLabel(CurrentPoint, CircleRadius, &MsrFiveOValue[ColorLabel][Few][0]);//寫字
			
        }
    }//for(Few = 0;Few<9;++Few)

	//填入中心點量測值

	//中心點[白色][所有值]       = 九點[白色][中心點][所有值]
	MsrCenter[0][i]              = MsrFiveOValue[0][2][i];
	// 9點[白色][點數][所有值]   = 九點[白色][點數][所有值]
    MsrNineOValue[0][5][i]       = MsrFiveOValue[0][2][i];
	// 49點[白色][點數][所有值]  = 九點[白色][點數][所有值]
    MsrFortyNineOValue[0][24][i] = MsrFiveOValue[0][2][i];
}

void CSwordDlg::Msr5Nits9P() 
{
    Sleep(500);
    int i;
    for(Few = 0; Few < 9; ++Few)
    {
		if(Few != 4)
		{

			CurrentPoint = Set5nits9Point(Few,m_floatFromEdge);					//定位點，In第幾點，Out座標（以九點計）
			DrawColorLabel(SetCtlPointQuadrant(CurrentPoint),CircleRadius);	//控制色塊

			Sleep(500);//量測動作停止或暫停（等使用者拿起量筒）
			//MsrAI(AutoSelectMode(GetBackGroundColor()),TRUE);				//偵測.動畫.畫圈
			if (MerOperation)
				MsrAI(AutoSelectMode(GetBackGroundColor()),TRUE);		//偵測.動畫.畫圈
			else
			MsrAct(TRUE);

			Invalidate();
			UpdateWindow();
        
			if(PassStopColor())//STOP
			{
				break;
			}
			else if(PassBackColor())//BACK
			{
				//顯示預覽數據
				for(i=0;i<Few;++i)
				{
					CurrentPoint = Set5nits9Point(i,6);//定位點，In第幾點，Out座標（以九點計）
					DrawMsrLabel(CurrentPoint, CircleRadius, &FiveNits[i][0]);//寫字
				}
				Few -= 2;
				if(Few < -1)
					break;
				else
					continue;
			}
			else
			{
				Sleep(100);
				m_ICa.Measure(0);
		//存值到記憶體
		//      FiveNits[n][value]
				FiveNits[Few][0] = static_cast<double>(m_IProbe.GetLv());

				FiveNits[Few][1] = static_cast<double>(m_IProbe.GetSx());
				FiveNits[Few][2] = static_cast<double>(m_IProbe.GetSy());

				FiveNits[Few][3] = static_cast<long>  (m_IProbe.GetT());
				FiveNits[Few][4] = static_cast<double>(m_IProbe.GetDuv());

				FiveNits[Few][5] = static_cast<double>(m_IProbe.GetUd());
				FiveNits[Few][6] = static_cast<double>(m_IProbe.GetVd());

				FiveNits[Few][7] = static_cast<double>(m_IProbe.GetX());
				FiveNits[Few][8] = static_cast<double>(m_IProbe.GetY());
				FiveNits[Few][9] = static_cast<double>(m_IProbe.GetZ());

				//FiveNits[Few][10] = static_cast<double>(m_IProbe.GetFlckrFMA());
		//顯示預覽數據
				for(i=0;i<=Few;++i)
				{
					CurrentPoint = Set5nits9Point(i,6);//定位點，In第幾點，Out座標（以九點計）
					DrawMsrLabel(CurrentPoint, CircleRadius, &FiveNits[i][0]);//寫字
				}
				DrawMsrLabel(CurrentPoint, CircleRadius, &FiveNits[Few][0]);//寫字
			}
		}
    }
}


void CSwordDlg::Msr1P() 
{
    //重畫視窗    
    Invalidate();
    UpdateWindow();

	m_ICa.SetDisplayMode(0);
	 
    CurrentPoint = Set9Point();								//定位圈圈的位置   （預設叫中心點）

// 	CString str;
// 	str.Format("%d x %d\nr = %d\np = (%d, %d)\nm_floatFromEdge = %f", \
// 		ScrmH, ScrmV, CircleRadius, CurrentPoint.x, CurrentPoint.y, m_floatFromEdge);
// 	MessageBox(str);

    //MsrAI(AutoSelectMode(GetBackGroundColor()));
	if (MerOperation)
		MsrAI(AutoSelectMode(GetBackGroundColor()));		//偵測.動畫.畫圈
	else
		MsrAct(FALSE);

    Invalidate();
    UpdateWindow();

    Sleep(100);
    m_ICa.Measure(0);
    ColorLabel = SelectColorLabel(GetBackGroundColor());	//用顏色標籤選擇儲存的記憶體
//存值到記憶體
//  MsrCenter[   Color  ][value]
    MsrCenter[ColorLabel][0] = static_cast<double>(m_IProbe.GetLv());

    MsrCenter[ColorLabel][1] = static_cast<double>(m_IProbe.GetSx());
    MsrCenter[ColorLabel][2] = static_cast<double>(m_IProbe.GetSy());

    MsrCenter[ColorLabel][3] = static_cast<long>  (m_IProbe.GetT());
    MsrCenter[ColorLabel][4] = static_cast<double>(m_IProbe.GetDuv());

    MsrCenter[ColorLabel][5] = static_cast<double>(m_IProbe.GetUd());
    MsrCenter[ColorLabel][6] = static_cast<double>(m_IProbe.GetVd());

    MsrCenter[ColorLabel][7] = static_cast<double>(m_IProbe.GetX());
    MsrCenter[ColorLabel][8] = static_cast<double>(m_IProbe.GetY());
    MsrCenter[ColorLabel][9] = static_cast<double>(m_IProbe.GetZ());

    //MsrCenter[ColorLabel][10] = static_cast<double>(m_IProbe.GetFlckrFMA());

//顯示預覽數據
    DrawMsrLabel(CurrentPoint, CircleRadius, &MsrCenter[ColorLabel][0]);
    int i;
    for(i=0;i<10;++i)//value
    {                 //九點[白色][中心點][所有值] = 中心點[白色][所有值]     
        //if(m_bool_1cvr9)    
			MsrNineOValue[0][4][i] = MsrCenter[0][i];   //填入中心點量測值
                             // 49點[白色][點數][所有值] = 中心點[白色][所有值]
       // if(m_bool_1cvr49)   
			MsrFortyNineOValue[0][24][i] = MsrCenter[0][i];  //填入49點量測值
		MsrTwentyFiveValue[12][i] = MsrCenter[4][i];//填入中心點量測值
    }
}


void CSwordDlg::MsrFine5Nits() 
{
    //設定背景色初始值
    int i=60;
	int j;
    SetBackGroundColor(BK_Other,RGB(i,i,i));
    m_Brush.DeleteObject(); 
    m_Brush.CreateSolidBrush(GetBackGroundColor());

    //重畫視窗   
    Invalidate(TRUE);
    UpdateWindow();

    m_ICa.SetDisplayMode(0);//Lv,x,y

    //定位圈圈的位置
    CurrentPoint = Set5nits9Point();//預設叫中心點
    //DrawACircle(CurrentPoint, CircleRadius);//畫圈
    //自動調整5nits
    if(MsrAI(AutoSelectMode(GetBackGroundColor())))//如果量筒有貼上螢幕
    {
        double  fLv = 0;
        for(j=0;j<2;++j)
		{
			while(fLv<5.0)//若亮度還在5以下，就...變亮
			{
				//變動背景顏色
				SetBackGroundColor(BK_Other,RGB(i,i,i));
				m_Brush.DeleteObject(); 
				m_Brush.CreateSolidBrush(GetBackGroundColor());
				Invalidate();
				UpdateWindow();
				//量測抓值
				m_ICa.Measure(0);
				fLv = m_IProbe.GetLv();
				i+=2;
			}        
			while(fLv>5.0)//若亮度還沒有到5以下，就減少
			{
				//變動背景顏色
				SetBackGroundColor(BK_Other,RGB(i,i,i));
				m_Brush.DeleteObject(); 
				m_Brush.CreateSolidBrush(BackGroundColor);
				Invalidate();
				UpdateWindow();
				Sleep(60);
				//量測抓值
				m_ICa.Measure(0);
				fLv = m_IProbe.GetLv();
				--i;
			}
		}
    //顯示預覽數據
    DrawMsrLabel(CurrentPoint, CircleRadius, &fLv);

	//存值到記憶體
//      FiveNits[n][value]
    FiveNits[4][0] = static_cast<double>(m_IProbe.GetLv());

    FiveNits[4][1] = static_cast<double>(m_IProbe.GetSx());
    FiveNits[4][2] = static_cast<double>(m_IProbe.GetSy());

    FiveNits[4][3] = static_cast<long>  (m_IProbe.GetT());
    FiveNits[4][4] = static_cast<double>(m_IProbe.GetDuv());

    FiveNits[4][5] = static_cast<double>(m_IProbe.GetUd());
    FiveNits[4][6] = static_cast<double>(m_IProbe.GetVd());

    FiveNits[4][7] = static_cast<double>(m_IProbe.GetX());
    FiveNits[4][8] = static_cast<double>(m_IProbe.GetY());
    FiveNits[4][9] = static_cast<double>(m_IProbe.GetZ());
	}
	//5Nits特別流程結束
}

void CSwordDlg::Msr49P() 
{
    //重畫視窗    
    Invalidate();
    UpdateWindow();

    m_ICa.SetDisplayMode(0);//設定為x,y,Lv顯示模式

    int i;
    for(Few = 0;Few < 49;++Few)
    {
        //定位圈圈的位置
        CurrentPoint = Set49Point(Few);						//定位點，In第幾點，Out座標（以九點計）
        DrawColorLabel(SetCtlPointQuadrant(CurrentPoint),CircleRadius);//控制色塊

        Sleep(500);//量測動作停止或暫停（等使用者拿起量筒）
        //MsrAI(AutoSelectMode(GetBackGroundColor()),TRUE);//偵測//動畫.畫圈
		if (MerOperation)
			MsrAI(AutoSelectMode(GetBackGroundColor()),TRUE);		//偵測.動畫.畫圈
		else
			MsrAct(TRUE);

        Invalidate();
        UpdateWindow();
        
		if(PassStopColor())//STOP
        {
            break;
        }
        else if(PassBackColor())//BACK
        {
            //顯示預覽數據
            for(i=0;i<Few;++i)
            {
                CurrentPoint = Set49Point(i);//定位點，In第幾點，Out座標（以九點計）
                DrawMsrLabel(CurrentPoint, CircleRadius, &MsrFortyNineOValue[ColorLabel][i][0]);//寫字
            }
			Few -= 2;
			if(Few < -1)
				break;
			else
				continue;       
        }
        else//存值到記憶體
        {
            Sleep(100);
            m_ICa.Measure(0);
			ColorLabel = SelectColorLabel(GetBackGroundColor());//用顏色標籤選擇儲存的記憶體
            //MsrFortyNineOValue[   Color  ][ n ][value]
            MsrFortyNineOValue[ColorLabel][Few][0] = static_cast<double>(m_IProbe.GetLv());
            MsrFortyNineOValue[ColorLabel][Few][1] = static_cast<double>(m_IProbe.GetSx());
            MsrFortyNineOValue[ColorLabel][Few][2] = static_cast<double>(m_IProbe.GetSy());

            MsrFortyNineOValue[ColorLabel][Few][3] = static_cast<long>  (m_IProbe.GetT());
            MsrFortyNineOValue[ColorLabel][Few][4] = static_cast<double>(m_IProbe.GetDuv());

            MsrFortyNineOValue[ColorLabel][Few][5] = static_cast<double>(m_IProbe.GetUd());
            MsrFortyNineOValue[ColorLabel][Few][6] = static_cast<double>(m_IProbe.GetVd());

            MsrFortyNineOValue[ColorLabel][Few][7] = static_cast<double>(m_IProbe.GetX());
            MsrFortyNineOValue[ColorLabel][Few][8] = static_cast<double>(m_IProbe.GetY());
            MsrFortyNineOValue[ColorLabel][Few][9] = static_cast<double>(m_IProbe.GetZ()); 
            
            //顯示預覽數據
            for(i=0;i<=Few;++i)
            {
                CurrentPoint = Set49Point(i);//定位點，In第幾點，Out座標（以九點計）
                DrawMsrLabel(CurrentPoint, CircleRadius, &MsrFortyNineOValue[ColorLabel][i][0]);//寫字
            }
            DrawMsrLabel(CurrentPoint, CircleRadius, &MsrFortyNineOValue[ColorLabel][Few][0]);//寫字
        }
    }//for(Few = 0;Few<9;++Few)
    for(i=0;i<10;++i)//value
    {
                      //中心點[白色][所有值] = 49點[白色][中心點][所有值]
        if(m_bool_49cvr1)    MsrCenter[0][i] = MsrFortyNineOValue[0][24][i];//填入中心點量測值
        if(m_bool_49cvr9)//填入49點量測值
        {
       // 九點[白色][點數][所有值] = 49點[白色][點數][所有值]
            MsrNineOValue[0][0][i] = MsrFortyNineOValue[0][ 8][i];
            MsrNineOValue[0][1][i] = MsrFortyNineOValue[0][10][i];
            MsrNineOValue[0][2][i] = MsrFortyNineOValue[0][12][i];

            MsrNineOValue[0][3][i] = MsrFortyNineOValue[0][22][i];
            MsrNineOValue[0][4][i] = MsrFortyNineOValue[0][24][i];
            MsrNineOValue[0][5][i] = MsrFortyNineOValue[0][26][i];

            MsrNineOValue[0][6][i] = MsrFortyNineOValue[0][36][i];
            MsrNineOValue[0][7][i] = MsrFortyNineOValue[0][38][i];
            MsrNineOValue[0][8][i] = MsrFortyNineOValue[0][40][i];
        }
    }    
}

// void CSwordDlg::Msr13P()
// {
// 	//重畫視窗    
//     Invalidate();
//     UpdateWindow();
// 
//     m_ICa.SetDisplayMode(0);//設定為x,y,Lv顯示模式
//   
//     int i;
//     for(Few = 0;Few < 13;++Few)
//     {
//         //定位圈圈的位置
//         CurrentPoint = SetD13Point(Few);//定位點，In第幾點，Out座標（以九點計）
//         DrawColorLabel(SetCtlPointQuadrant(CurrentPoint),CircleRadius);//控制色塊
// 
//         Sleep(500);//量測動作停止或暫停（等使用者拿起量筒）
//         //MsrAI(AutoSelectMode(GetBackGroundColor()),TRUE);//偵測//動畫.畫圈
// 		if (MerOperation)
// 			MsrAI(AutoSelectMode(GetBackGroundColor()),TRUE);		//偵測.動畫.畫圈
// 		else
// 			MsrAct(TRUE);
// 
//         Invalidate();
//         UpdateWindow();
//         
//         if(PassStopColor())//STOP
//         {
//             break;
//         }//BACK
//         else if(PassBackColor())
//         {
//             //顯示預覽數據
//             for(i=0;i<Few;++i)
//             {
//                 CurrentPoint = SetD13Point(i);//定位點，In第幾點，Out座標（以九點計）
//                 DrawMsrLabel(CurrentPoint, CircleRadius, &MsrThirteenValue[i][0]);//寫字
//             }
// 			Few -= 2;
// 			if(Few < -1)
// 				break;
// 			else
// 				continue;       
//         }
//         else
//         {
//             Sleep(100);
//             m_ICa.Measure(0);
// 	        ColorLabel = SelectColorLabel(GetBackGroundColor());//用顏色標籤選擇儲存的記憶體
// 
//             //存值到記憶體
//             //MsrThirteenValue[ n][value]
//             MsrThirteenValue[Few][0] = static_cast<double>(m_IProbe.GetLv());
// 
//             MsrThirteenValue[Few][1] = static_cast<double>(m_IProbe.GetSx());
//             MsrThirteenValue[Few][2] = static_cast<double>(m_IProbe.GetSy());
// 
//             MsrThirteenValue[Few][3] = static_cast<long>  (m_IProbe.GetT());
//             MsrThirteenValue[Few][4] = static_cast<double>(m_IProbe.GetDuv());
// 
//             MsrThirteenValue[Few][5] = static_cast<double>(m_IProbe.GetUd());
//             MsrThirteenValue[Few][6] = static_cast<double>(m_IProbe.GetVd());
// 
//             MsrThirteenValue[Few][7] = static_cast<double>(m_IProbe.GetX());
//             MsrThirteenValue[Few][8] = static_cast<double>(m_IProbe.GetY());
//             MsrThirteenValue[Few][9] = static_cast<double>(m_IProbe.GetZ()); 
//             //MsrThirteenValue[Few][10] = static_cast<double>(m_IProbe.GetFlckrFMA());
// 
//             //顯示預覽數據
//             for(i=0;i<=Few;++i)
//             {
//                 CurrentPoint = SetD13Point(i);//定位點，In第幾點，Out座標（以九點計）
//                 DrawMsrLabel(CurrentPoint, CircleRadius, &MsrThirteenValue[i][0]);//寫字
//             }
//             DrawMsrLabel(CurrentPoint, CircleRadius, &MsrThirteenValue[Few][0]);//寫字
//         }
//     }//for(Few = 0;Few<9;++Few)
// }

void CSwordDlg::Msr25P()
{
	//重畫視窗    
    Invalidate();
    UpdateWindow();
    
    m_ICa.SetDisplayMode(0);//設定為x,y,Lv顯示模式

    int i;
    for(Few = 0;Few < 25;++Few)
    {
		//定位圈圈的位置
        CurrentPoint = SetD25Point(Few);//定位點，In第幾點，Out座標（以九點計）
        DrawColorLabel(SetCtlPointQuadrant(CurrentPoint),CircleRadius);//控制色塊

        Sleep(500);//量測動作停止或暫停（等使用者拿起量筒）
        //MsrAI(AutoSelectMode(GetBackGroundColor()),TRUE);//偵測//動畫.畫圈
		if (MerOperation)
			MsrAI(AutoSelectMode(GetBackGroundColor()),TRUE);		//偵測.動畫.畫圈
		else
			MsrAct(TRUE);

        Invalidate();
        UpdateWindow();
        
        if(PassStopColor())//STOP
        {
            break;
        }//BACK
        else if(PassBackColor())
        {
            //顯示預覽數據
            for(i=0;i<=Few;++i)
            {
                CurrentPoint = SetD25Point(i);//定位點，In第幾點，Out座標（以九點計）
                DrawMsrLabel(CurrentPoint, CircleRadius, &MsrTwentyFiveValue[i][0]);//寫字
            }
			Few -= 2;
			if(Few < -1)
				break;
			else
				continue; 
        }
        else
        {
            Sleep(100);
            m_ICa.Measure(0);   
			ColorLabel = SelectColorLabel(GetBackGroundColor());//用顏色標籤選擇儲存的記憶體
            //存值到記憶體
            //MsrThirteenValue[ n][value]
            MsrTwentyFiveValue[Few][0] = static_cast<double>(m_IProbe.GetLv());

            MsrTwentyFiveValue[Few][1] = static_cast<double>(m_IProbe.GetSx());
            MsrTwentyFiveValue[Few][2] = static_cast<double>(m_IProbe.GetSy());

            MsrTwentyFiveValue[Few][3] = static_cast<long>  (m_IProbe.GetT());
            MsrTwentyFiveValue[Few][4] = static_cast<double>(m_IProbe.GetDuv());

            MsrTwentyFiveValue[Few][5] = static_cast<double>(m_IProbe.GetUd());
            MsrTwentyFiveValue[Few][6] = static_cast<double>(m_IProbe.GetVd());

            MsrTwentyFiveValue[Few][7] = static_cast<double>(m_IProbe.GetX());
            MsrTwentyFiveValue[Few][8] = static_cast<double>(m_IProbe.GetY());
            MsrTwentyFiveValue[Few][9] = static_cast<double>(m_IProbe.GetZ()); 

            //顯示預覽數據
            for(i=0;i<=Few;++i)
            {
                CurrentPoint = SetD25Point(i);//定位點，In第幾點，Out座標（以九點計）
                DrawMsrLabel(CurrentPoint, CircleRadius, &MsrTwentyFiveValue[i][0]);//寫字
            }
            DrawMsrLabel(CurrentPoint, CircleRadius, &MsrTwentyFiveValue[Few][0]);//寫字
        }
    }//for(Few = 0;Few<9;++Few)
    for(i=0;i<10;++i)//value
    {
        //中心點[黑色][所有值] = 25點[白色][中心點][所有值]
        MsrCenter[4][i] = MsrTwentyFiveValue[12][i];//填入中心點量測值
        //填入25點量測值
        // 九點[黑色][點數][所有值] = 25[點數][所有值]
        MsrNineOValue[4][0][i] = MsrTwentyFiveValue[0][i];
        MsrNineOValue[4][1][i] = MsrTwentyFiveValue[5][i];
        MsrNineOValue[4][2][i] = MsrTwentyFiveValue[6][i];

        MsrNineOValue[4][3][i] = MsrTwentyFiveValue[11][i];
        MsrNineOValue[4][4][i] = MsrTwentyFiveValue[12][i];
        MsrNineOValue[4][5][i] = MsrTwentyFiveValue[13][i];

        MsrNineOValue[4][6][i] = MsrTwentyFiveValue[14][i];
        MsrNineOValue[4][7][i] = MsrTwentyFiveValue[19][i];
        MsrNineOValue[4][8][i] = MsrTwentyFiveValue[20][i];
    }
}

//25+13點一次量完
// void CSwordDlg::Msr29P()
// {
// 	//重畫視窗    
//     Invalidate();
//     UpdateWindow();
//     
//     m_ICa.SetDisplayMode(0);//設定為x,y,Lv顯示模式
// 
//     int i;
//     for(Few = 0;Few < 29;++Few)
//     {
// 		//定位圈圈的位置
//         CurrentPoint = SetD29Point(Few);//定位點，In第幾點，Out座標（以九點計）
//         DrawColorLabel(SetCtlPointQuadrant(CurrentPoint),CircleRadius);//控制色塊
// 
//         Sleep(500);//量測動作停止或暫停（等使用者拿起量筒）
//         //MsrAI(AutoSelectMode(GetBackGroundColor()),TRUE);//偵測//動畫.畫圈
// 		if (MerOperation)
// 			MsrAI(AutoSelectMode(GetBackGroundColor()),TRUE);		//偵測.動畫.畫圈
// 		else
// 			MsrAct(TRUE);
// 
//         Invalidate();
//         UpdateWindow();
//         
//         if(PassStopColor())//STOP
//         {
//             break;
//         }//BACK
//         else if(PassBackColor())
//         {
//             //顯示預覽數據
//             for(i=0;i<=Few;++i)
//             {
//                 CurrentPoint = SetD29Point(i);//定位點，In第幾點，Out座標（以九點計）
//                 DrawMsrLabel(CurrentPoint, CircleRadius, &MsrTwentyNineValue[i][0]);//寫字
//             }
// 			Few -= 2;
// 			if(Few < -1)
// 				break;
// 			else
// 				continue; 
//         }
//         else
//         {
//             Sleep(100);
//             m_ICa.Measure(0);   
// 			ColorLabel = SelectColorLabel(GetBackGroundColor());//用顏色標籤選擇儲存的記憶體
//             //存值到記憶體
//           //MsrTwentyNineValue[ n][value]
//             MsrTwentyNineValue[Few][0] = static_cast<double>(m_IProbe.GetLv());
// 
//             MsrTwentyNineValue[Few][1] = static_cast<double>(m_IProbe.GetSx());
//             MsrTwentyNineValue[Few][2] = static_cast<double>(m_IProbe.GetSy());
// 
//             MsrTwentyNineValue[Few][3] = static_cast<long>  (m_IProbe.GetT());
//             MsrTwentyNineValue[Few][4] = static_cast<double>(m_IProbe.GetDuv());
// 
//             MsrTwentyNineValue[Few][5] = static_cast<double>(m_IProbe.GetUd());
//             MsrTwentyNineValue[Few][6] = static_cast<double>(m_IProbe.GetVd());
// 
//             MsrTwentyNineValue[Few][7] = static_cast<double>(m_IProbe.GetX());
//             MsrTwentyNineValue[Few][8] = static_cast<double>(m_IProbe.GetY());
//             MsrTwentyNineValue[Few][9] = static_cast<double>(m_IProbe.GetZ()); 
// 
//             //顯示預覽數據
//             for(i=0;i<=Few;++i)
//             {
//                 CurrentPoint = SetD29Point(i);//定位點，In第幾點，Out座標（以九點計）
//                 DrawMsrLabel(CurrentPoint, CircleRadius, &MsrTwentyNineValue[i][0]);//寫字
//             }
//             DrawMsrLabel(CurrentPoint, CircleRadius, &MsrTwentyNineValue[Few][0]);//寫字
//         }
//     }//for(Few = 0;Few<9;++Few)
// 
// 
// 	for(i=0;i<10;++i)
// 	{
// 	//13點[ n][value] = 25點[ n][value] = 29點[ n][value]
// 		MsrCenter[4][i] = MsrTwentyNineValue[12][i];
// 	//邊界九點
// 		MsrThirteenValue[0][i] = MsrTwentyNineValue[0][i];
// 		MsrThirteenValue[1][i] = MsrTwentyNineValue[5][i];
// 		MsrThirteenValue[2][i] = MsrTwentyNineValue[6][i];
// 
// 		MsrThirteenValue[3][i] = MsrTwentyNineValue[11][i];
// 		MsrThirteenValue[4][i] = MsrTwentyNineValue[12][i];
// 		MsrThirteenValue[5][i] = MsrTwentyNineValue[13][i];
// 
// 		MsrThirteenValue[6][i] = MsrTwentyNineValue[14][i];
// 		MsrThirteenValue[7][i] = MsrTwentyNineValue[19][i];
// 		MsrThirteenValue[8][i] = MsrTwentyNineValue[20][i];
// 	//四點
// 		MsrThirteenValue[9][i] = MsrTwentyNineValue[25][i];
// 		MsrThirteenValue[10][i] = MsrTwentyNineValue[26][i];
// 		MsrThirteenValue[11][i] = MsrTwentyNineValue[27][i];
// 		MsrThirteenValue[12][i] = MsrTwentyNineValue[28][i];
// 
// 
// 	//25點[ n][value] = 29點[ n][value]
// 	//十字點
// 		for(Few=0;Few<25;++Few)
// 			MsrTwentyFiveValue[Few][i] = MsrTwentyNineValue[Few][i];
// 	}
// }

//New!29+13點一次量完
// void CSwordDlg::Msr33P()
// {
// 	//重畫視窗    
//     Invalidate();
//     UpdateWindow();
//     
//     m_ICa.SetDisplayMode(0);//設定為x,y,Lv顯示模式
// 
//     int i;
//     for(Few = 0;Few < 33;++Few)
//     {
// 		//定位圈圈的位置
//         CurrentPoint = SetD33Point(Few);//定位點，In第幾點，Out座標（以九點計）
//         DrawColorLabel(SetCtlPointQuadrant(CurrentPoint),CircleRadius);//控制色塊
// 
//         Sleep(500);//量測動作停止或暫停（等使用者拿起量筒）
//         //MsrAI(AutoSelectMode(GetBackGroundColor()),TRUE);//偵測//動畫.畫圈
// 		if (MerOperation)
// 			MsrAI(AutoSelectMode(GetBackGroundColor()),TRUE);		//偵測.動畫.畫圈
// 		else
// 			MsrAct(TRUE);
// 
//         Invalidate();
//         UpdateWindow();
//         
//         if(PassStopColor())//STOP
//         {
//             break;
//         }//BACK
//         else if(PassBackColor())
//         {
//             //顯示預覽數據
//             for(i=0;i<=Few;++i)
//             {
//                 CurrentPoint = SetD33Point(i);//定位點，In第幾點，Out座標（以九點計）
//                 DrawMsrLabel(CurrentPoint, CircleRadius, &MsrThirtyThreeValue[i][0]);//寫字
//             }
// 			Few -= 2;
// 			if(Few < -1)
// 				break;
// 			else
// 				continue; 
//         }
//         else
//         {
//             Sleep(100);
//             m_ICa.Measure(0);   
// 			ColorLabel = SelectColorLabel(GetBackGroundColor());//用顏色標籤選擇儲存的記憶體
//             //存值到記憶體
//           //MsrThirtyThreeValue[ n][value]
//             MsrThirtyThreeValue[Few][0] = static_cast<double>(m_IProbe.GetLv());
// 
//             MsrThirtyThreeValue[Few][1] = static_cast<double>(m_IProbe.GetSx());
//             MsrThirtyThreeValue[Few][2] = static_cast<double>(m_IProbe.GetSy());
// 
//             MsrThirtyThreeValue[Few][3] = static_cast<long>  (m_IProbe.GetT());
//             MsrThirtyThreeValue[Few][4] = static_cast<double>(m_IProbe.GetDuv());
// 
//             MsrThirtyThreeValue[Few][5] = static_cast<double>(m_IProbe.GetUd());
//             MsrThirtyThreeValue[Few][6] = static_cast<double>(m_IProbe.GetVd());
// 
//             MsrThirtyThreeValue[Few][7] = static_cast<double>(m_IProbe.GetX());
//             MsrThirtyThreeValue[Few][8] = static_cast<double>(m_IProbe.GetY());
//             MsrThirtyThreeValue[Few][9] = static_cast<double>(m_IProbe.GetZ()); 
// 
//             //顯示預覽數據
//             for(i=0;i<=Few;++i)
//             {
//                 CurrentPoint = SetD33Point(i);//定位點，In第幾點，Out座標（以九點計）
//                 DrawMsrLabel(CurrentPoint, CircleRadius, &MsrThirtyThreeValue[i][0]);//寫字
//             }
//             DrawMsrLabel(CurrentPoint, CircleRadius, &MsrThirtyThreeValue[Few][0]);//寫字
//         }
//     }//for(Few = 0;Few<9;++Few)
// 
// 
// 	for(i=0;i<10;++i)
// 	{
// 	//13點[ n][value] = 25點[ n][value] = 29點[ n][value]
// 		MsrCenter[4][i] = MsrThirtyThreeValue[14][i];
// 	//邊界九點
// 		MsrThirteenValue[0][i] = MsrThirtyThreeValue[0][i];
// 		MsrThirteenValue[1][i] = MsrThirtyThreeValue[6][i];
// 		MsrThirteenValue[2][i] = MsrThirtyThreeValue[7][i];
// 
// 		MsrThirteenValue[3][i] = MsrThirtyThreeValue[13][i];
// 		MsrThirteenValue[4][i] = MsrThirtyThreeValue[14][i];
// 		MsrThirteenValue[5][i] = MsrThirtyThreeValue[15][i];
// 
// 		MsrThirteenValue[6][i] = MsrThirtyThreeValue[16][i];
// 		MsrThirteenValue[7][i] = MsrThirtyThreeValue[22][i];
// 		MsrThirteenValue[8][i] = MsrThirtyThreeValue[23][i];
// 	//四點
// 		MsrThirteenValue[9][i] = MsrThirtyThreeValue[29][i];
// 		MsrThirteenValue[10][i] = MsrThirtyThreeValue[30][i];
// 		MsrThirteenValue[11][i] = MsrThirtyThreeValue[31][i];
// 		MsrThirteenValue[12][i] = MsrThirtyThreeValue[32][i];
// 
// 
// 	//25點[ n][value] = 29點[ n][value]
// 	//十字點
// //		for(Few=0;Few<25;++Few)
// //			MsrTwentyFiveValue[Few][i] = MsrTwentyNineValue[Few][i];
// 	}
// }

void CSwordDlg::MsrFLIC()
{

    Invalidate();
    UpdateWindow();

	m_ICa.SetDisplayMode(6);//設定為Flicker顯示模式（其實是量測模式）

    //定位圈圈的位置
    CurrentPoint = Set9Point();//預設叫中心點

    //MsrAI(AutoSelectMode(GetBackGroundColor()));//偵測//動畫.畫圈
	if (MerOperation)
		MsrAI(AutoSelectMode(GetBackGroundColor()));		//偵測.動畫.畫圈
	else
		MsrAct(FALSE);

    Invalidate();
    UpdateWindow();

    Sleep(100);
    m_ICa.Measure(0);
    
	//存值到記憶體
	ColorLabel = SelectColorLabel(GetBackGroundColor());//用顏色標籤選擇儲存的記憶體
    //  MsrFLK[value]
    MsrFLK = static_cast<double>(m_IProbe.GetFlckrFMA());

    //顯示預覽數據
    DrawMsrLabel(CurrentPoint, CircleRadius, &MsrFLK);
	m_ICa.SetDisplayMode(0);//設定為x,y,Lv顯示模式
}


//----------------
//實驗
// void CSwordDlg::MsrLab()
// {
//     Invalidate(TRUE);
//     UpdateWindow();
// 
//     int i,j;
//     for(i=0;i<125;++i){
//     for(j=0;j<3;++j){
//         Lab[i][j] = -1.1;
//     }}
//     
//     time_t start,now;
//     time (&start);
//     double dif = 0.0;
//     CurrentPoint = Set9Point();//預設叫中心點
// 
//     if(MsrAI(AutoSelectMode(GetBackGroundColor())))
//     {
//         Invalidate();
//         UpdateWindow();
//         i = 0;
//         int par=120;//實驗時間長度（秒）
//         while(dif<par)
//         {
//             int sec;
//             sec = (int)dif;
//             
//             if((sec)%1 == 0)//每隔幾秒抓一次點
//             {
//                 m_ICa.Measure(0);
//                 //Lab[i][0] = m_IProbe.GetFlckrFMA();
//             
// 				float SxFristValue,SyFristValue,LvFristValue,deltaSx,deltaSy,deltaLv;
//                       SxFristValue = SyFristValue = LvFristValue = deltaSx = deltaSy = deltaLv = 0.0;
//                 //宣告誤差值計算空間
//                            
//                 //抓參考值
//                 m_ICa.Measure(0);
// 
//                 LvFristValue = m_IProbe.GetLv();
//                 SxFristValue = m_IProbe.GetSx();
//                 SyFristValue = m_IProbe.GetSy();
//                 //誤差取絕對值    
//                    m_ICa.Measure(0);
// 
// 				   deltaLv = ((LvFristValue-m_IProbe.GetLv())>=0) ? (LvFristValue-m_IProbe.GetLv()) : (m_IProbe.GetLv()-LvFristValue);
// 				   deltaSx = ((SxFristValue-m_IProbe.GetSx())>=0) ? (SxFristValue-m_IProbe.GetSx()) : (m_IProbe.GetSx()-SxFristValue);
// 				   deltaSy = ((SyFristValue-m_IProbe.GetSy())>=0) ? (SyFristValue-m_IProbe.GetSy()) : (m_IProbe.GetSy()-SyFristValue);
// 
//                 //         Lab[time][value]
//                 Lab[i][0] = static_cast<double>(deltaLv);
//                 Lab[i][1] = static_cast<double>(deltaSx);
//                 Lab[i][2] = static_cast<double>(deltaSy);
//                 ++i;//抓到的第幾筆資料
//             }
//             time (&now);
//             dif = difftime (now, start);
// 
//             m_strTestValue.Format("%f",((dif/par)*100));
//             UpdateData(FALSE);
//         }
//     }
//     m_strTestValue.Format("%f",dif);
//     UpdateData(FALSE);
// }