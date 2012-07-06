#include "stdafx.h"
//#include <comdef.h>
#include "Sword.h"
#include "SwordDlg.h"
//#include "excel.h"

/*
void CSwordDlg::Msr9P()	            �E�I�q��
void CSwordDlg::Msr5Nits9P()		5nits�q��
void CSwordDlg::Msr1P()             �����I�q��
void CSwordDlg::Msr5P()             5�I�q��
void CSwordDlg::MsrFine5Nits()      ��5nits���q��
void CSwordDlg::Msr49P()			49�I�q��
void CSwordDlg::Msr13P()			13�I�q��
void CSwordDlg::Msr25P()			25�I�q��
void CSwordDlg::Msr29P()			29�I�q��
void CSwordDlg::MsrFLIC()			FLK�q��
void CSwordDlg::MsrLab()			����q����
*/

void CSwordDlg::Msr9P()
{
    //���e����    
    Invalidate();
    UpdateWindow();

	m_ICa.SetDisplayMode(0);//�]�w��x,y,Lv��ܼҦ�

    int i;
    for(Few = 0;Few < 9;++Few)
    {
		CurrentPoint = (GetBackGroundColor() == RGB(0,0,0)) ? SetD9Point(Few) : Set9Point(Few,m_floatFromEdge);//�w���骺��m
        DrawColorLabel(SetCtlPointQuadrant(CurrentPoint),CircleRadius);						//������

        Sleep(500);//�q���ʧ@����μȰ��]���ϥΪ̮��_�q���^
        if (MerOperation)
			MsrAI(AutoSelectMode(GetBackGroundColor()),TRUE);		//����.�ʵe.�e��
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
            //��ܹw���ƾ�
			//CString str;	str.Format("Few=%d",Few);	MessageBox(str);
			for(i=0;i<Few;++i)
            {
				CurrentPoint = (GetBackGroundColor() == RGB(0,0,0)) ? SetD9Point(i) : Set9Point(i,m_floatFromEdge);
                DrawMsrLabel(CurrentPoint, CircleRadius, &MsrNineOValue[ColorLabel][i][0]);//�g�r
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
			ColorLabel = SelectColorLabel(GetBackGroundColor());//���C����ҿ���x�s���O����

            //�s�Ȩ�O����
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

            //��ܹw���ƾ�
            for(i=0;i<=Few;++i)
            {
				CurrentPoint = (GetBackGroundColor() == RGB(0,0,0)) ? SetD9Point(i) : Set9Point(i,m_floatFromEdge);
				DrawMsrLabel(CurrentPoint, CircleRadius, &MsrNineOValue[ColorLabel][i][0]);//�g�r
            }
            DrawMsrLabel(CurrentPoint, CircleRadius, &MsrNineOValue[ColorLabel][Few][0]);//�g�r

        }
    }//for(Few = 0;Few<9;++Few)
    for(i=0;i<10;++i)//value
    {
                     //�����I[�զ�][�Ҧ���] = �E�I[�զ�][�����I][�Ҧ���] 
		if(m_bool_9cvr1)    MsrCenter[0][i] = MsrNineOValue[0][4][i];//��J�����I�q����
        if(m_bool_9cvr49)//��J49�I�q����
        {
             // 49�I[�զ�][�I��][�Ҧ���] = �E�I[�զ�][�I��][�Ҧ���]
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
    //���e����    
    Invalidate();
    UpdateWindow();
	
	m_ICa.SetDisplayMode(0);//�]�w��x,y,Lv��ܼҦ�
	
    int i;
    for(Few = 0;Few < 5;++Few)
    {
		CurrentPoint = Set5Point(Few, m_float5FromEdge);//�w���骺��m
        DrawColorLabel(SetCtlPointQuadrant(CurrentPoint),CircleRadius);						//������
		
        Sleep(500);//�q���ʧ@����μȰ��]���ϥΪ̮��_�q���^
        if (MerOperation)
			MsrAI(AutoSelectMode(GetBackGroundColor()),TRUE);		//����.�ʵe.�e��
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
            //��ܹw���ƾ�
			//CString str;	str.Format("Few=%d",Few);	MessageBox(str);
			for(i=0;i<Few;++i)
            {
				CurrentPoint = Set5Point(i);
                DrawMsrLabel(CurrentPoint, CircleRadius, &MsrFiveOValue[ColorLabel][i][0]);//�g�r
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
			ColorLabel = SelectColorLabel(GetBackGroundColor());//���C����ҿ���x�s���O����
			
            //�s�Ȩ�O����
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


            //��ܹw���ƾ�
            for(i=0;i<=Few;++i)
            {
				CurrentPoint = Set5Point(i);
				DrawMsrLabel(CurrentPoint, CircleRadius, &MsrFiveOValue[ColorLabel][i][0]);//�g�r
            }
            DrawMsrLabel(CurrentPoint, CircleRadius, &MsrFiveOValue[ColorLabel][Few][0]);//�g�r
			
        }
    }//for(Few = 0;Few<9;++Few)

	//��J�����I�q����

	//�����I[�զ�][�Ҧ���]       = �E�I[�զ�][�����I][�Ҧ���]
	MsrCenter[0][i]              = MsrFiveOValue[0][2][i];
	// 9�I[�զ�][�I��][�Ҧ���]   = �E�I[�զ�][�I��][�Ҧ���]
    MsrNineOValue[0][5][i]       = MsrFiveOValue[0][2][i];
	// 49�I[�զ�][�I��][�Ҧ���]  = �E�I[�զ�][�I��][�Ҧ���]
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

			CurrentPoint = Set5nits9Point(Few,m_floatFromEdge);					//�w���I�AIn�ĴX�I�AOut�y�С]�H�E�I�p�^
			DrawColorLabel(SetCtlPointQuadrant(CurrentPoint),CircleRadius);	//������

			Sleep(500);//�q���ʧ@����μȰ��]���ϥΪ̮��_�q���^
			//MsrAI(AutoSelectMode(GetBackGroundColor()),TRUE);				//����.�ʵe.�e��
			if (MerOperation)
				MsrAI(AutoSelectMode(GetBackGroundColor()),TRUE);		//����.�ʵe.�e��
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
				//��ܹw���ƾ�
				for(i=0;i<Few;++i)
				{
					CurrentPoint = Set5nits9Point(i,6);//�w���I�AIn�ĴX�I�AOut�y�С]�H�E�I�p�^
					DrawMsrLabel(CurrentPoint, CircleRadius, &FiveNits[i][0]);//�g�r
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
		//�s�Ȩ�O����
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
		//��ܹw���ƾ�
				for(i=0;i<=Few;++i)
				{
					CurrentPoint = Set5nits9Point(i,6);//�w���I�AIn�ĴX�I�AOut�y�С]�H�E�I�p�^
					DrawMsrLabel(CurrentPoint, CircleRadius, &FiveNits[i][0]);//�g�r
				}
				DrawMsrLabel(CurrentPoint, CircleRadius, &FiveNits[Few][0]);//�g�r
			}
		}
    }
}


void CSwordDlg::Msr1P() 
{
    //���e����    
    Invalidate();
    UpdateWindow();

	m_ICa.SetDisplayMode(0);
	 
    CurrentPoint = Set9Point();								//�w���骺��m   �]�w�]�s�����I�^

// 	CString str;
// 	str.Format("%d x %d\nr = %d\np = (%d, %d)\nm_floatFromEdge = %f", \
// 		ScrmH, ScrmV, CircleRadius, CurrentPoint.x, CurrentPoint.y, m_floatFromEdge);
// 	MessageBox(str);

    //MsrAI(AutoSelectMode(GetBackGroundColor()));
	if (MerOperation)
		MsrAI(AutoSelectMode(GetBackGroundColor()));		//����.�ʵe.�e��
	else
		MsrAct(FALSE);

    Invalidate();
    UpdateWindow();

    Sleep(100);
    m_ICa.Measure(0);
    ColorLabel = SelectColorLabel(GetBackGroundColor());	//���C����ҿ���x�s���O����
//�s�Ȩ�O����
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

//��ܹw���ƾ�
    DrawMsrLabel(CurrentPoint, CircleRadius, &MsrCenter[ColorLabel][0]);
    int i;
    for(i=0;i<10;++i)//value
    {                 //�E�I[�զ�][�����I][�Ҧ���] = �����I[�զ�][�Ҧ���]     
        //if(m_bool_1cvr9)    
			MsrNineOValue[0][4][i] = MsrCenter[0][i];   //��J�����I�q����
                             // 49�I[�զ�][�I��][�Ҧ���] = �����I[�զ�][�Ҧ���]
       // if(m_bool_1cvr49)   
			MsrFortyNineOValue[0][24][i] = MsrCenter[0][i];  //��J49�I�q����
		MsrTwentyFiveValue[12][i] = MsrCenter[4][i];//��J�����I�q����
    }
}


void CSwordDlg::MsrFine5Nits() 
{
    //�]�w�I�����l��
    int i=60;
	int j;
    SetBackGroundColor(BK_Other,RGB(i,i,i));
    m_Brush.DeleteObject(); 
    m_Brush.CreateSolidBrush(GetBackGroundColor());

    //���e����   
    Invalidate(TRUE);
    UpdateWindow();

    m_ICa.SetDisplayMode(0);//Lv,x,y

    //�w���骺��m
    CurrentPoint = Set5nits9Point();//�w�]�s�����I
    //DrawACircle(CurrentPoint, CircleRadius);//�e��
    //�۰ʽվ�5nits
    if(MsrAI(AutoSelectMode(GetBackGroundColor())))//�p�G�q�����K�W�ù�
    {
        double  fLv = 0;
        for(j=0;j<2;++j)
		{
			while(fLv<5.0)//�Y�G���٦b5�H�U�A�N...�ܫG
			{
				//�ܰʭI���C��
				SetBackGroundColor(BK_Other,RGB(i,i,i));
				m_Brush.DeleteObject(); 
				m_Brush.CreateSolidBrush(GetBackGroundColor());
				Invalidate();
				UpdateWindow();
				//�q�����
				m_ICa.Measure(0);
				fLv = m_IProbe.GetLv();
				i+=2;
			}        
			while(fLv>5.0)//�Y�G���٨S����5�H�U�A�N���
			{
				//�ܰʭI���C��
				SetBackGroundColor(BK_Other,RGB(i,i,i));
				m_Brush.DeleteObject(); 
				m_Brush.CreateSolidBrush(BackGroundColor);
				Invalidate();
				UpdateWindow();
				Sleep(60);
				//�q�����
				m_ICa.Measure(0);
				fLv = m_IProbe.GetLv();
				--i;
			}
		}
    //��ܹw���ƾ�
    DrawMsrLabel(CurrentPoint, CircleRadius, &fLv);

	//�s�Ȩ�O����
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
	//5Nits�S�O�y�{����
}

void CSwordDlg::Msr49P() 
{
    //���e����    
    Invalidate();
    UpdateWindow();

    m_ICa.SetDisplayMode(0);//�]�w��x,y,Lv��ܼҦ�

    int i;
    for(Few = 0;Few < 49;++Few)
    {
        //�w���骺��m
        CurrentPoint = Set49Point(Few);						//�w���I�AIn�ĴX�I�AOut�y�С]�H�E�I�p�^
        DrawColorLabel(SetCtlPointQuadrant(CurrentPoint),CircleRadius);//������

        Sleep(500);//�q���ʧ@����μȰ��]���ϥΪ̮��_�q���^
        //MsrAI(AutoSelectMode(GetBackGroundColor()),TRUE);//����//�ʵe.�e��
		if (MerOperation)
			MsrAI(AutoSelectMode(GetBackGroundColor()),TRUE);		//����.�ʵe.�e��
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
            //��ܹw���ƾ�
            for(i=0;i<Few;++i)
            {
                CurrentPoint = Set49Point(i);//�w���I�AIn�ĴX�I�AOut�y�С]�H�E�I�p�^
                DrawMsrLabel(CurrentPoint, CircleRadius, &MsrFortyNineOValue[ColorLabel][i][0]);//�g�r
            }
			Few -= 2;
			if(Few < -1)
				break;
			else
				continue;       
        }
        else//�s�Ȩ�O����
        {
            Sleep(100);
            m_ICa.Measure(0);
			ColorLabel = SelectColorLabel(GetBackGroundColor());//���C����ҿ���x�s���O����
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
            
            //��ܹw���ƾ�
            for(i=0;i<=Few;++i)
            {
                CurrentPoint = Set49Point(i);//�w���I�AIn�ĴX�I�AOut�y�С]�H�E�I�p�^
                DrawMsrLabel(CurrentPoint, CircleRadius, &MsrFortyNineOValue[ColorLabel][i][0]);//�g�r
            }
            DrawMsrLabel(CurrentPoint, CircleRadius, &MsrFortyNineOValue[ColorLabel][Few][0]);//�g�r
        }
    }//for(Few = 0;Few<9;++Few)
    for(i=0;i<10;++i)//value
    {
                      //�����I[�զ�][�Ҧ���] = 49�I[�զ�][�����I][�Ҧ���]
        if(m_bool_49cvr1)    MsrCenter[0][i] = MsrFortyNineOValue[0][24][i];//��J�����I�q����
        if(m_bool_49cvr9)//��J49�I�q����
        {
       // �E�I[�զ�][�I��][�Ҧ���] = 49�I[�զ�][�I��][�Ҧ���]
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
// 	//���e����    
//     Invalidate();
//     UpdateWindow();
// 
//     m_ICa.SetDisplayMode(0);//�]�w��x,y,Lv��ܼҦ�
//   
//     int i;
//     for(Few = 0;Few < 13;++Few)
//     {
//         //�w���骺��m
//         CurrentPoint = SetD13Point(Few);//�w���I�AIn�ĴX�I�AOut�y�С]�H�E�I�p�^
//         DrawColorLabel(SetCtlPointQuadrant(CurrentPoint),CircleRadius);//������
// 
//         Sleep(500);//�q���ʧ@����μȰ��]���ϥΪ̮��_�q���^
//         //MsrAI(AutoSelectMode(GetBackGroundColor()),TRUE);//����//�ʵe.�e��
// 		if (MerOperation)
// 			MsrAI(AutoSelectMode(GetBackGroundColor()),TRUE);		//����.�ʵe.�e��
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
//             //��ܹw���ƾ�
//             for(i=0;i<Few;++i)
//             {
//                 CurrentPoint = SetD13Point(i);//�w���I�AIn�ĴX�I�AOut�y�С]�H�E�I�p�^
//                 DrawMsrLabel(CurrentPoint, CircleRadius, &MsrThirteenValue[i][0]);//�g�r
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
// 	        ColorLabel = SelectColorLabel(GetBackGroundColor());//���C����ҿ���x�s���O����
// 
//             //�s�Ȩ�O����
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
//             //��ܹw���ƾ�
//             for(i=0;i<=Few;++i)
//             {
//                 CurrentPoint = SetD13Point(i);//�w���I�AIn�ĴX�I�AOut�y�С]�H�E�I�p�^
//                 DrawMsrLabel(CurrentPoint, CircleRadius, &MsrThirteenValue[i][0]);//�g�r
//             }
//             DrawMsrLabel(CurrentPoint, CircleRadius, &MsrThirteenValue[Few][0]);//�g�r
//         }
//     }//for(Few = 0;Few<9;++Few)
// }

void CSwordDlg::Msr25P()
{
	//���e����    
    Invalidate();
    UpdateWindow();
    
    m_ICa.SetDisplayMode(0);//�]�w��x,y,Lv��ܼҦ�

    int i;
    for(Few = 0;Few < 25;++Few)
    {
		//�w���骺��m
        CurrentPoint = SetD25Point(Few);//�w���I�AIn�ĴX�I�AOut�y�С]�H�E�I�p�^
        DrawColorLabel(SetCtlPointQuadrant(CurrentPoint),CircleRadius);//������

        Sleep(500);//�q���ʧ@����μȰ��]���ϥΪ̮��_�q���^
        //MsrAI(AutoSelectMode(GetBackGroundColor()),TRUE);//����//�ʵe.�e��
		if (MerOperation)
			MsrAI(AutoSelectMode(GetBackGroundColor()),TRUE);		//����.�ʵe.�e��
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
            //��ܹw���ƾ�
            for(i=0;i<=Few;++i)
            {
                CurrentPoint = SetD25Point(i);//�w���I�AIn�ĴX�I�AOut�y�С]�H�E�I�p�^
                DrawMsrLabel(CurrentPoint, CircleRadius, &MsrTwentyFiveValue[i][0]);//�g�r
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
			ColorLabel = SelectColorLabel(GetBackGroundColor());//���C����ҿ���x�s���O����
            //�s�Ȩ�O����
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

            //��ܹw���ƾ�
            for(i=0;i<=Few;++i)
            {
                CurrentPoint = SetD25Point(i);//�w���I�AIn�ĴX�I�AOut�y�С]�H�E�I�p�^
                DrawMsrLabel(CurrentPoint, CircleRadius, &MsrTwentyFiveValue[i][0]);//�g�r
            }
            DrawMsrLabel(CurrentPoint, CircleRadius, &MsrTwentyFiveValue[Few][0]);//�g�r
        }
    }//for(Few = 0;Few<9;++Few)
    for(i=0;i<10;++i)//value
    {
        //�����I[�¦�][�Ҧ���] = 25�I[�զ�][�����I][�Ҧ���]
        MsrCenter[4][i] = MsrTwentyFiveValue[12][i];//��J�����I�q����
        //��J25�I�q����
        // �E�I[�¦�][�I��][�Ҧ���] = 25[�I��][�Ҧ���]
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

//25+13�I�@���q��
// void CSwordDlg::Msr29P()
// {
// 	//���e����    
//     Invalidate();
//     UpdateWindow();
//     
//     m_ICa.SetDisplayMode(0);//�]�w��x,y,Lv��ܼҦ�
// 
//     int i;
//     for(Few = 0;Few < 29;++Few)
//     {
// 		//�w���骺��m
//         CurrentPoint = SetD29Point(Few);//�w���I�AIn�ĴX�I�AOut�y�С]�H�E�I�p�^
//         DrawColorLabel(SetCtlPointQuadrant(CurrentPoint),CircleRadius);//������
// 
//         Sleep(500);//�q���ʧ@����μȰ��]���ϥΪ̮��_�q���^
//         //MsrAI(AutoSelectMode(GetBackGroundColor()),TRUE);//����//�ʵe.�e��
// 		if (MerOperation)
// 			MsrAI(AutoSelectMode(GetBackGroundColor()),TRUE);		//����.�ʵe.�e��
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
//             //��ܹw���ƾ�
//             for(i=0;i<=Few;++i)
//             {
//                 CurrentPoint = SetD29Point(i);//�w���I�AIn�ĴX�I�AOut�y�С]�H�E�I�p�^
//                 DrawMsrLabel(CurrentPoint, CircleRadius, &MsrTwentyNineValue[i][0]);//�g�r
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
// 			ColorLabel = SelectColorLabel(GetBackGroundColor());//���C����ҿ���x�s���O����
//             //�s�Ȩ�O����
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
//             //��ܹw���ƾ�
//             for(i=0;i<=Few;++i)
//             {
//                 CurrentPoint = SetD29Point(i);//�w���I�AIn�ĴX�I�AOut�y�С]�H�E�I�p�^
//                 DrawMsrLabel(CurrentPoint, CircleRadius, &MsrTwentyNineValue[i][0]);//�g�r
//             }
//             DrawMsrLabel(CurrentPoint, CircleRadius, &MsrTwentyNineValue[Few][0]);//�g�r
//         }
//     }//for(Few = 0;Few<9;++Few)
// 
// 
// 	for(i=0;i<10;++i)
// 	{
// 	//13�I[ n][value] = 25�I[ n][value] = 29�I[ n][value]
// 		MsrCenter[4][i] = MsrTwentyNineValue[12][i];
// 	//��ɤE�I
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
// 	//�|�I
// 		MsrThirteenValue[9][i] = MsrTwentyNineValue[25][i];
// 		MsrThirteenValue[10][i] = MsrTwentyNineValue[26][i];
// 		MsrThirteenValue[11][i] = MsrTwentyNineValue[27][i];
// 		MsrThirteenValue[12][i] = MsrTwentyNineValue[28][i];
// 
// 
// 	//25�I[ n][value] = 29�I[ n][value]
// 	//�Q�r�I
// 		for(Few=0;Few<25;++Few)
// 			MsrTwentyFiveValue[Few][i] = MsrTwentyNineValue[Few][i];
// 	}
// }

//New!29+13�I�@���q��
// void CSwordDlg::Msr33P()
// {
// 	//���e����    
//     Invalidate();
//     UpdateWindow();
//     
//     m_ICa.SetDisplayMode(0);//�]�w��x,y,Lv��ܼҦ�
// 
//     int i;
//     for(Few = 0;Few < 33;++Few)
//     {
// 		//�w���骺��m
//         CurrentPoint = SetD33Point(Few);//�w���I�AIn�ĴX�I�AOut�y�С]�H�E�I�p�^
//         DrawColorLabel(SetCtlPointQuadrant(CurrentPoint),CircleRadius);//������
// 
//         Sleep(500);//�q���ʧ@����μȰ��]���ϥΪ̮��_�q���^
//         //MsrAI(AutoSelectMode(GetBackGroundColor()),TRUE);//����//�ʵe.�e��
// 		if (MerOperation)
// 			MsrAI(AutoSelectMode(GetBackGroundColor()),TRUE);		//����.�ʵe.�e��
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
//             //��ܹw���ƾ�
//             for(i=0;i<=Few;++i)
//             {
//                 CurrentPoint = SetD33Point(i);//�w���I�AIn�ĴX�I�AOut�y�С]�H�E�I�p�^
//                 DrawMsrLabel(CurrentPoint, CircleRadius, &MsrThirtyThreeValue[i][0]);//�g�r
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
// 			ColorLabel = SelectColorLabel(GetBackGroundColor());//���C����ҿ���x�s���O����
//             //�s�Ȩ�O����
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
//             //��ܹw���ƾ�
//             for(i=0;i<=Few;++i)
//             {
//                 CurrentPoint = SetD33Point(i);//�w���I�AIn�ĴX�I�AOut�y�С]�H�E�I�p�^
//                 DrawMsrLabel(CurrentPoint, CircleRadius, &MsrThirtyThreeValue[i][0]);//�g�r
//             }
//             DrawMsrLabel(CurrentPoint, CircleRadius, &MsrThirtyThreeValue[Few][0]);//�g�r
//         }
//     }//for(Few = 0;Few<9;++Few)
// 
// 
// 	for(i=0;i<10;++i)
// 	{
// 	//13�I[ n][value] = 25�I[ n][value] = 29�I[ n][value]
// 		MsrCenter[4][i] = MsrThirtyThreeValue[14][i];
// 	//��ɤE�I
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
// 	//�|�I
// 		MsrThirteenValue[9][i] = MsrThirtyThreeValue[29][i];
// 		MsrThirteenValue[10][i] = MsrThirtyThreeValue[30][i];
// 		MsrThirteenValue[11][i] = MsrThirtyThreeValue[31][i];
// 		MsrThirteenValue[12][i] = MsrThirtyThreeValue[32][i];
// 
// 
// 	//25�I[ n][value] = 29�I[ n][value]
// 	//�Q�r�I
// //		for(Few=0;Few<25;++Few)
// //			MsrTwentyFiveValue[Few][i] = MsrTwentyNineValue[Few][i];
// 	}
// }

void CSwordDlg::MsrFLIC()
{

    Invalidate();
    UpdateWindow();

	m_ICa.SetDisplayMode(6);//�]�w��Flicker��ܼҦ��]���O�q���Ҧ��^

    //�w���骺��m
    CurrentPoint = Set9Point();//�w�]�s�����I

    //MsrAI(AutoSelectMode(GetBackGroundColor()));//����//�ʵe.�e��
	if (MerOperation)
		MsrAI(AutoSelectMode(GetBackGroundColor()));		//����.�ʵe.�e��
	else
		MsrAct(FALSE);

    Invalidate();
    UpdateWindow();

    Sleep(100);
    m_ICa.Measure(0);
    
	//�s�Ȩ�O����
	ColorLabel = SelectColorLabel(GetBackGroundColor());//���C����ҿ���x�s���O����
    //  MsrFLK[value]
    MsrFLK = static_cast<double>(m_IProbe.GetFlckrFMA());

    //��ܹw���ƾ�
    DrawMsrLabel(CurrentPoint, CircleRadius, &MsrFLK);
	m_ICa.SetDisplayMode(0);//�]�w��x,y,Lv��ܼҦ�
}


//----------------
//����
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
//     CurrentPoint = Set9Point();//�w�]�s�����I
// 
//     if(MsrAI(AutoSelectMode(GetBackGroundColor())))
//     {
//         Invalidate();
//         UpdateWindow();
//         i = 0;
//         int par=120;//����ɶ����ס]��^
//         while(dif<par)
//         {
//             int sec;
//             sec = (int)dif;
//             
//             if((sec)%1 == 0)//�C�j�X���@���I
//             {
//                 m_ICa.Measure(0);
//                 //Lab[i][0] = m_IProbe.GetFlckrFMA();
//             
// 				float SxFristValue,SyFristValue,LvFristValue,deltaSx,deltaSy,deltaLv;
//                       SxFristValue = SyFristValue = LvFristValue = deltaSx = deltaSy = deltaLv = 0.0;
//                 //�ŧi�~�t�ȭp��Ŷ�
//                            
//                 //��Ѧҭ�
//                 m_ICa.Measure(0);
// 
//                 LvFristValue = m_IProbe.GetLv();
//                 SxFristValue = m_IProbe.GetSx();
//                 SyFristValue = m_IProbe.GetSy();
//                 //�~�t�������    
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
//                 ++i;//��쪺�ĴX�����
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