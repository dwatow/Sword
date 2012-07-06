//-----------------------------------------------------------------------------------------------
//��i����
//    2nd_Source_LCM_optics data_modify Excel�ɮ榡
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
    objSheet.SetName("32''6K"); //�]�wsheet�W��

//�۰ʳ]�w
    /*���s�WSheet���y�k�a�I���M�]�u��s�W�T��*/
/*    
int Sheet;
short sBuf;
for(Sheet=0;Sheet<4;Sheet++)
{
    sBuf=(short)Sheet+1; 
    objSheet = objSheets.Add(3,4,5,1);
    objSheet = objSheets.GetItem(COleVariant(sBuf)); //�qsheet1�}�l

    ZeroMemory(buf,sizeof(buf));
    switch(Sheet)
    {
    case 0: strcpy(buf,"32''6K"); break;
    case 1: strcpy(buf,"37''6K"); break;
    case 2: strcpy(buf,"37''6.5K"); break;
    case 3: strcpy(buf,"40''5K"); break;
    case 4: strcpy(buf,"55''6K"); break;
    default: strcpy(buf,"�X���o�I"); break;
    }
    objSheet.SetName(buf); //�]�wsheet�W��

*/
//Step 4.�}�l�]�w���e

    //�]�wMODULE NO.
    range = objSheet.GetRange(COleVariant("A2"),COleVariant("A4"));//���A1:E1
    range.SetMergeCells(_variant_t(true));//�X���x�s��
    range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t("MODULE\nNO.")); //���e

    //�]�wLCM ID
    range = objSheet.GetRange(COleVariant("B2"),COleVariant("B4"));//���A2:E4
    range.SetMergeCells(_variant_t(true));//�X���x�s��
    range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t("LCM ID")); //���e
    
    //�]�wLED Light Bar ID
    range = objSheet.GetRange(COleVariant("C2"),COleVariant("E4"));//���A2:E4
    range.SetMergeCells(_variant_t(true));//�X���x�s��
    range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t("LED Light Bar ID")); //���e

    //�]�wmeasure point
    range = objSheet.GetRange(COleVariant("F2"),COleVariant("F4"));//���A2:E4
    range.SetMergeCells(_variant_t(true));//�X���x�s��
    range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t("measure point")); //���e

    //�]�wLight bar optics
    range = objSheet.GetRange(COleVariant("G2"),COleVariant("I2"));
    range.SetMergeCells(_variant_t(true));//�X���x�s��
    range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t("Light bar optics")); //���e
    
        range = objSheet.GetRange(COleVariant("G3"),COleVariant("G4"));
        range.SetMergeCells(_variant_t(true));//�X���x�s��
        range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t("Lv")); //���e

        range = objSheet.GetRange(COleVariant("H3"),COleVariant("H4"));
        range.SetMergeCells(_variant_t(true));//�X���x�s��
        range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t("x")); //���e

        range = objSheet.GetRange(COleVariant("I3"),COleVariant("I4"));
        range.SetMergeCells(_variant_t(true));//�X���x�s��
        range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t("y")); //���e

    //�]�wLumens / S-LED
    range = objSheet.GetRange(COleVariant("J2"),COleVariant("L2"));
    range.SetMergeCells(_variant_t(true));//�X���x�s��
    range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t("Lumens / S-LED")); //���e

        range = objSheet.GetRange(COleVariant("J3"),COleVariant("J4"));
        range.SetMergeCells(_variant_t(true));//�X���x�s��
        range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t("Avg. Lv")); //���e

        range = objSheet.GetRange(COleVariant("K3"),COleVariant("K4"));
        range.SetMergeCells(_variant_t(true));//�X���x�s��
        range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t("Avg. x")); //���e

        range = objSheet.GetRange(COleVariant("L3"),COleVariant("L4"));
        range.SetMergeCells(_variant_t(true));//�X���x�s��
        range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t("Avg. y")); //���e
    
    //�]�wCorrection factor
    range = objSheet.GetRange(COleVariant("M2"),COleVariant("O2"));
    range.SetMergeCells(_variant_t(true));//�X���x�s��
    range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t("Correction factor")); //���e

        range = objSheet.GetRange(COleVariant("M3"),COleVariant("M4"));
        range.SetMergeCells(_variant_t(true));//�X���x�s��
        range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t("Lv")); //���e

        range = objSheet.GetRange(COleVariant("N3"),COleVariant("N4"));
        range.SetMergeCells(_variant_t(true));//�X���x�s��
        range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t("x")); //���e

        range = objSheet.GetRange(COleVariant("O3"),COleVariant("O4"));
        range.SetMergeCells(_variant_t(true));//�X���x�s��
        range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t("y")); //���e

    //�]�wWHITE
    range = objSheet.GetRange(COleVariant("P2"),COleVariant("AB2"));
    range.SetMergeCells(_variant_t(true));//�X���x�s��
    range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t("WHITE")); //���e

        for(i=0;i<9;i++)
        {
            ZeroMemory(buf,sizeof(buf)); sprintf(buf,"%c3",'P'+i); 
            range = objSheet.GetRange(COleVariant(buf),COleVariant(buf));

                ZeroMemory(buf,sizeof(buf)); sprintf(buf,"%d",i+1);      
                range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf)); //���e
            
            ZeroMemory(buf,sizeof(buf)); sprintf(buf,"%c4",'P'+i);             
            range = objSheet.GetRange(COleVariant(buf),COleVariant(buf));

                ZeroMemory(buf,sizeof(buf));
                switch(i)
                {
                case 0: strcpy(buf,"���W"); break;
                case 1: strcpy(buf,"���W"); break;
                case 2: strcpy(buf,"�k�W"); break;
                case 3: strcpy(buf,"����"); break;
                case 4: strcpy(buf,"��"); break;
                case 5: strcpy(buf,"�k��"); break;
                case 6: strcpy(buf,"���U"); break;
                case 7: strcpy(buf,"���U"); break;
                case 8: strcpy(buf,"�k�U"); break;
                default: strcpy(buf,"�X���F�I"); break;
                }

                range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf)); //���e
        }
        range = objSheet.GetRange(COleVariant("Y3"),COleVariant("Y4"));
        range.SetMergeCells(_variant_t(true));//�X���x�s��
        range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t("Center x")); //���e

        range = objSheet.GetRange(COleVariant("Z3"),COleVariant("Z4"));
        range.SetMergeCells(_variant_t(true));//�X���x�s��
        range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t("Center y")); //���e

        range = objSheet.GetRange(COleVariant("AA3"),COleVariant("AA3"));
        range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t("Brightness")); //���e

        range = objSheet.GetRange(COleVariant("AA4"),COleVariant("AA4"));
        range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t("(AVG)")); //���e

        range = objSheet.GetRange(COleVariant("AB3"),COleVariant("AB4"));
        range.SetMergeCells(_variant_t(true));//�X���x�s��
        range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t("Uniformity")); //���e

    range = objSheet.GetRange(COleVariant("AC2"),COleVariant("AE3"));
    range.SetMergeCells(_variant_t(true));//�X���x�s��
    range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t("RED")); //���e

    range = objSheet.GetRange(COleVariant("AF2"),COleVariant("AH3"));
    range.SetMergeCells(_variant_t(true));//�X���x�s��
    range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t("GREEN")); //���e

    range = objSheet.GetRange(COleVariant("AI2"),COleVariant("AK3"));
    range.SetMergeCells(_variant_t(true));//�X���x�s��
    range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t("BLUE")); //���e

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
            default: strcpy(buf,"�X���F�I"); break;
        }
        range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf)); //���e
    }}

    range = objSheet.GetRange(COleVariant("A5"),COleVariant("A12"));
    range.SetMergeCells(_variant_t(true));//�X���x�s��
    range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t("1")); //���e
    range = objSheet.GetRange(COleVariant("B5"),COleVariant("B12"));
    range.SetMergeCells(_variant_t(true));//�X���x�s��
    
        range = objSheet.GetRange(COleVariant("C5"),COleVariant("C8"));
        range.SetMergeCells(_variant_t(true));//�X���x�s��
        range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t("1st")); //���e

        range = objSheet.GetRange(COleVariant("C9"),COleVariant("C12"));
        range.SetMergeCells(_variant_t(true));//�X���x�s��
        range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t("2nd")); //���e

    range = objSheet.GetRange(COleVariant("A13"),COleVariant("A20"));
    range.SetMergeCells(_variant_t(true));//�X���x�s��
    range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t("2")); //���e
    range = objSheet.GetRange(COleVariant("B13"),COleVariant("B20"));
    range.SetMergeCells(_variant_t(true));//�X���x�s��

        range = objSheet.GetRange(COleVariant("C13"),COleVariant("C16"));
        range.SetMergeCells(_variant_t(true));//�X���x�s��
        range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t("1st")); //���e

        range = objSheet.GetRange(COleVariant("C17"),COleVariant("C20"));
        range.SetMergeCells(_variant_t(true));//�X���x�s��
        range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t("2nd")); //���e

    range = objSheet.GetRange(COleVariant("A21"),COleVariant("A28"));
    range.SetMergeCells(_variant_t(true));//�X���x�s��
    range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t("3")); //���e
    range = objSheet.GetRange(COleVariant("B21"),COleVariant("B28"));
    range.SetMergeCells(_variant_t(true));//�X���x�s��

        range = objSheet.GetRange(COleVariant("C21"),COleVariant("C24"));
        range.SetMergeCells(_variant_t(true));//�X���x�s��
        range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t("1st")); //���e

        range = objSheet.GetRange(COleVariant("C25"),COleVariant("C28"));
        range.SetMergeCells(_variant_t(true));//�X���x�s��
        range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t("2nd")); //���e

    for(i=0;i<12;i++){
    for(j=0;j< 2;j++){
        ZeroMemory(buf,sizeof(buf)); 
        sprintf(buf,"F%d",j+5+(i*2)); 
        range = objSheet.GetRange(COleVariant(buf),COleVariant(buf));        
        
        ZeroMemory(buf,sizeof(buf));
        sprintf(buf,"%d",j+1); 
        range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf)); //���e
    
    }}
//-----------------------------------------------------------------------------------------------
//           ���r�񧹡I�U���O��J��ơI�зǳư}�C�I
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
    range.SetMergeCells(_variant_t(true));//�X���x�s��
    
    if(i<9)//�E�I�G�׭�
    {
        ZeroMemory(buf,sizeof(buf));
        sprintf(buf,"%3.2f",MsrNineOValue[3][i][2]);
        range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf)); //���e
    }
    else if(i==9)//�����Ix
    {
        ZeroMemory(buf,sizeof(buf));
        sprintf(buf,"%1.4f",MsrNineOValue[3][4][0]);
        range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf)); //���e
    }
    else if(i==10)//�����Iy
    {
        ZeroMemory(buf,sizeof(buf));
        sprintf(buf,"%1.4f",MsrNineOValue[3][4][1]);
        range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf)); //���e
    }
    else if(i==11)//�����G��
    {
        range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t("=AVERAGE(P30:X30)")); //���e
        range.SetNumberFormat(COleVariant("0.00"));
    }
    else if(i==12)//������
    {
        range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t("=MIN(P30:X30)/MAX(P30:X30)*100")); //���e
        range.SetNumberFormat(COleVariant("0"));
    }
}    

for(i = 0; i < 3; i++) //WRGB...�ܼ�
{
    for(j=0;j<3;j++)//Lv,x,y�ܼ�
    {
        ZeroMemory(buf1,sizeof(buf1));     sprintf(buf1,"A%c30",'C'+(i*3)+j); 
        ZeroMemory(buf2,sizeof(buf2));     sprintf(buf2,"A%c33",'C'+(i*3)+j); 
        range = objSheet.GetRange(COleVariant(buf1),COleVariant(buf2));
        range.SetMergeCells(_variant_t(true));//�X���x�s��

        if(j==0)//Lv
        {
            ZeroMemory(buf,sizeof(buf));
            sprintf(buf,"%3.2f",MsrCenter[i][j]);
            range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf)); //���e    
        }
        else//x,y
        {
            ZeroMemory(buf,sizeof(buf));
            sprintf(buf,"%1.4f",MsrCenter[i][j]);
            range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf)); //���e    
        }
    }
}

//-----------------------------------------------------------------------------------------------
//           �榡�����I�U�����C��M�ؽu
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