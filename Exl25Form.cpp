//-----------------------------------------------------------------------------------------------
//  SEC�n�D�q�����t�A25 ���N 13�I
//        25�I�ϧΤ� form Excel�ɮ榡
//-----------------------------------------------------------------------------------------------
#include "stdafx.h"
#include <comdef.h>
#include "Sword.h"
#include "SwordDlg.h"
//#include "excel.h"

#define FILE_NAME OQCForm.xls
#define FILE_PATH OQCFormPath

#define SELBOX1(x,y);                ZeroMemory(buf,sizeof(buf));sprintf(buf,"%c%d",x,y);range=objSheet.GetRange(COleVariant(buf),COleVariant(buf));
#define SELBOX2(x1,x2,y);            ZeroMemory(buf,sizeof(buf));sprintf(buf,"%c%c%d",x1,x2,y);range=objSheet.GetRange(COleVariant(buf),COleVariant(buf));
#define SETBOXDATA(Format, Data);    ZeroMemory(buf,sizeof(buf));sprintf(buf,Format,Data);range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf));


void CSwordDlg::OnSave25p() 
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

    //Step 1.�sExcel���ε{��
    if(!objApp.CreateDispatch("Excel.Application",&e))
    {//����
        CString str;
        str.Format("Excel CreateDispatch() failed w/err 0x%08lx", e.m_sc),
        AfxMessageBox(str, MB_SETFOREGROUND);
        return;
    }
    
    //Step 2.����Excel�ɮ׮榡

    //�����|�X��
    char FILE_PATH[MAX_PATH];
    //dwCurDirPathLen = GetCurrentDirectory(MAX_PATH,szCurrentDirectory);//��ثe�Ҧb���ؿ��]���|�^
    strcpy(FILE_PATH, szCurrentDirectory);
    strcat(FILE_PATH, "\\FILE_NAME");

    objBooks = objApp.GetWorkbooks();
    //objBook = objBooks.Add(VOptional);
    objBook.AttachDispatch(objBooks.Add(_variant_t(FILE_PATH))); //�}�Ҥ@�Ӥw�s�b���ɮ�



    //Step 3.�]�wSheet1
    //��ʳ]�w
    objSheets = objBook.GetWorksheets();

    objSheet = objSheets.GetItem(COleVariant((short)1)); //�qsheet1�}�l
    objSheet.SetName("Report"); //�]�wsheet�W��


//Step 4.�}�l�]�w���e

//-----------------------------------------------------------------------------------------------
//           ���r�񧹡I�U���O��J��ơI�зǳư}�C�I
//-----------------------------------------------------------------------------------------------

//��J���
//         MsrTwentyFiveValue[n][value]..as MsrThirteenValue

    range = objSheet.GetRange(COleVariant("C7"),COleVariant("C7"));    ZeroMemory(buf,sizeof(buf));    sprintf(buf,"%3.2f",MsrTwentyFiveValue[12][0]);    range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf));
    range = objSheet.GetRange(COleVariant("D7"),COleVariant("D7"));    ZeroMemory(buf,sizeof(buf));    sprintf(buf,"%1.4f",MsrTwentyFiveValue[12][1]);    range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf));
    range = objSheet.GetRange(COleVariant("E7"),COleVariant("E7"));    ZeroMemory(buf,sizeof(buf));    sprintf(buf,"%1.4f",MsrTwentyFiveValue[12][2]);    range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf));
	int i;
	for(i=0;i<21;i++)
		SELBOX1('F'+i,7);  SETBOXDATA("%3.4f", MsrTwentyFiveValue[i][0]);
	for(i=0;i<4;i++)
		SELBOX2('A','A'+i,7);  SETBOXDATA("%3.4f", MsrTwentyFiveValue[20+i][0]);
/*    range = objSheet.GetRange(COleVariant("F7"),COleVariant("F7"));    ZeroMemory(buf,sizeof(buf));    sprintf(buf,"%3.4f",MsrTwentyFiveValue[0][0]);    range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf));
    range = objSheet.GetRange(COleVariant("G7"),COleVariant("G7"));    ZeroMemory(buf,sizeof(buf));    sprintf(buf,"%3.4f",MsrTwentyFiveValue[1][0]);    range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf));
    range = objSheet.GetRange(COleVariant("H7"),COleVariant("H7"));    ZeroMemory(buf,sizeof(buf));    sprintf(buf,"%3.4f",MsrTwentyFiveValue[2][0]);    range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf));
    range = objSheet.GetRange(COleVariant("I7"),COleVariant("I7"));    ZeroMemory(buf,sizeof(buf));    sprintf(buf,"%3.4f",MsrTwentyFiveValue[3][0]);    range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf));
    range = objSheet.GetRange(COleVariant("J7"),COleVariant("J7"));    ZeroMemory(buf,sizeof(buf));    sprintf(buf,"%3.4f",MsrTwentyFiveValue[4][0]);    range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf));
    range = objSheet.GetRange(COleVariant("K7"),COleVariant("K7"));    ZeroMemory(buf,sizeof(buf));    sprintf(buf,"%3.4f",MsrTwentyFiveValue[5][0]);    range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf));
    range = objSheet.GetRange(COleVariant("L7"),COleVariant("L7"));    ZeroMemory(buf,sizeof(buf));    sprintf(buf,"%3.4f",MsrTwentyFiveValue[6][0]);    range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf));
    range = objSheet.GetRange(COleVariant("M7"),COleVariant("M7"));    ZeroMemory(buf,sizeof(buf));    sprintf(buf,"%3.4f",MsrTwentyFiveValue[7][0]);    range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf));
    range = objSheet.GetRange(COleVariant("N7"),COleVariant("N7"));    ZeroMemory(buf,sizeof(buf));    sprintf(buf,"%3.4f",MsrTwentyFiveValue[8][0]);    range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf));
    range = objSheet.GetRange(COleVariant("O7"),COleVariant("O7"));    ZeroMemory(buf,sizeof(buf));    sprintf(buf,"%3.4f",MsrTwentyFiveValue[9][0]);    range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf));
    range = objSheet.GetRange(COleVariant("P7"),COleVariant("P7"));    ZeroMemory(buf,sizeof(buf));    sprintf(buf,"%3.4f",MsrTwentyFiveValue[10][0]);    range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf));
    
    range = objSheet.GetRange(COleVariant("Q7"),COleVariant("Q7"));    ZeroMemory(buf,sizeof(buf));    sprintf(buf,"%3.4f",MsrTwentyFiveValue[11][0]);    range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf));
    range = objSheet.GetRange(COleVariant("R7"),COleVariant("R7"));    ZeroMemory(buf,sizeof(buf));    sprintf(buf,"%3.4f",MsrTwentyFiveValue[12][0]);    range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf));
    range = objSheet.GetRange(COleVariant("S7"),COleVariant("S7"));    ZeroMemory(buf,sizeof(buf));    sprintf(buf,"%3.4f",MsrTwentyFiveValue[12][0]);    range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf));
    range = objSheet.GetRange(COleVariant("T7"),COleVariant("T7"));    ZeroMemory(buf,sizeof(buf));    sprintf(buf,"%3.4f",MsrTwentyFiveValue[13][0]);    range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf));
    range = objSheet.GetRange(COleVariant("U7"),COleVariant("U7"));    ZeroMemory(buf,sizeof(buf));    sprintf(buf,"%3.4f",MsrTwentyFiveValue[14][0]);    range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf));
    range = objSheet.GetRange(COleVariant("V7"),COleVariant("V7"));    ZeroMemory(buf,sizeof(buf));    sprintf(buf,"%3.4f",MsrTwentyFiveValue[15][0]);    range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf));
    range = objSheet.GetRange(COleVariant("W7"),COleVariant("W7"));    ZeroMemory(buf,sizeof(buf));    sprintf(buf,"%3.4f",MsrTwentyFiveValue[16][0]);    range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf));
    range = objSheet.GetRange(COleVariant("X7"),COleVariant("X7"));    ZeroMemory(buf,sizeof(buf));    sprintf(buf,"%3.4f",MsrTwentyFiveValue[17][0]);    range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf));
    range = objSheet.GetRange(COleVariant("Y7"),COleVariant("Y7"));    ZeroMemory(buf,sizeof(buf));    sprintf(buf,"%3.4f",MsrTwentyFiveValue[18][0]);    range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf));
    range = objSheet.GetRange(COleVariant("Z7"),COleVariant("Z7"));    ZeroMemory(buf,sizeof(buf));    sprintf(buf,"%3.4f",MsrTwentyFiveValue[19][0]);    range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf));

    range = objSheet.GetRange(COleVariant("AA7"),COleVariant("AA7"));    ZeroMemory(buf,sizeof(buf));    sprintf(buf,"%3.4f",MsrTwentyFiveValue[20][0]);    range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf));
    range = objSheet.GetRange(COleVariant("AB7"),COleVariant("AB7"));    ZeroMemory(buf,sizeof(buf));    sprintf(buf,"%3.4f",MsrTwentyFiveValue[21][0]);    range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf));
    range = objSheet.GetRange(COleVariant("AC7"),COleVariant("AC7"));    ZeroMemory(buf,sizeof(buf));    sprintf(buf,"%3.4f",MsrTwentyFiveValue[22][0]);    range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf));
    range = objSheet.GetRange(COleVariant("AD7"),COleVariant("AD7"));    ZeroMemory(buf,sizeof(buf));    sprintf(buf,"%3.4f",MsrTwentyFiveValue[23][0]);    range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf));
*/
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


