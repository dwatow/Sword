//-----------------------------------------------------------------------------------------------
//  Michael�n�D��
//         �q26�o�� form Excel�ɮ榡
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
//Step 1.�sExcel���ε{��
	xlsFile t26Form;
    
//Step 2.����Excel�ɮ׮榡
    //�����|�X��
    char t26FormPath[MAX_PATH];
    //dwCurDirPathLen = GetCurrentDirectory(MAX_PATH,szCurrentDirectory);//��ثe�Ҧb���ؿ��]���|�^
    strcpy(t26FormPath, szCurrentDirectory);
    strcat(t26FormPath, "\\t26FORM.xls");

//Step 3.�]�wSheet1
	t26Form.Open(t26FormPath).SetSheetName(1,"Report");

//Step 4.�}�l�]�w���e

//-----------------------------------------------------------------------------------------------
//           ���r�񧹡I�U���O��J��ơI�зǳư}�C�I
//-----------------------------------------------------------------------------------------------

//��J���
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
//           �榡�����I�U�����C��M�ؽu
//-----------------------------------------------------------------------------------------------
	t26Form.SetVisible(TRUE);
//     objApp.SetUserControl(TRUE);

        SetBackGroundColor(); //�w�]�I���C��
        SetMaxWindow(0);
        ShowWinNormal();//�Y�p�n��ܪ��F��
        HideWinNormal();//�Y�p�n�������F�� 
        this->SetDlgItemInt(IDC_MAX,1);        
}