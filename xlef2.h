#ifndef __XLEF_H__
#define __XLEF_H__

//#include <comdef.h>
#include "excel.h"
#include "stdafx.h"

class xlsFile
{
    COleVariant VOptional;
    _Application objApp;
    Workbooks objBooks;
    _Workbook objBook; 
    Sheets objSheets;
    _Worksheet objSheet,objSheetT;
    Range range,col,row;//,oCell;//,range2,range3;
    //Interior cell;
    //Font font;
    COleException e;
	
    char buf[200];  //�Ȧs���r��
   // char buf1[200];
   // char buf2[200];
    //int i,j;  
 	 
public:
	xlsFile();
	~xlsFile();
	//xlsFile& �}�F�ɮפ���i�H�~����Sheet�M�R�W
	xlsFile& Open();
	xlsFile& Open(const char*);

	void SelectSheet(short, const char*);
	//xlsFile& ��F��l����i�H�~��U�uŪ�v�u�g�v���������
	xlsFile& SelectCell(char,int);	
	xlsFile& SelectCell(char,char,int);	

	void SetCell(int);
	void SetCell(double);
	void SetCell(long);	
	void SetCell(const char* );	
	void SetCell(CString );	

	void SetCell(const char*, int);
	void SetCell(const char*, double);
	void SetCell(const char*, long);
	void SetCell(const char*, const char* );
	void SetCell(const char*, CString );

	void SetVisible(bool);

	CString GetCell();
};

#endif