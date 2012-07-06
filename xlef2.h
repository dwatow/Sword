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
	
    char buf[200];  //暫存的字串
   // char buf1[200];
   // char buf2[200];
    //int i,j;  
 	 
public:
	xlsFile();
	~xlsFile();
	//xlsFile& 開了檔案之後可以繼續選擇Sheet和命名
	xlsFile& Open();
	xlsFile& Open(const char*);

	void SelectSheet(short, const char*);
	//xlsFile& 選了格子之後可以繼續下「讀」「寫」的成員函數
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