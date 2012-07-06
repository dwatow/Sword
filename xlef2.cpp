#include "stdafx.h"
#include "xlef.h"
#include <comdef.h>

xlsFile::xlsFile(): VOptional((long)DISP_E_PARAMNOTFOUND,VT_ERROR)
{
	ZeroMemory(buf,sizeof(buf));
//		ZeroMemory(buf1,sizeof(buf1));
//		ZeroMemory(buf2,sizeof(buf2));
	CString str;
	//Step 1.叫Excel應用程式
	if(!objApp.CreateDispatch("Excel.Application",&e))
	{
		str.Format("Excel CreateDispatch() failed w/err 0x%08lx", e.m_sc);
		AfxMessageBox(str, MB_SETFOREGROUND);
	}
};

xlsFile::~xlsFile()
{
	objApp.SetUserControl(TRUE);
	range.ReleaseDispatch();
	//chartobject.ReleaseDispatch();
	//chartobjects.ReleaseDispatch();
	objSheet.ReleaseDispatch();
	objSheets.ReleaseDispatch();
	objBook.ReleaseDispatch();
	objBooks.ReleaseDispatch();
	objApp.ReleaseDispatch();
}

//Open()
xlsFile& xlsFile::Open()
{
	objBooks = objApp.GetWorkbooks();
    objBook = objBooks.Add(VOptional);	//開新檔案
	objSheets = objBook.GetWorksheets();
	return *this;
}

xlsFile& xlsFile::Open(const char* path)
{
	objBooks = objApp.GetWorkbooks();
    objBook.AttachDispatch(objBooks.Add(_variant_t(path))); //開啟一個已存在的檔案
	objSheets = objBook.GetWorksheets();
	return *this;
}
//SelectSheet()
void xlsFile::SelectSheet(short SheetNumber,const char* SheetName)
{
	objSheet = objSheets.GetItem(COleVariant(SheetNumber)); //從sheet1開始
    objSheet.SetName(SheetName); //設定sheet名稱
}
//SelectCell()
xlsFile& xlsFile::SelectCell(char x, int y)
{
	ZeroMemory(buf,sizeof(buf));
	sprintf(buf,"%c%d",x,y);
	range=objSheet.GetRange(COleVariant(buf),COleVariant(buf));
	return *this;
}

xlsFile& xlsFile::SelectCell(char x1,char x2,int y)
{
	ZeroMemory(buf,sizeof(buf));
	sprintf(buf,"%c%c%d",x1,x2,y);
	range=objSheet.GetRange(COleVariant(buf),COleVariant(buf));
	return *this;
}
//SetCell()

void xlsFile::SetCell(int Data)
{
	ZeroMemory(buf,sizeof(buf));
	sprintf(buf,"%d",Data);
	range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf));
}

void xlsFile::SetCell(long Data)
{
	ZeroMemory(buf,sizeof(buf));
	sprintf(buf,"%d",Data);
	range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf));
}

void xlsFile::SetCell(double Data)
{
	ZeroMemory(buf,sizeof(buf));
	sprintf(buf,"%f",Data);
	range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf));
}

void xlsFile::SetCell(const char* Data)
{
	ZeroMemory(buf,sizeof(buf));
	sprintf(buf,"%s",Data);
	range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf));
}

void xlsFile::SetCell(CString Data)
{
	ZeroMemory(buf,sizeof(buf));
	sprintf(buf,"%s",Data);
	range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf));
}

void xlsFile::SetCell(const char* Format, int Data)
{
	ZeroMemory(buf,sizeof(buf));
	sprintf(buf,Format,Data);
	range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf));
}

void xlsFile::SetCell(const char* Format, double Data)
{
	ZeroMemory(buf,sizeof(buf));
	sprintf(buf,Format,Data);
	range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf));
}

void xlsFile::SetCell(const char* Format, long Data)
{
	ZeroMemory(buf,sizeof(buf));
	sprintf(buf,Format,Data);
	range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf));
}

void xlsFile::SetCell(const char* Format, const char* Data)
{
	ZeroMemory(buf,sizeof(buf));
	sprintf(buf,Format,Data);
	range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf));
}

void xlsFile::SetCell(const char* Format, CString Data)
{
	ZeroMemory(buf,sizeof(buf));
	sprintf(buf,Format,Data);
	range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf));
}
//SetVisible()
void xlsFile::SetVisible(bool a)
{
	objApp.SetVisible(a);    //顯示Excel檔
}
//GetCell()
CString xlsFile::GetCell()
{
	VARIANT cellvalue;
	cellvalue = range.GetItem(_variant_t((long)1), _variant_t((long)1)); 
	return (char*)_bstr_t(cellvalue);
}