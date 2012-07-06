// Machine generated IDispatch wrapper class(es) created by Microsoft Visual C++

// NOTE: Do not modify the contents of this file.  If this class is regenerated by
//  Microsoft Visual C++, your modifications will be overwritten.


#include "stdafx.h"
#include "_cacontrol.h"

/////////////////////////////////////////////////////////////////////////////
// C_CaControl

IMPLEMENT_DYNCREATE(C_CaControl, CWnd)

/////////////////////////////////////////////////////////////////////////////
// C_CaControl properties

/////////////////////////////////////////////////////////////////////////////
// C_CaControl operations

void C_CaControl::SetRefCa(LPDISPATCH* newValue)
{
	static BYTE parms[] =
		VTS_PDISPATCH;
	InvokeHelper(0x68030006, DISPATCH_PROPERTYPUTREF, VT_EMPTY, NULL, parms,
		 newValue);
}

LPDISPATCH C_CaControl::GetCa()
{
	LPDISPATCH result;
	InvokeHelper(0x68030006, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

void C_CaControl::SetRefProbe(LPDISPATCH* newValue)
{
	static BYTE parms[] =
		VTS_PDISPATCH;
	InvokeHelper(0x68030005, DISPATCH_PROPERTYPUTREF, VT_EMPTY, NULL, parms,
		 newValue);
}

LPDISPATCH C_CaControl::GetProbe()
{
	LPDISPATCH result;
	InvokeHelper(0x68030005, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

void C_CaControl::SetRefMemory(LPDISPATCH* newValue)
{
	static BYTE parms[] =
		VTS_PDISPATCH;
	InvokeHelper(0x68030004, DISPATCH_PROPERTYPUTREF, VT_EMPTY, NULL, parms,
		 newValue);
}

LPDISPATCH C_CaControl::GetMemory()
{
	LPDISPATCH result;
	InvokeHelper(0x68030004, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

void C_CaControl::SetStdColor(long* newValue)
{
	static BYTE parms[] =
		VTS_PI4;
	InvokeHelper(0x68030003, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 newValue);
}

long C_CaControl::GetStdColor()
{
	long result;
	InvokeHelper(0x68030003, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

void C_CaControl::SetVGType(BSTR* newValue)
{
	static BYTE parms[] =
		VTS_PBSTR;
	InvokeHelper(0x68030002, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 newValue);
}

CString C_CaControl::GetVGType()
{
	CString result;
	InvokeHelper(0x68030002, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

void C_CaControl::SetVGTiming(BSTR* newValue)
{
	static BYTE parms[] =
		VTS_PBSTR;
	InvokeHelper(0x68030001, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 newValue);
}

CString C_CaControl::GetVGTiming()
{
	CString result;
	InvokeHelper(0x68030001, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

void C_CaControl::SetVGPattern(BSTR* newValue)
{
	static BYTE parms[] =
		VTS_PBSTR;
	InvokeHelper(0x68030000, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 newValue);
}

CString C_CaControl::GetVGPattern()
{
	CString result;
	InvokeHelper(0x68030000, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

void C_CaControl::Load(BSTR* strDataName)
{
	static BYTE parms[] =
		VTS_PBSTR;
	InvokeHelper(0x6003000d, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 strDataName);
}

void C_CaControl::Save(BSTR* strDataName)
{
	static BYTE parms[] =
		VTS_PBSTR;
	InvokeHelper(0x6003000e, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 strDataName);
}

void C_CaControl::UpdateCaInfo()
{
	InvokeHelper(0x6003000f, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void C_CaControl::UpdateMemoryInfo()
{
	InvokeHelper(0x60030010, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}