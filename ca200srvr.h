// Machine generated IDispatch wrapper class(es) created with ClassWizard
/////////////////////////////////////////////////////////////////////////////
// ICa200 wrapper class
#ifndef CA200SRVR_H
#define CA200SRVR_H

class ICa200 : public COleDispatchDriver
{
public:
	ICa200() {}		// Calls COleDispatchDriver default constructor
	ICa200(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	ICa200(const ICa200& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attributes
public:

// Operations
public:
	LPDISPATCH GetCas();
	void SetConfiguration(long CaNumberVal, LPCTSTR ConnecStringVal, long PortVal, long BaudRateVal);
	void AutoConnect();
	LPDISPATCH GetSingleCa();
};
/////////////////////////////////////////////////////////////////////////////
// ICas wrapper class

class ICas : public COleDispatchDriver
{
public:
	ICas() {}		// Calls COleDispatchDriver default constructor
	ICas(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	ICas(const ICas& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attributes
public:

// Operations
public:
	LPDISPATCH GetItem(const VARIANT& IndexVal);
	long GetCount();
	void SendMsr();
	void ReceiveMsr();
	LPDISPATCH GetItemOfNumber(long CaNumberVal);
	void SetCaID(long CaNumberVal, LPCTSTR CaIDVal);
};
/////////////////////////////////////////////////////////////////////////////
// ICa wrapper class

class ICa : public COleDispatchDriver
{
public:
	ICa() {}		// Calls COleDispatchDriver default constructor
	ICa(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	ICa(const ICa& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attributes
public:

// Operations
public:
	LPDISPATCH GetProbes();
	LPDISPATCH GetOutputProbes();
	LPDISPATCH GetMemory();
	CString GetDisplayProbe();
	void SetDisplayProbe(LPCTSTR lpszNewValue);
	LPDISPATCH GetSingleProbe();
	float GetSyncMode();
	void SetSyncMode(float newValue);
	long GetDisplayMode();
	void SetDisplayMode(long nNewValue);
	long GetDisplayDigits();
	void SetDisplayDigits(long nNewValue);
	long GetAveragingMode();
	void SetAveragingMode(long nNewValue);
	long GetBrightnessUnit();
	void SetBrightnessUnit(long nNewValue);
	CString GetCAType();
	CString GetCAVersion();
	long GetNumber();
	CString GetPortID();
	CString GetId();
	void SetId(LPCTSTR lpszNewValue);
	void CalZero();
	void Measure(long newVal);
	void SetAnalyzerCalMode();
	void Enter();
	void SetAnalogRange(float Range1Val, float Range2Val);
	void SetPWROnStatus();
	long GetCalStandard();
	void SetCalStandard(long nNewValue);
	void ResetAnalyzerCalMode();
	void SetLvxyCalMode();
	void ResetLvxyCalMode();
	void SetAnalyzerCalData(long ColorVal);
	void SetDisplayProbe(long ProbeNumberVal);
	void SetLvxyCalData(long ColorVal, float XVal, float YVal, float LvVal);
	void SetRemoteMode(long nNewValue);
	void SetFMAAnalogRange(float RangeVal);
	void GetAnalogRange(float* Range1Val, float* Range2Val);
	void GetFMAAnalogRange(float* RangeVal);
};
/////////////////////////////////////////////////////////////////////////////
// IProbes wrapper class

class IProbes : public COleDispatchDriver
{
public:
	IProbes() {}		// Calls COleDispatchDriver default constructor
	IProbes(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	IProbes(const IProbes& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attributes
public:

// Operations
public:
	LPDISPATCH GetItem(const VARIANT& IndexVal);
	long GetCount();
	LPDISPATCH GetItemOfNumber(long ProbeNumber);
	void SetProbeID(long ProbeNumberVal, LPCTSTR ProbeIDVal);
};
/////////////////////////////////////////////////////////////////////////////
// IProbe wrapper class

class IProbe : public COleDispatchDriver
{
public:
	IProbe() {}		// Calls COleDispatchDriver default constructor
	IProbe(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	IProbe(const IProbe& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attributes
public:

// Operations
public:
	float GetX();
	float GetY();
	float GetZ();
	float GetFlckrJEITA();
	long GetNumber();
	CString GetId();
	void SetId(LPCTSTR lpszNewValue);
	CString GetSerialNO();
	float GetLv();
	float GetUd();
	float GetVd();
	long GetT();
	float GetDEUser();
	float GetFlckrFMA();
	float GetSy();
	float GetSx();
	float GetDuv();
	float GetUsUser();
	float GetVsUser();
	float GetLsUser();
	float GetLvfL();
	float GetR();
	float GetG();
	float GetB();
	long GetRd();
	long GetRad();
	long GetRfma();
	long GetRjeita();
	float GetSpectrum(long Frequency);
};
/////////////////////////////////////////////////////////////////////////////
// IOutputProbes wrapper class

class IOutputProbes : public COleDispatchDriver
{
public:
	IOutputProbes() {}		// Calls COleDispatchDriver default constructor
	IOutputProbes(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	IOutputProbes(const IOutputProbes& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attributes
public:

// Operations
public:
	LPDISPATCH GetItem(const VARIANT& ProbeIndexVal);
	long GetCount();
	void Add(LPCTSTR ProbeIDVal);
	LPDISPATCH GetItemOfNumber(long ProbeNumberVal);
	void AddAll();
	LPDISPATCH Clone();
	void RemoveAll();
};
/////////////////////////////////////////////////////////////////////////////
// IMemory wrapper class

class IMemory : public COleDispatchDriver
{
public:
	IMemory() {}		// Calls COleDispatchDriver default constructor
	IMemory(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	IMemory(const IMemory& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attributes
public:

// Operations
public:
	long GetChannelNO();
	void SetChannelNO(long nNewValue);
	CString GetChannelID();
	void SetID(LPCTSTR lpszNewValue);
	void GetReferenceColor(LPCTSTR ProbeIDVal, float* XVal, float* YVal, float* LvVal);
	void SetChannelID(LPCTSTR ChannelIDVal);
	void GetMemoryStatus(long ProbeNOVal, long* CalProbeSNOVal, long* RefProbeSNOVal, long* CalModeVal);
	long CheckCalData(long ProbeNOVal, LPCTSTR FileNameVal);
	void CopyToFile(long ProbeNOVal, LPCTSTR FileNameVal);
	void CopyFromFile(long ProbeNOVal, LPCTSTR FileNameVal);
};
/////////////////////////////////////////////////////////////////////////////
// _ICaEvents wrapper class

class _ICaEvents : public COleDispatchDriver
{
public:
	_ICaEvents() {}		// Calls COleDispatchDriver default constructor
	_ICaEvents(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	_ICaEvents(const _ICaEvents& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attributes
public:

// Operations
public:
	void ExeCalZero();
};
/////////////////////////////////////////////////////////////////////////////
// IProbeInfo wrapper class

class IProbeInfo : public COleDispatchDriver
{
public:
	IProbeInfo() {}		// Calls COleDispatchDriver default constructor
	IProbeInfo(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	IProbeInfo(const IProbeInfo& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attributes
public:

// Operations
public:
	CString GetTypeName();
	long GetTypeNO();
};
/////////////////////////////////////////////////////////////////////////////
// ICaOption wrapper class

class ICaOption : public COleDispatchDriver
{
public:
	ICaOption() {}		// Calls COleDispatchDriver default constructor
	ICaOption(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	ICaOption(const ICaOption& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attributes
public:

// Operations
public:
	CString Command(LPCTSTR CommandVal);
};

#endif