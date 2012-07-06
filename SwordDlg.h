// SwordDlg.h : header file
//
//{{AFX_INCLUDES()
#include "_xycontrol.h"
#include "_cacontrol.h"
#include "ca200srvr.h"    // Added by ClassView
//}}AFX_INCLUDES
#include "InitialDlg.h"
#include<cmath>
//#include <vector>

#if !defined(AFX_SWORDDLG_H__5B6D51A7_96D4_4809_A56B_01AE04E41FF1__INCLUDED_)
#define AFX_SWORDDLG_H__5B6D51A7_96D4_4809_A56B_01AE04E41FF1__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
/////////////////////////////////////////////////////////////////////////////
//BackGround Color in MsrColor and Menory Address
#define BK_White  0
#define BK_Red    1
#define BK_Green  2
#define BK_Blue   3
#define BK_Dark   4
#define BK_Other  5

//MsrPoint using in GetMsrPoint()
#define PO_CurrentPoint  0
#define PO_StopPoint     1
#define PO_BackPoint     2

//Color Value
#define CV_White  0x00FFFFFF
#define CV_Red    0x000000FF
#define CV_Green  0X0000FF00
#define CV_Blue   0x00FF0000
#define CV_Dark   0x00000000
#define CV_Flk    0x017F7F7F

//RGB(125,125,0);//RGB(R,G,B)
//0x000000FF     //0x00,B,G,R
/////////////////////////////////////////////////////////////////////////////
class CSwordDlgAutoProxy;

/////////////////////////////////////////////////////////////////////////////
// CSwordDlg dialog

class CSwordDlg : public CDialog
{
    DECLARE_DYNAMIC(CSwordDlg);
    friend class CSwordDlgAutoProxy;

// Construction
public:
	CPoint SetCtlPointQuadrant(CPoint point);

    CSwordDlg(CWnd* pParent = NULL);    // standard constructor
    virtual ~CSwordDlg();

//���󱱨�
//��l�ƪ�����
    CInitialDlg* pInitialDlg;
	//pDC->Rectangle(pRect);
//�������A    
    CButton *pBtmMax;
    CButton *pBtmOK;
    CButton *pBtmCANCEL;//pDC->SetBkMode(TRANSPARENT);
//���C��
    CButton *pBtmRED;
    CButton *pBtmGREEN;
    CButton *pBtmBLUE;
    CButton *pBtmWhite;
    CButton *pBtmBlack;
    CButton *pBtmUserColor;
//�����I�q�����
	CButton *p1Cvr9CheckBox;
	CButton *p1Cvr49CheckBox;
    CButton *pBtmCentrO;
//�E�I�q��    
	CButton *p9Cvr1CheckBox;
	CButton *p9Cvr49CheckBox;
    CButton *pBtmNineO;
    CStatic *pFE;
    CEdit   *pFromEdge;
//5�I�q��
	CStatic *p5FE;
    CEdit   *p5FromEdge;
//5Nits�q��
	CEdit   *pEditBackUalue;
    CButton *pBtmFiveNits;
    CSpinButtonCtrl* pSpin;
//Gamma
    CButton *pBtmGamma;
//49�I�q��
	CButton *p49Cvr1CheckBox;
	CButton *p49Cvr9CheckBox;
 	CButton *pBtmFNO;
//13�I�q��
//	CButton *pBtmThirteenO;
//25�I�q��
	CButton *pBtmTwentyFiveO;
//����
    CButton *pBtmAgingLab;
//���i
    CButton *pBtmSave1;
    CButton *pBtmSave2;
    CButton *pBtmAgingLabSave;
//	CButton *pBtmSave13p;
	CButton *pBtmSave25p;
	CButton *pBtmSaveT26;
	CButton *pBtmSaveOQC;
//�䥦
	CButton *pBtmRstDate;
    CStatic *pTextRredMe;
    CStatic *pMeasureFlowList;
	CStatic *pRePortList;
//	CButton *pBtmZeroCal;
//�q���w����
	CStatic *pDisplayValuesList;
//�w���ȤĿﶵ��
	CStatic *pMsrViewList;
    CButton *pLvCheckBox;
	CButton *pSXCheckBox;
	CButton *pSYCheckBox;

	CButton *pTCheckBox;
	CButton *pDUVCheckBox;
	
	CButton *pDUCheckBox;
	CButton *pDVCheckBox;
	
	CButton *pXCheckBox;
	CButton *pYCheckBox;
	CButton *pZCheckBox;
	CButton *pFlkCheckBox;
//Flicker Pattern
    CButton *pBtmFlkMsr;
	CStatic *pFlickerPatternList;
	CButton *pBtmFlkSubPix;
	CButton *pBtmFlk2LInv;
	CButton *pBtmFlk2L1Inv;
	CButton *pBtmFlkVS;
	CButton *pBtmFlkCol;
//��ĴX�x
	//CComboBox *pCmbSelTai;
	//CComboBox *pCmbSelFlkPtn;
//OQC
	CStatic   *pOQCTitle;
	CComboBox *pOQC5Nits;
	CComboBox *pOQCWCenter;
	CComboBox *pOQCRCenter;
	CComboBox *pOQCGCenter;
	CComboBox *pOQCBCenter;
	CComboBox *pOQCDCenter;
	CComboBox *pOQCW9P;
	CComboBox *pOQCW5P;
	CComboBox *pOQCR5P;
	CComboBox *pOQCG5P;
	CComboBox *pOQCB5P;
	CComboBox *pOQCD5P;
	CComboBox *pOQCW49P;
	CComboBox *pOQCD25P;
//	CComboBox *pOQCD33P;
	//CStatic   *pOQCFlickerStatic;
	//CComboBox *pOQCFlickerSel;     //pCmbSelFlkPtn
	CComboBox *pOQCTaiSel;         //pCmbSelTai
	CButton   *pOQCStartMsr;
	CComboBox *pOQCFormSel;
	CStatic   *psOQCW9P;
	CStatic   *psOQCW5P;
	CStatic   *psOQCR5P;
	CStatic   *psOQCG5P;
	CStatic   *psOQCB5P;
	CStatic   *psOQCD5P;
	CStatic   *psOQCW49P;
	CStatic   *psOQCD25P;
	CStatic   *psOQCRGB;
	CButton   *pOQCLoadBarCode;

	CStatic *pMsrOperation;
	CButton *pActEnter;
	CButton *pAutoMsr;
//	CButton *pBtmAutoMsrCal;

	BOOL MerOperation;

//�ۭq���
    void ShowColorController();  //��ܭI���ⱱ�
    void HideColorController();  //���íI���ⱱ�
    
    void ShowMsrItem();  //��ܶq������
    void HideMsrItem();  //���öq������

    void ShowMsrView();  //��ܹw���ȤĿﶵ��
    void HideMsrView();  //���ùw���ȤĿﶵ��

    void ShowWinControlItem();  //��ܶq������
    void HideWinControlItem();  //���öq������

    void ShowWinMaximum();
    void HideWinMaximum();

	void ShowOQCItem();  //���OQC����
	void HideOQCItem();  //����OQC����

    void ShowWinNormal();
    void HideWinNormal();

    void ShowReportBtm();
    void HideReportBtm();

	void SetOQCSaveSel();

	void EnabledReportBtm(BOOL bEnable = TRUE);

   // void SetSysBackGroundColor();

	void     SetBackGroundColor(int ColorFlag = 6, COLORREF BKDefaultColor = GetSysColor(COLOR_BTNFACE));
	//void     SetBackGroundColor(int ColorFlag = 6, COLORREF BKDefaultColor = RGB(236,233,216));
	const COLORREF GetBackGroundColor();

	BOOL   SetMsrPoint(CPoint MsrPoint, BOOL Mode = FALSE);
    const CPoint GetMsrPoint(int PointFlag = 0);

	const COLORREF GetStopColor();
	const COLORREF GetBackColor();

	BOOL PassStopColor();
	BOOL PassBackColor();

	void ChooseCircleRadius();
	void initValues();
	const CString FromSelector();

	//void OutOrigFile();
    //BOOL InitxyControl();
//�����
//         Lab[time][value]
//    double Lab[125][3];
//OQC�q���y�{��

	
	struct W9	{	double Lv;	double x;	double y;	long T;		double Duv;};
	struct W5	{	double Lv;	double x;	double y;	long T;		double Duv;};
	struct R5	{	double Lv;	double x;	double y;	long T;		double Duv;};
	struct G5	{	double Lv;	double x;	double y;	long T;		double Duv;};
	struct B5	{	double Lv;	double x;	double y;	long T;		double Duv;};
	struct D5	{	double Lv;	double x;	double y;	long T;		double Duv;};
	struct W49	{	double Lv;	double x;	double y;	};
	struct W1	{	double Lv;	double x;	double y;	double T;	double Duv;};
	struct D1	{	double Lv;	double x;	double y;	double T;	double Duv;};
	struct R1	{	double Lv;	double x;	double y;	double T;	double Duv;};
	struct G1	{	double Lv;	double x;	double y;	double T;	double Duv;};
	struct B1	{	double Lv;	double x;	double y;	double T;	double Duv;};
	struct D5nits1	{	double Lv;	double x;	double y;	double T;	double Duv;};
	
	struct D5nits9	{	double Lv;	double x;	double y;};
//	struct D13	{	double Lv;};
	struct D25	{	double Lv;};
//	struct D29	{	double Lv;};
	struct FLK1	{	double Flckr;};

	struct MsrDate
	{
		W5 w5[5];
		R5 r5[5];
		G5 g5[5];
		B5 b5[5];
		D5 d5[5];
		W9 w9[9];                    //W[0..8][Lv,x,y]
		W49 w49[49];                 //W[0..8][Lv,x,y]
		W1 w;
		D1 d;
		R1 r;                       //    R[Lv,x,y,T,duv]
		G1 g;                       //    G[Lv,x,y,T,duv]
		B1 b;                       //    B[Lv,x,y,T,duv]
		D5nits1 d5nits1;                   //5Nits[Lv,x,y,T,duv]
		D5nits9 d5nits9[9];               //D5N9[0..8]Lv
		W9 d9[9];
		D25 d25[25];                  //D25[0..24]Lv
		FLK1 flk;                     //FLK Flicker
		//CString ModuleName;
		int ChannelNO;
		int Current;
		char BarCode[128];
		char DriverDevice[128];
		char WorkNum[64];

	};
	struct MsrDate OQC_Date[10];
//Color = W,R,G,B,D,other
//Value = Lv,x,y,T,duv,du,dv,X,Y,Z

//  	   MsrNineOValue[Color][n][value]..
    double MsrNineOValue[6][9][10];
//         MsrFityNineOValue[Color][n][value]..
    double MsrFortyNineOValue[6][49][10];
//         MsrCenter[Color][value]..
    double MsrCenter[6][10];
//         FiveNits[n][value]..
    double FiveNits[9][10];  //Don't Care Color
//  	   MsrNineOValue[Color][n][value]..
	double MsrFiveOValue[6][5][10];
//	       MsrThirteenValue[n][value]..
//	double MsrThirteenValue[13][10];
//         MsrTwentyFiveValue[n][value]..
	double MsrTwentyFiveValue[25][10];
//         MsrTwentyNineValue[n][value]..as MsrThirteenValue
//    double MsrTwentyNineValue[29][10];
//         MsrThirtyThreeValue[n][value]..as MsrThirteenValue
//    double MsrThirtyThreeValue[33][10];
//         MsrFLK[value]..
	double MsrFLK;
//         Gamma[Color][ n ][value]..
    //double Gamma[6][256][11];
    
/*  ��=i,��=j
Color = RGBWD,other

   MsrNineOValue           5nitsValue               MsrCenter

   Lv  x   y  T   duv      Lv x   y   du  dv        xLv     y  du  dv
0                       0                       W0
1                       1                       R1
2                       2                       G2
3                       3                       B3
4                       4                       
5                       5
6                       6
7                       7
8                       8


   Gamma RGB              MsrFortyNineOValue
   Lv  x   y  du  dv      Lv  x   y  du  dv
0                       0
1                       1
...                    ...
255                    48
    
*/
        //�]�w�I���
        //NineO(int few);//�e�ĴX�Ӱ�]�H9�I���p�^
    CPoint Set9Point(int few = 4, float FromEdge = 2);  //�]�w�ĴX�I��pixel��m�]�H�E�I���p�^
	CPoint Set5nits9Point(int few = 4, float FromEdge = 6);  //�]�w�ĴX�I��pixel��m�]�H�E�I���p�^
	CPoint Set5Point(int few, float FromEdge = 0);
    CPoint SetD9Point(int few = 4);  //�]�w�ĴX�I��pixel��m�]�H�E�I���p�^
	CPoint Set49Point(int few);                     //���Ʀr�A��X�y�С]49�I�^
	CPoint SetD25Point(int few);
	CPoint SetD29Point(int few);
	CPoint SetD33Point(int few);
    //CPoint Set9Point();         //�w�]�s�����I

        //�e�F����
    //   �e���]�����I, �b�|�^
    void DrawACircle(CPoint point, int radius, int Percent = 6);//�e��
	//void DrawSight(CPoint point,int radius, int Percent=6);
    //   �e�w����T���ҡ]�����I,�t�X��Υb�|,�O�_�g�r�^
    void DrawMsrLabel(CPoint point, int radius, double *Data);  //��ܹw�����
	//	�e�X����Ϊ����
	void DrawColorLabel(CPoint StartPoint, int radius);

	//��
	//int ruler(int);
	int CmtoPixel(double);

	//�q���֤�
	void Msr1P();
	void Msr9P();
	void Msr25P();
	void Msr29P();
	void Msr33P();
	void Msr49P();

	void MsrFine5Nits();
	void Msr5Nits9P();
	void Msr5P();
	void MsrFLIC();
	void MsrLab();

// Dialog Data
    //{{AFX_DATA(CSwordDlg)
	enum { IDD = IDD_SWORD_DIALOG };
    C_CaControl    m_CCaControl;
    CString    m_strTestValue;
    float    m_floatFromEdge;
	BOOL	m_bool_du;
	BOOL	m_bool_dv;
	BOOL	m_bool_x;
	BOOL	m_bool_y;
	BOOL	m_bool_Lv;
	BOOL	m_bool_duv;
	BOOL	m_bool_T;
	BOOL	m_bool_X;
	BOOL	m_bool_Y;
	BOOL	m_bool_Z;
	BOOL	m_bool_49cvr1;
	BOOL	m_bool_49cvr9;
	BOOL	m_bool_9cvr1;
	BOOL	m_bool_9cvr49;
	BOOL	m_bool_1cvr49;
	BOOL	m_bool_1cvr9;
	BOOL	m_bool_oqc_5nits;
	BOOL	m_bool_oqc_b;
	BOOL	m_bool_oqc_g;
	BOOL	m_bool_oqc_r;
	BOOL	m_bool_oqc_w9;
	BOOL	m_bool_oqc_w5;
	BOOL	m_bool_oqc_r5;
	BOOL	m_bool_oqc_g5;
	BOOL	m_bool_oqc_b5;
	BOOL	m_bool_oqc_d5;
	BOOL	m_bool_oqc_d25;
	int		m_int_5nitsGray;
	BOOL	m_bool_oqc_w49;
	int		m_iActEnter;
	int		m_iAuto;
	BOOL	m_bool_oqc_d;
	BOOL	m_bool_oqc_w;
	float	m_float5FromEdge;
	//}}AFX_DATA

    // ClassWizard generated virtual function overrides
    //{{AFX_VIRTUAL(CSwordDlg)
    protected:
    virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
    //}}AFX_VIRTUAL

// Implementation
//�q���t��k
    BOOL MsrAI(int WorkMode, BOOL CtlMode = FALSE);
	BOOL MsrAct(BOOL CtlMode = FALSE);
    BOOL Pass(int PassMode, BOOL CtlMode);
//�������̤j�]�w
    const BOOL GetMaxWindow();
    void SetMaxWindow(bool flag);
//	CPaintDC pPDC;
protected:
    CSwordDlgAutoProxy* m_pAutoProxy;
    HICON m_hIcon;
   
    BOOL CanExit();

    // Generated message map functions
    //{{AFX_MSG(CSwordDlg)
    virtual BOOL OnInitDialog();
    afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
    afx_msg void OnPaint();
    afx_msg HCURSOR OnQueryDragIcon();
    afx_msg void OnClose();
    virtual void OnOK();
    virtual void OnCancel();
    afx_msg BOOL OnMouseWheel(UINT nFlags, short zDelta, CPoint pt);
    afx_msg void OnMouseMove(UINT nFlags, CPoint point);
    afx_msg HBRUSH OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor);
    afx_msg void OnRed();
    afx_msg void OnLButtonUp(UINT nFlags, CPoint point);
    afx_msg void OnGreen();
    afx_msg void OnBlue();
    afx_msg void OnWhite();
    afx_msg void OnBlack();
    afx_msg void OnUsercolor();
    afx_msg void OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags);
    afx_msg void OnKeyUp(UINT nChar, UINT nRepCnt, UINT nFlags);
    afx_msg void OnMsrFlowStart();
    afx_msg void OnButton2();
    afx_msg void OnSave1();
    afx_msg void OnMax();
    afx_msg void OnButton3();
    afx_msg void OnUpdateCacontrol1();
    afx_msg void OnButton4();
    afx_msg void OnSave2();
	afx_msg void OnButton5();
	afx_msg void OnZerocal();
	afx_msg void OnButton6();
	afx_msg void OnSave13p();
	afx_msg void OnButtonRst();
	afx_msg void OnCheck9cvr49();
	afx_msg void OnCheck9cvr1();
	afx_msg void OnCheck49cvr9();
	afx_msg void OnCheck49cvr1();
	afx_msg void OnCheckLv();
	afx_msg void OnCheckSx();
	afx_msg void OnCheckSy();
	afx_msg void OnCheckT();
	afx_msg void OnCHECKDuv();
	afx_msg void OnCheckDu();
	afx_msg void OnCheckDv();
	afx_msg void OnCheckX();
	afx_msg void OnCheckY();
	afx_msg void OnCheckZ();
	afx_msg void OnChange5nitsGray();
	afx_msg void OnOutofmemorySpin1(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnCheck1cvr9();
	afx_msg void OnCheck1cvr49();
	afx_msg void OnChangeFromEdge();
	afx_msg void OnButton7();
	afx_msg void OnSave25p();
	afx_msg void OnFlkSpix();
	afx_msg void OnFlk2l();
	afx_msg void OnFlk2l1();
	afx_msg void OnFlkVs();
	afx_msg void OnFlkCol();
	afx_msg void OnButton8();
	afx_msg void OnTempfor26();
	afx_msg void OnButtonMsrflow();
	afx_msg void OnCheckOqc5nits();
	afx_msg void OnCheckOqcr();
	afx_msg void OnCheckOqcg();
	afx_msg void OnCheckOqcb();
	afx_msg void OnCheckOqcw9();
	afx_msg void OnCheckOqcd25();
	afx_msg void OnButtonOqc();
	afx_msg void OnSelchangeComboSaveoqc();
	afx_msg void OnLoadBarCode();
	afx_msg void OnChangeEditOrigpath();
	afx_msg void OnSelchangeComboSeltai();
	afx_msg void OnRadioActenter();
	afx_msg void OnRadioAuto();
	afx_msg void OnCheckOqcw49();
	afx_msg void OnAutomsrcal();
	afx_msg void OnCheckOqcw5();
	afx_msg void OnCheckOqcd5();
	afx_msg void OnCheckOqcr5();
	afx_msg void OnCheckOqcg5();
	afx_msg void OnCheckOqcb5();
	afx_msg void OnCheckOqcw();
	afx_msg void OnCheckOqcd();
    DECLARE_EVENTSINK_MAP()
	//}}AFX_MSG
    DECLARE_MESSAGE_MAP()
    int AutoSelectMode(COLORREF color);//��¦�t�A
    const int SelectColorLabel(COLORREF BKColor);//��O����Ŷ��Ϊ�
private:
	void SECForm(MsrDate&);
	void SECForm();
	CBrush m_Brush;
    CFont m_font;
	//DWORD dwCurDirPathLen;
	char  szCurrentDirectory[MAX_PATH];
	char  szPath[MAX_PATH];
	CString BCFandODFPath;
    //CA-SDK�Ϊ��ܼ�
    IProbeInfo      m_IProbeInfo;
    long            m_lIndex;
    ICas            m_ICas;
    IProbes         m_IProbes;
    IOutputProbes   m_IOutputProbes;
    IProbe          m_IProbe;
    IMemory         m_IMemory;
    ICa200          m_ICa200;
    ICa             m_ICa;
    _ICaEvents      m__ICaEvents;
    //long m_lIndex;
    
	CPoint   CurrentPoint;
	CPoint   StopPoint;
	CPoint   BackPoint;

	double Brain;

    COLORREF BackGroundColor;
    BOOL MaxWindow;

    int CircleRadius;
    int MsrPass;
    int ScrmV;
    int ScrmH;
	
    int Few;
    int ColorLabel;

	int DiaplayCircle;
    int iLCMSize;
	int m_Tai;
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_SWORDDLG_H__5B6D51A7_96D4_4809_A56B_01AE04E41FF1__INCLUDED_)
