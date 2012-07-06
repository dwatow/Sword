/*
  �䥦�ɮ׻���

  ��ExcelForm
  �U�|Exl***.cpp    �s�ɪ��
  �uCheckBox.cpp    �ĤĪ���B�����J��
  �uDefMsrPoine.cpp �w�q�q���I��m
  �u
  �uHideAndShow.cpp ����֭n��ܡA�֭n����
  �u
  �uMsrItem.cpp     �w�q�q���y�{�]���n�^
  �|PatternType.cpp �w�q�q��Pattern���˦�
*/
// SwordDlg.cpp : implementation file
//
#include "stdafx.h"
#include <comdef.h>
#include <ctime>
#include "Sword.h"
#include "SwordDlg.h"
#include "DlgProxy.h"
#include <stdexcept>
#include "EnterDlg.h"
//#include "InitialDlg.h"
#include "excel.h"
//#include <time.h>

//#define SelectBox(x,y)              ZeroMemory(buf,sizeof(buf));sprintf(buf,"%c%d",x,y);range=objSheet.GetRange(COleVariant(buf),COleVariant(buf));
//#define SetBoxDate(Format, Data)    ZeroMemory(buf,sizeof(buf));sprintf(buf,Format,Data);range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf));
#define RADIUS 21 //cm

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CAboutDlg dialog used for App About

class CAboutDlg : public CDialog
{
public:
    CAboutDlg();

// Dialog Data
    //{{AFX_DATA(CAboutDlg)
    enum { IDD = IDD_ABOUTBOX };
    //}}AFX_DATA

    // ClassWizard generated virtual function overrides
    //{{AFX_VIRTUAL(CAboutDlg)
    protected:
    virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
    //}}AFX_VIRTUAL

// Implementation
protected:
    //{{AFX_MSG(CAboutDlg)
    //}}AFX_MSG
    DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialog(CAboutDlg::IDD)
{
    //{{AFX_DATA_INIT(CAboutDlg)
    //}}AFX_DATA_INIT
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
    CDialog::DoDataExchange(pDX);
    //{{AFX_DATA_MAP(CAboutDlg)
    //}}AFX_DATA_MAP
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialog)
    //{{AFX_MSG_MAP(CAboutDlg)
        // No message handlers
    //}}AFX_MSG_MAP
END_MESSAGE_MAP()


/////////////////////////////////////////////////////////////////////////////
// CSwordDlg dialog

IMPLEMENT_DYNAMIC(CSwordDlg, CDialog);

CSwordDlg::CSwordDlg(CWnd* pParent /*=NULL*/)
    : CDialog(CSwordDlg::IDD, pParent)
{
    //{{AFX_DATA_INIT(CSwordDlg)
    m_strTestValue = _T("");
    m_floatFromEdge = 6.0f;
    m_bool_du = FALSE;
    m_bool_dv = FALSE;
    m_bool_x = FALSE;
    m_bool_y = FALSE;
    m_bool_Lv = FALSE;
    m_bool_duv = FALSE;
    m_bool_T = FALSE;
    m_bool_X = FALSE;
    m_bool_Y = FALSE;
    m_bool_Z = FALSE;
    m_bool_49cvr1 = FALSE;
    m_bool_49cvr9 = FALSE;
    m_bool_9cvr1 = FALSE;
    m_bool_9cvr49 = FALSE;
    m_bool_1cvr49 = FALSE;
    m_bool_1cvr9 = FALSE;
    m_bool_oqc_5nits = FALSE;
    m_bool_oqc_b = FALSE;
    m_bool_oqc_g = FALSE;
    m_bool_oqc_r = FALSE;
    m_bool_oqc_w9 = FALSE;
    m_bool_oqc_w5 = FALSE;
    m_bool_oqc_r5 = FALSE;
    m_bool_oqc_g5 = FALSE;
    m_bool_oqc_b5 = FALSE;
    m_bool_oqc_d5 = FALSE;
    m_bool_oqc_d25 = FALSE;
    m_int_5nitsGray = 0;
    m_bool_oqc_w49 = FALSE;
    m_iActEnter = -1;
    m_iAuto = -1;
	m_bool_oqc_d = FALSE;
	m_bool_oqc_w = FALSE;
    m_bool_oqc_w49 = FALSE;
    m_bool_oqc_r5 = FALSE;
    m_bool_oqc_g5 = FALSE;
    m_bool_oqc_b5 = FALSE;
    m_bool_oqc_d5 = FALSE;
	m_float5FromEdge = 0.0f;
	//}}AFX_DATA_INIT
    // Note that LoadIcon does not require a subsequent DestroyIcon in Win32
    m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
    m_pAutoProxy = NULL;
}

CSwordDlg::~CSwordDlg()
{
    // If there is an automation proxy for this dialog, set
    //  its back pointer to this dialog to NULL, so it knows
    //  the dialog has been deleted.
    m_ICa.SetRemoteMode(0);//���u
    if (m_pAutoProxy != NULL)
        m_pAutoProxy->m_pDialog = NULL;
}

void CSwordDlg::DoDataExchange(CDataExchange* pDX)
{
    CDialog::DoDataExchange(pDX);
    //{{AFX_DATA_MAP(CSwordDlg)
    DDX_Control(pDX, IDC_CACONTROL1, m_CCaControl);
    DDX_Text(pDX, IDC_TESTVALUE, m_strTestValue);
    DDX_Text(pDX, IDC_FROM_EDGE, m_floatFromEdge);
    DDV_MinMaxFloat(pDX, m_floatFromEdge, 1.f, 50.f);
    DDX_Check(pDX, IDC_CHECK_DU, m_bool_du);
    DDX_Check(pDX, IDC_CHECK_DV, m_bool_dv);
    DDX_Check(pDX, IDC_CHECK_SX, m_bool_x);
    DDX_Check(pDX, IDC_CHECK_SY, m_bool_y);
    DDX_Check(pDX, IDC_CHECK_Lv, m_bool_Lv);
    DDX_Check(pDX, IDC_CHECK_Duv, m_bool_duv);
    DDX_Check(pDX, IDC_CHECK_T, m_bool_T);
    DDX_Check(pDX, IDC_CHECK_X, m_bool_X);
    DDX_Check(pDX, IDC_CHECK_Y, m_bool_Y);
    DDX_Check(pDX, IDC_CHECK_Z, m_bool_Z);
    DDX_Check(pDX, IDC_CHECK_49CVR1, m_bool_49cvr1);
    DDX_Check(pDX, IDC_CHECK_49CVR9, m_bool_49cvr9);
    DDX_Check(pDX, IDC_CHECK_9CVR1, m_bool_9cvr1);
    DDX_Check(pDX, IDC_CHECK_9CVR49, m_bool_9cvr49);
    DDX_Check(pDX, IDC_CHECK_1CVR49, m_bool_1cvr49);
    DDX_Check(pDX, IDC_CHECK_1CVR9, m_bool_1cvr9);
    DDX_Check(pDX, IDC_CHECK_OQC5NITS, m_bool_oqc_5nits);
    DDX_Check(pDX, IDC_CHECK_OQCB, m_bool_oqc_b);
    DDX_Check(pDX, IDC_CHECK_OQCG, m_bool_oqc_g);
    DDX_Check(pDX, IDC_CHECK_OQCR, m_bool_oqc_r);
    DDX_Check(pDX, IDC_CHECK_OQCW9, m_bool_oqc_w9);
    DDX_Check(pDX, IDC_CHECK_OQCW5, m_bool_oqc_w5);
    DDX_Check(pDX, IDC_CHECK_OQCR5, m_bool_oqc_r5);
    DDX_Check(pDX, IDC_CHECK_OQCG5, m_bool_oqc_g5);
    DDX_Check(pDX, IDC_CHECK_OQCB5, m_bool_oqc_b5);
    DDX_Check(pDX, IDC_CHECK_OQCD5, m_bool_oqc_d5);
    DDX_Check(pDX, IDC_CHECK_OQCD25, m_bool_oqc_d25);
    DDX_Text(pDX, IDC_5NITS_GRAY, m_int_5nitsGray);
    DDX_Check(pDX, IDC_CHECK_OQCW49, m_bool_oqc_w49);
    DDX_Radio(pDX, IDC_RADIO_ACTENTER, m_iActEnter);
    DDX_Radio(pDX, IDC_RADIO_AUTO, m_iAuto);
	DDX_Check(pDX, IDC_CHECK_OQCD, m_bool_oqc_d);
	DDX_Check(pDX, IDC_CHECK_OQCW, m_bool_oqc_w);
	DDX_Text(pDX, IDC_FROM_5EDGE, m_float5FromEdge);
	DDV_MinMaxFloat(pDX, m_float5FromEdge, 0.f, 50.f);
	//}}AFX_DATA_MAP
}

BEGIN_MESSAGE_MAP(CSwordDlg, CDialog)
    //{{AFX_MSG_MAP(CSwordDlg)
    ON_WM_SYSCOMMAND()
    ON_WM_PAINT()
    ON_WM_QUERYDRAGICON()
    ON_WM_CLOSE()
    ON_WM_MOUSEWHEEL()
    ON_WM_MOUSEMOVE()
    ON_WM_CTLCOLOR()
    ON_BN_CLICKED(IDC_RED, OnRed)
    ON_WM_LBUTTONUP()
    ON_BN_CLICKED(IDC_GREEN, OnGreen)
    ON_BN_CLICKED(IDC_BLUE, OnBlue)
    ON_BN_CLICKED(IDC_WHITE, OnWhite)
    ON_BN_CLICKED(IDC_BLACK, OnBlack)
    ON_BN_CLICKED(IDC_USERCOLOR, OnUsercolor)
    ON_WM_KEYDOWN()
    ON_WM_KEYUP()
    ON_BN_CLICKED(IDC_BUTTON1, OnMsrFlowStart)
    ON_BN_CLICKED(IDC_BUTTON2, OnButton2)
    ON_BN_CLICKED(IDC_SAVE1, OnSave1)
    ON_BN_CLICKED(IDC_MAX, OnMax)
    ON_BN_CLICKED(IDC_BUTTON3, OnButton3)
    ON_BN_CLICKED(IDC_BUTTON4, OnButton4)
    ON_BN_CLICKED(IDC_SAVE2, OnSave2)
    ON_BN_CLICKED(IDC_BUTTON5, OnButton5)
    ON_BN_CLICKED(IDC_BUTTON6, OnButton6)
    ON_BN_CLICKED(IDC_SAVE13P, OnSave13p)
    ON_BN_CLICKED(IDC_BUTTON_RST, OnButtonRst)
    ON_BN_CLICKED(IDC_CHECK_9CVR49, OnCheck9cvr49)
    ON_BN_CLICKED(IDC_CHECK_9CVR1, OnCheck9cvr1)
    ON_BN_CLICKED(IDC_CHECK_49CVR9, OnCheck49cvr9)
    ON_BN_CLICKED(IDC_CHECK_49CVR1, OnCheck49cvr1)
    ON_BN_CLICKED(IDC_CHECK_SX, OnCheckSx)
    ON_BN_CLICKED(IDC_CHECK_SY, OnCheckSy)
    ON_BN_CLICKED(IDC_CHECK_T, OnCheckT)
    ON_BN_CLICKED(IDC_CHECK_Duv, OnCHECKDuv)
    ON_BN_CLICKED(IDC_CHECK_DU, OnCheckDu)
    ON_BN_CLICKED(IDC_CHECK_DV, OnCheckDv)
    ON_BN_CLICKED(IDC_CHECK_X, OnCheckX)
    ON_BN_CLICKED(IDC_CHECK_Y, OnCheckY)
    ON_BN_CLICKED(IDC_CHECK_Z, OnCheckZ)
    ON_EN_CHANGE(IDC_5NITS_GRAY, OnChange5nitsGray)
    ON_NOTIFY(NM_OUTOFMEMORY, IDC_SPIN1, OnOutofmemorySpin1)
    ON_BN_CLICKED(IDC_CHECK_1CVR9, OnCheck1cvr9)
    ON_BN_CLICKED(IDC_CHECK_1CVR49, OnCheck1cvr49)
    ON_EN_CHANGE(IDC_FROM_EDGE, OnChangeFromEdge)
    ON_BN_CLICKED(IDC_BUTTON7, OnButton7)
    ON_BN_CLICKED(IDC_SAVE25P, OnSave25p)
    ON_BN_CLICKED(IDC_FLK_SPIX, OnFlkSpix)
    ON_BN_CLICKED(IDC_FLK_2L, OnFlk2l)
    ON_BN_CLICKED(IDC_FLK_2L1, OnFlk2l1)
    ON_BN_CLICKED(IDC_FLK_VS, OnFlkVs)
    ON_BN_CLICKED(IDC_FLK_COL, OnFlkCol)
    ON_BN_CLICKED(IDC_BUTTON8, OnButton8)
    ON_BN_CLICKED(IDC_TEMPFOR26, OnTempfor26)
    ON_BN_CLICKED(IDC_BUTTON_MSRFLOW, OnButtonMsrflow)
    ON_BN_CLICKED(IDC_CHECK_OQC5NITS, OnCheckOqc5nits)
    ON_BN_CLICKED(IDC_CHECK_OQCR, OnCheckOqcr)
    ON_BN_CLICKED(IDC_CHECK_OQCG, OnCheckOqcg)
    ON_BN_CLICKED(IDC_CHECK_OQCB, OnCheckOqcb)
    ON_BN_CLICKED(IDC_CHECK_OQCW9, OnCheckOqcw9)
    ON_BN_CLICKED(IDC_CHECK_OQCD25, OnCheckOqcd25)
    ON_BN_CLICKED(IDC_SAVEOQC, OnButtonOqc)
    ON_CBN_SELCHANGE(IDC_COMBO_SAVEOQC, OnSelchangeComboSaveoqc)
    ON_BN_CLICKED(IDC_BUTTON_LOADBC, OnLoadBarCode)
    ON_EN_CHANGE(IDC_EDIT_ORIGPATH, OnChangeEditOrigpath)
    ON_CBN_SELCHANGE(IDC_COMBO_SELTAI, OnSelchangeComboSeltai)
    ON_BN_CLICKED(IDC_RADIO_ACTENTER, OnRadioActenter)
    ON_BN_CLICKED(IDC_RADIO_AUTO, OnRadioAuto)
    ON_BN_CLICKED(IDC_CHECK_OQCW49, OnCheckOqcw49)
    ON_BN_CLICKED(IDC_AUTOMSRCAL, OnAutomsrcal)
    ON_BN_CLICKED(IDC_CHECK_OQCW5, OnCheckOqcw5)
    ON_BN_CLICKED(IDC_CHECK_OQCD5, OnCheckOqcd5)
    ON_BN_CLICKED(IDC_CHECK_OQCR5, OnCheckOqcr5)
    ON_BN_CLICKED(IDC_CHECK_OQCG5, OnCheckOqcg5)
    ON_BN_CLICKED(IDC_CHECK_OQCB5, OnCheckOqcb5)
	ON_BN_CLICKED(IDC_CHECK_OQCW, OnCheckOqcw)
    ON_BN_CLICKED(IDC_CHECK_Lv, OnCheckLv)
	ON_BN_CLICKED(IDC_CHECK_OQCD, OnCheckOqcd)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CSwordDlg message handlers

BOOL CSwordDlg::OnInitDialog()
{
    CDialog::OnInitDialog();

    // Add "About..." menu item to system menu.

    // IDM_ABOUTBOX must be in the system command range.
    ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
    ASSERT(IDM_ABOUTBOX < 0xF000);

    CMenu* pSysMenu = GetSystemMenu(FALSE);
    if (pSysMenu != NULL)
    {
        CString strAboutMenu;
        strAboutMenu.LoadString(IDS_ABOUTBOX);
        if (!strAboutMenu.IsEmpty())
        {
            pSysMenu->AppendMenu(MF_SEPARATOR);
            pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
        }
    }

    // Set the icon for this dialog.  The framework does this automatically
    //  when the application's main window is not a dialog
    SetIcon(m_hIcon, TRUE);            // Set big icon
    SetIcon(m_hIcon, FALSE);        // Set small icon
    // TODO: Add extra initialization here

/////////////////////////////////////////////////////////////////////////////////
//�}�l�ۭq��l��

   //dwCurDirPathLen = 
    GetCurrentDirectory(MAX_PATH,szCurrentDirectory);//��ثe�Ҧb���ؿ��]���|�^
    SHGetSpecialFolderPath(NULL,szPath,CSIDL_DESKTOP,0);//���o�ୱ���|
    BCFandODFPath.Format("%s",szPath);
    ScrmH = GetSystemMetrics(SM_CXSCREEN);
    ScrmV = GetSystemMetrics(SM_CYSCREEN);
    //����C�@�Ӥ��󪺫���
//�������A
    pBtmMax    = (CButton*)GetDlgItem(IDC_MAX);
    pBtmOK     = (CButton*)GetDlgItem(IDOK);
    pBtmCANCEL = (CButton*)GetDlgItem(IDCANCEL);
//���C��
    pBtmRED       = (CButton*)GetDlgItem(IDC_RED);
    pBtmGREEN     = (CButton*)GetDlgItem(IDC_GREEN);
    pBtmBLUE      = (CButton*)GetDlgItem(IDC_BLUE);
    pBtmWhite     = (CButton*)GetDlgItem(IDC_WHITE);
    pBtmBlack     = (CButton*)GetDlgItem(IDC_BLACK);
    pBtmUserColor = (CButton*)GetDlgItem(IDC_USERCOLOR);
//�E�I�q������
    p9Cvr1CheckBox  = (CButton *)GetDlgItem(IDC_CHECK_9CVR1);
    p9Cvr49CheckBox = (CButton *)GetDlgItem(IDC_CHECK_9CVR49);
    pFE             = (CStatic*)GetDlgItem(IDC_STATIC_FE);
    pFromEdge       = (CEdit*)GetDlgItem(IDC_FROM_EDGE);
    pBtmNineO       = (CButton*)GetDlgItem(IDC_BUTTON1);
//�����I�q�����
    p1Cvr9CheckBox  = (CButton *)GetDlgItem(IDC_CHECK_1CVR9);
    p1Cvr49CheckBox = (CButton *)GetDlgItem(IDC_CHECK_1CVR49);
    pBtmCentrO      = (CButton*)GetDlgItem(IDC_BUTTON2);
//5nits�q�����
    pEditBackUalue = (CEdit*)GetDlgItem(IDC_5NITS_GRAY);
    pBtmFiveNits   = (CButton*)GetDlgItem(IDC_BUTTON3);
    pSpin          = (CSpinButtonCtrl*)GetDlgItem(IDC_SPIN1);
//RGB Gamma�q��
    pBtmGamma      = (CButton*)GetDlgItem(IDC_BUTTON4);
//49�I�q��
    p49Cvr1CheckBox  = (CButton *)GetDlgItem(IDC_CHECK_49CVR1);
    p49Cvr9CheckBox  = (CButton *)GetDlgItem(IDC_CHECK_49CVR9);
    pBtmFNO          = (CButton*)GetDlgItem(IDC_BUTTON5);
//13�I�q��
//    pBtmThirteenO    = (CButton*)GetDlgItem(IDC_BUTTON6);
//25�I�q��
    pBtmTwentyFiveO  = (CButton*)GetDlgItem(IDC_BUTTON7);
//����
    pBtmAgingLab     = (CButton*)GetDlgItem(IDC_CA_AGING);
//���i
    pBtmSave1          = (CButton*)GetDlgItem(IDC_SAVE1);
    pBtmSave2          = (CButton*)GetDlgItem(IDC_SAVE2);
    pBtmAgingLabSave   = (CButton*)GetDlgItem(IDC_AGINGFROM);
//    pBtmSave13p        = (CButton*)GetDlgItem(IDC_SAVE13P);
    pBtmSave25p        = (CButton*)GetDlgItem(IDC_SAVE25P);
    pBtmSaveT26        = (CButton*)GetDlgItem(IDC_TEMPFOR26);
    pBtmSaveOQC        = (CButton*)GetDlgItem(IDC_SAVEOQC);
//�䥦    
    pBtmRstDate        = (CButton*)GetDlgItem(IDC_BUTTON_RST);
    pTextRredMe        = (CStatic*)GetDlgItem(IDC_STATIC_README);
    pMeasureFlowList   = (CStatic*)GetDlgItem(IDC_STATIC_MFL);
    pRePortList        = (CStatic*)GetDlgItem(IDC_STATIC_RPL);
//    pBtmZeroCal        = (CButton*)GetDlgItem(IDC_ZEROCAL);
//�q���w����
    pDisplayValuesList = (CStatic*)GetDlgItem(IDC_STATIC_DISVALUES);
//�w���ȤĿﶵ��
    pMsrViewList       = (CStatic*)GetDlgItem(IDC_STATIC_DISVALUES);
    pLvCheckBox        = (CButton*)GetDlgItem(IDC_CHECK_Lv);
    pSXCheckBox        = (CButton*)GetDlgItem(IDC_CHECK_SX);
    pSYCheckBox        = (CButton*)GetDlgItem(IDC_CHECK_SY);

    pTCheckBox         = (CButton*)GetDlgItem(IDC_CHECK_T);
    pDUVCheckBox       = (CButton*)GetDlgItem(IDC_CHECK_Duv);

    pDUCheckBox        = (CButton*)GetDlgItem(IDC_CHECK_DU);
    pDVCheckBox        = (CButton*)GetDlgItem(IDC_CHECK_DV);

    pXCheckBox         = (CButton*)GetDlgItem(IDC_CHECK_X);
    pYCheckBox         = (CButton*)GetDlgItem(IDC_CHECK_Y);
    pZCheckBox         = (CButton*)GetDlgItem(IDC_CHECK_Z);

//    pFlkCheckBox       = (CButton*)GetDlgItem(IDC_CHECK_FLK);
//Flicker Pattern
    pBtmFlkMsr         = (CButton*)GetDlgItem(IDC_BUTTON8);
    pBtmFlkSubPix      = (CButton*)GetDlgItem(IDC_FLK_SPIX);
    pBtmFlk2LInv       = (CButton*)GetDlgItem(IDC_FLK_2L);
    pBtmFlk2L1Inv      = (CButton*)GetDlgItem(IDC_FLK_2L1);
    pBtmFlkVS          = (CButton*)GetDlgItem(IDC_FLK_VS);
    pBtmFlkCol         = (CButton*)GetDlgItem(IDC_FLK_COL);
//OQC
    pOQCTitle          = (CStatic*)  GetDlgItem(IDC_STATIC_OQC);
    pOQC5Nits          = (CComboBox*)GetDlgItem(IDC_CHECK_OQC5NITS);
    
    pOQCWCenter        = (CComboBox*)GetDlgItem(IDC_CHECK_OQCW);
    pOQCRCenter        = (CComboBox*)GetDlgItem(IDC_CHECK_OQCR);
    pOQCGCenter        = (CComboBox*)GetDlgItem(IDC_CHECK_OQCG);
    pOQCBCenter        = (CComboBox*)GetDlgItem(IDC_CHECK_OQCB);
    pOQCDCenter        = (CComboBox*)GetDlgItem(IDC_CHECK_OQCD);

    psOQCRGB           = (CStatic*)  GetDlgItem(IDC_STATIC_OQCRGB);

    pOQCW9P            = (CComboBox*)GetDlgItem(IDC_CHECK_OQCW9);
    psOQCW9P           = (CStatic*)  GetDlgItem(IDC_STATIC_OQCW9);

//���I�q������
    pOQCW5P            = (CComboBox*)GetDlgItem(IDC_CHECK_OQCW5);
    pOQCR5P            = (CComboBox*)GetDlgItem(IDC_CHECK_OQCR5);
    pOQCG5P            = (CComboBox*)GetDlgItem(IDC_CHECK_OQCG5);
    pOQCB5P            = (CComboBox*)GetDlgItem(IDC_CHECK_OQCB5);
    pOQCD5P            = (CComboBox*)GetDlgItem(IDC_CHECK_OQCD5);

    psOQCW5P           = (CStatic*)  GetDlgItem(IDC_STATIC_OQCW5);
    psOQCR5P           = (CStatic*)  GetDlgItem(IDC_STATIC_OQCW5);
    psOQCG5P           = (CStatic*)  GetDlgItem(IDC_STATIC_OQCW5);
    psOQCB5P           = (CStatic*)  GetDlgItem(IDC_STATIC_OQCW5);
    psOQCD5P           = (CStatic*)  GetDlgItem(IDC_STATIC_OQCW5);
    p5FE               = (CStatic*)  GetDlgItem(IDC_STATIC_5FE);
    p5FromEdge         = (CEdit*)    GetDlgItem(IDC_FROM_5EDGE);

    pOQCW49P            = (CComboBox*)GetDlgItem(IDC_CHECK_OQCW49);
    psOQCW49P           = (CStatic*)  GetDlgItem(IDC_STATIC_OQCW49);

    pOQCD25P           = (CComboBox*)GetDlgItem(IDC_CHECK_OQCD25);
//    pOQCD33P           = (CComboBox*)GetDlgItem(IDC_CHECK_OQCD33);
    psOQCD25P          = (CStatic*)  GetDlgItem(IDC_STATIC_OQCD25);
    
    //pOQCFlickerStatic  = (CStatic*)  GetDlgItem(IDC_STATIC_OQCFLK);
    //pOQCFlickerSel     = (CComboBox*)GetDlgItem(IDC_COMBO_SELFLKPTM);  //pCmbSelFlkPtn
    
    pOQCTaiSel         = (CComboBox*)GetDlgItem(IDC_COMBO_SELTAI);     //pCmbSelTai
    pOQCFormSel        = (CComboBox*)GetDlgItem(IDC_COMBO_SAVEOQC);
    pOQCStartMsr       = (CButton*)GetDlgItem(IDC_BUTTON_MSRFLOW);
    pOQCLoadBarCode    = (CButton*)GetDlgItem(IDC_BUTTON_LOADBC);

    pMsrOperation      = (CStatic*)GetDlgItem(IDC_STATIC_MSROPITEM);
    pActEnter          = (CButton*)GetDlgItem(IDC_RADIO_ACTENTER);
    pAutoMsr           = (CButton*)GetDlgItem(IDC_RADIO_AUTO);
//    pBtmAutoMsrCal     = (CButton*)GetDlgItem(IDC_AUTOMSRCAL);

    MerOperation = TRUE;
    m_iActEnter = MerOperation?1:0;
    m_iAuto = MerOperation?0:1;
//�]�w�����r��
/*Example:
font.CreateFont(30,0,0,0,FW_BLACK,FALSE,FALSE,FALSE, 
GB2312_CHARSET,OUT_DEFAULT_PRECIS,CLIP_DEFAULT_PRECIS,   
DEFAULT_QUALITY,FIXED_PITCH   |   FF_MODERN, "�ؤ�淢");*/

    m_font.CreateFont(
        18,                    //���w�Q�n�����ס]�޿���^���r��C
        0,                     //���w�����e�ס]�޿���^�����r�Ū��r��C
        0,                     //���w���ס]�H0.1�׬����^�A�����צV�q�Px�b�������C 
        0,                     //���w���ס]�H0.1�׬����^�������r�Ű�u�MX�b�C
        FW_REGULAR,            //���w�r�骺�ʲӡ]�H���������A�Cñ�q1000�^�C 
        FALSE,                 //���w�r��O�_������C
        FALSE,                 //���w�r��O�_���U���u�C
        FALSE,                 //���w�r�餤���r�ŬO�_�O�簣�C
        SYMBOL_CHARSET,        //���w�r���]�Ÿ��B��ڡBBIG5�������^
        OUT_DEFAULT_PRECIS,    //���w�һݪ���X��סC
        CLIP_DEFAULT_PRECIS,   //���w�һݪ��Ť���סC �ſ��שw�q�F�p��ſ�r�Ū��ſ�ϰ�H�~�������C
        DEFAULT_QUALITY,       //���w�r�骺��X��q�A���w�q�p���ߪ�GDI�������դǰt�޿�r���ݩʪ��@�ӹ�ڪ����z�r��C
        FIXED_PITCH,           //���w���r�Z�M�s�ա]Family�^���r��C 
        "Webdings");           //�@�� CString�� �Ϋ��w���ŵ������r�Ŧ�A���w�r��W�٪��r��C

    pBtmOK->SetFont(&m_font);  //��
    pBtmMax->SetFont(&m_font); //��

    //��ܵ{����l�Ƶ���
    CString strinit;          //initial ��ܸ��J�r��

    pInitialDlg = new CInitialDlg;
    pInitialDlg->Create(IDD_INITIAL);
    pInitialDlg->ShowWindow(SW_SHOW);    
    pInitialDlg->UpdateWindow();

//    ::SetTimer(pInitialDlg->m_hWnd,1,1,NULL);

    //���JCa210    
    // CA-210 Initialzation
                        strinit = "��l��CA-210...";
                        pInitialDlg->SetString(strinit);    Sleep(1000);

    
    m_ICa200.CreateDispatch("CA200Srvr.Ca200.1");
    m_ICa200.m_bAutoRelease = TRUE;

                        strinit = "�]�wPC�PCA-210���s�u...";
                        pInitialDlg->SetString(strinit);

    m_ICa200.AutoConnect();

                        strinit = "�]�wPC�PCA-210���q�������s�u(1/2)...";
                        pInitialDlg->SetString(strinit);    Sleep(100);

    LPDISPATCH pICa = m_ICa200.GetSingleCa();
    m_ICa.AttachDispatch(pICa);

                        strinit = "�]�wPC�PCA-210���q�������s�u(2/2)...";
                        pInitialDlg->SetString(strinit);    Sleep(100);
 
    LPDISPATCH pIProbe = m_ICa.GetSingleProbe();
    m_IProbe.AttachDispatch(pIProbe);

                        strinit = "Ū��Channel���...";
                        pInitialDlg->SetString(strinit);    Sleep(100);
    
    LPDISPATCH pMemory = m_ICa.GetMemory();
    m_IMemory.AttachDispatch(pMemory);

                        strinit = "�]�wPC����CA-210����...";
                        pInitialDlg->SetString(strinit);    Sleep(100);

    m_CCaControl.SetRefCa(&m_ICa.m_lpDispatch);
    m_CCaControl.SetRefProbe(&m_IProbe.m_lpDispatch);
    m_CCaControl.SetRefMemory(&m_IMemory.m_lpDispatch);
    m_CCaControl.UpdateCaInfo();

                        strinit = "Ū��Chanel�W��...(�Y�N0-Cal�A�����0-Cal)";
                        pInitialDlg->SetString(strinit);

    m_CCaControl.UpdateMemoryInfo();

                        strinit = "��l�Ƹ�ƼȦs�O����Ŷ�...";
                        pInitialDlg->SetString(strinit);    Sleep(500);

    //��l�Ƹ�ƼȦs�}�C
    initValues();

                        //strinit = "�]�w SYNC MODE...";
                        //pInitialDlg->SetString(strinit);    Sleep(100);

    //�]�w SYNC MODE
    float fsetsyncmode;
    fsetsyncmode = 3.0;//UNIV
    m_ICa.SetSyncMode(fsetsyncmode);
    m_ICa.SetAveragingMode((long)1);//FAST

                        //strinit = "�]�w�����w�]�I����";
                        //pInitialDlg->SetString(strinit);    Sleep(100);

    SetBackGroundColor(); //�w�]�I���C��

    m_bool_1cvr9  = m_bool_1cvr49 = 0;
    m_bool_9cvr1  = m_bool_9cvr49 = m_bool_49cvr1 = m_bool_49cvr9 = 1;
    m_bool_oqc_r = m_bool_oqc_g = m_bool_oqc_b = m_bool_oqc_w9 = m_bool_oqc_w5 = m_bool_oqc_r5 = m_bool_oqc_g5 = m_bool_oqc_b5 = m_bool_oqc_d5 = m_bool_oqc_w49 = m_bool_oqc_d25 = m_bool_oqc_5nits = 0;

    UpdateData(FALSE);

//    ::KillTimer(pInitialDlg->m_hWnd,1);
//    Invalidate();

    pOQCTaiSel->SetCurSel(0);
//    pOQCFlickerSel->SetCurSel(5);

    pOQCFormSel->ResetContent();
    pOQCFormSel->AddString("None");
    pOQCFormSel->SetCurSel(0);
  

//    MessageBox();

    ChooseCircleRadius();
    SetOQCSaveSel();

    Brain = 0.005;

    //m_ICa.CalZero();
    try
    {
        strinit = "Zero-Cal...";
        pInitialDlg->SetString(strinit);    Sleep(500);

        CString str;
        str.Format("�n0-Cal�o~");
        if (AfxMessageBox(str, MB_YESNO) == 6)
            m_ICa.CalZero();
    }
    catch (CException* e)
    {
        TCHAR buf[255];
        e->GetErrorMessage(buf, 255);
        
        CString str;
        str.Format("�A���@���I�n0-Cal�o~");
        if (AfxMessageBox(str, MB_YESNO) == 6)
            m_ICa.CalZero();
    }
    pInitialDlg->ShowWindow(SW_HIDE);    
    pInitialDlg->DestroyWindow();
    delete pInitialDlg;  //������l�ƹ�ܮ�

    //////////////////////////////////////////////////////////////////////////
    SetMaxWindow(0);//�]�w�ٵ����٨S�����̤j�]����ȡ^

    return TRUE;  // return TRUE  unless you set the focus to a control
}


// void CSwordDlg::OnZerocal() 
// {
//     m_ICa.CalZero();
//     pBtmZeroCal  -> EnableWindow(0);
// 
//     SetMaxWindow(0);//�]�w�ٵ����٨S�����̤j�]����ȡ^
// }

void CSwordDlg::initValues() 
{
    int i,j,n;
    for(j=0;j<10;++j){   //Lv,x,y,T,duv,du,dv,X,Y,Z,flk
        for(i=0;i<6;++i)//W,R,G,B,D,other
        {
            for(n=0;n<9;++n)//n�]9�I�ơ^
            {                               //             [i=6][n=9][j=11]
                MsrNineOValue[i][n][j] = FiveNits[n][j] = MsrCenter[i][j] = 0;//FiveNits     [Color]   [value]
            }//for(n=0;n<9;++n)//n�]9�I�ơ^

            for(n=0;n<49;++n)//n�]49�I�^
            {                               //                       [i=6][n=49][j=11]
                MsrFortyNineOValue[i][n][j] = 0;//MsrFortyNineOValue[Color][n][value]
            }//for(n=0;n<49;++n)//n�]49�I�^

        }//for(i=0;i<6;++i)//W,R,G,B,D,other
//        for(i=0;i<13;++i)//n
//        {                      //                         [i=13][j=11]
//            MsrThirteenValue[i][j] = 0;//MsrThirteenValue[n][value]
//        }//for(i=0;i<13;++i)//n

        for(i=0;i<25;++i)//n
        {                      //                         [i=13][j=11]
            MsrTwentyFiveValue[i][j] = 0;//MsrThirteenValue[n][value]
        }//for(i=0;i<25;++i)//n
        MsrFLK = 0;//  MsrFLK[value]
    }
    
    for(i=0;i<10;++i)
    {
        OQC_Date[i].ChannelNO = OQC_Date[i].Current = 0;
        for(j=0;j<30;++j)
        {
            OQC_Date[i].BarCode[j] = OQC_Date[i].DriverDevice[j] = 0;
        }
        for(j=0;j<9;++j)
        {
            OQC_Date[i].w9[j].Lv = OQC_Date[i].w9[j].x  = OQC_Date[i].w9[j].y  = 0.0;
            OQC_Date[i].d5nits9[j].Lv = OQC_Date[i].d5nits9[j].x = OQC_Date[i].d5nits9[j].y = 0.0;
        }
        OQC_Date[i].w.Lv = OQC_Date[i].w.x = OQC_Date[i].w.y = OQC_Date[i].w.T = OQC_Date[i].w.Duv = 0.0;
        OQC_Date[i].r.Lv = OQC_Date[i].r.x = OQC_Date[i].r.y = OQC_Date[i].r.T = OQC_Date[i].r.Duv = 0.0;
        OQC_Date[i].g.Lv = OQC_Date[i].g.x = OQC_Date[i].g.y = OQC_Date[i].g.T = OQC_Date[i].g.Duv = 0.0;
        OQC_Date[i].b.Lv = OQC_Date[i].b.x = OQC_Date[i].b.y = OQC_Date[i].b.T = OQC_Date[i].b.Duv = 0.0;
        OQC_Date[i].d.Lv = OQC_Date[i].d.x = OQC_Date[i].d.y = OQC_Date[i].d.T = OQC_Date[i].d.Duv = 0.0;

        OQC_Date[i].d5nits1.Lv = OQC_Date[i].d5nits1.x = OQC_Date[i].d5nits1.y = OQC_Date[i].d5nits1.T = OQC_Date[i].d5nits1.Duv = 0.0;
        
        OQC_Date[i].flk.Flckr = 0.0;
        
        for(j=0;j<25;++j)
        {
            OQC_Date[i].d25[j].Lv = 0.0;
        }
//         for(j=0;j<13;++j)
//         {
//             OQC_Date[i].d13[j].Lv = 0.0;
//         }
    }

}



void CSwordDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
    if ((nID & 0xFFF0) == IDM_ABOUTBOX)
    {
        CAboutDlg dlgAbout;
        dlgAbout.DoModal();
    }
    else
    {
        CDialog::OnSysCommand(nID, lParam); //�s�X�D����
    }
}

// If you add a minimize button to your dialog, you will need the code below
//  to draw the icon.  For MFC applications using the document/view model,
//  this is automatically done for you by the framework.

void CSwordDlg::OnPaint() 
{
    if (IsIconic())
    {
        CPaintDC dc(this); // device context for painting

        SendMessage(WM_ICONERASEBKGND, (WPARAM) dc.GetSafeHdc(), 0);

        // Center icon in client rectangle
        int cxIcon = GetSystemMetrics(SM_CXICON);
        int cyIcon = GetSystemMetrics(SM_CYICON);
        CRect rect;
        GetClientRect(&rect);
        int x = (rect.Width() - cxIcon + 1) / 2;
        int y = (rect.Height() - cyIcon + 1) / 2;

        // Draw the icon
        dc.DrawIcon(x, y, m_hIcon);
    }
    else
    {
        CDialog::OnPaint();
    }
}

HBRUSH CSwordDlg::OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor) 
{
    HBRUSH hbr = CDialog::OnCtlColor(pDC, pWnd, nCtlColor);
    if (pWnd->GetDlgCtrlID() == IDC_CHECK_OQCW)     pDC->SetTextColor(CV_White);
    if (pWnd->GetDlgCtrlID() == IDC_CHECK_OQCD)     pDC->SetTextColor(CV_Dark);
    if (pWnd->GetDlgCtrlID() == IDC_CHECK_OQCR)       pDC->SetTextColor(CV_Red);
    if (pWnd->GetDlgCtrlID() == IDC_CHECK_OQCG)       pDC->SetTextColor(CV_Green);
    if (pWnd->GetDlgCtrlID() == IDC_CHECK_OQCB)       pDC->SetTextColor(CV_Blue);
    if (pWnd->GetDlgCtrlID() == IDC_CHECK_OQCW9)       pDC->SetTextColor(CV_White);
    if (pWnd->GetDlgCtrlID() == IDC_CHECK_OQCW5)       pDC->SetTextColor(CV_White);
    if (pWnd->GetDlgCtrlID() == IDC_CHECK_OQCR5)      pDC->SetTextColor(CV_Red);
    if (pWnd->GetDlgCtrlID() == IDC_CHECK_OQCG5)      pDC->SetTextColor(CV_Green);
    if (pWnd->GetDlgCtrlID() == IDC_CHECK_OQCB5)      pDC->SetTextColor(CV_Blue);
    if (pWnd->GetDlgCtrlID() == IDC_CHECK_OQCD5)      pDC->SetTextColor(CV_Dark);
    if (pWnd->GetDlgCtrlID() == IDC_CHECK_OQCW49)       pDC->SetTextColor(CV_White);
    if (pWnd->GetDlgCtrlID() == IDC_CHECK_OQCD25)     pDC->SetTextColor(CV_Dark);


//    if (pWnd->GetDlgCtrlID() == IDC_COMBO_SELFLKPTM)  pDC->SetBkColor (CV_White);
    if (pWnd->GetDlgCtrlID() == IDC_COMBO_SELTAI)       pDC->SetBkColor (CV_White);
    if(AutoSelectMode(GetBackGroundColor()))//���O�G
        pDC->SetTextColor(CV_White);
    pDC->SetBkMode(TRANSPARENT);//�R�A��r�I���z��
/**/

    return m_Brush;
}
// The system calls this to obtain the cursor to display while the user drags
//  the minimized window.
HCURSOR CSwordDlg::OnQueryDragIcon()
{
    return (HCURSOR) m_hIcon;
}

// Automation servers should not exit when a user closes the UI
//  if a controller still holds on to one of its objects.  These
//  message handlers make sure that if the proxy is still in use,
//  then the UI is hidden but the dialog remains around if it
//  is dismissed.

void CSwordDlg::OnClose() {    if (CanExit())    CDialog::OnClose();   }
void CSwordDlg::OnOK()    {    if (CanExit())    CDialog::OnOK();      }
void CSwordDlg::OnCancel(){    if (CanExit())    CDialog::OnCancel(); }

BOOL CSwordDlg::CanExit() 
{
    // If the proxy object is still around, then the automation
    //  controller is still holding on to this application.  Leave
    //  the dialog around, but hide its UI.
    int temp = MessageBox("����������A�����q����ƱN���۰ʾ�z�����!!\n�T�w�q���F?","�A���b���}Sword8",MB_YESNO);
        //6: Yes  , 7: No
    if(temp == 7)
    {
        return FALSE;
    }
    else
    {
        if (m_pAutoProxy != NULL)
        {
            ShowWindow(SW_HIDE);
            return FALSE;
        }
        return TRUE;
    }
}

//��j����
void CSwordDlg::OnMax() 
{
    // TODO: Add your control notification handler code here
    int Symbol;
    if(GetMaxWindow())//�Y�^�쥻���j�p
    {  
        SetBackGroundColor(); //�w�]�I���C��
        SetMaxWindow(0);
        ShowWinNormal();//�Y�p�n��ܪ��F��
        HideWinNormal();//�Y�p�n�������F��
        Symbol = 1;
    }
    else//���̤j
    {  
        SetMaxWindow(1);
        HideWinMaximum();//��j�n�������F��
        ShowWinMaximum();//��j�n��ܪ��F��
        Symbol = 2;
    }
    SetDlgItemInt(IDC_MAX, Symbol);
    UpdateData(TRUE);
//    NineO(); //��ܤ����I��
}
////////////////////////////////////////////////////////////////////
//�٨S�Ψ쪺���
void CSwordDlg::OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags) 
{
    // TODO: Add your message handler code here and/or call default
    CDialog::OnKeyDown(nChar, nRepCnt, nFlags);
}

void CSwordDlg::OnKeyUp(UINT nChar, UINT nRepCnt, UINT nFlags) 
{
    // TODO: Add your message handler code here and/or call default
    CDialog::OnKeyUp(nChar, nRepCnt, nFlags);
}

void CSwordDlg::OnUpdateCacontrol1() 
{
    // TODO: Add your control notification handler code here
    ChooseCircleRadius();
    SetOQCSaveSel();
}

void CSwordDlg::OnOutofmemorySpin1(NMHDR* pNMHDR, LRESULT* pResult) 
{
    // TODO: Add your control notification handler code here
    *pResult = 0;
}

BEGIN_EVENTSINK_MAP(CSwordDlg, CDialog)
    //{{AFX_EVENTSINK_MAP(CSwordDlg)
    ON_EVENT(CSwordDlg, IDC_CACONTROL1, 1 /* Update */, OnUpdateCacontrol1, VTS_NONE)
    //}}AFX_EVENTSINK_MAP
END_EVENTSINK_MAP()
////////////////////////////////////////////////////////////////////////////////////////
//�H�U�ۤv�g���U���U�˪��Ƶ{��

////////////////////////////////////////////////////////////////
//�e�ϩM�C��
//�q���y�{��

//�e��
void CSwordDlg::DrawACircle(CPoint point,int radius, int Percent) 
{
/*
    �b���e�X�Ӫ���O�H160*160�j�p
    �H�����I���Ѧ��I�e�X�Ӫ�
    ���ƬO�H���W�����Ѧ��I�A�b���ץ��~�t�A�N�M�w�n���Ѧ��I
    �A�������W�U��@�ӥb�|���Z���A�e�X��
*/
    //�]�w����C��=�P�I���ۤ��C��
    COLORREF CircleColor;
//    CircleColor = CV_White-GetBackGroundColor();
    
        switch(Percent)
        {

            case 0:
            case 1:    CircleColor = RGB(255, 0 ,255);  break;//��
            case 2:
            case 3:
            case 4:    CircleColor = RGB(255,255, 64);  break;//��
            case 5:
            case 6:    CircleColor = RGB (0 ,255,128);  break;//��
/*            case 7:
            case 0:    CircleColor = RGB(255, 0 ,255);  break;//��
            case 1:    CircleColor = RGB(255,255, 64);  break;//��
            case 2:    CircleColor = RGB (0 ,255,128);  break;//��
*/            default:   
                if(GetBackGroundColor()+RGB(25,25,25) >= CV_White)
                    CircleColor = CV_Dark;
                else
                    CircleColor = GetBackGroundColor()+RGB(25,25,25);  
                break;
        //�e�����^�w�]��
        }
    if(CircleColor == GetBackGroundColor())  CircleColor = RGB(128,128,128);

    //�]�w�e��ؼЪ���}
    //CWnd* pWndGrid = GetDlgItem(IDC_COLOR_PATTERN);
    //CDC* pDC = pWndGrid->GetDC();
    CDC* pDC = new CClientDC(this);

    //���ץ��~�t�A�q���W�ը줤�ߡ]�}��Υiø�϶��^
    CPoint StartDrawPoint(-radius,-radius);
    CPoint EndDrawPoint(radius,radius);
    CRect* pRect = new CRect(point+StartDrawPoint,point+EndDrawPoint);
    //�e�U�h�]�e�U��ΡA�_�I0,0�A���I0,0�^
/*                                         //             (0,-1)
        CPoint P10 (- radius,  - radius);//75 (-1,-1)+----+----+(+1,-1)
        CPoint P20 (- radius,  +   0  );//          |�@  |  ��|
        CPoint P30 (- radius,  + radius);//00        |  �@|��  |
        CPoint P40 (+   0   ,  + radius);//   (-1, 0)+--point--+(+1, 0)
        CPoint P50 (+ radius,  + radius);//25        |  ��|�@  |
        CPoint P60 (+ radius,  +   0  );//          |��  |  �@|
        CPoint P70 (+ radius,  - radius);//50 (-1,+1)+----+----+(+1,+1)
        CPoint P80 (+   0   ,  - radius);//             (0,+1)

        switch(Percent)
        {
            case 0:                                                          pDC->Arc(pRect,point + P80, point + P80);  break;
            case 1:{   if(BKType) pDC->FillSolidRect(pRect,GetBackGroundColor()); pDC->Arc(pRect,point + P10, point + P80);  break;   }
            case 2:{   if(BKType) pDC->FillSolidRect(pRect,GetBackGroundColor()); pDC->Arc(pRect,point + P20, point + P80);  break;   }
            case 3:{   if(BKType) pDC->FillSolidRect(pRect,GetBackGroundColor()); pDC->Arc(pRect,point + P30, point + P80);  break;   }
            case 4:{   if(BKType) pDC->FillSolidRect(pRect,GetBackGroundColor()); pDC->Arc(pRect,point + P40, point + P80);  break;   }
            case 5:{   if(BKType) pDC->FillSolidRect(pRect,GetBackGroundColor()); pDC->Arc(pRect,point + P50, point + P80);  break;   }
            case 6:{   if(BKType) pDC->FillSolidRect(pRect,GetBackGroundColor()); pDC->Arc(pRect,point + P60, point + P80);  break;   }
            case 7:{   if(BKType) pDC->FillSolidRect(pRect,GetBackGroundColor()); pDC->Arc(pRect,point + P70, point + P80);  break;   }
            case 8:               pDC->FillSolidRect(pRect,GetBackGroundColor());                                            break;
            default:                                                                                                   break;
        //�e�����^�w�]��
        }*/
    //�]�w�e��
    CPen aPen;//,aPen0;
    //aPen0.CreatePen(PS_SOLID,5,RGB(127,127,127)+GetBackGroundColor());//�T�w���
    aPen.CreatePen(PS_SOLID,5,CircleColor);//�ܦ�

    CPen *pOldPen = pDC ->SelectObject(&aPen);//�]�w�Ȧs�e���Ŷ�

    //pDC->SetBkMode(TRANSPARENT);
    //pDC->SetBkColor(GetBackGroundColor());

    //pDC->MoveTo(point + StartDrawPoint);
    //pDC->Ellipse(pRect); 
    pDC->Arc(pRect,point + StartDrawPoint, point + StartDrawPoint);   
    pDC->SelectObject(pOldPen);
    
    delete pDC, pRect;
}
/*
void CSwordDlg::DrawSight(CPoint point,int radius, int Percent) 
{
    COLORREF CircleColor;
   
    switch(Percent)
    {
        case 0:    CircleColor = RGB(255, 0 ,255);  break;//��
        case 1:    CircleColor = RGB(255, 0 ,255);  break;//��
        case 2:    CircleColor = RGB(255,255, 64);  break;//��
        case 3:    CircleColor = RGB(255,255, 64);  break;//��
        case 4:    CircleColor = RGB (0 ,255,128);  break;//��
        case 5:    CircleColor = RGB (0 ,255,128);  break;//��
        default:   
            if(GetBackGroundColor()+RGB(25,25,25) >= CV_White)
                CircleColor = CV_Dark;
            else
                CircleColor = GetBackGroundColor()+RGB(25,25,25);  
            break;
    //�e�����^�w�]��
    }
    if(CircleColor == GetBackGroundColor())  CircleColor = RGB(128,128,128);

    //�]�w�e��ؼЪ���}
    CDC* pDC = new CClientDC(this);

    //���ץ��~�t�A�q���W�ը줤�ߡ]�}��Υiø�϶��^
    CPoint StartDrawPoint(-radius,-radius);
    CPoint EndDrawPoint(radius,radius);
    CRect* pRect = new CRect(point+StartDrawPoint,point+EndDrawPoint);
    //�]�w�e��
    CPen aPen;//,aPen0;

    aPen.CreatePen(PS_SOLID,5,CircleColor);//�ܦ�

    CPen *pOldPen = pDC ->SelectObject(&aPen);//�]�w�Ȧs�e���Ŷ�

    pDC->SetBkMode(TRANSPARENT);
    pDC->SetBkColor(RGB(rand()%255,rand()%255,rand()%255));

    pDC->MoveTo(point + StartDrawPoint);
    pDC->Ellipse(pRect); 
    //pDC->Arc(pRect,point + StartDrawPoint, point + StartDrawPoint);   
    pDC->SelectObject(pOldPen);
    
    delete pDC, pRect;
}
*/
//�e�w���ȼ��ҡ]�I���bctlcolor()�]�w���z���^
void CSwordDlg::DrawMsrLabel(CPoint point, int radius, double *Data)
{
/*  ����ưt�XDrawACircle��ƪ���m�A�]�i�H�W�ߨϥ�
    �b���e�X�Ӫ���O�H92*48�j�p
    str:Lv = LLL.vv
         x = 0.xxxx
         y = 0.yyyy*/

    CString Lv;    Lv.Format("%3.2f",*(Data));//Small Lv
    
    CString SX;    SX.Format("%1.4f",*(Data+1));//Small x
    CString SY;    SY.Format("%1.4f",*(Data+2));//Small y

    CString T;      T.Format("%4.0f",*(Data+3));//T
    CString Duv;  Duv.Format("%1.4f",*(Data+4));//duv

    CString Du;    Du.Format("%1.4f",*(Data+5));//du
    CString Dv;    Dv.Format("%1.4f",*(Data+6));//dv

    CString X;     X.Format("%1.4f",*(Data+7));//X
    CString Y;     Y.Format("%1.4f",*(Data+8));//Y
    CString Z;     Z.Format("%1.4f",*(Data+9));//Z

    CString Flk;   Flk.Format("%3.2f",*(Data+10));//Z

    CString str;//��X�r��

    str = "";//�w���Ȫ���X�榡
        if(m_bool_Lv)   str = str + " Lv ="+Lv+"\n";
        
        if(m_bool_x)    str = str + "  x ="+SX+"\n";
        if(m_bool_y)    str = str + "  y ="+SY+"\n";

        if(m_bool_T)    str = str + "  T ="+T+"\n";
        if(m_bool_duv)  str = str + "�Guv="+Duv+"\n";

        if(m_bool_du)   str = str + "�Gu ="+Du+"\n";
        if(m_bool_dv)   str = str + "�Gv ="+Dv+"\n";

        if(m_bool_X)    str = str + " X ="+X+"\n";
        if(m_bool_Y)    str = str + " Y ="+Y+"\n";
        if(m_bool_Z)    str = str + " Z ="+Z+"\n";
//        if(m_bool_flk)  str = str + " Flicker ="+Flk+"\n";

    //�]�w�I
    int hight;

    hight = m_bool_Lv+m_bool_x+m_bool_y+m_bool_T+m_bool_duv+m_bool_du+m_bool_dv+
        m_bool_X+m_bool_Y+m_bool_Z;

    CSize ShiftSize(-radius,-radius);                //���w�I���r�ϰ쥪�W�����Z��
    CSize  RectSize(120,hight*16);                    //��r�ϰ�j�p80
    CPoint StartDrawPoint = point+ShiftSize;          //��������A�g�r�϶��}�l�I
    CPoint EndDrawPoint = RectSize+StartDrawPoint;    //�g�r�϶������I
    CRect rect(StartDrawPoint,EndDrawPoint);          //�]�w��r�϶�

    CDC* pDC = new CClientDC(this);

    pDC->SetBkMode(TRANSPARENT);
    pDC->SetTextColor(RGB(127,127,127)+GetBackGroundColor());
    pDC->DrawText(TEXT(str),-1, &rect, DT_LEFT | DT_VCENTER);

    delete pDC;
}

BOOL CSwordDlg::SetMsrPoint(CPoint MsrPoint, BOOL Mode)
{
    CurrentPoint = MsrPoint;
    if(Mode)
    {
        SetCtlPointQuadrant(MsrPoint);
        return 1;//�������I�^��1
    }
    else
        return 0;//�L�����I�^��0
}

const CPoint CSwordDlg::GetMsrPoint(int PointFlag)
{

    switch(PointFlag)
    {
        case 0:        return CurrentPoint;    break;
        case 1:        return StopPoint;       break;
        case 2:        return BackPoint;       break;
        default:       return CurrentPoint;    break;
    }
}

//�]�w�I����
void CSwordDlg::SetBackGroundColor(int ColorFlag , COLORREF BKDefaultColor)
{
    switch(ColorFlag)
    {
        case BK_White:    BackGroundColor = CV_White;        break;//W
        case BK_Red:      BackGroundColor = CV_Red;            break;//R
        case BK_Green:    BackGroundColor = CV_Green;        break;//G
        case BK_Blue:     BackGroundColor = CV_Blue;        break;//B
        case BK_Dark:     BackGroundColor = CV_Dark;        break;//D
        case BK_Other:    BackGroundColor = BKDefaultColor;    break;//Flag=5�A�⥦
        default:          BackGroundColor = BKDefaultColor;    break;//Flag>5�A�⥦
    }
    m_Brush.DeleteObject(); 
    m_Brush.CreateSolidBrush(BackGroundColor);
    Invalidate(TRUE);
}

const COLORREF CSwordDlg::GetBackGroundColor()
{
    return BackGroundColor;
}

//�κu������Pattern�j�p
BOOL CSwordDlg::OnMouseWheel(UINT nFlags, short zDelta, CPoint pt) 
{
    // TODO: Add your message handler code here and/or call default
    int Symbol;
    if(zDelta < 0)//�ԴN�Y�^�쥻���j�p
    {  
        if(GetMaxWindow())
            SetBackGroundColor(); //�w�]�I���C��
        SetMaxWindow(0);
        ShowWinNormal();//�Y�p�n��ܪ��F��
        HideWinNormal();//�Y�p�n�������F��
        Symbol = 1;
    }
    if(zDelta > 0)//���N���̤j
    {  
        SetMaxWindow(1);
        HideWinMaximum();//��j�n�������F��
        ShowWinMaximum();//��j�n��ܪ��F��
        Symbol = 2;
    }    
    this->SetDlgItemInt(IDC_MAX,Symbol);
    UpdateData(TRUE);
    return CDialog::OnMouseWheel(nFlags, zDelta, pt);
}

const BOOL CSwordDlg::GetMaxWindow()
{
    return MaxWindow;
}

void CSwordDlg::SetMaxWindow(bool flag)
{
    //ChooseCircleRadius();
    if(flag)    ShowWindow(SW_MAXIMIZE);
    else        ShowWindow(SW_RESTORE);
    //ShowWindow(SW_SHOW);
    MaxWindow = flag;
}

void CSwordDlg::OnMouseMove(UINT nFlags, CPoint point) 
{
    // TODO: Add your message handler code here and/or call default
    if(!GetMaxWindow())
    {
        RECT rect;
        ::GetWindowRect(m_hWnd,&rect);    //���o�����쥻����m
        //�ǰe�T���A�������}�l�H�۷ƹ����ʡA�����}�ƹ�����
        ::PostMessage(
            GetSafeHwnd(),
           WM_NCLBUTTONDOWN,
                  HTCAPTION,
           MAKELPARAM (rect.right - rect.left, rect.bottom - rect.top));
    }
    CDialog::OnMouseMove(nFlags, point);
}



void CSwordDlg::OnLButtonUp(UINT nFlags, CPoint point) 
{
    // TODO: Add your message handler code here and/or call default
    if(GetMaxWindow())
    {
        UpdateWindow();
        UpdateData(TRUE);
        if(DiaplayCircle == 0)//�S���
        {
            for(int i=0;i<9;++i)//�e�X9�I
            {
                CurrentPoint = Set9Point(i,m_floatFromEdge);
                DrawACircle(CurrentPoint, CircleRadius,6);
            }
            ++DiaplayCircle;
        }
        else if(DiaplayCircle == 1)
        {
            for(int i=0;i<49;++i)//�e�X49�I
            {
                CurrentPoint = Set49Point(i);
                DrawACircle(CurrentPoint, CircleRadius,6);
            }
            ++DiaplayCircle;    
        }
        else//����ɡA�|�������
        {
            Invalidate();
            DiaplayCircle = 0;
        }
    }
/*
    m_ICa.Measure(0); 
    DrawMsrLabel(point);

    //�I����A����
    //�]�w�e��ؼЪ���}
    CWnd* pWndGrid = GetDlgItem(IDC_COLOR_PATTERN);
    //CWnd* pWndGrid = GetDlgItem(IDD_PATTERN);
    CDC* pDC = pWndGrid->GetDC();
    //�]�w�e��
    CPen aPen;
    aPen.CreatePen(PS_INSIDEFRAME,5,CV_Dark);
    //�]�w�Ȧs�e���Ŷ�
    CPen* pOldPen = pDC ->SelectObject(&aPen);
    //���ץ��~�t�A�q���W�ը줤��
    CPoint OnePoint(80,80);
    CPoint TwoPoint(-80,-80);
    CRect* pRect = new CRect(point+TwoPoint,point+OnePoint);
    //�e�U�h
    CPoint Start(0,0);
    CPoint End(0,0);
    pDC->Arc(pRect,Start,End);
    //�e�����^�w�]��
    pDC->SelectObject(pOldPen);
*/
    CDialog::OnLButtonUp(nFlags, point);
}

//�]�w�����I�H����m
CPoint CSwordDlg::SetCtlPointQuadrant(CPoint point)
{
    int x,y;

         if((point.x  < ScrmH/2-1)&&(point.y  < ScrmV/2-1))  {  x = ScrmH/2; y = ScrmV/2;  }//�Ĥ@�H��
    else if((point.x  > ScrmH/2-1)&&(point.y  < ScrmV/2-1))  {  x =       0; y = ScrmV/2;  }//�ĤG�H��
    else if((point.x  < ScrmH/2-1)&&(point.y  > ScrmV/2-1))  {  x = ScrmH/2; y =       0;  }//�ĤT�H��
    else if((point.x  > ScrmH/2-1)&&(point.y  > ScrmV/2-1))  {  x =       0; y =       0;  }//�ĥ|�H��
    else if((point.x == ScrmH/2-1)||(point.y == ScrmV/2-1))  {  x =       0; y =       0;  }//�b�Q�r�u�W
    else                                                     {  x =   ScrmH; y =   ScrmV;  }

    CPoint CtlPoint(x,y);
    //CtlPoint.x = x;
    //CtlPoint.y = y;
    return CtlPoint;
}

//�e�X������
void CSwordDlg::DrawColorLabel(CPoint StartPoint, int radius) 
{
    CDC* pDC = new CClientDC(this);

    CPoint point1 = StartPoint;
    CPoint point2(point1.x + radius*2,point1.y);

    StopPoint.x = point1.x + radius;
    StopPoint.y = point1.y + radius;

    BackPoint.x = StopPoint.x + radius*2;
    BackPoint.y = StopPoint.y;

    CPoint EndDrawPoint(radius*2,radius*2);

    CRect* pRect1 = new CRect(point1,point1+EndDrawPoint);
    CRect* pRect2 = new CRect(point2,point2+EndDrawPoint);

    //�w�q���������C��
    pDC->FillSolidRect(pRect1,GetStopColor());
    pDC->FillSolidRect(pRect2,GetBackColor());
//      StartPoint--+------+         StartPoint��
//           |   �� |      |               �� �@�@��
//           +----point----+              �� �@    ��
//           |      |��    |               �� �@�@��
//           +------+-EndPoint                ��
    pDC->SetTextAlign(TA_CENTER | TA_BASELINE);
    pDC->SetTextColor(CV_Dark);
    pDC->SetBkMode(TRANSPARENT);
    pDC->TextOut(point1.x+radius,point1.y+radius,"����q��");
    pDC->TextOut(point2.x+radius,point2.y+radius,"�q���W�@��");

    delete pDC,pRect1,pRect2;
}
//---------------
//�M�w���������C��
const COLORREF CSwordDlg::GetStopColor()
{
    COLORREF StopColor;
         if(GetBackGroundColor() == CV_White)    StopColor = CV_Red;
    else if(GetBackGroundColor() == CV_Red )     StopColor = CV_Green;
    else if(GetBackGroundColor() == CV_Green)    StopColor = CV_Blue;
    else if(GetBackGroundColor() == CV_Blue)     StopColor = CV_Red;
    else if(GetBackGroundColor() == CV_Dark)     StopColor = CV_Green;
    else if(GetBackGroundColor() & 0x11000000 != 0x00000000)     StopColor = CV_Blue;
    else  StopColor = RGB(128,128,0);//����,R,G,B
    
    return StopColor;
}

const COLORREF CSwordDlg::GetBackColor()
{
    COLORREF BackColor;
         if(GetBackGroundColor() == CV_White)    BackColor = CV_Green;
    else if(GetBackGroundColor() == CV_Red )     BackColor = CV_Blue;
    else if(GetBackGroundColor() == CV_Green)    BackColor = CV_Red;
    else if(GetBackGroundColor() == CV_Blue)     BackColor = CV_Green;
    else if(GetBackGroundColor() == CV_Dark)     BackColor = CV_Blue;
    else if(GetBackGroundColor() & 0x11000000 != 0x00000000)     BackColor = CV_Red;
    else  BackColor = RGB(128,0,128);//����

    return BackColor;
}
//---------------
//���������P�_
BOOL CSwordDlg::PassStopColor()
{
         if(GetBackGroundColor() == CV_White)    
         {
            if ((double)m_IProbe.GetSx() > 0.5  &&
                (double)m_IProbe.GetSy() > 0.25 &&    (double)m_IProbe.GetSy() < 0.4  &&
                (double)m_IProbe.GetLv() > 10.5)
                return TRUE;//CV_Red
            else
                 return FALSE;
         }
    else if(GetBackGroundColor() == CV_Red ) 
    {
            if ((double)m_IProbe.GetSx() < 0.35 &&
                (double)m_IProbe.GetSy() > 0.55 &&
                (double)m_IProbe.GetLv() > 10.5)
                return TRUE;//CV_Green
            else
                return FALSE;
    }
    else if(GetBackGroundColor() == CV_Green)
    {
            if ((double)m_IProbe.GetSx() < 0.25 &&
                (double)m_IProbe.GetSy() < 0.3  &&
                (double)m_IProbe.GetLv() > 10.5)
                return TRUE;//CV_Blue
            else
                return FALSE;
    }
    else if(GetBackGroundColor() == CV_Blue) 
    {
            if ((double)m_IProbe.GetSx() > 0.5  &&
                (double)m_IProbe.GetSy() > 0.25 &&    (double)m_IProbe.GetSy() < 0.4  &&
                (double)m_IProbe.GetLv() > 10.5)
                return TRUE;//CV_Red
            else
                return FALSE;
    }
    else if(GetBackGroundColor() == CV_Dark) 
    {
            if ((double)m_IProbe.GetSx() < 0.35 &&
                (double)m_IProbe.GetSy() > 0.55 &&
                (double)m_IProbe.GetLv() > 10.5)
                return TRUE;//CV_Green
            else
                return FALSE;
    }
    else if(GetBackGroundColor() & 0x11000000 != 0x00000000) 
    {
            if ((double)m_IProbe.GetSx() < 0.25 &&
                (double)m_IProbe.GetSy() < 0.3  &&
                (double)m_IProbe.GetLv() > 10.5)
                return TRUE;//CV_Blue
            else
                return FALSE;
    }
    else
    {
            if ((double)m_IProbe.GetSx() > 0.35  &&    (double)m_IProbe.GetSx() < 0.5  &&
                (double)m_IProbe.GetSy() > 0.4   &&    (double)m_IProbe.GetSx() < 0.54 &&
                (double)m_IProbe.GetLv() > 10.5)
                return TRUE;//����
            else
                return FALSE;
    }
}

BOOL CSwordDlg::PassBackColor()
{
         if(GetBackGroundColor() == CV_White)
         {
            if ((double)m_IProbe.GetSx() < 0.35 &&
                (double)m_IProbe.GetSy() > 0.55 &&
                (double)m_IProbe.GetLv() > 10.5)
                return TRUE;//CV_Green
            else
                return FALSE;
         }
    else if(GetBackGroundColor() == CV_Red ) 
    {
            if ((double)m_IProbe.GetSx() < 0.25 &&
                (double)m_IProbe.GetSy() < 0.3  &&
                (double)m_IProbe.GetLv() > 10.5)
                return TRUE;//CV_Blue
            else
                return FALSE;
    }
    else if(GetBackGroundColor() == CV_Green)
    {
            if ((double)m_IProbe.GetSx() > 0.5  &&
                (double)m_IProbe.GetSy() > 0.25 &&    (double)m_IProbe.GetSy() < 0.4  &&
                (double)m_IProbe.GetLv() > 10.5)
                return TRUE;//CV_Red
            else
                return FALSE;
    }
    else if(GetBackGroundColor() == CV_Blue) 
    {
            if ((double)m_IProbe.GetSx() < 0.35 &&
                (double)m_IProbe.GetSy() > 0.55 &&
                (double)m_IProbe.GetLv() > 10.5)
                return TRUE;//CV_Green
            else
                return FALSE;
    }
    else if(GetBackGroundColor() == CV_Dark) 
    {
            if ((double)m_IProbe.GetSx() < 0.25 &&
                (double)m_IProbe.GetSy() < 0.3  &&
                (double)m_IProbe.GetLv() > 10.5)
                return TRUE;//CV_Blue
            else
                return FALSE;
    }
    else if(GetBackGroundColor() & 0x11000000 != 0x00000000) 
    {
            if ((double)m_IProbe.GetSx() > 0.5  &&
                (double)m_IProbe.GetSy() > 0.25 &&    (double)m_IProbe.GetSy() < 0.4  &&
                (double)m_IProbe.GetLv() > 10.5)
                return TRUE;//CV_Red
            else
                return FALSE;
    }
    else  
    {
            if ((double)m_IProbe.GetSx() >=0.25  &&    (double)m_IProbe.GetSx() < 0.5  &&
                (double)m_IProbe.GetSy() < 0.23  &&//    (double)m_IProbe.GetSx() < 0.5  &&
                (double)m_IProbe.GetLv() > 10.5)
                return TRUE;//����
            else
                return FALSE;
    }
}


//�����G��ηt��]�M�w�t��k�Ҧ��^
//�^�ǭ�
//    0:�G�I��
//    1:�t�I��
//    2:flk�I��
int CSwordDlg::AutoSelectMode(COLORREF color)
{
    int O = 0x000000FF & (color >>24);
    int B = 0x000000FF & (color >>16);
    int G = 0x000000FF & (color >> 8);
    int R = 0x000000FF & (color >> 0);

    //m_strTestValue.Format("AutoSelectMode ��I��");

       if(O!=0)                         return 2; //flk
    else if((R<25)&&(G<25)&&(B<25))        return 1; //�t
    //if((R<25)&&(G<25)&&(B<25))    return 1; //�t
    else                                return 0; //�G

}
//---------------
//Smart Measure ����
BOOL CSwordDlg::Pass(int PassMode, BOOL CtlMode)
{
// PassMode:  0=�k�s     1=+1     �䥦=�K��
// CtlMode:   0=�L����� 1=�������
    switch(PassMode)
    {
        //case -1: 
            //MsrPass = MsrPass; 
            //break;
        case  0: 
            //m_strTestValue.Format("Pass �k0");
            //UpdateData(FALSE);
            MsrPass = 0;
                       DrawACircle(CurrentPoint, CircleRadius, 0);//�e��
            if(CtlMode)DrawACircle(StopPoint,    CircleRadius, 0);//�e��
            if(CtlMode)DrawACircle(BackPoint,    CircleRadius, 0);//�e��
            break;
        case  1: 
            //m_strTestValue.Format("Pass +1");
            //UpdateData(FALSE);
             ++MsrPass;
           //�w�q���������C��
                 if(PassStopColor() && CtlMode)    DrawACircle(StopPoint,    CircleRadius, MsrPass);//�e��//STOP
            else if(PassBackColor() && CtlMode)    DrawACircle(BackPoint,    CircleRadius, MsrPass);//�e��//BACK    
            else                                   DrawACircle(CurrentPoint, CircleRadius, MsrPass);//�e��
            break;
        default:  
            break;
    }

    Sleep(10);
    //Sleep(m_int_5nitsGray);
    
    if(MsrPass < 6) //MsrPass < 3
        return TRUE;
    else                
        return FALSE;
}

//Smart Measure
BOOL CSwordDlg::MsrAI(int WorkMode, BOOL CtlMode)
// WorkMode:  �I�����ܼҦ�
//    0:�G�I��
//    1:�t�I��
//    2:flk�I��
// CtlMode:   0=�L����� 1=�������
{

    Pass(0, CtlMode);

        float XFristValue = 0.0, SxFristValue = 0.0,
              YFristValue = 0.0, SyFristValue = 0.0,
              ZFristValue = 0.0, LvFristValue = 0.0,
          //�ŧi�~�t�ȭp��Ŷ�
                   deltaX = 0.0, deltaSx = 0.0,
                   deltaY = 0.0, deltaSy = 0.0,
                   deltaZ = 0.0, deltaLv = 0.0,

                 FlkValue = 0.0, deltaFlk = 0.0; 
    do
    {
        int flag = 0;
        do 
        {
            try
            {
                m_ICa.Measure(0);
                flag = 0;
            }
            catch (CException* e)
            {
                TCHAR buf[255];
                e->GetErrorMessage(buf, 255);
                flag++;
            }
        } while (flag);

        if((double)m_IProbe.GetLv() < 0.009)
            MessageBox("�]�q��0-Cal�q���Ȱ��I");


/*        if(WorkMode == 0)//�G�I��
        {
//            Sleep(200);//�º������C�@�I

            //m_strTestValue.Format("MsrAI �G�I��");
            //UpdateData(FALSE);
            deltaSx = deltaSy = deltaLv = 0.0;
            //��Ѧҭ�
            LvFristValue = m_IProbe.GetLv();  
            SxFristValue = m_IProbe.GetSx();  
            SyFristValue = m_IProbe.GetSy();
    
            m_ICa.Measure(0);
            //�~�t�������
            deltaLv = ((LvFristValue-m_IProbe.GetLv())>=0) ? LvFristValue-m_IProbe.GetLv() : m_IProbe.GetLv()-LvFristValue;
            deltaSx = ((SxFristValue-m_IProbe.GetSx())>=0) ? LvFristValue-m_IProbe.GetLv() : m_IProbe.GetSx()-SxFristValue;
            deltaSy = ((SyFristValue-m_IProbe.GetSy())>=0) ? SyFristValue-m_IProbe.GetSy() : m_IProbe.GetSy()-SyFristValue;

            /*    ���e�Ƚվ�O��   if((deltaSx*deltaSy*deltaLv)<0.00000000001)
            Sx = 1.4f, deltaSx<0.00001
            Sy = 1.4f, deltaSy<0.00001
            Lv = 3.2f, deltaLv<0.001
                                            0.000000000001 �̦��]�w�ȡAPass�֭p�����A���Y�V�����P
                                            0.000000000005 �ĤG���]�w�ȡA�bPass���֭p�����AOK
                                                            �ĤG���]�w�ȡA�bPass���s�򤭦��A���Y�V�����P
                                            0.00000000001  �ĤT���]�w�ȡA�bPass���s�򤭦��ɡAOK
                                            0.00000000005  ����i������
                                            0.00000000025
                                            0.000000000000889 ���礧�᪺�����]AGING�^
                                            0.0000000000029 �̤j�ȡ]�x�^
                                            0.0000000000003 �����ȡ]�x�^
                                            0.0000000001 �]�������^
            */


            //if(WorkMode!=2)
            //{
            //CString temp;
            //temp.Format(" x=%1.4f \n y=%1.4f \n Lv=%4.2f \n total=%f",deltaSx, deltaSy, deltaLv, deltaSx*deltaSy*deltaLv);
            //MessageBox(temp);
//                if((deltaSx*deltaSy*deltaLv) <  0.00000000045)    Pass(1, CtlMode);
//                else                                             Pass(0, CtlMode);
            //}
            //else
            /*{
                                ���e�Ƚվ�O��   if((deltaSx*deltaSy*deltaLv)<0.00000000001)
            Sx = 1.4f, deltaSx<0.00001
            Sy = 1.4f, deltaSy<0.00001
            Lv = 3.2f, deltaLv<0.001
                                                0.0001        �̦��]�w�ȡA���P
                                                 0.00000001    ���I��
            
                //if((deltaSx*deltaSy*deltaLv) <  0.00000005)    Pass(1, CtlMode);
                //else                                             Pass(0, CtlMode);
            }*/
//        }
//        else 
        if(WorkMode == 1 || WorkMode == 0)//�t�I�� �� �䥦
        {
            //m_strTestValue.Format("MsrAI �t�I��");
            //UpdateData(FALSE);
            deltaX = deltaY = deltaZ = 0.0;
            //��Ѧҭ�
            XFristValue = m_IProbe.GetX();  
            YFristValue = m_IProbe.GetY();  
            ZFristValue = m_IProbe.GetZ();
    
            int flag = 0;
            do 
            {
                try
                {
                    m_ICa.Measure(0);
                    flag = 0;
                }
                catch (CException* e)
                {
                    TCHAR buf[255];
                    e->GetErrorMessage(buf, 255);
                    flag++;
                }
            } while (flag);
            //�~�t�������
            deltaX = ((XFristValue-m_IProbe.GetX())>=0) ? XFristValue-m_IProbe.GetX() : m_IProbe.GetX() - XFristValue;
            deltaY = ((YFristValue-m_IProbe.GetY())>=0) ? YFristValue-m_IProbe.GetY() : m_IProbe.GetY() - YFristValue;
            deltaZ = ((ZFristValue-m_IProbe.GetZ())>=0) ? ZFristValue-m_IProbe.GetZ() : m_IProbe.GetZ() - ZFristValue;

            /*    ���e�Ƚվ�O��   
                                    0.001
                                    0.05  ���P�A�q���t�׫ܧ�
                                    0.045 ���P�A�q���t�ק�
                                    0.425 �Ӻ�A�q���|�ܤ��}��
                                    0.435 �٦n�A���O���o���I�[
                                    0.443 ��~�ѹ껡�A�ٺ��֪�
                                    0.415 ����e�ϥΪ���
                                    0.03  �T�}�[���絲�G
                X = 
                Y = 
                Z = 
                0.00_00_00_0001  
                0.00_00_00_00005
                0.00_00_00_1  �hLMO�ɧ諸�]�٤����A�C�F�I�^
                0.005
            */
            if((deltaX*deltaY*deltaZ) <  Brain)    Pass(1, CtlMode);
            else                                   Pass(0, CtlMode);
        }
        else if(WorkMode == 2)//flk�I��
        {
//            Sleep(200);//�º������C�@�I

            //m_strTestValue.Format("MsrAI flk�I��");
            //UpdateData(FALSE);
            deltaFlk = 0.0;
        //��Ѧҭ�
            FlkValue = m_IProbe.GetFlckrFMA();
            int flag = 0;
            do 
            {
                try
                {
                    m_ICa.Measure(0);
                    flag = 0;
                }
                catch (CException* e)
                {
                    TCHAR buf[255];
                    e->GetErrorMessage(buf, 255);
                    flag++;
                }
            } while (flag);
        //�~�t�������
            deltaFlk = ((FlkValue-m_IProbe.GetFlckrFMA())>=0) ? FlkValue-m_IProbe.GetFlckrFMA() : m_IProbe.GetFlckrFMA() - FlkValue;

            //if((FlkValue-m_IProbe.GetFlckrFMA())>=0) deltaFlk = FlkValue-m_IProbe.GetFlckrFMA();
            //else                                     deltaFlk = m_IProbe.GetFlckrFMA() - FlkValue;
        //    ���e�Ƚվ�O��   
        //                        0.01 �򥻨S��
        //                        0.1  �Ӻ�  
            if(deltaFlk <  0.9 && (double)m_IProbe.GetFlckrFMA() > 2.0)    Pass(1, CtlMode);
            else                                                           Pass(0, CtlMode);
        }
        else
            return FALSE;
        
    }
    while(Pass(-1, CtlMode));
    return TRUE;
}

BOOL CSwordDlg::MsrAct(BOOL CtlMode)
{
    DrawACircle(CurrentPoint, CircleRadius, 0);
    if(CtlMode) DrawACircle(StopPoint,    CircleRadius, 0);//�e��//STOP
    if(CtlMode) DrawACircle(BackPoint,    CircleRadius, 0);//�e��//BACK
    CEnterDlg *aDlg = new CEnterDlg;
    int Act;
    do{
        Act = 0;


        aDlg->DoModal();

        m_ICa.Measure(0);
        if((double)m_IProbe.GetLv() < 0.009)
        {
            MessageBox("�]�q��0-Cal�q���Ȱ��I");
            Act = 1;
        }
    }while(Act);
    delete(aDlg);
    
    m_ICa.Measure(0);
    return TRUE;
}

//�w�q���j�p�]��LCM�ؤo�^
void CSwordDlg::ChooseCircleRadius()
{
//CircleRadius
/*
    int mmX,mmY;
    int pixelX,pixelY;
    int RadiusX,RadiusY;
    int mm10X,mm10Y;

    CString strr;

    CDC* pDC = new CClientDC(this);
    mmX    = pDC->GetDeviceCaps(HORZSIZE);
    pixelX = pDC->GetDeviceCaps(HORZRES);

    mmY    = pDC->GetDeviceCaps(VERTSIZE);
    pixelY = pDC->GetDeviceCaps(VERTRES);

    RadiusX = (int)((float)((float)pixelX*(float)RADIUS)/(float)(mmX));
    RadiusY = (int)((float)((float)pixelY*(float)RADIUS)/(float)(mmY));

    CircleRadius = (RadiusX+RadiusY)/2;

    mm10X = (int)((float)((float)pixelX*(float)100)/(float)(mmX));
    mm10Y = (int)((float)((float)pixelY*(float)100)/(float)(mmY));

    SetPixelPitch((int)((mm10X+mm10Y)/2)/1.414);
    delete pDC;
*/
//CircleRadius
    ScrmH = GetSystemMetrics(SM_CXSCREEN);
    ScrmV = GetSystemMetrics(SM_CYSCREEN);

    iLCMSize = atoi(m_IMemory.GetChannelID().Left(2));

    if(ScrmV >= 1000)//FHD
    {
        switch(iLCMSize)
        {    //                            ���٨S��Check

            case 55: CircleRadius = 41; break;//55F11 
            case 50: CircleRadius = 43; break;//50
            case 46: CircleRadius = 46; break;//46F11/F61
            case 40: CircleRadius = 53; break;//40F11
            case 37: CircleRadius = 56; break;//37F11
            case 31: 
            case 32: CircleRadius = 64; break;//315F11/F61

            default: CircleRadius = 83; break;//�H24�T����ù�����¦�]�}�o�̨ϥΪ��ù��^
/*
            case 55: CircleRadius = 41; SetPixelPitch(146); SetHalf45mm(1); break;//55F11 
            case 46: CircleRadius = 46; SetPixelPitch(133); SetHalf45mm(1); break;//46F11/F61
            case 40: CircleRadius = 53; SetPixelPitch(150); SetHalf45mm(1); break;//40F11
            case 37: CircleRadius = 56; SetPixelPitch(166); SetHalf45mm(1); break;//37F11
            case 31: 
            case 32: CircleRadius = 64; SetPixelPitch(194); SetHalf45mm(1); break;//315F11/F61

            default: CircleRadius = 83; SetPixelPitch(261); SetHalf45mm(1); break;//�H24�T����ù�����¦�]�}�o�̨ϥΪ��ù��^
*/
        }
    }
    else//HD
    {
        switch(iLCMSize)
        {    //                            ���٨S��Check

            case 31: 
            case 32: CircleRadius = 47; break;//315H61
            case 26: CircleRadius = 58; break;//315H61
            case 23: CircleRadius = 58; break;//26H61
            case 21: CircleRadius = 67; break;//216H61

            default: CircleRadius = 90; break;//�H13�T���q�ù�����¦�]�}�o�̨ϥΪ��ù��^
/*
            case 31: 
            case 32: CircleRadius = 47; SetPixelPitch(277); SetHalf45mm(1); break;//315H61
            case 26: CircleRadius = 58; SetPixelPitch(167); SetHalf45mm(1); break;//315H61
            case 23: CircleRadius = 58; SetPixelPitch(277); SetHalf45mm(1); break;//26H61
            case 21: CircleRadius = 67; SetPixelPitch(277); SetHalf45mm(1); break;//216H61

            default: CircleRadius = 90; SetPixelPitch(261); SetHalf45mm(1); break;//�H13�T���q�ù�����¦�]�}�o�̨ϥΪ��ù��^
*/
        }
    }    
/*
     Check ���p
    +----+------------+------------+
    |inch|CircleRadius| PixelPitch |
    +-FHD+------------+------------+
    |  55|            |            |
    |  46|20110316 CLT|20110316 CLT|
    |  47|20110316 CLT|20110316 CLT|
    |  40|            |            |
    |  37|            |            |
    |  32|20110118 LMO|            |
    |  31|            |            |
    |  24|�}�o��PC CLT|�}�o��PC CLT|
    +--HD+------------+------------+
    |  32|20110118 LMO|            |
    |  31|20110118 LMO|            |
    |  26|20110316 CLT|20110316 CLT|
    |  23|            |            |
    |  21|            |            |
    |  13|�}�o��NB CLT|�}�o��NB CLT|
    +----+------------+------------+
*/
}

int CSwordDlg::CmtoPixel(double cm)
{
    double hight = ScrmV;
    double width = ScrmH;
    double tzda = hight/width;
    double Vpixel = iLCMSize * sin (atan(tzda))  * 2.54;
    //cout << "�e:" << iLCMSize * cos (atan(tzda))  * 2.54;

    double Pixel = ScrmV*cm/Vpixel;

    return (int)Pixel;
}

void CSwordDlg::SetOQCSaveSel()    
{
    pOQCFormSel->ResetContent();

    char *pbuff1;
    char *pbuff2;

//---------------------------------------
    WIN32_FIND_DATA FindFileData;
    HANDLE hListFile;
    CHAR szFilePath[MAX_PATH];

    lstrcpy(szFilePath,szCurrentDirectory);        //�ثe�����ɩҦb���|
    lstrcat(szFilePath, "\\*.xls");    //�ثe�����ɩҦb���|\*.xls

    hListFile = FindFirstFile(szFilePath, &FindFileData);

    CString temp("");
    CString tempbuff("");

    if(hListFile == INVALID_HANDLE_VALUE)
    {
        temp.Format("���~%d", GetLastError());
        MessageBox(temp);
    }
    else
    {
        do{
            temp.Format("%s\n",  FindFileData.cFileName);
            tempbuff += temp;
        }while(FindNextFile(hListFile, &FindFileData));
    }

    pbuff1=tempbuff.GetBuffer(0);
//---------------------------------------
    pbuff2 = strtok (pbuff1,"\t\n");  //pbuff1 ���� pbuff2

    CString CutStr;

    CutStr.Format("SEC Form");
    pOQCFormSel->AddString(CutStr);
    CutStr.Format("WRGBD5�I�{�ɪ��");
    pOQCFormSel->AddString(CutStr);

    do{
        CutStr.Format("%s",pbuff2);            //�r����Ȧs
        //if(iLCMSize == atoi(x.Left(2)))        //��Ʀr�I�U�ӡA�çP�_���O���O�ŦX�n��
            pOQCFormSel->AddString(CutStr);        //�N�[�J

        pbuff2 = strtok (NULL, "\t\n");  //�U�@��
    }while (pbuff2 != NULL);

    pOQCFormSel->SetCurSel(0);
}

//�I������ҿ���
const int CSwordDlg::SelectColorLabel(COLORREF BKColor)
{
    switch(BKColor)
    {
        case CV_White:    return 0; break;//W
        case CV_Red:      return 1; break;//R
        case CV_Green:    return 2; break;//G
        case CV_Blue:     return 3; break;//B
        case CV_Dark:     return 4; break;//D
        default:          return 5; break;
    }
}

void CSwordDlg::OnButtonRst() 
{
    // TODO: Add your control notification handler code here
    //��l�Ƹ�ƼȦs�}�C
    EnabledReportBtm(0);
    initValues();
}

void CSwordDlg::OnChange5nitsGray() 
{
    // TODO: If this is a RICHEDIT control, the control will not
    // send this notification unless you override the CDialog::OnInitDialog()
    // function and call CRichEditCtrl().SetEventMask()
    // with the ENM_CHANGE flag ORed into the mask.
    
    // TODO: Add your control notification handler code here
    UpdateData(TRUE);
/*  
    SetBackGroundColor(BK_Other,RGB(m_int_5nitsGray,m_int_5nitsGray,m_int_5nitsGray));
    m_Brush.DeleteObject(); 
    m_Brush.CreateSolidBrush(GetBackGroundColor());
    Invalidate(TRUE);
*/
}

void CSwordDlg::OnButtonMsrflow() 
{
    if(pOQCTaiSel->GetCurSel() != 10)
    {
        // TODO: Add your control notification handler code here
        OnMax();
        HideWinControlItem();//���å��W������Ӥp���
    
        //flk
        //���
        if(m_bool_oqc_w){    OnWhite();     Msr1P();   }    //���⤤���I
        if(m_bool_oqc_r){    OnRed();       Msr1P();   }    //���⤤���I
        if(m_bool_oqc_g){    OnGreen();     Msr1P();   }    //��⤤���I
        if(m_bool_oqc_b){    OnBlue();      Msr1P();   }    //�Ŧ⤤���I
        if(m_bool_oqc_d){    OnBlack();     Msr1P();   }    //���⤤���I

        //if(pOQC_Flicker_Sel->GetCurSel() < 5){switch(pOQC_Flicker_Sel->GetCurSel()){case 0:  CSwordDlg::OnFlkSpix(); break;case 1:  CSwordDlg::OnFlk2l();   break;case 2:  CSwordDlg::OnFlk2l1();  break;case 3:  CSwordDlg::OnFlkCol();  break;case 4:  CSwordDlg::OnFlkVs();   break;default:  break;}CSwordDlg::OnButton8(); //FLK�q��}

        //5nits
        if(m_bool_oqc_5nits)
        {
            MsrFine5Nits();  
            Msr5Nits9P(); 
        }
        
        //������
        if (m_bool_oqc_w5)  {    OnWhite();    Msr5P(); } //�զ�5�I
        if (m_bool_oqc_r5)  {    OnRed();    Msr5P(); } //�զ�5�I
        if (m_bool_oqc_g5)  {    OnGreen();    Msr5P(); } //�զ�5�I
        if (m_bool_oqc_b5)  {    OnBlue();    Msr5P(); } //�զ�5�I
        if (m_bool_oqc_d5)  {    OnBlack();    Msr5P(); } //�զ�5�I
        if (m_bool_oqc_w9)  {    OnWhite();    Msr9P(); } //�զ�9�I
        if (m_bool_oqc_w49)    {    OnWhite(); Msr49P();    }       //�զ�49�I
    //    CSwordDlg::OnWhite();   CSwordDlg::OnButton5();

        //�t�A�|��
    //    CSwordDlg::OnBlack();   CSwordDlg::OnButton6();       //�¦�13�I
        if(m_bool_oqc_d25){     OnBlack();    Msr25P();  }     //�¦�25�I

//        if(m_bool_oqc_d33){     OnBlack();    Msr33P();  }     //�¦�25�I

        SetBackGroundColor();
        ShowWinControlItem();//��ܥ��W�����F��
    //-----------------------------------------------------------------------------------
    //              ���q������, ����ƳB�z
    //-----------------------------------------------------------------------------------
        int Tai = pOQCTaiSel->GetCurSel();
        int i;

        //Excel��
        //�q�����c��
        for(i=0;i<9;++i)
        {
            OQC_Date[Tai].w9[i].Lv  = MsrNineOValue[0][i][0];
            OQC_Date[Tai].w9[i].x   = MsrNineOValue[0][i][1];
            OQC_Date[Tai].w9[i].y   = MsrNineOValue[0][i][2];
            OQC_Date[Tai].w9[i].T   = MsrNineOValue[0][i][3];

//             CString temp;
//             temp.Format("OnButtonMsrflow()��q���A�q�����c��\nOQC_Date[Tai].w9[i].T = MsrNineOValue[0][i][3];\n%d\n%f", 
//                  OQC_Date[Tai].w9[i].T, MsrNineOValue[0][i][3]);
//             MessageBox(temp);

            OQC_Date[Tai].w9[i].Duv = MsrNineOValue[0][i][4];
        }

        for(i=0;i<49;++i)
        {   
            //MsrFortyNineOValue[   Color  ][ n ][value]
            OQC_Date[Tai].w49[i].Lv = MsrFortyNineOValue[0][i][0];
        }

        for(i=0;i<5;++i)
        {   
            //MsrFiveOValue[   Color  ][ n ][value]
            OQC_Date[Tai].w5[i].Lv = MsrFiveOValue[0][i][0];
            OQC_Date[Tai].w5[i].x  = MsrFiveOValue[0][i][1];
            OQC_Date[Tai].w5[i].y  = MsrFiveOValue[0][i][2];
            OQC_Date[Tai].w5[i].T  = MsrFiveOValue[0][i][3];
        }
        
        for(i=0;i<5;++i)
        {   
            //MsrFiveOValue[   Color  ][ n ][value]
            OQC_Date[Tai].r5[i].Lv = MsrFiveOValue[1][i][0];
            OQC_Date[Tai].r5[i].x  = MsrFiveOValue[1][i][1];
            OQC_Date[Tai].r5[i].y  = MsrFiveOValue[1][i][2];
            OQC_Date[Tai].r5[i].T  = MsrFiveOValue[1][i][3];
        }

        
        for(i=0;i<5;++i)
        {   
            //MsrFiveOValue[   Color  ][ n ][value]
            OQC_Date[Tai].g5[i].Lv = MsrFiveOValue[2][i][0];
            OQC_Date[Tai].g5[i].x  = MsrFiveOValue[2][i][1];
            OQC_Date[Tai].g5[i].y  = MsrFiveOValue[2][i][2];
            OQC_Date[Tai].g5[i].T  = MsrFiveOValue[2][i][3];
        }

        for(i=0;i<5;++i)
        {   
            //MsrFiveOValue[   Color  ][ n ][value]
            OQC_Date[Tai].b5[i].Lv = MsrFiveOValue[3][i][0];
            OQC_Date[Tai].b5[i].x  = MsrFiveOValue[3][i][1];
            OQC_Date[Tai].b5[i].y  = MsrFiveOValue[3][i][2];
            OQC_Date[Tai].b5[i].T  = MsrFiveOValue[3][i][3];
        }
    
        for(i=0;i<5;++i)
        {   
            //MsrFiveOValue[   Color  ][ n ][value]
            OQC_Date[Tai].d5[i].Lv = MsrFiveOValue[4][i][0];
            OQC_Date[Tai].d5[i].x  = MsrFiveOValue[4][i][1];
            OQC_Date[Tai].d5[i].y  = MsrFiveOValue[4][i][2];
            OQC_Date[Tai].d5[i].T  = MsrFiveOValue[4][i][3];
        }

        //W1(x,y)
        OQC_Date[Tai].w.Lv= MsrCenter[0][0];
        OQC_Date[Tai].w.x = MsrCenter[0][1];
        OQC_Date[Tai].w.y = MsrCenter[0][2];

        //R1,G1,B1=i
        //Lv,x,y
        OQC_Date[Tai].r.Lv = MsrCenter[1][0];
        OQC_Date[Tai].r.x  = MsrCenter[1][1];
        OQC_Date[Tai].r.y  = MsrCenter[1][2];
        OQC_Date[Tai].r.T  = MsrCenter[1][3];
        OQC_Date[Tai].r.Duv= MsrCenter[1][4];

        
        OQC_Date[Tai].g.Lv = MsrCenter[2][0];
        OQC_Date[Tai].g.x  = MsrCenter[2][1];
        OQC_Date[Tai].g.y  = MsrCenter[2][2];
        OQC_Date[Tai].g.T  = MsrCenter[1][3];
        OQC_Date[Tai].g.Duv= MsrCenter[1][4];
        
        OQC_Date[Tai].b.Lv = MsrCenter[3][0];
        OQC_Date[Tai].b.x  = MsrCenter[3][1];
        OQC_Date[Tai].b.y  = MsrCenter[3][2];
        OQC_Date[Tai].b.T  = MsrCenter[3][3];
        OQC_Date[Tai].b.Duv= MsrCenter[3][4];

        //D1
        OQC_Date[Tai].d.Lv = MsrCenter[4][0];
        OQC_Date[Tai].d.x  = MsrCenter[4][1];
        OQC_Date[Tai].d.y  = MsrCenter[4][2];
        OQC_Date[Tai].d.T  = MsrCenter[4][3];
        OQC_Date[Tai].d.Duv= MsrCenter[4][4];

        //Fnits1
        OQC_Date[Tai].d5nits1.Lv = FiveNits[4][0];
        OQC_Date[Tai].d5nits1.x  = FiveNits[4][1];
        OQC_Date[Tai].d5nits1.y  = FiveNits[4][2];
        OQC_Date[Tai].d5nits1.T  = FiveNits[4][3];
        OQC_Date[Tai].d5nits1.Duv= FiveNits[4][4];

        //d9
        for(i=0;i<9;++i)
        {                          //MsrNineOValue[Color][n][value]
            OQC_Date[Tai].d9[i].Lv = MsrNineOValue[4][i][0];
        }
        //D13
//         for(i=0;i<13;++i)
//         {
//             OQC_Date[Tai].d13[i].Lv = MsrThirteenValue[i][0];
//         }

        //D25
        for(i=0;i<25;++i)
        {
            OQC_Date[Tai].d25[i].Lv = MsrTwentyFiveValue[i][0];
        }
        //5nits
        for(i=0;i<9;++i)
        {
            OQC_Date[Tai].d5nits9[i].Lv = FiveNits[i][0];
        }

        OQC_Date[Tai].ChannelNO = m_IMemory.GetChannelNO();
    //-----------------------------------------------------------------------------------
        CString OriginalData;
        CString Tempstr;

    //�g�J��l�����
        //W9(Lv)
        Tempstr.FormatMessage("�զ�E�I(Lv)\n");            OriginalData += (Tempstr);
        
        for(i=0;i<9;++i)
        {
            Tempstr.Format("%f",OQC_Date[Tai].w9[i].Lv);    OriginalData += (Tempstr + "\t");
        }
        
        Tempstr.FormatMessage("\n\n");                        OriginalData += (Tempstr);
        Tempstr.FormatMessage("�զ�E�I(x)\n");            OriginalData += (Tempstr);
        
        for(i=0;i<9;++i)
        {
            Tempstr.Format("%f",OQC_Date[Tai].w9[i].x);    OriginalData += (Tempstr + "\t");
        }
        
        Tempstr.FormatMessage("\n\n");                        OriginalData += (Tempstr);
        Tempstr.FormatMessage("�զ�E�I(y)\n");            OriginalData += (Tempstr);
        
        for(i=0;i<9;++i)
        {
            Tempstr.Format("%f",OQC_Date[Tai].w9[i].y);    OriginalData += (Tempstr + "\t");
        }
        
        Tempstr.FormatMessage("\n\n");                        OriginalData += (Tempstr);
        Tempstr.FormatMessage("�զ�E�I(T)\n");            OriginalData += (Tempstr);
        
        for(i=0;i<9;++i)
        {
            Tempstr.Format("%d",OQC_Date[Tai].w9[i].T);    OriginalData += (Tempstr + "\t");

//             CString temp;
//             temp.Format("OnButtonMsrflow()��q���A���b�g�J��l�����\nTempstr.Format(\"%%d\",OQC_Date[Tai].w9[i].T);\nOriginalData += (Tempstr + \"\\t\");\n%s\n%d", Tempstr, OQC_Date[Tai].w9[i].T);
//             MessageBox(temp);

        }
        
        Tempstr.FormatMessage("\n\n");                        OriginalData += (Tempstr);
        Tempstr.FormatMessage("�զ�E�I(Duv)\n");            OriginalData += (Tempstr);
        
        for(i=0;i<9;++i)
        {
            Tempstr.Format("%f",OQC_Date[Tai].w9[i].Duv);    OriginalData += (Tempstr + "\t");
        }
        
        Tempstr.FormatMessage("\n\n");                        OriginalData += (Tempstr);
        Tempstr.FormatMessage("�զ�49�I(Lv)\n");            OriginalData += (Tempstr);
        
        for(i=0;i<49;++i)
        {
            Tempstr.Format("%f",OQC_Date[Tai].w49[i].Lv);    OriginalData += (Tempstr + "\t");
        }
        
        Tempstr.FormatMessage("\n\n");                        OriginalData += (Tempstr);
        Tempstr.FormatMessage("�զ⤤���I(x,y)\n");            OriginalData += (Tempstr);

            Tempstr.Format("%f",OQC_Date[Tai].w.Lv);    OriginalData += (Tempstr + "\t");
            Tempstr.Format("%f",OQC_Date[Tai].w.x);        OriginalData += (Tempstr + "\t");
            Tempstr.Format("%f",OQC_Date[Tai].w.y);        OriginalData += (Tempstr + "\t");

        Tempstr.FormatMessage("\n\n");                        OriginalData += (Tempstr);
        Tempstr.FormatMessage("���⤤���I(Lv,x,y)\n");        OriginalData += (Tempstr);
        
            Tempstr.Format("%f",OQC_Date[Tai].r.Lv);                OriginalData += (Tempstr + "\t");
            Tempstr.Format("%f",OQC_Date[Tai].r.x);                OriginalData += (Tempstr + "\t");
            Tempstr.Format("%f",OQC_Date[Tai].r.y);                OriginalData += (Tempstr + "\t");
        
        Tempstr.FormatMessage("\n\n");                        OriginalData += (Tempstr);
        Tempstr.FormatMessage("��⤤���I(Lv,x,y)\n");        OriginalData += (Tempstr);
        
            Tempstr.Format("%f",OQC_Date[Tai].g.Lv);                OriginalData += (Tempstr + "\t");
            Tempstr.Format("%f",OQC_Date[Tai].g.x);                OriginalData += (Tempstr + "\t");
            Tempstr.Format("%f",OQC_Date[Tai].g.y);                OriginalData += (Tempstr + "\t");

        Tempstr.FormatMessage("\n\n");                        OriginalData += (Tempstr);
        Tempstr.FormatMessage("�Ŧ⤤���I(Lv,x,y)\n");        OriginalData += (Tempstr);

            Tempstr.Format("%f",OQC_Date[Tai].b.Lv);                OriginalData += (Tempstr + "\t");
            Tempstr.Format("%f",OQC_Date[Tai].b.x);                OriginalData += (Tempstr + "\t");
            Tempstr.Format("%f",OQC_Date[Tai].b.y);                OriginalData += (Tempstr + "\t");

        Tempstr.FormatMessage("\n\n");                        OriginalData += (Tempstr);
        Tempstr.FormatMessage("�¦⤤���I(Lv,x,y)\n");        OriginalData += (Tempstr);

            Tempstr.Format("%f",OQC_Date[Tai].d.Lv);                OriginalData += (Tempstr + "\t");    
            Tempstr.Format("%f",OQC_Date[Tai].d.x);                OriginalData += (Tempstr + "\t");
            Tempstr.Format("%f",OQC_Date[Tai].d.y);                OriginalData += (Tempstr + "\t");    
        
        Tempstr.FormatMessage("\n\n");                        OriginalData += (Tempstr);
        Tempstr.FormatMessage("5Nits�����I(Lv,x,y)\n");        OriginalData += (Tempstr);

            Tempstr.Format("%f",OQC_Date[Tai].d5nits1.Lv);                OriginalData += (Tempstr + "\t");    
            Tempstr.Format("%f",OQC_Date[Tai].d5nits1.x);                OriginalData += (Tempstr + "\t");
            Tempstr.Format("%f",OQC_Date[Tai].d5nits1.y);                OriginalData += (Tempstr + "\t");    

//         Tempstr.FormatMessage("\n\n");                        OriginalData += (Tempstr);
//         Tempstr.FormatMessage("�¦�13�I(Lv)\n");            OriginalData += (Tempstr);
// 
//         for(i=0;i<13;++i)
//         {
//             Tempstr.Format("%f",OQC_Date[Tai].d13[i].Lv);
//             Tempstr.Format("");    OriginalData += (Tempstr + "\t");
//         }

        Tempstr.FormatMessage("\n\n");                        OriginalData += (Tempstr);
        Tempstr.FormatMessage("�զ�5�I(Lv)\n");                OriginalData += (Tempstr);

        for(i=0;i<5;++i)
        {
            Tempstr.Format("%f",OQC_Date[Tai].w5[i].Lv);    OriginalData += (Tempstr + "\t");
        }

        Tempstr.FormatMessage("\n\n");                        OriginalData += (Tempstr);
        Tempstr.FormatMessage("�զ�5�I(x)\n");                OriginalData += (Tempstr);
        
        for(i=0;i<5;++i)
        {
            Tempstr.Format("%f",OQC_Date[Tai].w5[i].x);        OriginalData += (Tempstr + "\t");
        }

        Tempstr.FormatMessage("\n\n");                        OriginalData += (Tempstr);
        Tempstr.FormatMessage("�զ�5�I(y)\n");                OriginalData += (Tempstr);
        
        for(i=0;i<5;++i)
        {
            Tempstr.Format("%f",OQC_Date[Tai].w5[i].y);        OriginalData += (Tempstr + "\t");
        }

        Tempstr.FormatMessage("\n\n");                        OriginalData += (Tempstr);
        Tempstr.FormatMessage("�¦�25�I(Lv)\n");            OriginalData += (Tempstr);

        for(i=0;i<25;++i)
        {
            Tempstr.Format("%f",OQC_Date[Tai].d25[i].Lv);    OriginalData += (Tempstr + "\t");
        }
        //5nits
        Tempstr.FormatMessage("\n\n");                        OriginalData += (Tempstr);
        Tempstr.FormatMessage("5nits(Lv)\n");                OriginalData += (Tempstr);

        for(i=0;i<9;++i)
        {
            Tempstr.Format("%f",OQC_Date[Tai].d5nits9[i].Lv);            OriginalData += (Tempstr + "\t");
        }
        Tempstr.FormatMessage("\n\n");                        OriginalData += (Tempstr);
        Tempstr.FormatMessage("�䥦(BarCode, DriverDevice, Current, ChannelNO)\n");            OriginalData += (Tempstr);
        
        Tempstr.Format("%s",OQC_Date[Tai].BarCode);            OriginalData += (Tempstr + "\t");
        Tempstr.Format("%s",OQC_Date[Tai].DriverDevice);    OriginalData += (Tempstr + "\t");
        Tempstr.Format("%d",OQC_Date[Tai].Current);            OriginalData += (Tempstr + "\t");
        Tempstr.Format("%d",OQC_Date[Tai].ChannelNO);        OriginalData += (Tempstr + "\t");

    //----------------------------------------------------------------------------------------
        //struct  �ĴX�x.�q������[�I].��
    //    TCHAR szPath[MAX_PATH]; 
        CString DesktopPath;
        CString pathName;
        CString SizeAndNumber;    

        //if(BCFandODFPath == NULL)
        //{
            //static int NoBarCodeN = 1;
            //SHGetSpecialFolderPath(NULL,szPath,CSIDL_DESKTOP,0);//���o�ୱ���|

            //DesktopPath.Format("%s\\NoBarCode-%d.txt",szPath, NoBarCodeN);

            //pathName = DesktopPath;
            //NoBarCodeN++;
        //}
        //else
        //{
            DesktopPath.Format("%s",BCFandODFPath);                    //Original File path = Bar Code File path �˦��r��

            SizeAndNumber.Format("%d-%s",iLCMSize, OQC_Date[Tai].BarCode);//���o���|

            pathName = DesktopPath + "\\" + SizeAndNumber + ".txt";
        

            CFile myFile;
            myFile.Open (pathName, CFile::modeCreate  | CFile::modeWrite);//�}��
            myFile.Write(OriginalData, strlen (OriginalData));//�ɮפ��e�ಾ�� pbuf
            myFile.Close();

            pathName = "��l������x�s�b:\n" + pathName;
            AfxMessageBox(pathName);
        //}

        //m_strTestValue.Format("%s",pathName);
        //UpdateData(FALSE);

    //-----------------------------------------------------------------------------------
        pBtmSaveOQC  -> EnableWindow(1);
        pOQCFormSel  -> EnableWindow(1);
        OnMax();
        
        
        if(pOQCTaiSel->GetCurSel()<9)
            pOQCTaiSel   -> SetCurSel(pOQCTaiSel->GetCurSel()+1);
        else
        {
            pOQCTaiSel   -> SetCurSel(0);
            //pOQCStartMsr-> EnableWindow(0);
            //pOQCLoadBarCode->EnableWindow(1);
            MessageBox("10�x�q���o!");
        }
        
        
//         CString temp;
//         temp.Format("OnButtonMsrflow()��q���A���X��l�����\n%d", OQC_Date[Tai].w9[i].T);
//         MessageBox(temp);

    }
    else//-------����----------------------------------------------------------
    {
        int Tai;
        char szFilters[]= _T("MyType Files (*.txt)|*.txt|All Files (*.*)|*.*||");    //   |��,  ||��;

        char szSelectFileBuffer[0x10000] = {0};

        // Create an Open dialog; the default file name extension is ".my".
       CFileDialog fileDlg (
           TRUE,                                //bOpenFileDialog    �]�m�� TRUE�G�}�����ɡFFALSE�t�s�s��
           "txt",                                //lpszDefExt        �ɮ����������e
           "*.txt",                                //lpszFileName        �ɮצW��
          OFN_ALLOWMULTISELECT,                    //dwFlags           
          szFilters,                            //szFilters         �s���L�o��
          this                                    //pParentWnd        ����������
                                                //dwSize            �w�]0�G�̧@�~�t�ιw�]����
         );

       //���w�Ǧ^
       fileDlg.m_ofn.lpstrFile = szSelectFileBuffer;
       //�]�wbuffer�j�p
       fileDlg.m_ofn.nMaxFile = 0x10000;

        if (fileDlg.DoModal ()==IDOK) 
        {
            //--------------------------------------------
            //Ū�h��
            Tai = 0;
            CString szPathName;
            //CString szFileName;
            POSITION pos = fileDlg.GetStartPosition();
            while(pos != NULL)
            {
                
                szPathName = fileDlg.GetNextPathName(pos);//�z�L���ШӨ��o�ɮצW�� �ɮצW�٥]�t������|
                //�ɮ׸��| = szPathName
                //------------------Ū��-----------------
                char pbuf1[2048];//�浧�ɮפj�p
                char *pbuf2;

                unsigned int sizeFile;

                // Implement opening and reading file in here.
                CFile myFile;

                myFile.Open (szPathName, CFile::shareExclusive  | CFile::modeRead);

                    //szFileName = myFile.GetFileName().Mid(3,22);       //���o�ɦW
                    sizeFile   = myFile.GetLength();//���o�ɮפj�p
                    if(sizeFile > 2048)
                        MessageBox("�ɮ׶W�L2KB");
                    else
                        myFile.Read(pbuf1, sizeFile);//�ɮפ��e�ಾ�� pbuf

                myFile.Close();
                
                //------------------�N���e�নpbuf1�A����pbuf2-----------------

                pbuf2 = strtok (pbuf1,"\t\n");

                //---------------------------
                CString x;
                int point;

                //�̮榡��J�O����
                //--------------------------------�@���h�ɡA�@�ɦh�����ۦP-------------------------
                //    OQC_Date[i];

                if(pbuf2 != NULL)    pbuf2 = strtok (NULL, "\t\n");/*�զ�E�I(Lv)*/
                point=0;
                while (pbuf2 != NULL && point < 9)
                {    
                    x.Format("%s",pbuf2);
                    OQC_Date[Tai].w9[point].Lv = strtod(x, &pbuf2);    
                    pbuf2 = strtok (NULL, "\t\n");
                    ++point;
                }
                if(pbuf2 != NULL)pbuf2 = strtok (NULL, "\t\n");
                if(pbuf2 != NULL)pbuf2 = strtok (NULL, "\t\n");

                if(pbuf2 != NULL)    pbuf2 = strtok (NULL, "\t\n");/*�զ�E�I(x)*/
                point=0;
                while (pbuf2 != NULL && point < 9)
                {    
                    x.Format("%s",pbuf2);    
                    OQC_Date[Tai].w9[point].x = strtod(x, &pbuf2);    
                    pbuf2 = strtok (NULL, "\t\n");
                    ++point;
                }
                if(pbuf2 != NULL)pbuf2 = strtok (NULL, "\t\n");
                if(pbuf2 != NULL)pbuf2 = strtok (NULL, "\t\n");
                
                if(pbuf2 != NULL)    pbuf2 = strtok (NULL, "\t\n");/*�զ�E�I(y)*/
                point=0;
                while (pbuf2 != NULL && point < 9)
                {    
                    x.Format("%s",pbuf2);    
                    OQC_Date[Tai].w9[point].y = strtod(x, &pbuf2);    
                    pbuf2 = strtok (NULL, "\t\n");
                    ++point;
                }
                if(pbuf2 != NULL)pbuf2 = strtok (NULL, "\t\n");
                if(pbuf2 != NULL)pbuf2 = strtok (NULL, "\t\n");
                
                if(pbuf2 != NULL)    pbuf2 = strtok (NULL, "\t\n");/*�զ�E�I(T)*/
                point=0;
                while (pbuf2 != NULL && point < 9)
                {    
                    x.Format("%s",pbuf2);    
                    OQC_Date[Tai].w9[point].T = strtod(x, &pbuf2);    
                    pbuf2 = strtok (NULL, "\t\n");
                    ++point;
                }
                if(pbuf2 != NULL)pbuf2 = strtok (NULL, "\t\n");
                if(pbuf2 != NULL)pbuf2 = strtok (NULL, "\t\n");
                
                if(pbuf2 != NULL)    pbuf2 = strtok (NULL, "\t\n");/*�զ�9�I(Duv)*/
                point=0;
                while (pbuf2 != NULL && point < 9)
                {    
                    x.Format("%s",pbuf2);    
                    OQC_Date[Tai].w9[point].Duv = strtod(x, &pbuf2);    
                    pbuf2 = strtok (NULL, "\t\n");
                    ++point;
                }
                if(pbuf2 != NULL)pbuf2 = strtok (NULL, "\t\n");
                if(pbuf2 != NULL)pbuf2 = strtok (NULL, "\t\n");
                
                if(pbuf2 != NULL)    pbuf2 = strtok (NULL, "\t\n");/*�զ�49�I(Lv)*/
                point=0;
                while (pbuf2 != NULL && point < 49)
                {    
                    x.Format("%s",pbuf2);    
                    OQC_Date[Tai].w49[point].Lv = strtod(x, &pbuf2);    
                    pbuf2 = strtok (NULL, "\t\n");
                    ++point;
                }
                if(pbuf2 != NULL)pbuf2 = strtok (NULL, "\t\n");
                if(pbuf2 != NULL)pbuf2 = strtok (NULL, "\t\n");


                if(pbuf2 != NULL)pbuf2 = strtok (NULL, "\t\n");/*�զ⤤���I(x,y)*/    

                if(pbuf2 != NULL){    x.Format("%s",pbuf2);    OQC_Date[Tai].w.Lv = strtod(x, &pbuf2);    pbuf2 = strtok (NULL, "\t\n");    }
                if(pbuf2 != NULL){    x.Format("%s",pbuf2);    OQC_Date[Tai].w.x  = strtod(x, &pbuf2);    pbuf2 = strtok (NULL, "\t\n");    }
                if(pbuf2 != NULL){    x.Format("%s",pbuf2);    OQC_Date[Tai].w.y  = strtod(x, &pbuf2);    pbuf2 = strtok (NULL, "\t\n");    }

                if(pbuf2 != NULL)pbuf2 = strtok (NULL, "\t\n");
                if(pbuf2 != NULL)pbuf2 = strtok (NULL, "\t\n");
                if(pbuf2 != NULL)pbuf2 = strtok (NULL, "\t\n");
                //if(pbuf2 != NULL){MessageBox(pbuf2,"���⤤���I(Lv,x,y)");    /*���⤤���I(Lv,x,y)*/        pbuf2 = strtok (NULL, "\t\n");    }
                if(pbuf2 != NULL){    x.Format("%s",pbuf2);    OQC_Date[Tai].r.Lv = strtod(x, &pbuf2);    pbuf2 = strtok (NULL, "\t\n");    }
                if(pbuf2 != NULL){    x.Format("%s",pbuf2);    OQC_Date[Tai].r.x = strtod(x, &pbuf2);    pbuf2 = strtok (NULL, "\t\n");    }
                if(pbuf2 != NULL){    x.Format("%s",pbuf2);    OQC_Date[Tai].r.y = strtod(x, &pbuf2);    pbuf2 = strtok (NULL, "\t\n");    }

                if(pbuf2 != NULL)pbuf2 = strtok (NULL, "\t\n");
                if(pbuf2 != NULL)pbuf2 = strtok (NULL, "\t\n");
                if(pbuf2 != NULL)pbuf2 = strtok (NULL, "\t\n");
                //if(pbuf2 != NULL){MessageBox(pbuf2,"��⤤���I(Lv,x,y)");    /*��⤤���I(Lv,x,y)*/        pbuf2 = strtok (NULL, "\t\n");    }
                if(pbuf2 != NULL){    x.Format("%s",pbuf2);    OQC_Date[Tai].g.Lv = strtod(x, &pbuf2);    pbuf2 = strtok (NULL, "\t\n");    }
                if(pbuf2 != NULL){    x.Format("%s",pbuf2);    OQC_Date[Tai].g.x = strtod(x, &pbuf2);    pbuf2 = strtok (NULL, "\t\n");    }
                if(pbuf2 != NULL){    x.Format("%s",pbuf2);    OQC_Date[Tai].g.y = strtod(x, &pbuf2);    pbuf2 = strtok (NULL, "\t\n");    }

                if(pbuf2 != NULL)pbuf2 = strtok (NULL, "\t\n");
                if(pbuf2 != NULL)pbuf2 = strtok (NULL, "\t\n");
                if(pbuf2 != NULL)pbuf2 = strtok (NULL, "\t\n");
                //if(pbuf2 != NULL){MessageBox(pbuf2,"�Ŧ⤤���I(Lv,x,y)");    /*�Ŧ⤤���I(Lv,x,y)*/        pbuf2 = strtok (NULL, "\t\n");    }
                if(pbuf2 != NULL){    x.Format("%s",pbuf2);    OQC_Date[Tai].b.Lv = strtod(x, &pbuf2);    pbuf2 = strtok (NULL, "\t\n");    }
                if(pbuf2 != NULL){    x.Format("%s",pbuf2);    OQC_Date[Tai].b.x = strtod(x, &pbuf2);    pbuf2 = strtok (NULL, "\t\n");    }
                if(pbuf2 != NULL){    x.Format("%s",pbuf2);    OQC_Date[Tai].b.y = strtod(x, &pbuf2);    pbuf2 = strtok (NULL, "\t\n");    }

                if(pbuf2 != NULL)pbuf2 = strtok (NULL, "\t\n");
                if(pbuf2 != NULL)pbuf2 = strtok (NULL, "\t\n");
                if(pbuf2 != NULL)pbuf2 = strtok (NULL, "\t\n");
                //if(pbuf2 != NULL){MessageBox(pbuf2,"�¦⤤���I(Lv,x,y)");    /*�¦⤤���I(Lv,x,y)*/        pbuf2 = strtok (NULL, "\t\n");    }
                if(pbuf2 != NULL){    x.Format("%s",pbuf2);    OQC_Date[Tai].d.Lv = strtod(x, &pbuf2);    pbuf2 = strtok (NULL, "\t\n");    }
                if(pbuf2 != NULL){    x.Format("%s",pbuf2);    OQC_Date[Tai].d.x = strtod(x, &pbuf2);    pbuf2 = strtok (NULL, "\t\n");    }
                if(pbuf2 != NULL){    x.Format("%s",pbuf2);    OQC_Date[Tai].d.y = strtod(x, &pbuf2);    pbuf2 = strtok (NULL, "\t\n");    }

                if(pbuf2 != NULL)pbuf2 = strtok (NULL, "\t\n");
                if(pbuf2 != NULL)pbuf2 = strtok (NULL, "\t\n");
                if(pbuf2 != NULL)pbuf2 = strtok (NULL, "\t\n");
                //if(pbuf2 != NULL){MessageBox(pbuf2,"5Nits�����I(Lv,x,y)");    /*5Nits�����I(Lv,x,y)*/        pbuf2 = strtok (NULL, "\t\n");    }
                if(pbuf2 != NULL){    x.Format("%s",pbuf2);    OQC_Date[Tai].d5nits1.Lv = strtod(x, &pbuf2);    pbuf2 = strtok (NULL, "\t\n");    }
                if(pbuf2 != NULL){    x.Format("%s",pbuf2);    OQC_Date[Tai].d5nits1.x = strtod(x, &pbuf2);    pbuf2 = strtok (NULL, "\t\n");    }
                if(pbuf2 != NULL){    x.Format("%s",pbuf2);    OQC_Date[Tai].d5nits1.y = strtod(x, &pbuf2);    pbuf2 = strtok (NULL, "\t\n");    }

                if(pbuf2 != NULL)pbuf2 = strtok (NULL, "\t\n");
                if(pbuf2 != NULL)pbuf2 = strtok (NULL, "\t\n");
                if(pbuf2 != NULL)pbuf2 = strtok (NULL, "\t\n");
                //if(pbuf2 != NULL){MessageBox(pbuf2,"�¦�13�I(Lv)");    /*�¦�13�I(Lv)*/        pbuf2 = strtok (NULL, "\t\n");    }
//                 point=0;
//                 while (pbuf2 != NULL && point < 13)
//                 {    
//                     x.Format("%s",pbuf2);    
//                     //OQC_Date[Tai].d13[point].Lv = strtod(x, &pbuf2);    
//                     pbuf2 = strtok (NULL, "\t\n");
//                     ++point;
//                 }
//                 
//                 if(pbuf2 != NULL)pbuf2 = strtok (NULL, "\t\n");
//                 if(pbuf2 != NULL)pbuf2 = strtok (NULL, "\t\n");
//                 if(pbuf2 != NULL)pbuf2 = strtok (NULL, "\t\n");
                //if(pbuf2 != NULL){MessageBox(pbuf2,"�զ�5�I(Lv)");    /*�զ�5�I(Lv)*/        pbuf2 = strtok (NULL, "\t\n");    }
                point=0;
                while (pbuf2 != NULL && point < 5)
                {    
                    x.Format("%s",pbuf2);    
                    OQC_Date[Tai].w5[point].Lv = strtod(x, &pbuf2);    
                    pbuf2 = strtok (NULL, "\t\n");
                    ++point;
                }
                
                if(pbuf2 != NULL)pbuf2 = strtok (NULL, "\t\n");
                if(pbuf2 != NULL)pbuf2 = strtok (NULL, "\t\n");
                if(pbuf2 != NULL)pbuf2 = strtok (NULL, "\t\n");
                //if(pbuf2 != NULL){MessageBox(pbuf2,"�զ�5�I(Lv)");    /*�զ�5�I(x)*/        pbuf2 = strtok (NULL, "\t\n");    }
                point=0;
                while (pbuf2 != NULL && point < 5)
                {    
                    x.Format("%s",pbuf2);    
                    OQC_Date[Tai].w5[point].Lv = strtod(x, &pbuf2);    
                    pbuf2 = strtok (NULL, "\t\n");
                    ++point;
                }
                
                if(pbuf2 != NULL)pbuf2 = strtok (NULL, "\t\n");
                if(pbuf2 != NULL)pbuf2 = strtok (NULL, "\t\n");
                if(pbuf2 != NULL)pbuf2 = strtok (NULL, "\t\n");
                //if(pbuf2 != NULL){MessageBox(pbuf2,"�զ�5�I(Lv)");    /*�զ�5�Iy)*/        pbuf2 = strtok (NULL, "\t\n");    }
                point=0;
                while (pbuf2 != NULL && point < 5)
                {    
                    x.Format("%s",pbuf2);    
                    OQC_Date[Tai].w5[point].Lv = strtod(x, &pbuf2);    
                    pbuf2 = strtok (NULL, "\t\n");
                    ++point;
                }
                
                if(pbuf2 != NULL)pbuf2 = strtok (NULL, "\t\n");
                if(pbuf2 != NULL)pbuf2 = strtok (NULL, "\t\n");
                if(pbuf2 != NULL)pbuf2 = strtok (NULL, "\t\n");
                //if(pbuf2 != NULL){MessageBox(pbuf2,"�¦�25�I(Lv)");    /*�¦�25�I(Lv)*/        pbuf2 = strtok (NULL, "\t\n");    }
                point=0;
                while (pbuf2 != NULL && point < 25)
                {    
                    x.Format("%s",pbuf2);    
                    OQC_Date[Tai].d25[point].Lv = strtod(x, &pbuf2);    
                    pbuf2 = strtok (NULL, "\t\n");
                    ++point;
                }
                
                if(pbuf2 != NULL)pbuf2 = strtok (NULL, "\t\n");
                if(pbuf2 != NULL)pbuf2 = strtok (NULL, "\t\n");
                if(pbuf2 != NULL)pbuf2 = strtok (NULL, "\t\n");
                //if(pbuf2 != NULL){MessageBox(pbuf2,"5nits(Lv)");    /*5nits(Lv)*/        pbuf2 = strtok (NULL, "\t\n");    }
                point=0;
                while (pbuf2 != NULL && point < 9)
                {    
                    x.Format("%s",pbuf2);    
                    OQC_Date[Tai].d5nits9[point].Lv = strtod(x, &pbuf2);    
                    pbuf2 = strtok (NULL, "\t\n");
                    ++point;
                }

                //CString temp;

                if(pbuf2 != NULL)pbuf2 = strtok (NULL, "\t\n");
                if(pbuf2 != NULL)pbuf2 = strtok (NULL, "\t\n");
                if(pbuf2 != NULL)pbuf2 = strtok (NULL, "\t\n");
                //if(pbuf2 != NULL){MessageBox(pbuf2,"other");        /*BarCode DirverDevice Current ChannelNo*/        pbuf2 = strtok (NULL, "\t\n");    }

                if(pbuf2 != NULL){    strcpy(OQC_Date[Tai].BarCode        ,pbuf2);    pbuf2 = strtok (NULL, "\t\n");    }
                if(pbuf2 != NULL){    strcpy(OQC_Date[Tai].DriverDevice    ,pbuf2);    pbuf2 = strtok (NULL, "\t\n");    }
                
                if(pbuf2 != NULL){    OQC_Date[Tai].Current   = atoi(pbuf2);    pbuf2 = strtok (NULL, "\t\n");    }
                if(pbuf2 != NULL){    OQC_Date[Tai].ChannelNO = atoi(pbuf2);    pbuf2 = strtok (NULL, "\t\n");    }


                //sprintf(buf,"%d",OQC_Date[Tai].ChannelNO);
                while (pbuf2 != NULL)
                {    
                    pbuf2 = strtok (NULL, "\t\n");
                }
                //--------------------------------�@���h�ɡA�@�ɦh�����ۦP-------------------------
                ++Tai;

            }//�C�@�Ӹ��|�B�z����
            m_Tai = Tai;
            pBtmSaveOQC  -> EnableWindow(1);
            pOQCFormSel  -> EnableWindow(1);
        }//��w�ܦh�ɮ׫��UOK
    CString tranMsg;
    tranMsg.Format("���\���J%d���ɮ�\n��w���榡\n�A�uOutput Form�v", Tai);
    MessageBox(tranMsg);
    }//���ɼҦ�
}
/*
void CSwordDlg::OutOrigFile()
{
    // TODO: Add your control notification handler code here
    //int Tai = pOQCTaiSel->GetCurSel();

}
*/

//�E�I�q��
void CSwordDlg::OnMsrFlowStart() 
{
    //�]�w�������F��
    HideColorController();
    HideOQCItem();
    HideMsrItem();
    HideMsrView();
    HideReportBtm();
    HideWinControlItem();

    Msr9P();
//�q���L�{�����A����ܦ^�Ӫ��F��
    ShowColorController();
    ShowOQCItem();
    ShowMsrItem();
    ShowMsrView();
    ShowReportBtm();
    ShowWinControlItem();
    EnabledReportBtm();
}

//�����I�q���]���]�w����^
void CSwordDlg::OnButton2() 
{
//�]�w�������F��
    HideColorController();
    HideOQCItem();
    HideMsrItem();
    HideMsrView();
    HideWinControlItem();
    HideReportBtm();
    Msr1P();
//�q���L�{�����A����ܦ^�Ӫ��F��
    ShowColorController();
    ShowWinControlItem();
    ShowOQCItem();
    ShowMsrItem();
    ShowMsrView();
    ShowReportBtm();
    EnabledReportBtm();
}


//��5nits���q���]���]�w����^
//�t��5nits��9�I�q��
void CSwordDlg::OnButton3() 
{
//�]�w�������F��
    HideColorController();
    HideOQCItem();
    HideMsrItem();
    HideMsrView();
    HideWinControlItem();
    HideReportBtm();
    //��5Nits
    MsrFine5Nits();

    //�i�J5nits���E�I�q��
    Msr5Nits9P();

    //�q���L�{�����A����ܦ^�Ӫ��F��
    ShowColorController();
    ShowWinControlItem();
    ShowOQCItem();
    ShowMsrItem();
    ShowMsrView();
    ShowReportBtm();
    EnabledReportBtm();
}

//Gamma�q���]���]�w����^
//�]�q�`�A�ڬO���q�`�աI���|�[�]�T���[
//�ӥB�ڳ]�w�ܧ֡I��49�I�֡I
void CSwordDlg::OnButton4() 
{
    // TODO: Add your control notification handler code here
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
    char buf1[200];
    char buf2[200];

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

    //�����|�X��
    char GammaPath[MAX_PATH];
    //dwCurDirPathLen = GetCurrentDirectory(MAX_PATH,szCurrentDirectory);//��ثe�Ҧb���ؿ��]���|�^
    strcpy(GammaPath, szCurrentDirectory);
    strcat(GammaPath, "\\Gamma.xls");
    
    objBooks = objApp.GetWorkbooks();
    //objBook = objBooks.Add(VOptional);
    objBook.AttachDispatch(objBooks.Add(_variant_t(GammaPath))); //�}�Ҥ@�Ӥw�s�b���ɮ�

//Step 3.�]�wSheet1
//��ʳ]�w
    objSheets = objBook.GetWorksheets();

    objSheet = objSheets.GetItem(COleVariant((short)1)); //�qsheet1�}�l
    objSheet.SetName("Report"); //�]�wsheet�W��


//Step 4.�}�l�]�w���e

//-----------------------------------------------------------------------------------------------
//           ���r�񧹡I�U���O��J��ơI�зǳư}�C�I
//-----------------------------------------------------------------------------------------------

    
//�]�w�������F��
    HideColorController();
    HideOQCItem();
    HideMsrItem();
    HideMsrView();
    HideWinControlItem();
    HideReportBtm();

    Invalidate(TRUE);
    UpdateWindow();

    m_ICa.SetDisplayMode(0);
//�w���骺��m
    CurrentPoint = Set9Point();//�w�]�s�����I
    if(MsrAI(AutoSelectMode(GetBackGroundColor())))//�p�G�q�����K�W�ù�
    {
        int q=0;   //�ⶥ
        int r=0;   //�C�����
        float Lv = 0;

        for(r=0;r<4;r++){  //�C�����
        for(q=0;q<=255;q++){//�ⶥ
            //�ο���x�s���O����A�M�w�������C��
            //�ܰʭI���C�� R - G - B
            switch(r)//�C�����
            {
                 case BK_White: SetBackGroundColor(BK_Other,RGB(q,q,q)); break;//w
                 case BK_Red:   SetBackGroundColor(BK_Other,RGB(q,0,0)); break;//R
                 case BK_Green: SetBackGroundColor(BK_Other,RGB(0,q,0)); break;//G
                 case BK_Blue:  SetBackGroundColor(BK_Other,RGB(0,0,q)); break;//B
                default: SetBackGroundColor(BK_Other,RGB(rand()%255+1,rand()%255+1,rand()%255+1));
            }
            m_Brush.DeleteObject(); 
            m_Brush.CreateSolidBrush(GetBackGroundColor());
            Invalidate();
            UpdateWindow();
            //20  �i�H
            //50  �]5nits �]�w100�A�Q���o�ӭ����Ӥ]�i�H
            Sleep(100);//�`�N�����ɶ��վ�o�ӭȡ]�̫G���X���D���̷t�^
            m_ICa.Measure(0);//�q��
            
            //RGB�����I
            switch(r)
            {
            case 0://R
            //L
                        ZeroMemory(buf,sizeof(buf));    sprintf(buf,"%c%d",'E',4+q);
                                                         range = objSheet.GetRange(COleVariant(buf),COleVariant(buf));
             
                        ZeroMemory(buf,sizeof(buf));    Lv=m_IProbe.GetLv();
                                                        sprintf(buf,"%3.2f",Lv); //L
                                                        range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf));
            //x
                        ZeroMemory(buf,sizeof(buf));    sprintf(buf,"%c%d",'F',4+q);
                                                        range = objSheet.GetRange(COleVariant(buf),COleVariant(buf));
             
                        ZeroMemory(buf,sizeof(buf));    Lv=m_IProbe.GetLv();
                                                        sprintf(buf,"%3.2f",Lv); //L
                                                        range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf));
            //y
                        ZeroMemory(buf,sizeof(buf));    sprintf(buf,"%c%d",'G',4+q);
                                                        range = objSheet.GetRange(COleVariant(buf),COleVariant(buf));
             
                        ZeroMemory(buf,sizeof(buf));    Lv=m_IProbe.GetLv();
                                                        sprintf(buf,"%3.2f",Lv); //L
                                                        range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf));
                        break;
            case 1://G
            //L
                        ZeroMemory(buf,sizeof(buf));    sprintf(buf,"%c%d",'L',4+q);
                                                        range = objSheet.GetRange(COleVariant(buf),COleVariant(buf));
             
                        ZeroMemory(buf,sizeof(buf));    Lv=m_IProbe.GetLv();
                                                        sprintf(buf,"%3.2f",Lv); //L
                                                        range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf));
            //x
                        ZeroMemory(buf,sizeof(buf));    sprintf(buf,"%c%d",'M',4+q);
                                                        range = objSheet.GetRange(COleVariant(buf),COleVariant(buf));
             
                        ZeroMemory(buf,sizeof(buf));    Lv=m_IProbe.GetLv();
                                                        sprintf(buf,"%3.2f",Lv); //L
                                                        range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf));
            //y
                        ZeroMemory(buf,sizeof(buf));    sprintf(buf,"%c%d",'N',4+q);
                                                        range = objSheet.GetRange(COleVariant(buf),COleVariant(buf));
             
                        ZeroMemory(buf,sizeof(buf));    Lv=m_IProbe.GetLv();
                                                        sprintf(buf,"%3.2f",Lv); //L
                                                        range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf));
                        break;
            case 2://B
            //L
                        ZeroMemory(buf,sizeof(buf));    sprintf(buf,"%c%d",'S',4+q);
                                                        range = objSheet.GetRange(COleVariant(buf),COleVariant(buf));
             
                        ZeroMemory(buf,sizeof(buf));    Lv=m_IProbe.GetLv();
                                                        sprintf(buf,"%3.2f",Lv); //L
                                                        range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf));
            //x
                        ZeroMemory(buf,sizeof(buf));    sprintf(buf,"%c%d",'T',4+q);
                                                        range = objSheet.GetRange(COleVariant(buf),COleVariant(buf));
             
                        ZeroMemory(buf,sizeof(buf));    Lv=m_IProbe.GetLv();
                                                        sprintf(buf,"%3.2f",Lv); //L
                                                        range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf));
            //y
                        ZeroMemory(buf,sizeof(buf));    sprintf(buf,"%c%d",'U',4+q);
                                                        range = objSheet.GetRange(COleVariant(buf),COleVariant(buf));
             
                        ZeroMemory(buf,sizeof(buf));    Lv=m_IProbe.GetLv();
                                                        sprintf(buf,"%3.2f",Lv); //L
                                                        range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf));
                        break;
            default://W
            //L
                        ZeroMemory(buf,sizeof(buf));    sprintf(buf,"%c%d",'X',4+q);
                                                        range = objSheet.GetRange(COleVariant(buf),COleVariant(buf));
             
                        ZeroMemory(buf,sizeof(buf));    Lv=m_IProbe.GetLv();
                                                        sprintf(buf,"%3.2f",Lv); //L
                                                        range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf));
            //x
                        ZeroMemory(buf,sizeof(buf));    sprintf(buf,"%c%d",'Y',4+q);
                                                        range = objSheet.GetRange(COleVariant(buf),COleVariant(buf));
             
                        ZeroMemory(buf,sizeof(buf));    Lv=m_IProbe.GetLv();
                                                        sprintf(buf,"%3.2f",Lv); //L
                                                        range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf));
            //y
                        ZeroMemory(buf,sizeof(buf));    sprintf(buf,"%c%d",'Z',4+q);
                                                        range = objSheet.GetRange(COleVariant(buf),COleVariant(buf));
             
                        ZeroMemory(buf,sizeof(buf));    Lv=m_IProbe.GetLv();
                                                        sprintf(buf,"%3.2f",Lv); //L
                                                        range.SetItem(_variant_t((long)1),_variant_t((long)1),_variant_t(buf));
                        break;
            }
        }}

    }
    //�q���L�{�����A����ܦ^�Ӫ��F��
    ShowColorController();
    ShowWinControlItem();
    ShowOQCItem();
    ShowMsrItem();
    ShowMsrView();
    ShowReportBtm();
    EnabledReportBtm();
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
//-----------------------------------------------------------------------------------------------
//-----------------------------------------------------------------------------------------------
}
//49�I�q��
void CSwordDlg::OnButton5() 
{
    //�]�w�������F��
    HideColorController();
    HideOQCItem();
    HideMsrItem();
    HideMsrView();
    HideReportBtm();
    HideWinControlItem();
    Msr49P();
//�q���L�{�����A����ܦ^�Ӫ��F��
    ShowColorController();
    ShowOQCItem();
    ShowMsrItem();
    ShowMsrView();
    ShowReportBtm();
    ShowWinControlItem();    
    EnabledReportBtm();
}



//13�I�q��
void CSwordDlg::OnButton6() 
{
//�]�w�������F��
    HideColorController();
    HideOQCItem();
    HideMsrItem();
    HideMsrView();
    HideReportBtm();
    HideWinControlItem();

//    Msr13P();
    //�q���L�{�����A����ܦ^�Ӫ��F��
    ShowColorController();
    ShowOQCItem();
    ShowMsrItem();
    ShowMsrView();
    ShowReportBtm();
    ShowWinControlItem();    
    EnabledReportBtm();    
}


//�q��25�I
void CSwordDlg::OnButton7() 
{
//�]�w�������F��
    HideColorController();
    HideOQCItem();
    HideMsrItem();
    HideMsrView();
    HideReportBtm();
    HideWinControlItem();

    Msr25P();//Msr33P()

    //�q���L�{�����A����ܦ^�Ӫ��F��
    ShowColorController();
    ShowOQCItem();
    ShowMsrItem();
    ShowMsrView();
    ShowReportBtm();
    ShowWinControlItem();    
    EnabledReportBtm();    
}

//FLK�q���]���]����^
//�]flk�ȤӤjerror
//�ҥH�bPattern�|�e��
void CSwordDlg::OnButton8() 
{
    //�]�w�������F��
    HideColorController();
    HideOQCItem();
    HideMsrItem();
    HideMsrView();
    HideWinControlItem();
    HideReportBtm();

    MsrFLIC();

    //�q���L�{�����A����ܦ^�Ӫ��F��
    m_ICa.SetDisplayMode(0);
    ShowColorController();
    ShowWinControlItem();
    ShowOQCItem();
    ShowMsrItem();
    ShowMsrView();
    ShowReportBtm();
    EnabledReportBtm();    
    
}

//����q���Ρ]���α���a�H�^XD
// void CSwordDlg::OnCaAging() 
// {
//     //�]�w�������F��
//     HideColorController();
//     HideOQCItem();
//     HideMsrItem();
//     HideMsrView();
//     HideWinControlItem();
//     HideReportBtm();
// 
//     MsrLab();
// 
//     //�q���L�{�����A����ܦ^�Ӫ��F��
//     ShowColorController();
//     ShowWinControlItem();
//     ShowOQCItem();
//     ShowMsrItem();
//     ShowMsrView();
//     ShowReportBtm();
//     pBtmAgingLabSave-> EnableWindow();
// }

void CSwordDlg::OnSelchangeComboSaveoqc() 
{
    // TODO: Add your control notification handler code here
    FromSelector();
}

void CSwordDlg::OnLoadBarCode() 
{
    // TODO: Add your control notification handler code here
    initValues();
    char szFilters[]="MyType Files (*.txt)|*.txt|All Files (*.*)|*.*||";

    CFileDialog fileDlg (
        TRUE,                                    //bOpenFileDialog    �]�m�� TRUE�G�}�����ɡFFALSE�t�s�s��
        "txt",                                    //lpszDefExt        �ɮ����������e
        "*.txt",                                //lpszFileName        �ɮצW��
        OFN_SHAREAWARE,                            //dwFlags           
        szFilters,                                //szFilters         �s���L�o��
        this                                    //pParentWnd        ����������
                                                //dwSize            �w�]0�G�̧@�~�t�ιw�]����
     );

    // Display the file dialog. When user clicks OK, fileDlg.DoModal() 
    // returns IDOK.
    if (fileDlg.DoModal ()==IDOK) 
    {
        //pOQCStartMsr-> EnableWindow(1);
        //pOQCLoadBarCode->EnableWindow(0);
        TCHAR pbuf[600];
        TCHAR *pch = 0;
        UINT i;
        for(i=0;i<600;++i)
            pbuf[i]=0;

        unsigned int sizeFile;

        CString pathName = fileDlg.GetPathName();
        CString fileName = fileDlg.GetFileName();

        UINT n = pathName.GetLength() - fileName.GetLength();

        BCFandODFPath = pathName.Left(n);
        // Implement opening and reading file in here.
        CFile myFile;

        myFile.Open (pathName, CFile::shareExclusive  | CFile::modeRead);

            sizeFile = myFile.GetLength();//���o�ɮפj�p
            if(sizeFile > 600) 
                //MessageBox("�ɮפj��600�Ӧ줸��");
                MessageBox("�o�ɮשM���`�����@�ˡI�ˬd�@�U�a�I");
            else
                myFile.Read(pbuf, sizeFile);//�ɮפ��e�ಾ�� pbuf
                //*(pbuf+sizeFile)='\0';//�����Ÿ�
        myFile.Close();

        CString BarCodeFile;
        CString LineBuffer;
        BarCodeFile.Format("%s", pbuf);
        BarCodeFile.TrimRight();
        CString temp;

        int bcfBegin = 0, bcfLngth = 0;
        for (i = 0; bcfLngth != -1; ++i)
        {
            bcfLngth = BarCodeFile.Right(BarCodeFile.GetLength() - bcfBegin).Find("\n");
            LineBuffer.Format("%s", BarCodeFile.Mid(bcfBegin, bcfLngth));
            LineBuffer.TrimRight();

            if (bcfLngth - bcfBegin != 0)
            {
                int lBegin = 0, lLength = 0;
                for (int item = 0; /*item < 4 &&*/ lLength != -1; ++item)
                {
                    lLength = LineBuffer.Right(LineBuffer.GetLength() - lBegin).Find("\t");
                    switch (item)
                    {
                        case 0:    strcpy(OQC_Date[i].BarCode,      LineBuffer.Mid(lBegin, lLength));     break;
                        case 1:    strcpy(OQC_Date[i].WorkNum,      LineBuffer.Mid(lBegin, lLength));     break;
                        case 2:    strcpy(OQC_Date[i].DriverDevice, LineBuffer.Mid(lBegin, lLength));     break;
                        case 3:    OQC_Date[i].Current = atoi(LineBuffer.Right(LineBuffer.GetLength() - lBegin));     break;
                    }
                    lBegin = lBegin + lLength + 1;
                }
            }
//             temp.Format("%d", bcfLngth);
//             MessageBox(temp);
            bcfBegin = bcfBegin + bcfLngth + 1;
        }
        temp.Format("���\�פJ%d��BarCode", i);
        MessageBox(temp);
    }
}
void CSwordDlg::OnChangeEditOrigpath() 
{
    // TODO: If this is a RICHEDIT control, the control will not
    // send this notification unless you override the CDialog::OnInitDialog()
    // function and call CRichEditCtrl().SetEventMask()
    // with the ENM_CHANGE flag ORed into the mask.
    
    // TODO: Add your control notification handler code here
    UpdateData(TRUE);
}

void CSwordDlg::OnSelchangeComboSeltai() 
{
    // TODO: Add your control notification handler code here
    if(pOQCTaiSel->GetCurSel() == 10)
        pOQCStartMsr->SetWindowText ("���J�����ɮ�");
    else
        pOQCStartMsr->SetWindowText ("�q��");
}


void CSwordDlg::OnRadioActenter() 
{
    // TODO: Add your control notification handler code here
    MerOperation = FALSE;
    m_iActEnter = MerOperation?1:0;
    m_iAuto = MerOperation?0:1;
    UpdateData(FALSE);
}

void CSwordDlg::OnRadioAuto() 
{
    // TODO: Add your control notification handler code here
    MerOperation = TRUE;
    m_iActEnter = MerOperation?1:0;
    m_iAuto = MerOperation?0:1;
    UpdateData(FALSE);
}


void CSwordDlg::OnAutomsrcal() 
{
    // TODO: Add your control notification handler code here
    OnMax();
    HideWinControlItem();//���å��W������Ӥp���
    OnWhite();
    m_ICa.SetDisplayMode(0);
    CString Msg;
    double MaxValue = Brain, MinValue = Brain;
    //CTime BeginT, EndT;
    time_t BeginT, EndT;
    double DeltaT = 100;
//     Invalidate();
//     UpdateWindow();
    Brain = 0.0;

    do 
    {        
        CurrentPoint = Set9Point();        Pass(0, 0);

        float XFristValue = 0.0, deltaX = 0.0, YFristValue = 0.0, deltaY = 0.0, ZFristValue = 0.0, deltaZ = 0.0;

        BeginT = (float)clock();
        do {
            int flag = 0;
            do{try {m_ICa.Measure(0);    flag = 0;}catch (CException* e) {TCHAR buf[255];e->GetErrorMessage(buf, 255);flag++;}}while (flag);
            deltaX = deltaY = deltaZ = 0.0;
            //��Ѧҭ�
            XFristValue = m_IProbe.GetX();  
            YFristValue = m_IProbe.GetY();  
            ZFristValue = m_IProbe.GetZ();

            flag = 0;
            do{try {m_ICa.Measure(0);    flag = 0;}catch (CException* e) {TCHAR buf[255];e->GetErrorMessage(buf, 255);flag++;}}while (flag);

            //�~�t�������
            deltaX = ((XFristValue-m_IProbe.GetX())>=0) ? XFristValue-m_IProbe.GetX() : m_IProbe.GetX() - XFristValue;
            deltaY = ((YFristValue-m_IProbe.GetY())>=0) ? YFristValue-m_IProbe.GetY() : m_IProbe.GetY() - YFristValue;
            deltaZ = ((ZFristValue-m_IProbe.GetZ())>=0) ? ZFristValue-m_IProbe.GetZ() : m_IProbe.GetZ() - ZFristValue;

            if((deltaX*deltaY*deltaZ) < Brain)
                Pass(1, 0);
            else           
            {
                Pass(0, 0);
                Brain = Brain + 0.0000001;
            }
            ////////////////////////////////////////////////////
            EndT = (float)clock();
            DeltaT = EndT - BeginT;
            Msg.Format(_T("��O�ɶ�:%u��\n%lf\n%lf\n�N��\n"), DeltaT, deltaX*deltaY*deltaZ, Brain);

            CPoint StartDrawPoint(CurrentPoint.x + CircleRadius, CurrentPoint.y - CircleRadius);                //���w�I���r�ϰ쥪�W�����Z��
            CPoint EndDrawPoint(StartDrawPoint.x + 500, StartDrawPoint.y + 100);    //�g�r�϶������I
            CRect rect(StartDrawPoint,EndDrawPoint);          //�]�w��r�϶�
            
            //DeltaT = EndT.GetSecond() - BeginT.GetSecond();
            //DeltaT = difftime(EndT, BeginT);
            
            CDC* pDC = new CClientDC(this);
            
//             InvalidateRect(&rect, TRUE);
             UpdateWindow();
            
            pDC->Rectangle(&rect);
            pDC->SetTextColor(RGB(0,0,0));
            //pDC->SetBkColor(RGB(192, 192, 192));
            pDC->DrawText(TEXT(Msg),-1, &rect, DT_RIGHT | DT_TOP | DT_PATH_ELLIPSIS);
        }
        while(Pass(-1, 0));
        EndT = (float)clock();
        DeltaT = difftime(EndT, BeginT);

//        MessageBox("�q���@��");
            //EndT = CTime::GetCurrentTime();
    } while (DeltaT < 1.0);
    MessageBox(Msg);

    SetBackGroundColor();
    ShowWinControlItem();//��ܥ��W�����F��
    OnMax();
}


