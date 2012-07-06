/*

void CSwordDlg::ShowColorController()			//���Pattern����
void CSwordDlg::HideColorController()			//����Pattern����
void CSwordDlg::ShowMsrItem()					//��ܶq������
void CSwordDlg::HideMsrItem()					//���öq������
void CSwordDlg::ShowMsrView()					//��ܶq���w�����ء]���Ī��^
void CSwordDlg::HideMsrView()					//���öq���w�����ء]���Ī��^
void CSwordDlg::ShowWinControlItem()			//��ܥ��W���]���̤j�M�^�С^���F��
void CSwordDlg::HideWinControlItem()			//���å��W���]���̤j�M�^�С^���F��
void CSwordDlg::ShowReportBtm()					//��ܿ�XExcel���i����
void CSwordDlg::HideReportBtm()					//���ÿ�XExcel���i����
void CSwordDlg::EnabledReportBtm(BOOL bEnable)	//Enable��XExcel���i����
void CSwordDlg::ShowWinMaximum()				//��j�n��ܪ��F��
void CSwordDlg::HideWinMaximum()				//��j�n�������F��
void CSwordDlg::ShowWinNormal()					//�Y�p�n��ܪ��F��
void CSwordDlg::HideWinNormal()					//�Y�p�n�������F��
*/

#include "stdafx.h"
#include "Sword.h"
#include "SwordDlg.h"

void CSwordDlg::ShowColorController()
{/*
    pBtmRED      -> ShowWindow(SW_SHOW);
    pBtmGREEN    -> ShowWindow(SW_SHOW);
    pBtmBLUE     -> ShowWindow(SW_SHOW);
    pBtmWhite    -> ShowWindow(SW_SHOW);
    pBtmBlack    -> ShowWindow(SW_SHOW);
    pBtmUserColor-> ShowWindow(SW_SHOW);

    pBtmFlkSubPix-> ShowWindow(SW_SHOW);
    pBtmFlk2LInv -> ShowWindow(SW_SHOW);
    pBtmFlk2L1Inv-> ShowWindow(SW_SHOW);
    pBtmFlkVS    -> ShowWindow(SW_SHOW);
    pBtmFlkCol   -> ShowWindow(SW_SHOW);

*/}

void CSwordDlg::HideColorController()
{/*
    pBtmRED      -> ShowWindow(SW_HIDE);
    pBtmGREEN    -> ShowWindow(SW_HIDE);
    pBtmBLUE     -> ShowWindow(SW_HIDE);
    pBtmWhite    -> ShowWindow(SW_HIDE);
    pBtmBlack    -> ShowWindow(SW_HIDE);
    pBtmUserColor-> ShowWindow(SW_HIDE);

    pBtmFlkSubPix-> ShowWindow(SW_HIDE);
    pBtmFlk2LInv -> ShowWindow(SW_HIDE);
    pBtmFlk2L1Inv-> ShowWindow(SW_HIDE);
    pBtmFlkVS    -> ShowWindow(SW_HIDE);
    pBtmFlkCol   -> ShowWindow(SW_HIDE);
*/}

void CSwordDlg::ShowOQCItem()
{
	pOQCTitle          -> ShowWindow(SW_SHOW);
	pOQC5Nits          -> ShowWindow(SW_SHOW);
	pOQCWCenter        -> ShowWindow(SW_SHOW);
	pOQCRCenter        -> ShowWindow(SW_SHOW);
	pOQCGCenter        -> ShowWindow(SW_SHOW);
	pOQCBCenter        -> ShowWindow(SW_SHOW);
	pOQCDCenter        -> ShowWindow(SW_SHOW);
	pOQCW9P            -> ShowWindow(SW_SHOW);
	pOQCW5P            -> ShowWindow(SW_SHOW);
	pOQCR5P            -> ShowWindow(SW_SHOW);
	pOQCG5P            -> ShowWindow(SW_SHOW);
	pOQCB5P            -> ShowWindow(SW_SHOW);
	pOQCD5P            -> ShowWindow(SW_SHOW);
	pOQCW49P            -> ShowWindow(SW_SHOW);
	pOQCD25P           -> ShowWindow(SW_SHOW);
	//pOQCD33P           -> ShowWindow(SW_SHOW);
	//pOQCFlickerStatic  -> ShowWindow(SW_SHOW);
	//pOQCFlickerSel     -> ShowWindow(SW_SHOW);
	pOQCTaiSel         -> ShowWindow(SW_SHOW);
	pOQCStartMsr       -> ShowWindow(SW_SHOW);
//	pBtmAutoMsrCal     -> ShowWindow(SW_SHOW);
    pFE                -> ShowWindow(SW_SHOW);
    pFromEdge          -> ShowWindow(SW_SHOW);
    p5FE               -> ShowWindow(SW_SHOW);
    p5FromEdge         -> ShowWindow(SW_SHOW);
	pOQCFormSel        -> ShowWindow(SW_SHOW);
	psOQCRGB         -> ShowWindow(SW_SHOW);
	psOQCW9P         -> ShowWindow(SW_SHOW);
	psOQCW5P         -> ShowWindow(SW_SHOW);
	psOQCR5P         -> ShowWindow(SW_SHOW);
	psOQCG5P         -> ShowWindow(SW_SHOW);
	psOQCB5P         -> ShowWindow(SW_SHOW);
	psOQCD5P         -> ShowWindow(SW_SHOW);
	psOQCW49P         -> ShowWindow(SW_SHOW);
	psOQCD25P         -> ShowWindow(SW_SHOW);
	pOQCLoadBarCode   -> ShowWindow(SW_SHOW);
	pMsrOperation     -> ShowWindow(SW_SHOW);
	pActEnter         -> ShowWindow(SW_SHOW);
	pAutoMsr          -> ShowWindow(SW_SHOW);
}

void CSwordDlg::HideOQCItem()
{
	pOQCTitle          -> ShowWindow(SW_HIDE);
	pOQC5Nits          -> ShowWindow(SW_HIDE);
	pOQCWCenter        -> ShowWindow(SW_HIDE);
	pOQCRCenter        -> ShowWindow(SW_HIDE);
	pOQCGCenter        -> ShowWindow(SW_HIDE);
	pOQCBCenter        -> ShowWindow(SW_HIDE);
	pOQCDCenter        -> ShowWindow(SW_HIDE);
	pOQCW9P            -> ShowWindow(SW_HIDE);
	pOQCW5P            -> ShowWindow(SW_HIDE);
	pOQCR5P            -> ShowWindow(SW_HIDE);
	pOQCG5P            -> ShowWindow(SW_HIDE);
	pOQCB5P            -> ShowWindow(SW_HIDE);
	pOQCD5P            -> ShowWindow(SW_HIDE);
	pOQCW49P            -> ShowWindow(SW_HIDE);
	pOQCD25P           -> ShowWindow(SW_HIDE);
	//pOQCD33P           -> ShowWindow(SW_HIDE);
	//pOQCFlickerStatic  -> ShowWindow(SW_HIDE);
	//pOQCFlickerSel     -> ShowWindow(SW_HIDE);
	pOQCTaiSel         -> ShowWindow(SW_HIDE);
	pOQCStartMsr       -> ShowWindow(SW_HIDE);
//	pBtmAutoMsrCal     -> ShowWindow(SW_HIDE);
    pFE                -> ShowWindow(SW_HIDE);
    pFromEdge          -> ShowWindow(SW_HIDE);
    p5FE               -> ShowWindow(SW_HIDE);
    p5FromEdge         -> ShowWindow(SW_HIDE);
	pOQCFormSel        -> ShowWindow(SW_HIDE);
	psOQCRGB         -> ShowWindow(SW_HIDE);
	psOQCW9P         -> ShowWindow(SW_HIDE);
	psOQCW5P         -> ShowWindow(SW_HIDE);
	psOQCR5P         -> ShowWindow(SW_HIDE);
	psOQCG5P         -> ShowWindow(SW_HIDE);
	psOQCB5P         -> ShowWindow(SW_HIDE);
	psOQCD5P         -> ShowWindow(SW_HIDE);
	psOQCW49P         -> ShowWindow(SW_HIDE);
	psOQCD25P         -> ShowWindow(SW_HIDE);
	pOQCLoadBarCode   -> ShowWindow(SW_HIDE);
	pMsrOperation     -> ShowWindow(SW_HIDE);
	pActEnter         -> ShowWindow(SW_HIDE);
	pAutoMsr          -> ShowWindow(SW_HIDE);
}

void CSwordDlg::ShowMsrItem()
{/*
    pEditBackUalue      -> ShowWindow(SW_SHOW);
    pFE                 -> ShowWindow(SW_SHOW);
    pFromEdge           -> ShowWindow(SW_SHOW);
    pBtmNineO           -> ShowWindow(SW_SHOW);
    pBtmCentrO          -> ShowWindow(SW_SHOW);
    pBtmFiveNits        -> ShowWindow(SW_SHOW);
    pBtmGamma           -> ShowWindow(SW_SHOW);
    pMeasureFlowList    -> ShowWindow(SW_SHOW);
    pBtmAgingLab        -> ShowWindow(SW_SHOW);
    pBtmFNO             -> ShowWindow(SW_SHOW);
    pBtmThirteenO       -> ShowWindow(SW_SHOW);
    pBtmTwentyFiveO     -> ShowWindow(SW_SHOW);
    pBtmFlkMsr          -> ShowWindow(SW_SHOW);

    p1Cvr9CheckBox      -> ShowWindow(SW_SHOW);
    p1Cvr49CheckBox     -> ShowWindow(SW_SHOW);
    p9Cvr1CheckBox      -> ShowWindow(SW_SHOW);
    p9Cvr49CheckBox     -> ShowWindow(SW_SHOW);
    p49Cvr1CheckBox     -> ShowWindow(SW_SHOW);
    p49Cvr9CheckBox     -> ShowWindow(SW_SHOW);
*/}

void CSwordDlg::HideMsrItem()
{/*
    pEditBackUalue      -> ShowWindow(SW_HIDE);
    pFE                 -> ShowWindow(SW_HIDE);
    pFromEdge           -> ShowWindow(SW_HIDE);
    pBtmNineO           -> ShowWindow(SW_HIDE);
    pBtmCentrO          -> ShowWindow(SW_HIDE);
    pBtmFiveNits        -> ShowWindow(SW_HIDE);
    pBtmGamma           -> ShowWindow(SW_HIDE);
    pMeasureFlowList    -> ShowWindow(SW_HIDE);
    pBtmAgingLab        -> ShowWindow(SW_HIDE);
    pBtmFNO             -> ShowWindow(SW_HIDE);
    pBtmThirteenO       -> ShowWindow(SW_HIDE);
    pBtmTwentyFiveO     -> ShowWindow(SW_HIDE);
    pBtmFlkMsr          -> ShowWindow(SW_HIDE);

    p1Cvr9CheckBox      -> ShowWindow(SW_HIDE);
    p1Cvr49CheckBox     -> ShowWindow(SW_HIDE);
    p9Cvr1CheckBox      -> ShowWindow(SW_HIDE);
    p9Cvr49CheckBox     -> ShowWindow(SW_HIDE);
    p49Cvr1CheckBox     -> ShowWindow(SW_HIDE);
    p49Cvr9CheckBox     -> ShowWindow(SW_HIDE);
*/}

void CSwordDlg::ShowMsrView()
{
    pMsrViewList        -> ShowWindow(SW_SHOW);
    pLvCheckBox         -> ShowWindow(SW_SHOW);
    pSXCheckBox         -> ShowWindow(SW_SHOW);
    pSYCheckBox         -> ShowWindow(SW_SHOW);
    pTCheckBox          -> ShowWindow(SW_SHOW);
    pDUVCheckBox        -> ShowWindow(SW_SHOW);
    pDUCheckBox         -> ShowWindow(SW_SHOW);
    pDVCheckBox         -> ShowWindow(SW_SHOW);
    pXCheckBox          -> ShowWindow(SW_SHOW);
    pYCheckBox          -> ShowWindow(SW_SHOW);
    pZCheckBox          -> ShowWindow(SW_SHOW);
    //pFlkCheckBox        -> ShowWindow(SW_SHOW);
}

void CSwordDlg::HideMsrView()
{
	pMsrViewList        -> ShowWindow(SW_HIDE);
    pLvCheckBox         -> ShowWindow(SW_HIDE);
    pSXCheckBox         -> ShowWindow(SW_HIDE);
    pSYCheckBox         -> ShowWindow(SW_HIDE);
    pTCheckBox          -> ShowWindow(SW_HIDE);
    pDUVCheckBox        -> ShowWindow(SW_HIDE);
    pDUCheckBox         -> ShowWindow(SW_HIDE);
    pDVCheckBox         -> ShowWindow(SW_HIDE);
    pXCheckBox          -> ShowWindow(SW_HIDE);
    pYCheckBox          -> ShowWindow(SW_HIDE);
    pZCheckBox          -> ShowWindow(SW_HIDE);
    //pFlkCheckBox        -> ShowWindow(SW_HIDE);
}

void CSwordDlg::ShowWinControlItem()
{
    pBtmMax        -> ShowWindow(SW_SHOW);
    pBtmOK         -> ShowWindow(SW_SHOW);
}
void CSwordDlg::HideWinControlItem()
{
    pBtmMax        -> ShowWindow(SW_HIDE);
    pBtmOK         -> ShowWindow(SW_HIDE);
}
void CSwordDlg::ShowReportBtm()//��ܿ�XExcel���i����
{/*
    pRePortList  -> ShowWindow(SW_SHOW);
    pBtmSave1    -> ShowWindow(SW_SHOW);
    pBtmSave2    -> ShowWindow(SW_SHOW);
    pBtmAgingLabSave-> ShowWindow(SW_SHOW);
    pBtmSave13p  -> ShowWindow(SW_SHOW);
    pBtmSave25p  -> ShowWindow(SW_SHOW);
    pBtmSaveT26  -> ShowWindow(SW_SHOW);
    pBtmRstDate  -> ShowWindow(SW_SHOW);
*/	pBtmSaveOQC  -> ShowWindow(SW_SHOW);

}
void CSwordDlg::HideReportBtm()//���ÿ�XExcel���i����
{/*
    pRePortList  -> ShowWindow(SW_HIDE);
    pBtmSave1    -> ShowWindow(SW_HIDE);
    pBtmSave2    -> ShowWindow(SW_HIDE);
    pBtmAgingLabSave-> ShowWindow(SW_HIDE);
    pBtmSave13p  -> ShowWindow(SW_HIDE);
    pBtmSave25p  -> ShowWindow(SW_HIDE);
    pBtmSaveT26  -> ShowWindow(SW_HIDE);
    pBtmRstDate  -> ShowWindow(SW_HIDE);
*/	pBtmSaveOQC  -> ShowWindow(SW_HIDE);
}

void CSwordDlg::EnabledReportBtm(BOOL bEnable)//�}���XExcel���i����
{/*
    pBtmSave1    -> EnableWindow(bEnable);
    pBtmSave2    -> EnableWindow(bEnable);
    pBtmSave13p  -> EnableWindow(bEnable);
    pBtmSave25p  -> EnableWindow(bEnable);
    pBtmSaveT26  -> EnableWindow(bEnable);
    pBtmRstDate  -> EnableWindow(bEnable);
*/	pBtmSaveOQC  -> EnableWindow(bEnable);
}

void CSwordDlg::ShowWinMaximum()//��j�n��ܪ��F��
{
/*
	ShowOQCItem();
    ShowMsrItem();
    ShowMsrView();
    ShowWinControlItem();
    ShowColorController();
*/}
void CSwordDlg::HideWinMaximum()//��j�n�������F��
{
    m_CCaControl.ShowWindow(SW_HIDE);
    pTextRredMe -> ShowWindow(SW_HIDE);//�@���N�������u���u�����̤j�v
//    pBtmZeroCal -> ShowWindow(SW_HIDE);
//    ShowReportBtm();

	HideColorController();
	HideOQCItem();
    HideMsrItem();
    HideMsrView();
    HideReportBtm();
}

void CSwordDlg::ShowWinNormal()//�Y�p�n��ܪ��F��
{
    m_CCaControl.ShowWindow(SW_SHOW);
    pTextRredMe -> ShowWindow(SW_SHOW);
 //   pBtmZeroCal -> ShowWindow(SW_SHOW);
    //ShowWinControlItem();

	ShowReportBtm();
	ShowOQCItem();
    ShowMsrItem();
    ShowMsrView();
    ShowWinControlItem();
    ShowColorController();
}

void CSwordDlg::HideWinNormal()//�Y�p�n�������F��
{
/*
    HideColorController();
	HideOQCItem();
    HideMsrItem();
    HideMsrView();
    HideReportBtm();
*/
}