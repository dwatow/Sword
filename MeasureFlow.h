//MeasureFlow.h

class MsrFlow
{
    int few;
    CPoint point;
    
	//材料

	//設定點函數
	//NineO(int few);//畫第幾個圈（以9點為計）
	CPoint SetNineOPoint(int few);//設定第幾點的pixel位置（以九點為計）
    CPoint SetNineOPoint();//預設叫中心點

	//畫東西函數
    void DrawACircle(CPoint point);//畫圈
	void DrawMsrLabel(CPoint point);//預覽資料
};