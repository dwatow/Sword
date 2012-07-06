#include "stdafx.h"
#include "Sword.h"
#include "SwordDlg.h"
/*
readme
CPoint CSwordDlg::Set9Point(int few, float FromEdge)	9點量測
CPoint CSwordDlg::SetD9Point(int few, float FromEdge)	暗態9點
CPoint CSwordDlg::Set49Point(int few)					49點
CPoint CSwordDlg::SetD25Point(int few)					暗態25點
CPoint CSwordDlg::SetD13Point(int few)					暗態13點
*/

CPoint CSwordDlg::Set9Point(int few, float FromEdge)
{
//運算第幾個（以九點為計）

    //ScrmV 螢幕垂直pixel數
    //ScrmH 螢幕水平pixel數
    int LeftEdge   = static_cast<int>(ScrmH / FromEdge);//六分之一左邊緣
    int TopEdge    = static_cast<int>(ScrmV / FromEdge);//六分之一上邊緣
    int RightEdge  = ScrmH - LeftEdge;
    int BottomEdge = ScrmV - TopEdge;
    int CenterH    = ScrmH/2;
    int CenterV    = ScrmV/2;
/*
+------------------------------+
|                              |
|    00        01        02    |
|                              |
|                              |
|    03        04        05    |
|                              |
|                              |
|    06        07        08    |
|                              |
+-----------------------------+
*/
    CPoint 
        Point0(LeftEdge ,TopEdge),
        Point1(CenterH  ,TopEdge),
        Point2(RightEdge,TopEdge),
        Point3(LeftEdge ,CenterV),
        Point4(CenterH  ,CenterV),
        Point5(RightEdge,CenterV),
        Point6(LeftEdge ,BottomEdge),
        Point7(CenterH  ,BottomEdge),
        Point8(RightEdge,BottomEdge),
        PointD(CenterH  ,CenterV);

//回傳一個點
    switch(few)
    {
        case 0: return Point0; break;
        case 1: return Point1; break;
        case 2: return Point2; break;

        case 3: return Point3; break;
        case 4: return Point4; break;//中心
        case 5: return Point5; break;

        case 6: return Point6; break;
        case 7: return Point7; break;
        case 8: return Point8; break;

        default: PointD; return PointD; break;
    }
}

// CPoint CSwordDlg::Set5Point(int few)
// {
// //運算第幾個（以九點為計）
// 
//     //ScrmV 螢幕垂直pixel數
//     //ScrmH 螢幕水平pixel數
//     int LeftEdge   = CircleRadius;//左邊緣
//     int TopEdge    = CircleRadius;//上邊緣
//     int RightEdge  = ScrmH - LeftEdge;
//     int BottomEdge = ScrmV - TopEdge;
//     int CenterH    = ScrmH/2;
//     int CenterV    = ScrmV/2;
// /*
// +------------------------------+
// |00                          01|
// |                              |  
// |                              |  
// |                              |
// |              02              |
// |                              |
// |                              |  
// |                              |  
// |03                          05|
// +------------------------------+
// */
// 
//     CPoint 
//         Point0(LeftEdge ,TopEdge),
//         Point1(RightEdge,TopEdge),
//         Point2(CenterH  ,CenterV),
//         Point3(LeftEdge ,BottomEdge),
//         Point4(RightEdge,BottomEdge),
//         PointD(CenterH  ,CenterV);
// //回傳一個點
//     switch(few)
//     {
//         case 0: return Point0; break;
//         case 1: return Point1; break;
//         case 2: return Point2; break;//中心
//         case 3: return Point3; break;
//         case 4: return Point4; break;
// 
//         default: PointD; return PointD; break;
//     }
// }

CPoint CSwordDlg::Set5Point(int few, float FromEdge)
{
//運算第幾個（以五點為計）

    //ScrmV 螢幕垂直pixel數
    //ScrmH 螢幕水平pixel數
    int LeftEdge   = FromEdge ? static_cast<int>(ScrmH / FromEdge) : CircleRadius;//左邊緣
    int TopEdge    = FromEdge ? static_cast<int>(ScrmV / FromEdge) : CircleRadius;//上邊緣
    int RightEdge  = ScrmH - LeftEdge;
    int BottomEdge = ScrmV - TopEdge;
    int CenterH    = ScrmH/2;
    int CenterV    = ScrmV/2;
/*
+------------------------------+
|00                          01|
|                              |  
|                              |  
|                              |
|              02              |
|                              |
|                              |  
|                              |  
|03                          05|
+------------------------------+
*/

    CPoint 
        Point0(LeftEdge ,TopEdge),
        Point1(RightEdge,TopEdge),
        Point2(CenterH  ,CenterV),
        Point3(LeftEdge ,BottomEdge),
        Point4(RightEdge,BottomEdge),
        PointD(CenterH  ,CenterV);
//回傳一個點
    switch(few)
    {
        case 0: return Point0; break;
        case 1: return Point1; break;
        case 2: return Point2; break;//中心
        case 3: return Point3; break;
        case 4: return Point4; break;

        default: PointD; return PointD; break;
    }
}

CPoint CSwordDlg::Set5nits9Point(int few, float FromEdge)
{
//運算第幾個（以九點為計）

    //ScrmV 螢幕垂直pixel數
    //ScrmH 螢幕水平pixel數
    int LeftEdge   = CmtoPixel(10);
    int TopEdge    = static_cast<int>(ScrmV / FromEdge);//六分之一上邊緣
    int RightEdge  = ScrmH - LeftEdge;
    int BottomEdge = ScrmV - TopEdge;
    int CenterH    = ScrmH/2;
    int CenterV    = ScrmV/2;
/*5nits
+------------------------------+
|                              |
|    00        01        02    |
|                              |
|                              |
|    03        04        05    |
|                              |
|                              |
|    06        07        08    |
|                              |
+-----------------------------+
*/
    CPoint 
        Point0(LeftEdge ,TopEdge),
        Point1(CenterH  ,TopEdge),
        Point2(RightEdge,TopEdge),
        Point3(LeftEdge ,CenterV),
        Point4(CenterH  ,CenterV),  //中心點不量
        Point5(RightEdge,CenterV),
        Point6(LeftEdge ,BottomEdge),
        Point7(CenterH  ,BottomEdge),
        Point8(RightEdge,BottomEdge),

        PointD(CenterH  ,CenterV);

//回傳一個點
    switch(few)
    {
        case 0: return Point0; break;
        case 1: return Point1; break;
        case 2: return Point2; break;

        case 3: return Point3; break;
        case 4: return Point4; break;//中心
        case 5: return Point5; break;

        case 6: return Point6; break;
        case 7: return Point7; break;
		case 8: return Point8; break;

        default: PointD; return PointD; break;
    }
}


CPoint CSwordDlg::SetD9Point(int few)
{
//運算第幾個（以九點為計）

    //ScrmV 螢幕垂直pixel數
    //ScrmH 螢幕水平pixel數
    int LeftEdge   = CircleRadius;//左邊緣
    int TopEdge    = CircleRadius;//上邊緣
    int RightEdge  = ScrmH - LeftEdge;
    int BottomEdge = ScrmV - TopEdge;
    int CenterH    = ScrmH/2;
    int CenterV    = ScrmV/2;
/*
+------------------------------+
|00            01            02|
|                              |  
|                              |  
|                              |
|03            04            05|
|                              |
|                              |  
|                              |  
|06            07            08|
+------------------------------+
*/

    CPoint 
        Point0(LeftEdge ,TopEdge),
        Point1(CenterH  ,TopEdge),
        Point2(RightEdge,TopEdge),
        Point3(LeftEdge ,CenterV),
        Point4(CenterH  ,CenterV),
        Point5(RightEdge,CenterV),
        Point6(LeftEdge ,BottomEdge),
        Point7(CenterH  ,BottomEdge),
        Point8(RightEdge,BottomEdge),
        PointD(CenterH  ,CenterV);
//回傳一個點
    switch(few)
    {
        case 0: return Point0; break;
        case 1: return Point1; break;
        case 2: return Point2; break;

        case 3: return Point3; break;
        case 4: return Point4; break;//中心
        case 5: return Point5; break;

        case 6: return Point6; break;
        case 7: return Point7; break;
        case 8: return Point8; break;

        default: PointD; return PointD; break;
    }
}


// CPoint CSwordDlg::SetD13Point(int few)
// {
// //運算第幾個（以九點為計）
// 
//     //ScrmV 螢幕垂直pixel數
//     //ScrmH 螢幕水平pixel數
//     int LeftEdge  = CircleRadius;//左邊緣
//     int TopEdge   = CircleRadius;//上邊緣
// 
//     int RightEdge  = ScrmH - LeftEdge;
//     int BottomEdge = ScrmV - TopEdge;
// 
//     int CenterH = ScrmH/2;
//     int CenterV = ScrmV/2;
// 
//     int LeftQuarterH = ScrmH/4;
//     int TopQuarterV  = ScrmV/4;
// 
//     int RightQuarterH   = ScrmH - ScrmH/4;
//     int BottomQuarterV  = ScrmV - ScrmV/4;
// /*
// +------------------------------+
// |00            01            02|
// |                              |  
// |       09            10       |  
// |                              |
// |03            04            05|
// |                              |
// |       11            12       |  
// |                              |  
// |06            07            08|
// +------------------------------+
// */
//     CPoint 
//         Point00(LeftEdge ,TopEdge),
//         Point01(CenterH  ,TopEdge),
//         Point02(RightEdge,TopEdge),
// 
//         Point03(LeftEdge ,CenterV),
//         Point04(CenterH  ,CenterV),
//         Point05(RightEdge,CenterV),
// 
//         Point06(LeftEdge ,BottomEdge),
//         Point07(CenterH  ,BottomEdge),
//         Point08(RightEdge,BottomEdge),
// 
//             Point09(LeftQuarterH ,TopQuarterV),
//             Point10(RightQuarterH  ,TopQuarterV),
// 
//             Point11(LeftQuarterH,BottomQuarterV),
//             Point12(RightQuarterH ,BottomQuarterV),
// 
//         PointD(CenterH  ,CenterV);
// 
// //回傳一個點
//     switch(few)
//     {
//         case 0: return Point00; break;
//         case 1: return Point01; break;
//         case 2: return Point02; break;
// 
//         case 3: return Point03; break;
//         case 4: return Point04; break;//中心
//         case 5: return Point05; break;
// 
//         case 6: return Point06; break;
//         case 7: return Point07; break;
//         case 8: return Point08; break;
// 
// 		    case 9:  return Point09; break;
// 			case 10: return Point10; break;
// 			case 11: return Point11; break;
// 			case 12: return Point12; break;
// 
//         default: PointD; return PointD; break;
//     }
// }


CPoint CSwordDlg::SetD25Point(int few)
{
    //ScrmV 螢幕垂直pixel數
    //ScrmH 螢幕水平pixel數

    int LeftEdge   = CircleRadius;
    int Left_10    = LeftEdge + CmtoPixel(5);
    int Left_20    = Left_10  + CmtoPixel(5);

    int RightEdge  = ScrmH - LeftEdge;
    int Right_10   = RightEdge - CmtoPixel(5);
    int Right_20   = Right_10  - CmtoPixel(5);
    
    int CenterH    = ScrmH/2;
    int CenterV    = ScrmV/2;

    int TopEdge    = CircleRadius;//上邊緣
    int Top_10     = TopEdge + CmtoPixel(5);
    int Top_20     = Top_10  + CmtoPixel(5);

    int BottomEdge = ScrmV - TopEdge;
    int Bottom_10  = BottomEdge - CmtoPixel(5);
    int Bottom_20  = Bottom_10  - CmtoPixel(5);

/* L1L2                  R2R1
+------------------------------+
|00  01        05        07  06|
|  04                      10  |T1
|02  03                  09  08|T2
|                              |
|11            12            13|
|                              |
|16  17                  23  25|B2
|  18                      24  |B1
|14  15        19        21  20|
+------------------------------+
*/
    CPoint 
        Point00(LeftEdge, TopEdge),        Point01(Left_20, TopEdge),

                    Point04(Left_10, Top_10),

        Point02(LeftEdge, Top_20),        Point03(Left_20, Top_20),
                        
                            
        Point05(CenterH, TopEdge),


        Point07(Right_20, TopEdge),        Point06(RightEdge, TopEdge),
        
                    Point10(Right_10 ,Top_10),

        Point09(Right_20, Top_20),        Point08(RightEdge,Top_20),
        //--------------------
        Point11(LeftEdge, CenterV),

        Point12(CenterH, CenterV),

        Point13(RightEdge, CenterV),
        //--------------------
        Point16(LeftEdge ,Bottom_20),    Point17(Left_20    ,Bottom_20),

                    Point18(Left_10, Bottom_10),

        Point14(LeftEdge,BottomEdge),    Point15(Left_20    ,BottomEdge),
        
        Point19(CenterH ,BottomEdge),

        Point23(Right_20, Bottom_20),    Point22(RightEdge  ,Bottom_20),

                    Point24(Right_10, Bottom_10),

        Point21(Right_20 ,BottomEdge),    Point20(RightEdge  ,BottomEdge),
        //--------------------        
        PointD(CenterH  ,CenterV);

//回傳一個點
    switch(few)
    {
        case 0: return Point00; break;
        case 1: return Point01; break;
        case 2: return Point02; break;
        case 3: return Point03; break;
        case 4: return Point04; break;

        case 5: return Point05; break;

        case 6: return Point06; break;
        case 7: return Point07; break;
        case 8: return Point08; break;
        case 9: return Point09; break;
        case 10: return Point10; break;
//----------
        case 11: return Point11; break;

        case 12: return Point12; break;//中心

        case 13: return Point13; break;
//----------
        case 14: return Point14; break;
        case 15: return Point15; break;
        case 16: return Point16; break;
        case 17: return Point17; break;
        case 18: return Point18; break;

        case 19: return Point19; break;
        
        case 20: return Point20; break;
        case 21: return Point21; break;
        case 22: return Point22; break;
        case 23: return Point23; break;
        case 24: return Point24; break;

        default: PointD; return PointD; break;
    }
}

//25+13的組合
CPoint CSwordDlg::SetD29Point(int few)
{
//運算第幾個（以九點為計）

    //ScrmV 螢幕垂直pixel數
    //ScrmH 螢幕水平pixel數

    int LeftEdge   = CircleRadius;
    int Left_10    = LeftEdge + CmtoPixel(5);
    int Left_20    = Left_10  + CmtoPixel(5);

    int RightEdge  = ScrmH - LeftEdge;
    int Right_10   = RightEdge - CmtoPixel(5);
    int Right_20   = Right_10  - CmtoPixel(5);
    
    int CenterH    = ScrmH/2;
    int CenterV    = ScrmV/2;

    int TopEdge    = CircleRadius;//上邊緣
    int Top_10     = TopEdge + CmtoPixel(5);
    int Top_20     = Top_10  + CmtoPixel(5);

    int BottomEdge = ScrmV - TopEdge;
    int Bottom_10  = BottomEdge - CmtoPixel(5);
    int Bottom_20  = Bottom_10  - CmtoPixel(5);

    int LeftQuarterH = ScrmH/4;
    int TopQuarterV  = ScrmV/4;

    int RightQuarterH   = ScrmH - ScrmH/4;
    int BottomQuarterV  = ScrmV - ScrmV/4;
/* L1L2                  R2R1
+------------------------------+
|00  01        05        07  06|
|  04                      10  |T1
|02  03  25          26  09  08|T2
|                              |
|11            12            13|
|                              |
|16  17  27          28  23  25|B2
|  18                      24  |B1
|14  15        19        21  20|
+------------------------------+
*/
    CPoint 
        Point00(LeftEdge, TopEdge),        Point01(Left_20, TopEdge),

                    Point04(Left_10, Top_10),

        Point02(LeftEdge, Top_20),        Point03(Left_20, Top_20),
                        
                            
        Point05(CenterH, TopEdge),


        Point07(Right_20, TopEdge),        Point06(RightEdge, TopEdge),
        
                    Point10(Right_10 ,Top_10),

        Point09(Right_20, Top_20),        Point08(RightEdge,Top_20),
        //--------------------
        Point11(LeftEdge, CenterV),

        Point12(CenterH, CenterV),

        Point13(RightEdge, CenterV),
        //--------------------
        Point16(LeftEdge ,Bottom_20),    Point17(Left_20    ,Bottom_20),

                    Point18(Left_10, Bottom_10),

        Point14(LeftEdge,BottomEdge),    Point15(Left_20    ,BottomEdge),
        
        Point19(CenterH ,BottomEdge),

        Point23(Right_20, Bottom_20),    Point22(RightEdge  ,Bottom_20),

                    Point24(Right_10, Bottom_10),

        Point21(Right_20 ,BottomEdge),    Point20(RightEdge  ,BottomEdge),
        //--------------------        
        Point25(LeftQuarterH ,TopQuarterV),        Point26(RightQuarterH  ,TopQuarterV),

        Point27(LeftQuarterH,BottomQuarterV),        Point28(RightQuarterH ,BottomQuarterV),        
        //--------------------  
		PointD(CenterH  ,CenterV);

//回傳一個點
    switch(few)
    {
        case 0: return Point00; break;
        case 1: return Point01; break;
        case 2: return Point02; break;
        case 3: return Point03; break;
        case 4: return Point04; break;

        case 5: return Point05; break;

        case 6: return Point06; break;
        case 7: return Point07; break;
        case 8: return Point08; break;
        case 9: return Point09; break;
        case 10: return Point10; break;
//----------
        case 11: return Point11; break;

        case 12: return Point12; break;//中心

        case 13: return Point13; break;
//----------
        case 14: return Point14; break;
        case 15: return Point15; break;
        case 16: return Point16; break;
        case 17: return Point17; break;
        case 18: return Point18; break;

        case 19: return Point19; break;
        
        case 20: return Point20; break;
        case 21: return Point21; break;
        case 22: return Point22; break;
        case 23: return Point23; break;
        case 24: return Point24; break;
//----------
		case 25: return Point25; break;
        case 26: return Point26; break;
		case 27: return Point27; break;
        case 28: return Point28; break;

        default: PointD; return PointD; break;
    }
}


//New!29+13的組合
CPoint CSwordDlg::SetD33Point(int few)
{
//運算第幾個（以九點為計）

    //ScrmV 螢幕垂直pixel數
    //ScrmH 螢幕水平pixel數

    int LeftEdge   = CmtoPixel(4.5);
    int Left_10    = LeftEdge + CmtoPixel(9);
    int Left_20    = Left_10  + CmtoPixel(9);

    int RightEdge  = ScrmH - CmtoPixel(4.5);
    int Right_10   = RightEdge - CmtoPixel(9);
    int Right_20   = Right_10  - CmtoPixel(9);

    int TopEdge    = CmtoPixel(4.5);
    int Top_10     = TopEdge + CmtoPixel(9);
    int Top_20     = Top_10  + CmtoPixel(9);

    int BottomEdge = ScrmV - CmtoPixel(4.5);
    int Bottom_10  = BottomEdge - CmtoPixel(9);
    int Bottom_20  = Bottom_10  - CmtoPixel(9);
    
    int CenterH    = ScrmH/2;
    int CenterV    = ScrmV/2;

    int LeftQuarterH = ScrmH/4;
    int TopQuarterV  = ScrmV/4;

    int RightQuarterH   = ScrmH - ScrmH/4;
    int BottomQuarterV  = ScrmV - ScrmV/4;

/* L1L2                  R2R1
+------------------------------+
|000102        06        090807|
|0304                      1110|T1
|05      29          30  09  12|T2
|                              |
|13            14            15|
|                              |
|21      31          32      28|B2
|1920                      2726|B1
|161718        22        252423|
+------------------------------+
*/
    CPoint 
        Point00(LeftEdge, TopEdge),    Point01(Left_10, TopEdge),       Point02(Left_20, TopEdge),
		Point03(LeftEdge,  Top_10),    Point04(Left_10, Top_10),
        Point05(LeftEdge,  Top_20),
        
        Point06(CenterH, TopEdge),
                            
        Point09(Right_20, TopEdge),    Point08(Right_10,TopEdge),       Point07(RightEdge, TopEdge),        
                                       Point11(Right_10, Top_10),       Point10(RightEdge ,Top_10),
                                                                        Point12(RightEdge, Top_20),
                
        //--------------------
        Point13(LeftEdge, CenterV),    Point14(CenterH, CenterV),       Point15(RightEdge, CenterV),
        //--------------------
        Point21(LeftEdge ,Bottom_20),
        Point19(LeftEdge ,Bottom_10),  Point20(Left_10  ,Bottom_10),
        Point16(LeftEdge, BottomEdge), Point17(Left_10, BottomEdge),    Point18(Left_20, BottomEdge),

        Point22(CenterH  ,BottomEdge),

                                                                        Point28(RightEdge ,Bottom_20),
                                        Point26(Right_10  ,Bottom_10),  Point27(RightEdge, Bottom_10),
        Point25(Right_20 ,BottomEdge),  Point24(Right_10, BottomEdge),  Point23(RightEdge, BottomEdge),    
    
        //--------------------        
        Point29(LeftQuarterH ,TopQuarterV),        Point30(RightQuarterH  ,TopQuarterV),

        Point31(LeftQuarterH,BottomQuarterV),        Point32(RightQuarterH ,BottomQuarterV),        
        //--------------------  

		PointD(CenterH  ,CenterV);

//回傳一個點
    switch(few)
    {
        case 0: return Point00; break;
        case 1: return Point01; break;
        case 2: return Point02; break;
        case 3: return Point03; break;
        case 4: return Point04; break;

        case 5: return Point05; break;

        case 6: return Point06; break;
        case 7: return Point07; break;
        case 8: return Point08; break;
        case 9: return Point09; break;
        case 10: return Point10; break;
//----------
        case 11: return Point11; break;

        case 12: return Point12; break;//中心

        case 13: return Point13; break;
//----------
        case 14: return Point14; break;
        case 15: return Point15; break;
        case 16: return Point16; break;
        case 17: return Point17; break;
        case 18: return Point18; break;

        case 19: return Point19; break;
        
        case 20: return Point20; break;
        case 21: return Point21; break;
        case 22: return Point22; break;
        case 23: return Point23; break;
        case 24: return Point24; break;
//----------
		case 25: return Point25; break;
        case 26: return Point26; break;
		case 27: return Point27; break;
        case 28: return Point28; break;
//----------
		case 29: return Point29; break;
        case 30: return Point30; break;
		case 31: return Point31; break;
        case 32: return Point32; break;
//----------
        default: PointD; return PointD; break;
    }
}

CPoint CSwordDlg::Set49Point(int few)
{
//運算第幾個（以49點為計）

    //ScrmV 螢幕垂直pixel數
    //ScrmH 螢幕水平pixel數
    int LeftEdge  =  ScrmH / 12;   //左數1直排
    int DLeftEdge =  LeftEdge*2;   //左數2直排
    int FLeftEdge = DLeftEdge*2;   //左數3直排

    int TopEdge   = ScrmV / 12;    //上數1橫排
    int DTopEdge  =  TopEdge*2;    //上數2橫排
    int FTopEdge  = DTopEdge*2;    //上數2橫排

    int RightEdge  = ScrmH -  LeftEdge;   //右數1直排
    int DRightEdge = ScrmH - DLeftEdge;   //右數2直排
    int FRightEdge = ScrmH - FLeftEdge;   //右數3直排

    int BottomEdge = ScrmV -  TopEdge;    //下數1橫排
    int DBottomEdge= ScrmV - DTopEdge;    //下數2橫排
    int FBottomEdge= ScrmV - FTopEdge;    //下數3橫排

    int CenterH = ScrmH/2;  //水平中心
    int CenterV = ScrmV/2;  //垂直中心
/* L1L2                  R2R1
+------------------------------+
|                              |
| 00 01  02    03    04  05 06 |T1
| 10 11  12    13    14  15 16 |T2
| 20 21  22    23    24  25 26 |
| 30 31  32    33    34  35 36 |
| 40 41  42    43    44  45 46 |
| 50 51  52    53    54  55 56 |B2
| 60 61  62    63    64  65 66 |B1
|                              |
+------------------------------+
*/
    CPoint 
        Point00(LeftEdge  ,TopEdge),
        Point01(DLeftEdge ,TopEdge),
        Point02(FLeftEdge ,TopEdge),
        Point03(CenterH   ,TopEdge),
        Point04(FRightEdge,TopEdge),
        Point05(DRightEdge,TopEdge),
        Point06(RightEdge ,TopEdge),

        Point10(LeftEdge  ,DTopEdge),
        Point11(DLeftEdge ,DTopEdge),
        Point12(FLeftEdge ,DTopEdge),
        Point13(CenterH   ,DTopEdge),
        Point14(FRightEdge,DTopEdge),
        Point15(DRightEdge,DTopEdge),
        Point16(RightEdge ,DTopEdge),

        Point20(LeftEdge  ,FTopEdge),
        Point21(DLeftEdge ,FTopEdge),
        Point22(FLeftEdge ,FTopEdge),
        Point23(CenterH   ,FTopEdge),
        Point24(FRightEdge,FTopEdge),
        Point25(DRightEdge,FTopEdge),
        Point26(RightEdge ,FTopEdge),

        Point30(LeftEdge  ,CenterV),
        Point31(DLeftEdge ,CenterV),
        Point32(FLeftEdge ,CenterV),
        Point33(CenterH   ,CenterV),//中心點
        Point34(FRightEdge,CenterV),
        Point35(DRightEdge,CenterV),
        Point36(RightEdge ,CenterV),

        Point40(LeftEdge  ,FBottomEdge),
        Point41(DLeftEdge ,FBottomEdge),
        Point42(FLeftEdge ,FBottomEdge),
        Point43(CenterH   ,FBottomEdge),
        Point44(FRightEdge,FBottomEdge),
        Point45(DRightEdge,FBottomEdge),
        Point46(RightEdge ,FBottomEdge),

        Point50(LeftEdge  ,DBottomEdge),
        Point51(DLeftEdge ,DBottomEdge),
        Point52(FLeftEdge ,DBottomEdge),
        Point53(CenterH   ,DBottomEdge),
        Point54(FRightEdge,DBottomEdge),
        Point55(DRightEdge,DBottomEdge),
        Point56(RightEdge ,DBottomEdge),

        Point60(LeftEdge  ,BottomEdge),
        Point61(DLeftEdge ,BottomEdge),
        Point62(FLeftEdge ,BottomEdge),
        Point63(CenterH   ,BottomEdge),
        Point64(FRightEdge,BottomEdge),
        Point65(DRightEdge,BottomEdge),
        Point66(RightEdge ,BottomEdge),

        PointD(CenterH  ,CenterV);

//回傳一個點
    switch(few)
    {
        case  0:  return Point00;  break;
        case  1:  return Point01;  break;
        case  2:  return Point02;  break;
        case  3:  return Point03;  break;
        case  4:  return Point04;  break;
        case  5:  return Point05;  break;
        case  6:  return Point06;  break;

        case  7:  return Point10;  break;
        case  8:  return Point11;  break;
        case  9:  return Point12;  break;
        case 10:  return Point13;  break;
        case 11:  return Point14;  break;
        case 12:  return Point15;  break;
        case 13:  return Point16;  break;

        case 14:  return Point20;  break;
        case 15:  return Point21;  break;
        case 16:  return Point22;  break;
        case 17:  return Point23;  break;
        case 18:  return Point24;  break;
        case 19:  return Point25;  break;
        case 20:  return Point26;  break;

        case 21:  return Point30;  break;
        case 22:  return Point31;  break;
        case 23:  return Point32;  break;
        case 24:  return Point33;  break;//中心點
        case 25:  return Point34;  break;
        case 26:  return Point35;  break;
        case 27:  return Point36;  break;

        case 28:  return Point40;  break;
        case 29:  return Point41;  break;
        case 30:  return Point42;  break;
        case 31:  return Point43;  break;
        case 32:  return Point44;  break;
        case 33:  return Point45;  break;
        case 34:  return Point46;  break;

        case 35:  return Point50;  break;
        case 36:  return Point51;  break;
        case 37:  return Point52;  break;
        case 38:  return Point53;  break;
        case 39:  return Point54;  break;
        case 40:  return Point55;  break;
        case 41:  return Point56;  break;

        case 42:  return Point60;  break;
        case 43:  return Point61;  break;
        case 44:  return Point62;  break;
        case 45:  return Point63;  break;
        case 46:  return Point64;  break;
        case 47:  return Point65;  break;
        case 48:  return Point66;  break;

        default: PointD; return PointD; break;
    }
}