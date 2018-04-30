// 30Apr2018 addex CExcel->RunMacro
// 19Apr2017 added Sheet[].AddRow
// 14Jun2016 added CExcel->SaveAsTab
// 07Jul2015 added Bold property for CRange
// 17Feb2015 added CopySheet, Sheet[].Select()
// 23Jan2015 added CRange, FontSize, Alignment
// 14-Oct added Strike and BgColor
// 5-OCT Visible property added
#include "olectrls.hpp"


// https://msdn.microsoft.com/en-us/library/office/ff198017.aspx
typedef enum XlFileFormat
{
  xlText = -4158
} XlFileFormat;

typedef enum XlBordersIndex
{
  xlInsideHorizontal = 12,
  xlInsideVertical = 11,
  xlDiagonalDown = 5,
  xlDiagonalUp = 6,
  xlEdgeBottom = 9,
  xlEdgeLeft = 7,
  xlEdgeRight = 10,
  xlEdgeTop = 8
} XlBordersIndex;

typedef enum XlLineStyle
{
  xlContinuous = 1, 
  xlDash = 0xFFFFEFED, 
  xlDashDot = 4,
  xlDashDotDot = 5, 
  xlDot = 0xFFFFEFEA,
  xlDouble = 0xFFFFEFE9,
  xlSlantDashDot = 13, 
  xlLineStyleNone = 0xFFFFEFD2
} XlLineStyle;

// *********************************************************************//
// Declaration of Enumerations defined in Type Library
// *********************************************************************//
typedef enum Constants
{
  xlAll = 0xFFFFEFF8, 
  xlAutomatic = 0xFFFFEFF7, 
  xlBoth = 1, 
  xlCenter = 0xFFFFEFF4,
  xlChecker = 9, 
  xlCircle = 8, 
  xlCorner = 2,
  xlCrissCross = 16, 
  xlCross = 4, 
  xlDiamond = 2, 
  xlDistributed = 0xFFFFEFEB, 
  xlDoubleAccounting = 5, 
  xlFixedValue = 1, 
  xlFormats = 0xFFFFEFE6,
  xlGray16 = 17, 
  xlGray8 = 18, 
  xlGrid = 15, 
  xlHigh = 0xFFFFEFE1, 
  xlInside = 2, 
  xlJustify = 0xFFFFEFDE,
  xlLightDown = 13,
  xlLightHorizontal = 11, 
  xlLightUp = 14,
  xlLightVertical = 12, 
  xlLow = 0xFFFFEFDA, 
  xlManual = 0xFFFFEFD9, 
  xlMinusValues = 3, 
  xlModule = 0xFFFFEFD3, 
  xlNextToAxis = 4, 
  xlNone = 0xFFFFEFD2, 
  xlNotes = 0xFFFFEFD0, 
  xlOff = 0xFFFFEFCE, 
  xlOn = 1, 
  xlPercent = 2, 
  xlPlus = 9, 
  xlPlusValues = 2, 
  xlSemiGray75 = 10, 
  xlShowLabel = 4, 
  xlShowLabelAndPercent = 5, 
  xlShowPercent = 3, 
  xlShowValue = 2, 
  xlSimple = 0xFFFFEFC6, 
  xlSingle = 2, 
  xlSingleAccounting = 4, 
  xlSolid = 1, 
  xlSquare = 1,
  xlStar = 5,
  xlStError = 4, 
  xlToolbarButton = 2, 
  xlTriangle = 3, 
  xlGray25 = 0xFFFFEFE4, 
  xlGray50 = 0xFFFFEFE3, 
  xlGray75 = 0xFFFFEFE2, 
  xlBottom = 0xFFFFEFF5,
  xlLeft = 0xFFFFEFDD, 
  xlRight = 0xFFFFEFC8,
  xlTop = 0xFFFFEFC0, 
  xl3DBar = 0xFFFFEFFD, 
  xl3DSurface = 0xFFFFEFF9, 
  xlBar = 2, 
  xlColumn = 3, 
  xlCombination = 0xFFFFEFF1, 
  xlCustom = 0xFFFFEFEE,
  xlDefaultAutoFormat = 0xFFFFFFFF, 
  xlMaximum = 2, 
  xlMinimum = 4, 
  xlOpaque = 3, 
  xlTransparent = 2, 
  xlBidi = 0xFFFFEC78, 
  xlLatin = 0xFFFFEC77, 
  xlContext = 0xFFFFEC76, 
  xlLTR = 0xFFFFEC75, 
  xlRTL = 0xFFFFEC74, 
  xlFullScript = 1, 
  xlPartialScript = 2, 
  xlMixedScript = 3, 
  xlMixedAuthorizedScript = 4, 
  xlVisualCursor = 2, 
  xlLogicalCursor = 1, 
  xlSystem = 1,
  xlPartial = 3, 
  xlHindiNumerals = 3, 
  xlBidiCalendar = 3, 
  xlGregorian = 2, 
  xlComplete = 4, 
  xlScale = 3, 
  xlClosed = 3,
  xlColor1 = 7, 
  xlColor2 = 8,
  xlColor3 = 9, 
  xlConstants = 2, 
  xlContents = 2, 
  xlBelow = 1, 
  xlCascade = 7, 
  xlCenterAcrossSelection = 7,
  xlChart4 = 2, 
  xlChartSeries = 17, 
  xlChartShort = 6, 
  xlChartTitles = 18, 
  xlClassic1 = 1, 
  xlClassic2 = 2, 
  xlClassic3 = 3, 
  xl3DEffects1 = 13, 
  xl3DEffects2 = 14, 
  xlAbove = 0, 
  xlAccounting1 = 4,
  xlAccounting2 = 5, 
  xlAccounting3 = 6, 
  xlAccounting4 = 17, 
  xlAdd = 2, 
  xlDebugCodePane = 13, 
  xlDesktop = 9, 
  xlDirect = 1,
  xlDivide = 5, 
  xlDoubleClosed = 5, 
  xlDoubleOpen = 4, 
  xlDoubleQuote = 1, 
  xlEntireChart = 20, 
  xlExcelMenus = 1, 
  xlExtended = 3,
  xlFill = 5, 
  xlFirst = 0,
  xlFloating = 5, 
  xlFormula = 5, 
  xlGeneral = 1, 
  xlGridline = 22, 
  xlIcons = 1, 
  xlImmediatePane = 12, 
  xlInteger = 2, 
  xlLast = 1, 
  xlLastCell = 11, 
  xlList1 = 10, 
  xlList2 = 11, 
  xlList3 = 12, 
  xlLocalFormat1 = 15, 
  xlLocalFormat2 = 16, 
  xlLong = 3, 
  xlLotusHelp = 2, 
  xlMacrosheetCell = 7, 
  xlMixed = 2, 
  xlMultiply = 4, 
  xlNarrow = 1, 
  xlNoDocuments = 3, 
  xlOpen = 2, 
  xlOutside = 3, 
  xlReference = 4,
  xlSemiautomatic = 2, 
  xlShort = 1, 
  xlSingleQuote = 2,
  xlStrict = 2, 
  xlSubtract = 3, 
  xlTextBox = 16, 
  xlTiled = 1,
  xlTitleBar = 8, 
  xlToolbar = 1,
  xlVisible = 12, 
  xlWatchPane = 11, 
  xlWide = 3, 
  xlWorkbookTab = 6, 
  xlWorksheet4 = 1, 
  xlWorksheetCell = 3, 
  xlWorksheetShort = 5, 
  xlAllExceptBorders = 7, 
  xlLeftToRight = 2, 
  xlTopToBottom = 1, 
  xlVeryHidden = 2, 
  xlDrawingObject = 14
} Constants;


class CExcel{
  private:
  Variant Ex;

  class CRange{
    private:
    Variant range;
    void __fastcall SetBold(int value){
      range.OlePropertyGet("Font").OlePropertySet("Bold",value);
    }
    int __fastcall GetBold(){
      return range.OlePropertyGet("Font").OlePropertyGet("Bold");
    }

    public:
    __property int Bold = {read=GetBold,write=SetBold};

    void MergeCells(){
    
     range.OlePropertySet("MergeCells",1);

    }
    void WrapText(){
      range.OlePropertySet("WrapText", 1);
    }
    void AutoFilter(){
      range.OleProcedure("AutoFilter");
    }    

    void SetBorder1(){
      range.OlePropertyGet("Borders").OlePropertyGet("Item",xlEdgeRight).OlePropertySet("LineStyle", xlContinuous);
      range.OlePropertyGet("Borders").OlePropertyGet("Item",xlEdgeLeft).OlePropertySet("LineStyle", xlContinuous);
      range.OlePropertyGet("Borders").OlePropertyGet("Item",xlEdgeBottom).OlePropertySet("LineStyle", xlContinuous);
      range.OlePropertyGet("Borders").OlePropertyGet("Item",xlEdgeTop).OlePropertySet("LineStyle", xlContinuous);
      range.OlePropertyGet("Borders").OlePropertyGet("Item",xlInsideVertical).OlePropertySet("LineStyle", xlContinuous);
      range.OlePropertyGet("Borders").OlePropertyGet("Item",xlInsideHorizontal).OlePropertySet("LineStyle", xlContinuous);
    }

    void SetBorder(XlBordersIndex bi,XlLineStyle ls){
      range.OlePropertyGet("Borders").OlePropertyGet("Item",bi).OlePropertySet("LineStyle", ls);
    }  

    CRange(int row1,int col1,int row2,int col2,Variant Sheet){//"A1:C10"
      range = Sheet.OlePropertyGet("Range",StringToOleStr(AnsiString().sprintf("%c%i:%c%i",'A'+col1-1,row1,'A'+col2-1,row2)));
    }

  };

  class CCell{
    private:
    Variant cell;
    void __fastcall SetBold(int value){
      cell.OlePropertyGet("Font").OlePropertySet("Bold",value);
    }
    int __fastcall GetBold(){
      return cell.OlePropertyGet("Font").OlePropertyGet("Bold");
    }
    void __fastcall SetStrike(int value){
      cell.OlePropertyGet("Font").OlePropertySet("Strikethrough",value);
    }
    int __fastcall GetStrike(){
      return cell.OlePropertyGet("Font").OlePropertyGet("Strikethrough");
    }
    
    void __fastcall SetFontSize(int value){
      cell.OlePropertyGet("Font").OlePropertySet("Size",value);
    }
    int __fastcall GetFontSize(){
      return cell.OlePropertyGet("Font").OlePropertyGet("Size");
    }    

    void __fastcall SetBgColor(TColor value){
      cell.OlePropertyGet("Interior").OlePropertySet("Color",value);
    }
    TColor __fastcall GetBgColor(){
      return (TColor)(int)cell.OlePropertyGet("Interior").OlePropertyGet("Color");
    }


    void __fastcall SetStr(AnsiString value){
      cell.OlePropertySet("Value",StringToOleStr(value));
    }
    AnsiString __fastcall GetStr(){
      return cell.OlePropertyGet("Value");
    }

    void __fastcall SetInt(int value){
      cell.OlePropertySet("Value",value);
    }
    int __fastcall GetInt(){
      return cell.OlePropertyGet("Value");
    }

    void __fastcall SetHorizontalAlignment(int value){
      cell.OlePropertySet("HorizontalAlignment",value);
    }
    int __fastcall GetHorizontalAlignment(){
      return cell.OlePropertyGet("HorizontalAlignment");
    }

    void __fastcall SetVerticalAlignment(int value){
      cell.OlePropertySet("VerticalAlignment",value);
    }
    int __fastcall GetVerticalAlignment(){
      return cell.OlePropertyGet("VerticalAlignment");
    }




    
    void __fastcall SetFormat(AnsiString value){
      cell.OlePropertySet("NumberFormat",StringToOleStr(value));
    }
    AnsiString __fastcall GetFormat(){
      return cell.OlePropertyGet("NumberFormat");
    }

    void __fastcall SetDT(TDateTime value){
      cell.OlePropertySet("Value",value);
    }
    TDateTime __fastcall GetDT(){
      TDateTime d1;
      AnsiString st = cell.OlePropertyGet("Value");
      try{
        d1 = (TDateTime)cell.OlePropertyGet("Value");
      }
      catch(...){
        ShowMessage((AnsiString)"Problem converting "+st+" as datetime.");
      }
      return d1;
    }

    void __fastcall SetFlo(double value){
      cell.OlePropertySet("Value",value);
    }
    double __fastcall GetFlo(){
      return cell.OlePropertyGet("Value");
    }
    public:
    __property AnsiString Str = {read=GetStr,write=SetStr};
    __property AnsiString Format = {read=GetFormat,write=SetFormat};
    __property TDateTime DT = {read=GetDT,write=SetDT};
    __property double Flo = {read=GetFlo,write=SetFlo};
    __property int Int = {read=GetInt,write=SetInt};
    __property int Bold = {read=GetBold,write=SetBold};
    __property int Strike = {read=GetStrike,write=SetStrike};
    __property int FontSize = {read=GetFontSize,write=SetFontSize};
    __property TColor BgColor = {read=GetBgColor,write=SetBgColor};

    __property int HorizontalAlignment = {read=GetHorizontalAlignment,write=SetHorizontalAlignment};
    __property int VerticalAlignment = {read=GetVerticalAlignment,write=SetVerticalAlignment};




    CCell(int i,int j,Variant Sheet){
      cell = Sheet.OlePropertyGet("Cells",i,j);
    }

  };

  class CSheet{
    private:
    Variant sheet;
    CCell __fastcall GetCell(int i,int j){
      return CCell(i,j,sheet);
    }
    CRange __fastcall GetRange(int row1,int col1,int row2,int col2){
      return CRange(row1,col1,row2,col2,sheet);
    }
    void __fastcall SetName(AnsiString value){
      sheet.OlePropertySet("Name",StringToOleStr(value));
    }
    AnsiString __fastcall GetName(){
      return sheet.OlePropertyGet("Name");
    }

    public:
    __property CCell Cell[int i][int j] = {read=GetCell};
    __property AnsiString Name = {read=GetName,write=SetName};
    __property CRange Range[int row1][int col1][int row2][int col2] = {read=GetRange};
    void Clear(){
      sheet.OlePropertyGet("Cells").OleProcedure("Clear");
    }
    void Select(){
      sheet.OleProcedure("Select");
    }
    void AutoFit(int j){
      char c = 'A'+j-1;
      sheet.OlePropertyGet("Columns",StringToOleStr((AnsiString)char(c)+":"+char(c))).OlePropertyGet("EntireColumn").OleProcedure("AutoFit");
    }

    void AddRow(int j){
      sheet.OlePropertyGet ("Rows", j). OleProcedure ("Insert");
    }

    CSheet(int i,Variant Ex){
      sheet = Ex.OlePropertyGet("ActiveWorkbook").OlePropertyGet("Sheets", i);
    }
    CSheet(AnsiString st,Variant Ex){
      sheet = Ex.OlePropertyGet("ActiveWorkbook").OlePropertyGet("Sheets", StringToOleStr(st));
    }


  };


  AnsiString fname;
  AnsiString password;



  CSheet __fastcall GetSheet(int i){
    return CSheet(i,Ex);
  }
  CSheet __fastcall GetSheetByName(AnsiString st){
    return CSheet(st,Ex);
  }

  int __fastcall GetSheetsCount(){
    return Ex.OlePropertyGet("Workbooks").OlePropertyGet("Item", 1).OlePropertyGet("Worksheets").OlePropertyGet("Count");
  }


    void __fastcall SetVisible(int value){
      Ex.OlePropertySet("Visible",value);
    }
    int __fastcall GetVisible(){
      return Ex.OlePropertyGet("Visible");
    }

  public:

  __property int Count = {read=GetSheetsCount};
  __property CSheet Sheet[int i] = {read=GetSheet};
  __property CSheet SheetByName[AnsiString st] = {read=GetSheetByName};
  __property int Visible = {read=GetVisible,write=SetVisible};
  void AddSheet(){
    Ex.OlePropertyGet("ActiveWorkbook").OlePropertyGet("Sheets").OleProcedure("Add");
  }

  void CopySheet(AnsiString fname,AnsiString SheetName){
    Ex.OlePropertyGet("Workbooks").OleProcedure("Open",StringToOleStr(fname),EmptyParam,EmptyParam,EmptyParam,EmptyParam);
    Ex.OlePropertySet("Visible",1);

//    ShowMessage(Ex.OlePropertyGet("Workbooks").OlePropertyGet("Item", 1).OlePropertyGet("Sheets", 1).
//      OlePropertyGet("Cells",1,1).OlePropertyGet("Value"));
//    ShowMessage(Ex.OlePropertyGet("Workbooks").OlePropertyGet("Item", 2).OlePropertyGet("Sheets", StringToOleStr(SheetName)).
//      OlePropertyGet("Cells",1,1).OlePropertyGet("Value"));
    Ex.OlePropertyGet("Workbooks").OlePropertyGet("Item", 2).OlePropertyGet("Sheets", StringToOleStr(SheetName)).
      OleProcedure("Copy",EmptyParam,Ex.OlePropertyGet("Workbooks").OlePropertyGet("Item", 1).OlePropertyGet("Sheets", 1));

    Ex.OlePropertySet("DisplayAlerts",0);
    Ex.OlePropertyGet("Workbooks").OlePropertyGet("Item", 2).OleProcedure("Close");
  }


  void RunMacro(AnsiString st){
    Ex.OleProcedure("Run",StringToOleStr(st));
/*    Variant v2 = v1.OlePropertyGet("VBProject");
    ShowMessage(v2.OlePropertyGet("Name"));

    /*
    Variant VBComp = VBComp.OlePropertyGet("VBComponents");
    VBComp = VBComp.OleFunction("Add", 1);
    VBComp.OlePropertySet("Name", "MyModule777");
    VBComp.OlePropertyGet("CodeModule").OleProcedure("AddFromString", StringToOleStr(st));
      */

   }
    void Freeze11(){
      Ex.OlePropertyGet("ActiveWindow").OlePropertySet("SplitColumn", 1);
      Ex.OlePropertyGet("ActiveWindow").OlePropertySet("SplitRow", 1);
      Ex.OlePropertyGet("ActiveWindow").OlePropertySet("FreezePanes", 1);
    }

  int CloseExcelOnDelete;

  void SaveAsTab(AnsiString f1){
    Ex.OlePropertySet("DisplayAlerts",0);
    Ex.OlePropertyGet("ActiveWorkbook").OleProcedure("Saveas",StringToOleStr(f1),xlText);
  }


  CExcel(AnsiString _fname, AnsiString _password=""){
    fname = _fname;
    password = _password;
    CloseExcelOnDelete = 1;
    Ex=Variant::CreateObject("Excel.Application");
    if(FileExists(fname))
      Ex.OlePropertyGet("Workbooks").OleProcedure("Open",StringToOleStr(fname),EmptyParam,EmptyParam,EmptyParam,StringToOleStr(password));
    else
      Ex.OlePropertyGet("Workbooks").OleFunction("Add", 1);

  }

  ~CExcel(){
    Ex.OlePropertySet("DisplayAlerts",0);
    if(FileExists(fname))
      Ex.OlePropertyGet("ActiveWorkbook").OleProcedure("Save");
    else
      Ex.OlePropertyGet("ActiveWorkbook").OleProcedure("Saveas",StringToOleStr(fname));
    if(CloseExcelOnDelete){
      Ex.OlePropertyGet("ActiveWorkbook").OleProcedure("Close");
      Ex.OleFunction("Quit");
    }
    Ex = Unassigned;
  }
};
