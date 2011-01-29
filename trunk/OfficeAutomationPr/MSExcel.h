#pragma once

class CMSExcel
{
protected:
	HRESULT m_hr;
	IDispatch*	m_pEApp;
	IDispatch*  m_pBooks;
	IDispatch*	m_pActiveBook;
private:
	HRESULT Initialize(bool bVisible=true);
public:
	//Constructor
	CMSExcel(void);
	//Destructor
	~CMSExcel(void);
	//Sets visible application
	HRESULT SetVisible(bool bVisible=true);
	//Creates new Excel book
	HRESULT NewExcelBook(bool bVisible=true);
	//Opens Excel book
	HRESULT OpenExcelBook(LPCTSTR szFilename, bool bVisible=true);
	//Saves as 
	HRESULT SaveAs(LPCTSTR szFilename, int nSaveAs=40);
	//Protect with password
	HRESULT ProtectExcelWorkbook(LPCTSTR szPassword);
	//UnProtect wordbook
	HRESULT UnProtectExcelWorkbook(LPCTSTR szPassword);
	//Protect just one sheet
	HRESULT ProtectExcelSheet(int nSheetNo, LPCTSTR szPassword);
	//Unprotect one sheet
	HRESULT UnProtectExcelSheet(int nSheetNo, LPCTSTR szPassword);
	//Sets cell format. Range is string of Range ex "$I$2:$P$20" or "I2:P20"
	HRESULT SetExcelCellFormat(LPCTSTR szRange, LPCTSTR szFormat);
	//Sets name of Sheet
	HRESULT SetExcelSheetName(int nSheetNo, LPCTSTR szSheetName);
	//Gets value
	HRESULT GetExcelValue(LPCTSTR szCell, CString &sValue);
	//Sets background colour
	HRESULT SetExcelBackgroundColor(LPCTSTR szRange, COLORREF crColor, int nPattern);
	//Sets Excel Font
	HRESULT SetExcelFont(LPCTSTR szRange, LPCTSTR szName, int nSize, COLORREF crColor, bool bBold, bool bItalic);
	//Sets value 
	HRESULT SetExcelValue(LPCTSTR szRange,LPCTSTR szValue,bool bAutoFit, int nAlignment);
	//Sets border stile
	HRESULT SetExcelBorder(LPCTSTR szRange, int nStyle);
	//Merges cells of range. Write Range as "$I$2:$P$20" or "I2:P20"
	HRESULT MergeExcelCells(LPCTSTR szRange);
	//Auto fit Excel column
	HRESULT AutoFitExcelColumn(LPCTSTR szColumn);
	//Add chart
	HRESULT AddExcelChart(LPCTSTR szRange, LPCTSTR szTitle, int nChartType, int nLeft, int nTop, int nWidth, int nHeight, int nRangeSheet, int nChartSheet);
	//Moves cursor
	HRESULT MoveCursor(int nDirection);
	//Gets active cell number
	HRESULT GetActiveCell(int &nRow, int &nCol);
	//Sets active cell text
	HRESULT SetActiveCellText(LPCTSTR szText);
	//Quits application
	HRESULT Quit();
};
