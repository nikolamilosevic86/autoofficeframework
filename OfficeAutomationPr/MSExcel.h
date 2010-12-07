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
	CMSExcel(void);
	~CMSExcel(void);
	HRESULT SetVisible(bool bVisible=true);
	HRESULT NewExcelBook(bool bVisible=true);
	HRESULT OpenExcelBook(LPCTSTR szFilename, bool bVisible=true);
	HRESULT SaveAs(LPCTSTR szFilename, int nSaveAs=40);
	HRESULT ProtectExcelWorkbook(LPCTSTR szPassword);
	HRESULT UnProtectExcelWorkbook(LPCTSTR szPassword);
	HRESULT ProtectExcelSheet(int nSheetNo, LPCTSTR szPassword);
	HRESULT UnProtectExcelSheet(int nSheetNo, LPCTSTR szPassword);
	HRESULT SetExcelCellFormat(LPCTSTR szRange, LPCTSTR szFormat);
	HRESULT SetExcelSheetName(int nSheetNo, LPCTSTR szSheetName);
	HRESULT GetExcelValue(LPCTSTR szCell, CString &sValue);
	HRESULT SetExcelBackgroundColor(LPCTSTR szRange, COLORREF crColor, int nPattern);
	HRESULT SetExcelFont(LPCTSTR szRange, LPCTSTR szName, int nSize, COLORREF crColor, bool bBold, bool bItalic);
	HRESULT SetExcelValue(LPCTSTR szRange,LPCTSTR szValue,bool bAutoFit, int nAlignment);
	HRESULT SetExcelBorder(LPCTSTR szRange, int nStyle);
	HRESULT MergeExcelCells(LPCTSTR szRange);
	HRESULT AutoFitExcelColumn(LPCTSTR szColumn);
	HRESULT AddExcelChart(LPCTSTR szRange, LPCTSTR szTitle, int nChartType, int nLeft, int nTop, int nWidth, int nHeight, int nRangeSheet, int nChartSheet);
	HRESULT MoveCursor(int nDirection);
	HRESULT GetActiveCell(int &nRow, int &nCol);
	HRESULT SetActiveCellText(LPCTSTR szText);
	HRESULT Quit();
};
