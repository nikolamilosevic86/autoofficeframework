# Auto Office Library #

## Introduction ##

Auto Office Frmaevork is framework for automating functions of MS Office applications. It works via Component Object Model, and via COM interfaces interacts with MS Office Functions. First realease version supports large scale of MS Word functionlaity, and some functionality of MS Excel. Other programs from MS Office will be supported in later versions. Also Auto Office Framework is supported in all versions of MS Office since MS Office 2003 (tested), and maybe earlier (not tested).

## Motivation ##

Many applications can be extended to create reports, or to use MS Office documents. It can be used for reports, or can be used as templates for creating documents reading some database. There are some frameworks for interaction with MS Office applications but in languages like Java or .NET programming languages. If someone want to interact with MS Office via native C++, there is a problem. To solve that problem is created Auto Office Framework for MFC C++.


## Details ##

The library contains classes for each office applications that are able to ineract with that MS Office component. In first (0.1) version of Auto Office Framework there are just two, CMSWord and CMSExcel classes. CMSWord class is used to interact with MS Office Word application. **This is the class interface that can be used:**

```
class CMSWord
{
protected:
        IDispatch*      m_pWApp;
        IDispatch*  m_pDocuments;
        IDispatch*      m_pActiveDocument;
        IDispatch*  pDocApp;
        HRESULT         m_hr;
        HRESULT Initialize(bool bVisible=true);
public:
        //Constructor creates a instance of object
        CMSWord();
        //Destructor
        ~CMSWord();
        //Sets visibility of active document
        HRESULT SetVisible(bool bVisible=true);
        //Quits the MS Word. Closes process
        HRESULT Quit();
        //Opens document and set document's visibility
        HRESULT OpenDocument(LPCTSTR szFilename, bool bVisible=true);
        //Open new empty document
        HRESULT NewDocument(bool bVisible=true);
        //Closes active document with or without saving
        HRESULT CloseActiveDocument(bool bSave=true);
        //Finds first next text in active document as specified 
        HRESULT FindFirst(LPCTSTR szText);
        //Finds first next text in active document as specified. Returns false if end of document reached.
        bool FindFirstBool(LPCTSTR szText);
        //Close all documents
        HRESULT CloseDocuments(bool bSave=true);
        //Copies selected text into clipboard
        HRESULT Copy();
        //Pastes from Clipboard to active document
        HRESULT Paste();
        //Activate document by specified id
        HRESULT ActivateDocumentById(int id);
        //Returns number of opened documents in the process. There can be more MS Word processes, and the framework won't see documents controled by other processes.
        int CountDocuments();
        //Sets selected text. Replace it with specified text.
        HRESULT SetSelectionText(LPCTSTR szText);
        //Inserts picture from path specified in argument.
        HRESULT InserPicture(LPCTSTR szFilename);
        //Inserts text in active document.
        HRESULT InserText(LPCTSTR szText);
        // Inserts MS Word file into active document with all formating.
        HRESULT InsertFile(LPCTSTR szFilename);
        //Adds comment in ballon. Text is specified in argument
        HRESULT AddComment(LPCTSTR szComment);
        //Moves cursor. 2 is forward, 1 i backward. Selection is true then it selects text as it moves
        HRESULT MoveCursor(int nDirection=2,bool bSelection=false);
        //Delete char forward or backward
        HRESULT DeleteChar(bool bBack=false);
        //Sets bold for next inserted chars or for selected text.
        HRESULT SetBold(bool bBold=false);
        //Sets italic for next inserted chars or for selected text.
        HRESULT SetItalic(bool bItalic=false);
        //Sets underline for next inserted chars or for selected text.
        HRESULT SetUnderline(bool bUnderline=false);
        //Check spelling
        HRESULT CheckSpelling(LPCTSTR szWord, bool &bResult);
        //check grammer
        HRESULT CheckGrammer(LPCTSTR szString, bool &bResult);
        //sets font as specified
        HRESULT SetFont(LPCTSTR szFontName, int nSize, bool bBold, bool bItalic,COLORREF crColor);
        //Gets string of specified size
        CString GetString(int nlenght);
        //Get selected string
        CString GetSelectedString();
        //Saves file
        HRESULT SaveFile(LPCTSTR czFileName);

};
```

**Interface of CMSExcel class is:**
```
class CMSExcel
{
protected:
        HRESULT m_hr;
        IDispatch*      m_pEApp;
        IDispatch*  m_pBooks;
        IDispatch*      m_pActiveBook;
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
//This class will be commented soon

```
Also all of this methods from classes uses OLEMethod function. This is the function that uses COM to gain access to some remote objects and to invoke methods or to get or set some property. Coding in COM requires a lot of code, and this function enables to reuse the great amount of that code.

## Code samples ##

Code samples will be added later.