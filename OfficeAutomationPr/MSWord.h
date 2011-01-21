#pragma once

#include <list>
#include <vector>
using namespace std;


class CMSWord
{
protected:
	IDispatch*	m_pWApp;
	IDispatch*  m_pDocuments;
	IDispatch*	m_pActiveDocument;
	IDispatch*  pDocApp;
	HRESULT		m_hr;
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
	//Saves file in path given by parameter
	HRESULT SaveFile(LPCTSTR czFileName);
	//sets Align Justify
	HRESULT AlignJustify();
	//sets Align Left
	HRESULT AlignLeft();
	//sets Align Right
	HRESULT AlignRight();
	//sets Align Center
	HRESULT AlignCenter();



};