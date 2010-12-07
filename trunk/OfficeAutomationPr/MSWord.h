#pragma once

class CMSWord
{
protected:
	IDispatch*	m_pWApp;
	IDispatch*  m_pDocuments;
	IDispatch*	m_pActiveDocument;
	HRESULT		m_hr;
	HRESULT Initialize(bool bVisible=true);
public:
	CMSWord();
	~CMSWord();
	HRESULT SetVisible(bool bVisible=true);
	HRESULT Quit();
	HRESULT OpenDocument(LPCTSTR szFilename, bool bVisible=true);
	HRESULT NewDocument(bool bVisible=true);
	HRESULT CloseActiveDocument(bool bSave=true);
	HRESULT FindFirst(LPCTSTR szText);
	bool FindFirstBool(LPCTSTR szText);
	HRESULT CloseDocuments(bool bSave=true);
	HRESULT Copy();
	HRESULT Paste();
	HRESULT ActivateDocumentById(int id);
	int CountDocuments();
	HRESULT SetSelectionText(LPCTSTR szText);
	HRESULT Replace(LPCTSTR szText,LPCTSTR szReplacementText,bool ReplaceAll);
	HRESULT InserPicture(LPCTSTR szFilename);
	HRESULT InserText(LPCTSTR szText);
	HRESULT AddComment(LPCTSTR szComment);
	HRESULT MoveCursor(int nDirection=2,bool bSelection=false);
	HRESULT DeleteChar(bool bBack=false);
	HRESULT CheckSpelling(LPCTSTR szWord, bool &bResult);
	HRESULT CheckGrammer(LPCTSTR szString, bool &bResult);
	HRESULT SetFont(LPCTSTR szFontName, int nSize, bool bBold, bool bItalic,COLORREF crColor);
};