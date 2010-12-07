#include "stdafx.h"
#include "MSWord.h"
#include "OLEMethod.h"

CMSWord::CMSWord()
{
	m_pWApp=NULL;
	m_pDocuments=NULL;
	m_pActiveDocument=NULL;
}

HRESULT CMSWord::Initialize(bool bVisible)
{
	CoInitialize(NULL);
	CLSID clsid;
	m_hr = CLSIDFromProgID(L"Word.Application", &clsid);
	if(SUCCEEDED(m_hr))
	{
		m_hr = CoCreateInstance(clsid, NULL, CLSCTX_LOCAL_SERVER, IID_IDispatch, (void **)&m_pWApp);
		if(FAILED(m_hr)) m_pWApp=NULL;
	}
	{
		m_hr=SetVisible(bVisible);
	}
	return m_hr;
}



HRESULT CMSWord::SetVisible(bool bVisible)
{
/*	DISPID dispID;
	VARIANT pvResult;
	LPOLESTR ptName=_T("Visible");
	m_hr = m_pWApp->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
	if(SUCCEEDED(m_hr))
	{
		VARIANT x;
		x.vt = VT_I4;
		x.lVal =bVisible?1:0;
		DISPID prop=DISPATCH_PROPERTYPUT;

		DISPPARAMS dp = { NULL,NULL,0,0 };
		dp.cArgs =1;
		dp.rgvarg =&x;
		dp.cNamedArgs=1;
		dp.rgdispidNamedArgs= &prop;
		m_hr = m_pWApp->Invoke(dispID, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_PROPERTYPUT, 
								&dp, &pvResult, NULL, NULL);
	}*/
	VARIANT x;
	x.vt = VT_I4;
	x.lVal = bVisible;
	m_hr=OLEMethod(DISPATCH_PROPERTYPUT, NULL, m_pWApp, L"Visible", 1, x);

	return m_hr;
}

HRESULT CMSWord::OpenDocument(LPCTSTR szFilename, bool bVisible)
{
	if(m_pWApp==NULL) 
	{
		if(FAILED(m_hr=Initialize(bVisible)))
			return m_hr;
	}
	COleVariant vFname(szFilename);
	VARIANT fname=vFname.Detach();
	// GetDocuments
	{
		VARIANT result;
		VariantInit(&result);
		m_hr=OLEMethod(DISPATCH_PROPERTYGET, &result, m_pWApp, L"Documents", 0);
		m_pDocuments= result.pdispVal;
	}
	// OpenDocument
	{
		VARIANT result;
        VariantInit(&result);
		m_hr=OLEMethod(DISPATCH_METHOD, &result, m_pDocuments, L"Open", 1, fname);
		m_pActiveDocument = result.pdispVal;
	}
	return m_hr;
}

HRESULT CMSWord::ActivateDocumentById(int id){

	if(m_pWApp==NULL) return E_FAIL;

	{
		VARIANT result;
		VariantInit(&result);
		m_hr=OLEMethod(DISPATCH_PROPERTYGET, &result, m_pWApp, L"Documents", 0);
		m_pDocuments = result.pdispVal;
	}	
	IDispatch *pItem;
	{
		VARIANT result;
		VariantInit(&result);
		COleVariant varIndex;
		varIndex.lVal=id;
		varIndex.vt=VT_I4;
		m_hr=OLEMethod(DISPATCH_METHOD, &result, m_pDocuments, L"Item", 1,varIndex);
		pItem = result.pdispVal;
	}
	m_hr=OLEMethod(DISPATCH_METHOD, NULL, pItem, L"Activate", 0);
	return m_hr;

}

int CMSWord::CountDocuments(){

	int cnt;
	if(m_pWApp==NULL) 
	{
		//if(FAILED(m_hr=Initialize(bVisible)))
			return -5;
	}

	{
		VARIANT result;
		VariantInit(&result);
		m_hr=OLEMethod(DISPATCH_PROPERTYGET, &result, m_pWApp, L"Documents", 0);
		m_pDocuments = result.pdispVal;
	}	

	{
		VARIANT result;
		VariantInit(&result);
		m_hr=OLEMethod(DISPATCH_PROPERTYGET, &result, m_pDocuments, L"Count", 0);
		cnt = result.lVal;
	}
	return cnt;
}

HRESULT CMSWord::NewDocument(bool bVisible)
{
	if(m_pWApp==NULL) 
	{
		if(FAILED(m_hr=Initialize(bVisible)))
			return m_hr;
	}

	{
		VARIANT result;
		VariantInit(&result);
		m_hr=OLEMethod(DISPATCH_PROPERTYGET, &result, m_pWApp, L"Documents", 0);
		m_pDocuments = result.pdispVal;
	}	

	{
		VARIANT result;
		VariantInit(&result);
		m_hr=OLEMethod(DISPATCH_METHOD, &result, m_pDocuments, L"Add", 0);
		m_pActiveDocument = result.pdispVal;
	}
	return m_hr;
}

// bSave=true will autosave the document. bSave=false will close document without saving.
HRESULT CMSWord::CloseDocuments(bool bSave)
{
	if(m_pWApp==NULL) return E_FAIL;
	{
		VARIANT result;
        VariantInit(&result);
		COleVariant varSave;
		varSave.boolVal=bSave;
		m_hr=OLEMethod(DISPATCH_METHOD, &result, m_pDocuments, L"Close", 1, varSave.Detach());
		m_pDocuments=NULL;
		m_pActiveDocument=NULL;
	}	 
	return m_hr;
}

HRESULT CMSWord::DeleteChar(bool bBack)
{

	if(m_pWApp==NULL) return E_FAIL;
	IDispatch *pDocApp;
	{  
		VARIANT result;
		VariantInit(&result);
		OLEMethod(DISPATCH_PROPERTYGET, &result, m_pActiveDocument, L"Application", 0);
		pDocApp= result.pdispVal;
	}
	IDispatch *pSelection;
	{
		VARIANT result;
		VariantInit(&result);
		OLEMethod(DISPATCH_PROPERTYGET, &result, pDocApp, L"Selection", 0);
		pSelection=result.pdispVal;
	}

	{
		VARIANT wdCharacter, Count;
		wdCharacter.vt =VT_I4;
		wdCharacter.lVal =1;
		Count.vt =VT_I4;
		Count.lVal =1;
		if(bBack)
			OLEMethod(DISPATCH_METHOD, NULL, pSelection, L"TypeBackspace", 0);
		else
			OLEMethod(DISPATCH_METHOD, NULL, pSelection, L"Delete", 2,Count,wdCharacter);
	}
	return m_hr;
}

HRESULT CMSWord::CheckSpelling(LPCTSTR szWord, bool &bResult)
{
	if(!m_pWApp || !m_pActiveDocument) return E_FAIL;
	IDispatch *pDocApp;
	{  
		VARIANT result;
		VariantInit(&result);
		OLEMethod(DISPATCH_PROPERTYGET, &result, m_pActiveDocument, L"Application", 0);
		pDocApp= result.pdispVal;
	}
	{
		VARIANT result;
		VariantInit(&result);
		COleVariant varWord(szWord);
		OLEMethod(DISPATCH_METHOD, &result,pDocApp,L"CheckSpelling",1,varWord.Detach());
		bResult=result.boolVal;
	}
	return m_hr;
}

HRESULT CMSWord::CheckGrammer(LPCTSTR szString, bool &bResult)
{
	if(!m_pWApp || !m_pActiveDocument) return E_FAIL;
	IDispatch *pDocApp;
	{  
		VARIANT result;
		VariantInit(&result);
		OLEMethod(DISPATCH_PROPERTYGET, &result, m_pActiveDocument, L"Application", 0);
		pDocApp= result.pdispVal;
	}
	{
		VARIANT result;
		VariantInit(&result);
		COleVariant varString(szString);
		OLEMethod(DISPATCH_METHOD, &result,pDocApp,L"CheckGrammar",1,varString.Detach());
		bResult=result.boolVal;
	}
	return m_hr;
}


HRESULT CMSWord::Copy(){

	if(!m_pWApp || !m_pActiveDocument) return E_FAIL;
	IDispatch *pDocApp;
	{  
		VARIANT result;
		VariantInit(&result);
		OLEMethod(DISPATCH_PROPERTYGET, &result, m_pActiveDocument, L"Application", 0);
		pDocApp= result.pdispVal;
	}
	IDispatch *pSelection;
	{
		VARIANT result;
		VariantInit(&result);
		OLEMethod(DISPATCH_PROPERTYGET, &result, pDocApp, L"Selection", 0);
		pSelection=result.pdispVal;
	}
	{
		VARIANT result;
		VariantInit(&result);
		OLEMethod(DISPATCH_METHOD, &result, pSelection, L"Copy", 0);
	}
	pDocApp->Release();
	pSelection->Release();
	return m_hr;

}

HRESULT CMSWord::Paste(){

	if(!m_pWApp || !m_pActiveDocument) return E_FAIL;
	IDispatch *pDocApp;
	{  
		VARIANT result;
		VariantInit(&result);
		OLEMethod(DISPATCH_PROPERTYGET, &result, m_pActiveDocument, L"Application", 0);
		pDocApp= result.pdispVal;
	}
	IDispatch *pSelection;
	{
		VARIANT result;
		VariantInit(&result);
		OLEMethod(DISPATCH_PROPERTYGET, &result, pDocApp, L"Selection", 0);
		pSelection=result.pdispVal;
	}
	{
		VARIANT wdPasteDefault;
		wdPasteDefault.vt =VT_I4;
		wdPasteDefault.lVal =0;
		VARIANT result;
		VariantInit(&result);
		OLEMethod(DISPATCH_METHOD, &result, pSelection, L"PasteAndFormat", 1,wdPasteDefault);
	}
	pDocApp->Release();
	pSelection->Release();
	return m_hr;

}

HRESULT CMSWord::InserPicture(LPCTSTR szFilename)
{
	if(!m_pWApp || !m_pActiveDocument) return E_FAIL;
	IDispatch *pDocApp;
	{  
		VARIANT result;
		VariantInit(&result);
		OLEMethod(DISPATCH_PROPERTYGET, &result, m_pActiveDocument, L"Application", 0);
		pDocApp= result.pdispVal;
	}
	IDispatch *pSelection;
	{
		VARIANT result;
		VariantInit(&result);
		OLEMethod(DISPATCH_PROPERTYGET, &result, pDocApp, L"Selection", 0);
		pSelection=result.pdispVal;
	}
	IDispatch *pInlineShapes;
	{
		VARIANT result;
		VariantInit(&result);
		OLEMethod(DISPATCH_PROPERTYGET, &result, pSelection, L"InlineShapes", 0);
		pInlineShapes=result.pdispVal;
	}
	{
		COleVariant varFile(szFilename);
		COleVariant varLink((BYTE)0);
		COleVariant varSave((BYTE)1);
		OLEMethod(DISPATCH_METHOD,NULL,pInlineShapes,L"AddPicture",3,varSave.Detach(),varLink.Detach(),varFile.Detach());
	}
	return m_hr;
}

HRESULT CMSWord::InserText(LPCTSTR szText)
{
	if(!m_pWApp || !m_pActiveDocument) return E_FAIL;
	IDispatch *pDocApp;
	{  
		VARIANT result;
		VariantInit(&result);
		OLEMethod(DISPATCH_PROPERTYGET, &result, m_pActiveDocument, L"Application", 0);
		pDocApp= result.pdispVal;
	}
	IDispatch *pSelection;
	{
		VARIANT result;
		VariantInit(&result);
		OLEMethod(DISPATCH_PROPERTYGET, &result, pDocApp, L"Selection", 0);
		pSelection=result.pdispVal;
	}
	
	{
		COleVariant varText(szText);
		OLEMethod(DISPATCH_METHOD,NULL,pSelection,L"TypeText",1,varText.Detach());
	}
	pDocApp->Release();
	pSelection->Release();
	return m_hr;
}

HRESULT CMSWord::SetFont(LPCTSTR szFontName, int nSize, bool bBold, bool bItalic,COLORREF crColor)
{
	if(!m_pWApp || !m_pActiveDocument) return E_FAIL;
	IDispatch *pDocApp;
	{  
		VARIANT result;
		VariantInit(&result);
		OLEMethod(DISPATCH_PROPERTYGET, &result, m_pActiveDocument, L"Application", 0);
		pDocApp= result.pdispVal;
	}
	IDispatch *pSelection;
	{
		VARIANT result;
		VariantInit(&result);
		OLEMethod(DISPATCH_PROPERTYGET, &result, pDocApp, L"Selection", 0);
		pSelection=result.pdispVal;
	}
	IDispatch *pFont;
	{
		VARIANT result;
		VariantInit(&result);
		OLEMethod(DISPATCH_PROPERTYGET, &result, pSelection, L"Font", 0);
		pFont=result.pdispVal;
	}
	{
		COleVariant oleName(szFontName);
		m_hr=OLEMethod(DISPATCH_PROPERTYPUT, NULL, pFont, L"Name", 1, oleName.Detach());
		VARIANT x;
		x.vt = VT_I4;
		x.lVal = nSize;
		m_hr=OLEMethod(DISPATCH_PROPERTYPUT, NULL, pFont, L"Size", 1, x);
		x.lVal = crColor;
		m_hr=OLEMethod(DISPATCH_PROPERTYPUT, NULL, pFont, L"Color", 1, x);
		x.lVal = bBold?1:0;
		m_hr=OLEMethod(DISPATCH_PROPERTYPUT, NULL, pFont, L"Bold", 1, x);
		x.lVal = bItalic?1:0;
		m_hr=OLEMethod(DISPATCH_PROPERTYPUT, NULL, pFont, L"Italic", 1, x);
	}
	pFont->Release();
	pSelection->Release();
	pDocApp->Release();
	return m_hr;
}

HRESULT CMSWord::MoveCursor(int nDirection,bool bSelection)
{
	if(!m_pWApp || !m_pActiveDocument) return E_FAIL;
	IDispatch *pDocApp;
	{  
		VARIANT result;
		VariantInit(&result);
		OLEMethod(DISPATCH_PROPERTYGET, &result, m_pActiveDocument, L"Application", 0);
		pDocApp= result.pdispVal;
	}
	IDispatch *pSelection;
	{
		VARIANT result;
		VariantInit(&result);
		OLEMethod(DISPATCH_PROPERTYGET, &result, pDocApp, L"Selection", 0);
		pSelection=result.pdispVal;
	}

	{
		VARIANT wdCharacter, wdExtend,Count;
		wdCharacter.vt =VT_I4;
		wdCharacter.lVal =1;
		wdExtend.vt =VT_I4;
		wdExtend.lVal =bSelection?1:0;
		Count.vt =VT_I4;
		Count.lVal =1;
		switch(nDirection)
		{
			case 1:
				OLEMethod(DISPATCH_METHOD, NULL, pSelection, L"MoveLeft", 3,wdExtend,Count,wdCharacter);
				break;
			case 2:
				OLEMethod(DISPATCH_METHOD, NULL, pSelection, L"MoveRight", 3,wdExtend,Count,wdCharacter);
				break;
			case 3:
				wdCharacter.lVal =5;
				OLEMethod(DISPATCH_METHOD, NULL, pSelection, L"MoveUp", 3,wdExtend,Count,wdCharacter);
				break;
			case 4:
				wdCharacter.lVal =5;
				OLEMethod(DISPATCH_METHOD, NULL, pSelection, L"MoveDown", 3,wdExtend,Count,wdCharacter);
				break;
		}
	}
	pDocApp->Release();
	pSelection->Release();
	return m_hr;
}

HRESULT CMSWord::AddComment(LPCTSTR szComment)
{
	if(!m_pWApp || !m_pActiveDocument) return E_FAIL;
	IDispatch *pDocApp;
	{  
		VARIANT result;
		VariantInit(&result);
		OLEMethod(DISPATCH_PROPERTYGET, &result, m_pActiveDocument, L"Application", 0);
		pDocApp= result.pdispVal;
	}
	IDispatch *pSelection;
	{
		VARIANT result;
		VariantInit(&result);
		OLEMethod(DISPATCH_PROPERTYGET, &result, pDocApp, L"Selection", 0);
		pSelection=result.pdispVal;
	}
	IDispatch *pRange;
	{
		VARIANT result;
		VariantInit(&result);
		OLEMethod(DISPATCH_PROPERTYGET, &result, pSelection, L"Range", 0);
		pRange = result.pdispVal;
	}
	IDispatch *pComments;
	{  
		VARIANT result;
		VariantInit(&result);
		OLEMethod(DISPATCH_PROPERTYGET, &result, m_pActiveDocument, L"Comments", 0);
		pComments = result.pdispVal;
	}
	{
		COleVariant sCmt(szComment);
		VARIANT result;
        VariantInit(&result);
		VARIANT varRange;
		varRange.vt = VT_DISPATCH;
		varRange.pdispVal = pRange;
		OLEMethod(DISPATCH_METHOD, &result, pComments, L"Add", 2, sCmt.Detach(), varRange);
	}
	pRange->Release();
	pComments->Release();
	pDocApp->Release();
	pSelection->Release();
	return m_hr;
}

HRESULT CMSWord::FindFirst(LPCTSTR szText)
{
	if(m_pWApp==NULL) return E_FAIL;
	IDispatch *pDocApp;
	{  
		VARIANT result;
		VariantInit(&result);
		OLEMethod(DISPATCH_PROPERTYGET, &result, m_pActiveDocument, L"Application", 0);
		pDocApp= result.pdispVal;
	}
	IDispatch *pSelection;
	{
		VARIANT result;
		VariantInit(&result);
		OLEMethod(DISPATCH_PROPERTYGET, &result, pDocApp, L"Selection", 0);
		pSelection=result.pdispVal;
	}
	IDispatch *pFind;
	{
		VARIANT result;
		VariantInit(&result);
		OLEMethod(DISPATCH_PROPERTYGET, &result, pSelection, L"Find", 0);
		pFind=result.pdispVal;

	}
	{
	OLEMethod(DISPATCH_METHOD, NULL, pFind, L"ClearFormatting", 0);
	}
	IDispatch *pReplacement;
	{
		VARIANT result;
		VariantInit(&result);
		OLEMethod(DISPATCH_PROPERTYGET, &result, pFind, L"Replacement", 0);
		pReplacement=result.pdispVal;

	}

	{
		COleVariant varText(szText);
		OLEMethod(DISPATCH_PROPERTYPUT,NULL,pFind,L"Text",1,varText.Detach());
		char* txt="";
		CString str(txt);
		LPCTSTR lpStr = (LPCTSTR)str;
		COleVariant varReplacement(lpStr);
		OLEMethod(DISPATCH_PROPERTYPUT,NULL,pReplacement,L"Text",1,varReplacement.Detach());
		COleVariant varForward((BYTE)1);
		OLEMethod(DISPATCH_PROPERTYPUT,NULL,pFind,L"Forward",1,varForward.Detach());
		VARIANT wdFindContinue;
		wdFindContinue.vt =VT_I4;
		wdFindContinue.lVal =1;
		OLEMethod(DISPATCH_PROPERTYPUT,NULL,pFind,L"Wrap",1,wdFindContinue);
		COleVariant varFormat((BYTE)0);
		OLEMethod(DISPATCH_PROPERTYPUT,NULL,pFind,L"Format",1,varFormat.Detach());
		COleVariant varMatchCase((BYTE)0);
		OLEMethod(DISPATCH_PROPERTYPUT,NULL,pFind,L"MatchCase",1,varMatchCase.Detach());
		COleVariant varMatchWholeWord((BYTE)0);
		OLEMethod(DISPATCH_PROPERTYPUT,NULL,pFind,L"MatchWholeWord",1,varMatchWholeWord.Detach());
		COleVariant varMatchWildcards((BYTE)0);
		OLEMethod(DISPATCH_PROPERTYPUT,NULL,pFind,L"MatchWildcards",1,varMatchWildcards.Detach());
		COleVariant varMatchSoundsLike((BYTE)0);
		OLEMethod(DISPATCH_PROPERTYPUT,NULL,pFind,L"MatchSoundsLike",1,varMatchSoundsLike.Detach());
		COleVariant varMatchAllWordForms((BYTE)0);
		OLEMethod(DISPATCH_PROPERTYPUT,NULL,pFind,L"MatchAllWordForms",1,varMatchAllWordForms.Detach());
		VARIANT res;
		VariantInit(&res);
		OLEMethod(DISPATCH_METHOD,&res,pFind,L"Execute",0);
		

		
	}
	pDocApp->Release();
	pSelection->Release();
	pReplacement->Release();
	pFind->Release();
	return m_hr;
}


bool CMSWord::FindFirstBool(LPCTSTR szText)
{
	bool retur;
	if(m_pWApp==NULL) return E_FAIL;
	IDispatch *pDocApp;
	{  
		VARIANT result;
		VariantInit(&result);
		OLEMethod(DISPATCH_PROPERTYGET, &result, m_pActiveDocument, L"Application", 0);
		pDocApp= result.pdispVal;
	}
	IDispatch *pSelection;
	{
		VARIANT result;
		VariantInit(&result);
		OLEMethod(DISPATCH_PROPERTYGET, &result, pDocApp, L"Selection", 0);
		pSelection=result.pdispVal;
	}
	IDispatch *pFind;
	{
		VARIANT result;
		VariantInit(&result);
		OLEMethod(DISPATCH_PROPERTYGET, &result, pSelection, L"Find", 0);
		pFind=result.pdispVal;

	}
	{
	OLEMethod(DISPATCH_METHOD, NULL, pFind, L"ClearFormatting", 0);
	}
	IDispatch *pReplacement;
	{
		VARIANT result;
		VariantInit(&result);
		OLEMethod(DISPATCH_PROPERTYGET, &result, pFind, L"Replacement", 0);
		pReplacement=result.pdispVal;

	}

	{
		COleVariant varText(szText);
		OLEMethod(DISPATCH_PROPERTYPUT,NULL,pFind,L"Text",1,varText.Detach());
		char* txt="";
		CString str(txt);
		LPCTSTR lpStr = (LPCTSTR)str;
		COleVariant varReplacement(lpStr);
		OLEMethod(DISPATCH_PROPERTYPUT,NULL,pReplacement,L"Text",1,varReplacement.Detach());
		COleVariant varForward((BYTE)1);
		OLEMethod(DISPATCH_PROPERTYPUT,NULL,pFind,L"Forward",1,varForward.Detach());
		VARIANT wdFindContinue;
		wdFindContinue.vt =VT_I4;
		wdFindContinue.lVal =1;
		OLEMethod(DISPATCH_PROPERTYPUT,NULL,pFind,L"Wrap",1,wdFindContinue);
		COleVariant varFormat((BYTE)0);
		OLEMethod(DISPATCH_PROPERTYPUT,NULL,pFind,L"Format",1,varFormat.Detach());
		COleVariant varMatchCase((BYTE)0);
		OLEMethod(DISPATCH_PROPERTYPUT,NULL,pFind,L"MatchCase",1,varMatchCase.Detach());
		COleVariant varMatchWholeWord((BYTE)0);
		OLEMethod(DISPATCH_PROPERTYPUT,NULL,pFind,L"MatchWholeWord",1,varMatchWholeWord.Detach());
		COleVariant varMatchWildcards((BYTE)0);
		OLEMethod(DISPATCH_PROPERTYPUT,NULL,pFind,L"MatchWildcards",1,varMatchWildcards.Detach());
		COleVariant varMatchSoundsLike((BYTE)0);
		OLEMethod(DISPATCH_PROPERTYPUT,NULL,pFind,L"MatchSoundsLike",1,varMatchSoundsLike.Detach());
		COleVariant varMatchAllWordForms((BYTE)0);
		OLEMethod(DISPATCH_PROPERTYPUT,NULL,pFind,L"MatchAllWordForms",1,varMatchAllWordForms.Detach());
		VARIANT res;
		VariantInit(&res);
		OLEMethod(DISPATCH_METHOD,&res,pFind,L"Execute",0);
		if(res.boolVal==-1)retur=true;else retur=false;
		
		

		
	}
	pDocApp->Release();
	pSelection->Release();
	pReplacement->Release();
	pFind->Release();
	return retur;
}


HRESULT CMSWord::Replace(LPCTSTR szText,LPCTSTR szReplacementText,bool ReplaceAll)
{
	if(m_pWApp==NULL) return E_FAIL;
	IDispatch *pDocApp;
	{  
		VARIANT result;
		VariantInit(&result);
		OLEMethod(DISPATCH_PROPERTYGET, &result, m_pActiveDocument, L"Application", 0);
		pDocApp= result.pdispVal;
	}
	IDispatch *pSelection;
	{
		VARIANT result;
		VariantInit(&result);
		OLEMethod(DISPATCH_PROPERTYGET, &result, pDocApp, L"Selection", 0);
		pSelection=result.pdispVal;
	}
	IDispatch *pFind;
	{
		VARIANT result;
		VariantInit(&result);
		OLEMethod(DISPATCH_PROPERTYGET, &result, pSelection, L"Find", 0);
		pFind=result.pdispVal;

	}
	{
	OLEMethod(DISPATCH_METHOD, NULL, pFind, L"ClearFormatting", 0);
	}
	IDispatch *pReplacement;
	{
		VARIANT result;
		VariantInit(&result);
		OLEMethod(DISPATCH_PROPERTYGET, &result, pFind, L"Replacement", 0);
		pReplacement=result.pdispVal;

	}
	OLEMethod(DISPATCH_METHOD, NULL, pReplacement, L"ClearFormatting", 0);

	{
		COleVariant varText(szText);
		OLEMethod(DISPATCH_PROPERTYPUT,NULL,pFind,L"Text",1,varText.Detach());
		//char* txt="";
		//CString str(txt);
		//LPCTSTR lpStr = (LPCTSTR)str;
		COleVariant varReplacement(szReplacementText);
		OLEMethod(DISPATCH_PROPERTYPUT,NULL,pReplacement,L"Text",1,varReplacement.Detach());
		COleVariant varForward((BYTE)1);
		OLEMethod(DISPATCH_PROPERTYPUT,NULL,pFind,L"Forward",1,varForward.Detach());
		VARIANT wdFindContinue;
		wdFindContinue.vt =VT_I4;
		wdFindContinue.lVal =1;
		OLEMethod(DISPATCH_PROPERTYPUT,NULL,pFind,L"Wrap",1,wdFindContinue);
		COleVariant varFormat((BYTE)0);
		OLEMethod(DISPATCH_PROPERTYPUT,NULL,pFind,L"Format",1,varFormat.Detach());
		COleVariant varMatchCase((BYTE)0);
		OLEMethod(DISPATCH_PROPERTYPUT,NULL,pFind,L"MatchCase",1,varMatchCase.Detach());
		COleVariant varMatchWholeWord((BYTE)0);
		OLEMethod(DISPATCH_PROPERTYPUT,NULL,pFind,L"MatchWholeWord",1,varMatchWholeWord.Detach());
		COleVariant varMatchWildcards((BYTE)0);
		OLEMethod(DISPATCH_PROPERTYPUT,NULL,pFind,L"MatchWildcards",1,varMatchWildcards.Detach());
		COleVariant varMatchSoundsLike((BYTE)0);
		OLEMethod(DISPATCH_PROPERTYPUT,NULL,pFind,L"MatchSoundsLike",1,varMatchSoundsLike.Detach());
		COleVariant varMatchAllWordForms((BYTE)0);
		OLEMethod(DISPATCH_PROPERTYPUT,NULL,pFind,L"MatchAllWordForms",1,varMatchAllWordForms.Detach());
		VARIANT wdReplaceAll;
		COleVariant r1((BYTE)0);
		COleVariant r2((BYTE)1);
		COleVariant r3((BYTE)1);
		COleVariant r4((BYTE)1);
		
		if(ReplaceAll){
		
		wdReplaceAll.vt =VT_I4;
		wdReplaceAll.lVal =2;
		//OLEMethod(DISPATCH_PROPERTYPUT,NULL,pFind,L"Replace",1,wdReplaceAll);
		}else{
		
		wdReplaceAll.vt =VT_I4;
		wdReplaceAll.lVal =1;
		//OLEMethod(DISPATCH_PROPERTYPUT,NULL,pFind,L"Replace",1,wdReplaceOne);
		}
		
		OLEMethod(DISPATCH_METHOD,NULL,pFind,L"Execute",0);
		
	}
	pDocApp->Release();
	pSelection->Release();
	pReplacement->Release();
	pFind->Release();
	return m_hr;
}

/*
HRESULT CMSWord::SetSelectionText(LPCTSTR szText)
{
	if(m_pWApp==NULL) return E_FAIL;
	IDispatch *pDocApp;
	{  
		VARIANT result;
		VariantInit(&result);
		OLEMethod(DISPATCH_PROPERTYGET, &result, m_pActiveDocument, L"Application", 0);
		pDocApp= result.pdispVal;
	}
	IDispatch *pSelection;
	{
		VARIANT result;
		VariantInit(&result);
		OLEMethod(DISPATCH_PROPERTYGET, &result, pDocApp, L"Selection", 0);
		pSelection=result.pdispVal;
	}

	{
		COleVariant sTxt(szText);
		VARIANT text=sTxt.Detach();
		OLEMethod(DISPATCH_METHOD, NULL, pSelection, L"TypeText", 1,text);

	}
	pDocApp->Release();
	pSelection->Release();
	return m_hr;
}*/

// bSave=true will autosave the document. bSave=false will close document without saving.
HRESULT CMSWord::CloseActiveDocument(bool bSave)
{
	if(m_pWApp==NULL) return E_FAIL;
	{
		VARIANT result;
        VariantInit(&result);
		COleVariant varSave;
		varSave.boolVal=bSave;
		m_hr=OLEMethod(DISPATCH_METHOD, &result, m_pActiveDocument, L"Close", 1, varSave.Detach());
		m_pActiveDocument=NULL;
	}	 
	return m_hr;
}

HRESULT CMSWord::Quit()
{
	if(m_pWApp==NULL) return E_FAIL;
	DISPID dispID;
	LPOLESTR ptName=_T("Quit");
	m_hr = m_pWApp->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
	
	if(SUCCEEDED(m_hr))
	{
		DISPPARAMS dp = { NULL, NULL, 0, 0 };
		m_hr = m_pWApp->Invoke(dispID, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_METHOD, 
									&dp, NULL, NULL, NULL);
	}
	return m_hr;
}

CMSWord::~CMSWord()
{
	Quit();
	if(m_pWApp!=NULL) 
	{
		m_pWApp->Release();
		m_pWApp=NULL;
	}
	CoUninitialize();
}

