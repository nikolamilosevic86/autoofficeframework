#include "stdafx.h"
#include "SDictionary.h"
#include "OLEMethod.h"

CSDictionary::CSDictionary()
{
	m_pSDict=NULL;
}

CSDictionary::~CSDictionary()
{
	if(m_pSDict) m_pSDict->Release();
	CoUninitialize();
}

HRESULT CSDictionary::Initialize()
{
	CoInitialize(NULL);
	CLSID clsid;
	HRESULT hr=CLSIDFromProgID(L"Scripting.Dictionary", &clsid);

	if(SUCCEEDED(hr))
	{
		hr = CoCreateInstance(clsid, NULL, CLSCTX_INPROC_SERVER, IID_IDispatch, (void **)&m_pSDict);
		if(FAILED(hr)) m_pSDict=NULL;
	}
	return hr;
}


HRESULT CSDictionary::Add(LPCTSTR szKey, LPCTSTR szItem)
{
	HRESULT hr=E_FAIL;
	COleVariant key(szKey);
	COleVariant item(szItem);
	{
		hr=OLEMethod(DISPATCH_METHOD, NULL, m_pSDict, L"Add", 2,item.Detach(),key.Detach());
	}
	return hr;
}

bool CSDictionary::Exists(LPCTSTR szKey)
{
	COleVariant key(szKey);
	VARIANT result;
	VariantInit(&result);
	OLEMethod(DISPATCH_METHOD, &result, m_pSDict, L"Exists", 1,key.Detach());
	return result.boolVal;
}

HRESULT CSDictionary::Remove(LPCTSTR szKey)
{
	HRESULT hr=E_FAIL;
	COleVariant key(szKey);
	{
		hr=OLEMethod(DISPATCH_METHOD, NULL, m_pSDict, L"Remove", 1,key.Detach());
	}
	return hr;
}

HRESULT CSDictionary::RemoveAll()
{
	HRESULT hr=E_FAIL;
	{
		hr=OLEMethod(DISPATCH_METHOD, NULL, m_pSDict, L"RemoveAll", 0);
	}
	return hr;
}

HRESULT CSDictionary::SetItem(LPCTSTR szKey, LPCTSTR szItem)
{
	HRESULT hr=E_FAIL;
	return hr;
}

HRESULT CSDictionary::SetKey(LPCTSTR szOldKey, LPCTSTR szNewKey)
{
	HRESULT hr=E_FAIL;
	return hr;
}

LPCTSTR CSDictionary::GetItem(LPCTSTR szKey)
{
	VARIANT result;
	VariantInit(&result);
	COleVariant key(szKey);
	OLEMethod(DISPATCH_PROPERTYGET, &result, m_pSDict, L"Item", 1,key.Detach());
	return result.bstrVal;
}

int CSDictionary::GetCount()
{
	VARIANT result;
	VariantInit(&result);
	OLEMethod(DISPATCH_PROPERTYGET, &result, m_pSDict, L"Count", 0);
	return result.iVal;
}
