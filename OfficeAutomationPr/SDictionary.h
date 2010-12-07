#pragma once

class CSDictionary
{
protected:
	IDispatch *m_pSDict;
public:
	CSDictionary();
	~CSDictionary();
	HRESULT Initialize();
	HRESULT Add(LPCTSTR szKey, LPCTSTR szItem);
	bool Exists(LPCTSTR szKey);
	HRESULT Remove(LPCTSTR szKey);
	HRESULT RemoveAll();
	HRESULT SetItem(LPCTSTR szKey, LPCTSTR szItem);
	HRESULT SetKey(LPCTSTR szOldKey, LPCTSTR szNewKey);
	LPCTSTR GetItem(LPCTSTR szKey);
	int GetCount();
};