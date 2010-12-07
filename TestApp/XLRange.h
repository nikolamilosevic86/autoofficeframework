#pragma once


// CXLRange dialog

class CXLRange : public CDialog
{
	DECLARE_DYNAMIC(CXLRange)

public:
	CXLRange(CWnd* pParent = NULL);   // standard constructor
	virtual ~CXLRange();

// Dialog Data
	enum { IDD = IDD_XLRANGE };
	
protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	TCHAR m_szValue[255];
	TCHAR m_szTitle[255];
	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedOk();
	CString GetValue();
	void SetTitle(LPCTSTR szTitle);
	void SetValue(LPCTSTR szValue);
	virtual BOOL OnInitDialog();
};
