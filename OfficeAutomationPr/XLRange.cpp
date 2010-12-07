// XLRange.cpp : implementation file
//

#include "stdafx.h"
#include "OLEAuto.h"
#include "XLRange.h"


// CXLRange dialog

IMPLEMENT_DYNAMIC(CXLRange, CDialog)

CXLRange::CXLRange(CWnd* pParent /*=NULL*/)
	: CDialog(CXLRange::IDD, pParent)
{
	_tcscpy(m_szTitle,_T("Enter XL Range"));
	_tcscpy(m_szValue,_T(""));
}

CXLRange::~CXLRange()
{
}

void CXLRange::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
}


BEGIN_MESSAGE_MAP(CXLRange, CDialog)
	ON_BN_CLICKED(IDOK, &CXLRange::OnBnClickedOk)
END_MESSAGE_MAP()


// CXLRange message handlers

void CXLRange::OnBnClickedOk()
{
	GetDlgItemText(IDC_EDIT1,m_szValue,255);
	OnOK();
}

void CXLRange::SetTitle(LPCTSTR szTitle)
{
	_tcscpy(m_szTitle,szTitle);
}

void CXLRange::SetValue(LPCTSTR szValue)
{
	_tcscpy(m_szValue,szValue);
}

CString CXLRange::GetValue()
{
	return m_szValue;
}

BOOL CXLRange::OnInitDialog()
{
	CDialog::OnInitDialog();

	SetDlgItemText(IDC_EDIT1,m_szValue);
	SetWindowText(m_szTitle);
	return TRUE;  // return TRUE unless you set the focus to a control
	// EXCEPTION: OCX Property Pages should return FALSE
}
