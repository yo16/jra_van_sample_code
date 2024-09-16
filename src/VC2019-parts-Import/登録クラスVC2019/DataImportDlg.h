// DataImportDlg.h : ヘッダー ファイル
//
#include "cjvlink.h"
#pragma once


// CDataImportDlg ダイアログ
class CDataImportDlg : public CDialog
{
// コンストラクション
public:
	CDataImportDlg(CWnd* pParent = NULL);	// 標準コンストラクタ

// ダイアログ データ
	enum { IDD = IDD_DATAIMPORT_DIALOG };
	CJVLink	m_jvlink1;

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV サポート


// 実装
protected:
	HICON m_hIcon;

	// 生成された、メッセージ割り当て関数
	virtual BOOL OnInitDialog();
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	CEdit	m_txtFromDate;
	CEdit	m_txtDataSpec;
	CEdit	m_txtCount;
	int		m_iRadio;
	void JVClosing();
	void JVReading();
void CDataImportDlg::PumpMessages();
DECLARE_MESSAGE_MAP()

public:
	afx_msg void OnBnClickedButton1();
	afx_msg void OnBnClickedButton2();
	afx_msg void OnBnClickedOk();
	afx_msg void OnEnChangeEdit1();
	afx_msg void OnBnClickedRadiotujo();
	afx_msg void OnBnClickedRadiokonshu();
	afx_msg void OnBnClickedRadiosetup();
	afx_msg void OnBnClickedRadioreal();
};
