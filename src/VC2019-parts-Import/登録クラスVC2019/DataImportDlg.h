// DataImportDlg.h : �w�b�_�[ �t�@�C��
//
#include "cjvlink.h"
#pragma once


// CDataImportDlg �_�C�A���O
class CDataImportDlg : public CDialog
{
// �R���X�g���N�V����
public:
	CDataImportDlg(CWnd* pParent = NULL);	// �W���R���X�g���N�^

// �_�C�A���O �f�[�^
	enum { IDD = IDD_DATAIMPORT_DIALOG };
	CJVLink	m_jvlink1;

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV �T�|�[�g


// ����
protected:
	HICON m_hIcon;

	// �������ꂽ�A���b�Z�[�W���蓖�Ċ֐�
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
