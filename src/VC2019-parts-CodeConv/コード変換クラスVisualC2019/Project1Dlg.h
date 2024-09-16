// Project1Dlg.h : �w�b�_�[ �t�@�C��
//

#if !defined(AFX_PROJECT1DLG_H__2275A368_7152_4D46_939A_F5A1456BD5D4__INCLUDED_)
#define AFX_PROJECT1DLG_H__2275A368_7152_4D46_939A_F5A1456BD5D4__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

class CProject1DlgAutoProxy;

/////////////////////////////////////////////////////////////////////////////
// CProject1Dlg �_�C�A���O

class CProject1Dlg : public CDialog
{
	DECLARE_DYNAMIC(CProject1Dlg);
	friend class CProject1DlgAutoProxy;

// �\�z
public:
	CProject1Dlg(CWnd* pParent = NULL);	// �W���̃R���X�g���N�^
	virtual ~CProject1Dlg();

// �_�C�A���O �f�[�^
	//{{AFX_DATA(CProject1Dlg)
	enum { IDD = IDD_PROJECT1_DIALOG };
	CString	m_eCode;
	CString	m_eValue;
	CString	m_eNum;
	CString	m_eOut;
	//}}AFX_DATA

	// ClassWizard �͉��z�֐��̃I�[�o�[���C�h�𐶐����܂��B
	//{{AFX_VIRTUAL(CProject1Dlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV �̃T�|�[�g
	//}}AFX_VIRTUAL

// �C���v�������e�[�V����
protected:
	CProject1DlgAutoProxy* m_pAutoProxy;
	HICON m_hIcon;

	BOOL CanExit();

	// �������ꂽ���b�Z�[�W �}�b�v�֐�
	//{{AFX_MSG(CProject1Dlg)
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	afx_msg void OnClose();
	virtual void OnOK();
	virtual void OnCancel();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ �͑O�s�̒��O�ɒǉ��̐錾��}�����܂��B

#endif // !defined(AFX_PROJECT1DLG_H__2275A368_7152_4D46_939A_F5A1456BD5D4__INCLUDED_)
