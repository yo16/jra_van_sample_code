// DlgProxy.h : �w�b�_�[ �t�@�C��
//

#if !defined(AFX_DLGPROXY_H__957DA897_FACB_4B75_B5A7_58B0DCBA72D6__INCLUDED_)
#define AFX_DLGPROXY_H__957DA897_FACB_4B75_B5A7_58B0DCBA72D6__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

class CProject1Dlg;

/////////////////////////////////////////////////////////////////////////////
// CProject1DlgAutoProxy �R�}���h �^�[�Q�b�g

class CProject1DlgAutoProxy : public CCmdTarget
{
	DECLARE_DYNCREATE(CProject1DlgAutoProxy)

	CProject1DlgAutoProxy();           // ���I�����Ŏg�p����� protected �R���X�g���N�^

// �A�g���r���[�g
public:
	CProject1Dlg* m_pDialog;

// �I�y���[�V����
public:

// �I�[�o�[���C�h
	// ClassWizard �͉��z�֐��̃I�[�o�[���C�h�𐶐����܂�
	//{{AFX_VIRTUAL(CProject1DlgAutoProxy)
	public:
	virtual void OnFinalRelease();
	//}}AFX_VIRTUAL

// �C���v�������e�[�V����
protected:
	virtual ~CProject1DlgAutoProxy();

	// �������ꂽ���b�Z�[�W �}�b�v�֐�
	//{{AFX_MSG(CProject1DlgAutoProxy)
		// ���� - ClassWizard �͂��̈ʒu�Ƀ����o�֐���ǉ��܂��͍폜���܂��B
	//}}AFX_MSG

	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(CProject1DlgAutoProxy)

	// �������ꂽ OLE �f�B�X�p�b�` �}�b�v�֐�
	//{{AFX_DISPATCH(CProject1DlgAutoProxy)
		// ���� - ClassWizard �͂��̈ʒu�Ƀ����o�֐���ǉ��܂��͍폜���܂��B
	//}}AFX_DISPATCH
	DECLARE_DISPATCH_MAP()
	DECLARE_INTERFACE_MAP()
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ �͑O�s�̒��O�ɒǉ��̐錾��}�����܂��B

#endif // !defined(AFX_DLGPROXY_H__957DA897_FACB_4B75_B5A7_58B0DCBA72D6__INCLUDED_)
