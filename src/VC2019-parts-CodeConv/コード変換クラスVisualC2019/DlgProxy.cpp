// DlgProxy.cpp : �C���v�������e�[�V���� �t�@�C��
//

#include "stdafx.h"
#include "Project1.h"
#include "DlgProxy.h"
#include "Project1Dlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CProject1DlgAutoProxy

IMPLEMENT_DYNCREATE(CProject1DlgAutoProxy, CCmdTarget)

CProject1DlgAutoProxy::CProject1DlgAutoProxy()
{
	EnableAutomation();
	
	// �I�[�g���[�V���� �I�u�W�F�N�g���A�N�e�B�u�ł������A�A�v���P�[�V������ 
	// ���s��Ԃɂ��Ă��������A�R���X�g���N�^�� AfxOleLockApp ���Ăяo���܂��B
	AfxOleLockApp();

	// �A�v���P�[�V�����̃��C�� �E�B���h�E �|�C���^��ʂ��ă_�C�A���O
	// �փA�N�Z�X���܂��B�v���L�V�̓����|�C���^����_�C�A���O�ւ̃|�C
	// ���^��ݒ肵�A�_�C�A���O�̖߂�|�C���^�����̃v���L�V�֐ݒ肵��
	// ���B
	ASSERT (AfxGetApp()->m_pMainWnd != NULL);
	ASSERT_VALID (AfxGetApp()->m_pMainWnd);
	ASSERT_KINDOF(CProject1Dlg, AfxGetApp()->m_pMainWnd);
	m_pDialog = (CProject1Dlg*) AfxGetApp()->m_pMainWnd;
	m_pDialog->m_pAutoProxy = this;
}

CProject1DlgAutoProxy::~CProject1DlgAutoProxy()
{
	// ���ׂẴI�u�W�F�N�g���I�[�g���[�V�����ō쐬���ꂽ�ꍇ�ɃA�v���P�[�V����
	// ���I�����邽�߂ɁA�f�X�g���N�^�� AfxOleUnlockApp ���Ăяo���܂��B
	// ���̏����̊ԂɁA���C�� �_�C�A���O��j�󂵂܂��B
	if (m_pDialog != NULL)
		m_pDialog->m_pAutoProxy = NULL;
	AfxOleUnlockApp();
}

void CProject1DlgAutoProxy::OnFinalRelease()
{
	// �I�[�g���[�V���� �I�u�W�F�N�g�ɑ΂���Ō�̎Q�Ƃ��������鎞��
	// OnFinalRelease ���Ăяo����܂��B��{�N���X�͎����I�ɃI�u�W�F�N
	// �g���폜���܂��B��{�N���X���Ăяo���O�ɁA�I�u�W�F�N�g�ŕK�v�ȓ�
	// �ʂȌ㏈����ǉ����Ă��������B

	CCmdTarget::OnFinalRelease();
}

BEGIN_MESSAGE_MAP(CProject1DlgAutoProxy, CCmdTarget)
	//{{AFX_MSG_MAP(CProject1DlgAutoProxy)
		// ���� - ClassWizard �͂��̈ʒu�Ƀ}�b�s���O�p�̃}�N����ǉ��܂��͍폜���܂��B
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

BEGIN_DISPATCH_MAP(CProject1DlgAutoProxy, CCmdTarget)
	//{{AFX_DISPATCH_MAP(CProject1DlgAutoProxy)
		// ���� - ClassWizard �͂��̈ʒu�Ƀ}�b�s���O�p�̃}�N����ǉ��܂��͍폜���܂��B
	//}}AFX_DISPATCH_MAP
END_DISPATCH_MAP()

// ����: VBA ����^�C�v �Z�[�t�ȃo�C���f�B���O���T�|�[�g���邽�߂� IID_IProject1
//  �̃T�|�[�g��ǉ����܂��B���� IID �� .ODL �t�@�C�����̃f�B�X�p�b�` �C���^�[�t�F�C�X 
//  �փA�^�b�`����� GUID �ƍ��v���Ȃ���΂Ȃ�܂���B

// {D32FF675-669B-4A8C-B480-DA7B832CBA00}
static const IID IID_IProject1 =
{ 0xd32ff675, 0x669b, 0x4a8c, { 0xb4, 0x80, 0xda, 0x7b, 0x83, 0x2c, 0xba, 0x0 } };

BEGIN_INTERFACE_MAP(CProject1DlgAutoProxy, CCmdTarget)
	INTERFACE_PART(CProject1DlgAutoProxy, IID_IProject1, Dispatch)
END_INTERFACE_MAP()

// IMPLEMENT_OLECREATE2 �}�N���͂��̃v���W�F�N�g�� StdAfx.h ���ɒ�`����Ă��܂��B
// {4805CCB7-8A33-4029-ACB0-7EE3162679F8}
IMPLEMENT_OLECREATE2(CProject1DlgAutoProxy, "Project1.Application", 0x4805ccb7, 0x8a33, 0x4029, 0xac, 0xb0, 0x7e, 0xe3, 0x16, 0x26, 0x79, 0xf8)

/////////////////////////////////////////////////////////////////////////////
// CProject1DlgAutoProxy ���b�Z�[�W �n���h��
