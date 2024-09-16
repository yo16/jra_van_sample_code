// Project1.cpp : �A�v���P�[�V�����p�N���X�̒�`���s���܂��B
//

#include "stdafx.h"
#include "Project1.h"
#include "Project1Dlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CProject1App

BEGIN_MESSAGE_MAP(CProject1App, CWinApp)
	//{{AFX_MSG_MAP(CProject1App)
		// ���� - ClassWizard �͂��̈ʒu�Ƀ}�b�s���O�p�̃}�N����ǉ��܂��͍폜���܂��B
		//        ���̈ʒu�ɐ��������R�[�h��ҏW���Ȃ��ł��������B
	//}}AFX_MSG
	ON_COMMAND(ID_HELP, CWinApp::OnHelp)
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CProject1App �N���X�̍\�z

CProject1App::CProject1App()
{
	// TODO: ���̈ʒu�ɍ\�z�p�̃R�[�h��ǉ����Ă��������B
	// ������ InitInstance ���̏d�v�ȏ��������������ׂċL�q���Ă��������B
}

/////////////////////////////////////////////////////////////////////////////
// �B��� CProject1App �I�u�W�F�N�g

CProject1App theApp;

/////////////////////////////////////////////////////////////////////////////
// CProject1App �N���X�̏�����

BOOL CProject1App::InitInstance()
{
	// OLE ���C�u�����̏�����
	if (!AfxOleInit())
	{
		AfxMessageBox(IDP_OLE_INIT_FAILED);
		return FALSE;
	}

	AfxEnableControlContainer();

	// �W���I�ȏ���������
	// ���������̋@�\���g�p�����A���s�t�@�C���̃T�C�Y��������������
	//  ��Έȉ��̓���̏��������[�`���̒�����s�K�v�Ȃ��̂��폜����
	//  ���������B
#if _MSC_VER <= 1200
	#ifdef _AFXDLL
		Enable3dControls();			// ���L DLL ���� MFC ���g���ꍇ�͂������R�[�����Ă��������B
	#else
		Enable3dControlsStatic();	// MFC �ƐÓI�Ƀ����N����ꍇ�͂������R�[�����Ă��������B
	#endif
#endif
	// OLE �T�[�o�[�Ƃ��ċN������Ă��鎞�ɂ̓R�}���h���C������͂��܂��B
	if (RunEmbedded() || RunAutomated())
	{
		// ���s����Ƃ��ׂĂ� OLE �T�[�o�[ �t�@�N�g�� ��o�^���܂��B
		//  ���̃A�v���P�[�V��������I�u�W�F�N�g����邽�߂� OLE ���C�u�������g�p�\�ɂ��܂��B
		COleTemplateServer::RegisterAll();
	}
	else
	{
		// �V�X�e�� ���W�X�g�������Ă��ăT�[�o�[ �A�v���P�[�V�������X�^���h �A������
		// �N�����ꂽ���ɂ́A�V�X�e�� ���W�X�g�����X�V���Ă��������B
		COleObjectFactory::UpdateRegistryAll();
	}

	CProject1Dlg dlg;
	m_pMainWnd = &dlg;
	int nResponse = dlg.DoModal();
	if (nResponse == IDOK)
	{
		// TODO: �_�C�A���O�� <OK> �ŏ����ꂽ���̃R�[�h��
		//       �L�q���Ă��������B
	}
	else if (nResponse == IDCANCEL)
	{
		// TODO: �_�C�A���O�� <��ݾ�> �ŏ����ꂽ���̃R�[�h��
		//       �L�q���Ă��������B
	}

	// �_�C�A���O�������Ă���A�v���P�[�V�����̃��b�Z�[�W �|���v���J�n������́A
	// �A�v���P�[�V�������I�����邽�߂� FALSE ��Ԃ��Ă��������B
	return FALSE;
}
