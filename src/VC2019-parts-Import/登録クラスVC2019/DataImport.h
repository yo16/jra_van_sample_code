// DataImport.h : PROJECT_NAME �A�v���P�[�V�����̃��C�� �w�b�_�[ �t�@�C���ł��B
//

#pragma once

#ifndef __AFXWIN_H__
	#error include 'stdafx.h' before including this file for PCH
#endif

#include "resource.h"		// ���C�� �V���{��


// CDataImportApp:
// ���̃N���X�̎����ɂ��ẮADataImport.cpp ���Q�Ƃ��Ă��������B
//

class CDataImportApp : public CWinApp
{
public:
	CDataImportApp();

// �I�[�o�[���C�h
	public:
	virtual BOOL InitInstance();

// ����

	DECLARE_MESSAGE_MAP()
};

extern CDataImportApp theApp;
