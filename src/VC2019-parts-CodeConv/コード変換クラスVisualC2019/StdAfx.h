// stdafx.h : �W���̃V�X�e�� �C���N���[�h �t�@�C���A
//            �܂��͎Q�Ɖ񐔂������A�����܂�ύX����Ȃ�
//            �v���W�F�N�g��p�̃C���N���[�h �t�@�C�����L�q���܂��B
//

#if !defined(AFX_STDAFX_H__B1689EC0_3BFF_4F6C_8618_907791D9E3BF__INCLUDED_)
#define AFX_STDAFX_H__B1689EC0_3BFF_4F6C_8618_907791D9E3BF__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#define VC_EXTRALEAN		// Windows �w�b�_�[����w�ǎg�p����Ȃ��X�^�b�t�����O���܂��B
#define NO_WARN_MBCS_MFC_DEPRECATION

#include <afxwin.h>         // MFC �̃R�A����ѕW���R���|�[�l���g
#include <afxext.h>         // MFC �̊g������
#include <afxdisp.h>        // MFC �̃I�[�g���[�V���� �N���X
#include <afxdtctl.h>		// MFC �� Internet Explorer 4 �R���� �R���g���[�� �T�|�[�g
#ifndef _AFX_NO_AFXCMN_SUPPORT
#include <afxcmn.h>			// MFC �� Windows �R���� �R���g���[�� �T�|�[�g
#endif // _AFX_NO_AFXCMN_SUPPORT


// ���̃}�N���� bMultiInstance �p�����[�^�p�� COleObjectFactory �R���X�g���N�^
// TRUE ��n���ȊO�� IMPLEMENT_OLECREATE �Ɠ����ł��B
// �I�[�g���[�V���� �R���g���[���ɂ���ėv�������e�I�[�g���[�V���� �v���L�V
// �I�u�W�F�N�g�ɑ΂��ċN�������悤�ɂ��̃A�v���P�[�V�����̃C���X�^���X�𕪂��܂��B
#ifndef IMPLEMENT_OLECREATE2
#define IMPLEMENT_OLECREATE2(class_name, external_name, l, w1, w2, b1, b2, b3, b4, b5, b6, b7, b8) \
	AFX_DATADEF COleObjectFactory class_name::factory(class_name::guid, \
		RUNTIME_CLASS(class_name), TRUE, _T(external_name)); \
	const AFX_DATADEF GUID class_name::guid = \
		{ l, w1, w2, { b1, b2, b3, b4, b5, b6, b7, b8 } };
#endif // IMPLEMENT_OLECREATE2

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ �͑O�s�̒��O�ɒǉ��̐錾��}�����܂��B

#endif // !defined(AFX_STDAFX_H__B1689EC0_3BFF_4F6C_8618_907791D9E3BF__INCLUDED_)
