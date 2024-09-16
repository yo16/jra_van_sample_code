// ControlDB.h : PROJECT_NAME アプリケーションのメイン ヘッダー ファイルです。
//

#pragma once

#ifndef __AFXWIN_H__
	#error include 'stdafx.h' before including this file for PCH
#endif

#include "resource.h"		// メイン シンボル


// CControlDBApp:
// このクラスの実装については、ControlDB.cpp を参照してください。
//

class CControlDBApp : public CWinApp
{
public:
	CControlDBApp();

// オーバーライド
	public:
	virtual BOOL InitInstance();

// 実装

	DECLARE_MESSAGE_MAP()
};

extern CControlDBApp theApp;
