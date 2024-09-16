// DataImport.h : PROJECT_NAME アプリケーションのメイン ヘッダー ファイルです。
//

#pragma once

#ifndef __AFXWIN_H__
	#error include 'stdafx.h' before including this file for PCH
#endif

#include "resource.h"		// メイン シンボル


// CDataImportApp:
// このクラスの実装については、DataImport.cpp を参照してください。
//

class CDataImportApp : public CWinApp
{
public:
	CDataImportApp();

// オーバーライド
	public:
	virtual BOOL InitInstance();

// 実装

	DECLARE_MESSAGE_MAP()
};

extern CDataImportApp theApp;
