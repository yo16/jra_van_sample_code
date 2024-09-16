/*=======================================================================
  JRA-VAN Data Lab.プログラミングパーツ「データ取り込みパーツ　メインクラス」

	   作成: JRA-VAN ソフトウェア工房  2003年6月26日

========================================================================
   (C) Copyright JRA SYSTEM SERVICE CO.,LTD. 2007 All rights reserved
========================================================================*/
// DataImport.cpp : アプリケーションのクラス動作を定義します。
//

#include "stdafx.h"
#include "DataImport.h"
#include "DataImportDlg.h"
#include "CJVLink.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CDataImportApp

BEGIN_MESSAGE_MAP(CDataImportApp, CWinApp)
	ON_COMMAND(ID_HELP, CWinApp::OnHelp)
END_MESSAGE_MAP()


// CDataImportApp コンストラクション

CDataImportApp::CDataImportApp()
{
	// TODO: この位置に構築用コードを追加してください。
	// ここに InitInstance 中の重要な初期化処理をすべて記述してください。
}


// 唯一の CDataImportApp オブジェクトです。

CDataImportApp theApp;


// CDataImportApp 初期化

BOOL CDataImportApp::InitInstance()
{
	// アプリケーション　マニフェストが　visual スタイルを有効にするために、
	// ComCtl32.dll バージョン 6　以降の使用を指定する場合は、
	// Windows XP に　InitCommonControls() が必要です。さもなければ、ウィンドウ作成はすべて失敗します。
	InitCommonControls();

	CWinApp::InitInstance();

	AfxEnableControlContainer();


	CDataImportDlg dlg;
	m_pMainWnd = &dlg;
	INT_PTR nResponse = dlg.DoModal();
	if (nResponse == IDOK)
	{
		// TODO: ダイアログが <OK> で消された時のコードを
		//       記述してください。
	}
	else if (nResponse == IDCANCEL)
	{
		// TODO: ダイアログが <キャンセル> で消された時のコードを
		//       記述してください。
	}

	// ダイアログは閉じられました。アプリケーションのメッセージ ポンプを開始しないで
	// アプリケーションを終了するために FALSE を返してください。
	return FALSE;
}
