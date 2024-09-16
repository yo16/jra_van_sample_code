// Project1.cpp : アプリケーション用クラスの定義を行います。
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
		// メモ - ClassWizard はこの位置にマッピング用のマクロを追加または削除します。
		//        この位置に生成されるコードを編集しないでください。
	//}}AFX_MSG
	ON_COMMAND(ID_HELP, CWinApp::OnHelp)
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CProject1App クラスの構築

CProject1App::CProject1App()
{
	// TODO: この位置に構築用のコードを追加してください。
	// ここに InitInstance 中の重要な初期化処理をすべて記述してください。
}

/////////////////////////////////////////////////////////////////////////////
// 唯一の CProject1App オブジェクト

CProject1App theApp;

/////////////////////////////////////////////////////////////////////////////
// CProject1App クラスの初期化

BOOL CProject1App::InitInstance()
{
	// OLE ライブラリの初期化
	if (!AfxOleInit())
	{
		AfxMessageBox(IDP_OLE_INIT_FAILED);
		return FALSE;
	}

	AfxEnableControlContainer();

	// 標準的な初期化処理
	// もしこれらの機能を使用せず、実行ファイルのサイズを小さくしたけ
	//  れば以下の特定の初期化ルーチンの中から不必要なものを削除して
	//  ください。
#if _MSC_VER <= 1200
	#ifdef _AFXDLL
		Enable3dControls();			// 共有 DLL 内で MFC を使う場合はここをコールしてください。
	#else
		Enable3dControlsStatic();	// MFC と静的にリンクする場合はここをコールしてください。
	#endif
#endif
	// OLE サーバーとして起動されている時にはコマンドラインを解析します。
	if (RunEmbedded() || RunAutomated())
	{
		// 実行するとすべての OLE サーバー ファクトリ を登録します。
		//  他のアプリケーションからオブジェクトを作るために OLE ライブラリを使用可能にします。
		COleTemplateServer::RegisterAll();
	}
	else
	{
		// システム レジストリが壊れていてサーバー アプリケーションがスタンド アロンで
		// 起動された時には、システム レジストリを更新してください。
		COleObjectFactory::UpdateRegistryAll();
	}

	CProject1Dlg dlg;
	m_pMainWnd = &dlg;
	int nResponse = dlg.DoModal();
	if (nResponse == IDOK)
	{
		// TODO: ダイアログが <OK> で消された時のコードを
		//       記述してください。
	}
	else if (nResponse == IDCANCEL)
	{
		// TODO: ダイアログが <ｷｬﾝｾﾙ> で消された時のコードを
		//       記述してください。
	}

	// ダイアログが閉じられてからアプリケーションのメッセージ ポンプを開始するよりは、
	// アプリケーションを終了するために FALSE を返してください。
	return FALSE;
}
