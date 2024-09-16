// DlgProxy.cpp : インプリメンテーション ファイル
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
	
	// オートメーション オブジェクトがアクティブである限り、アプリケーションを 
	// 実行状態にしてください、コンストラクタは AfxOleLockApp を呼び出します。
	AfxOleLockApp();

	// アプリケーションのメイン ウィンドウ ポインタを通してダイアログ
	// へアクセスします。プロキシの内部ポインタからダイアログへのポイ
	// ンタを設定し、ダイアログの戻りポインタをこのプロキシへ設定しま
	// す。
	ASSERT (AfxGetApp()->m_pMainWnd != NULL);
	ASSERT_VALID (AfxGetApp()->m_pMainWnd);
	ASSERT_KINDOF(CProject1Dlg, AfxGetApp()->m_pMainWnd);
	m_pDialog = (CProject1Dlg*) AfxGetApp()->m_pMainWnd;
	m_pDialog->m_pAutoProxy = this;
}

CProject1DlgAutoProxy::~CProject1DlgAutoProxy()
{
	// すべてのオブジェクトがオートメーションで作成された場合にアプリケーション
	// を終了するために、デストラクタが AfxOleUnlockApp を呼び出すます。
	// 他の処理の間に、メイン ダイアログを破壊します。
	if (m_pDialog != NULL)
		m_pDialog->m_pAutoProxy = NULL;
	AfxOleUnlockApp();
}

void CProject1DlgAutoProxy::OnFinalRelease()
{
	// オートメーション オブジェクトに対する最後の参照が解放される時に
	// OnFinalRelease が呼び出されます。基本クラスは自動的にオブジェク
	// トを削除します。基本クラスを呼び出す前に、オブジェクトで必要な特
	// 別な後処理を追加してください。

	CCmdTarget::OnFinalRelease();
}

BEGIN_MESSAGE_MAP(CProject1DlgAutoProxy, CCmdTarget)
	//{{AFX_MSG_MAP(CProject1DlgAutoProxy)
		// メモ - ClassWizard はこの位置にマッピング用のマクロを追加または削除します。
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

BEGIN_DISPATCH_MAP(CProject1DlgAutoProxy, CCmdTarget)
	//{{AFX_DISPATCH_MAP(CProject1DlgAutoProxy)
		// メモ - ClassWizard はこの位置にマッピング用のマクロを追加または削除します。
	//}}AFX_DISPATCH_MAP
END_DISPATCH_MAP()

// メモ: VBA からタイプ セーフなバインディングをサポートするために IID_IProject1
//  のサポートを追加します。この IID は .ODL ファイル内のディスパッチ インターフェイス 
//  へアタッチされる GUID と合致しなければなりません。

// {D32FF675-669B-4A8C-B480-DA7B832CBA00}
static const IID IID_IProject1 =
{ 0xd32ff675, 0x669b, 0x4a8c, { 0xb4, 0x80, 0xda, 0x7b, 0x83, 0x2c, 0xba, 0x0 } };

BEGIN_INTERFACE_MAP(CProject1DlgAutoProxy, CCmdTarget)
	INTERFACE_PART(CProject1DlgAutoProxy, IID_IProject1, Dispatch)
END_INTERFACE_MAP()

// IMPLEMENT_OLECREATE2 マクロはこのプロジェクトの StdAfx.h 内に定義されています。
// {4805CCB7-8A33-4029-ACB0-7EE3162679F8}
IMPLEMENT_OLECREATE2(CProject1DlgAutoProxy, "Project1.Application", 0x4805ccb7, 0x8a33, 0x4029, 0xac, 0xb0, 0x7e, 0xe3, 0x16, 0x26, 0x79, 0xf8)

/////////////////////////////////////////////////////////////////////////////
// CProject1DlgAutoProxy メッセージ ハンドラ
