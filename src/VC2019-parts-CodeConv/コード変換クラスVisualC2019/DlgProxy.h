// DlgProxy.h : ヘッダー ファイル
//

#if !defined(AFX_DLGPROXY_H__957DA897_FACB_4B75_B5A7_58B0DCBA72D6__INCLUDED_)
#define AFX_DLGPROXY_H__957DA897_FACB_4B75_B5A7_58B0DCBA72D6__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

class CProject1Dlg;

/////////////////////////////////////////////////////////////////////////////
// CProject1DlgAutoProxy コマンド ターゲット

class CProject1DlgAutoProxy : public CCmdTarget
{
	DECLARE_DYNCREATE(CProject1DlgAutoProxy)

	CProject1DlgAutoProxy();           // 動的生成で使用される protected コンストラクタ

// アトリビュート
public:
	CProject1Dlg* m_pDialog;

// オペレーション
public:

// オーバーライド
	// ClassWizard は仮想関数のオーバーライドを生成します
	//{{AFX_VIRTUAL(CProject1DlgAutoProxy)
	public:
	virtual void OnFinalRelease();
	//}}AFX_VIRTUAL

// インプリメンテーション
protected:
	virtual ~CProject1DlgAutoProxy();

	// 生成されたメッセージ マップ関数
	//{{AFX_MSG(CProject1DlgAutoProxy)
		// メモ - ClassWizard はこの位置にメンバ関数を追加または削除します。
	//}}AFX_MSG

	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(CProject1DlgAutoProxy)

	// 生成された OLE ディスパッチ マップ関数
	//{{AFX_DISPATCH(CProject1DlgAutoProxy)
		// メモ - ClassWizard はこの位置にメンバ関数を追加または削除します。
	//}}AFX_DISPATCH
	DECLARE_DISPATCH_MAP()
	DECLARE_INTERFACE_MAP()
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ は前行の直前に追加の宣言を挿入します。

#endif // !defined(AFX_DLGPROXY_H__957DA897_FACB_4B75_B5A7_58B0DCBA72D6__INCLUDED_)
