// Project1.h : PROJECT1 アプリケーションのメイン ヘッダー ファイルです。
//

#if !defined(AFX_PROJECT1_H__F0F7C725_BFAE_41D5_813D_2F8BF818829A__INCLUDED_)
#define AFX_PROJECT1_H__F0F7C725_BFAE_41D5_813D_2F8BF818829A__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#ifndef __AFXWIN_H__
	#error include 'stdafx.h' before including this file for PCH
#endif

#include "resource.h"		// メイン シンボル

/////////////////////////////////////////////////////////////////////////////
// CProject1App:
// このクラスの動作の定義に関しては Project1.cpp ファイルを参照してください。
//

class CProject1App : public CWinApp
{
public:
	CProject1App();

// オーバーライド
	// ClassWizard は仮想関数のオーバーライドを生成します。
	//{{AFX_VIRTUAL(CProject1App)
	public:
	virtual BOOL InitInstance();
	//}}AFX_VIRTUAL

// インプリメンテーション

	//{{AFX_MSG(CProject1App)
		// メモ - ClassWizard はこの位置にメンバ関数を追加または削除します。
		//        この位置に生成されるコードを編集しないでください。
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};


/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ は前行の直前に追加の宣言を挿入します。

#endif // !defined(AFX_PROJECT1_H__F0F7C725_BFAE_41D5_813D_2F8BF818829A__INCLUDED_)
