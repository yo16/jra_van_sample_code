// Project1Dlg.cpp : インプリメンテーション ファイル
//

#include "stdafx.h"
#include "Project1.h"
#include "Project1Dlg.h"
#include "DlgProxy.h"
#include "CodeCnv.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// アプリケーションのバージョン情報で使われている CAboutDlg ダイアログ

class CAboutDlg : public CDialog
{
public:
	CAboutDlg();

// ダイアログ データ
	//{{AFX_DATA(CAboutDlg)
	enum { IDD = IDD_ABOUTBOX };
	//}}AFX_DATA

	// ClassWizard は仮想関数のオーバーライドを生成します
	//{{AFX_VIRTUAL(CAboutDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV のサポート
	//}}AFX_VIRTUAL

// インプリメンテーション
protected:
	//{{AFX_MSG(CAboutDlg)
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialog(CAboutDlg::IDD)
{
	//{{AFX_DATA_INIT(CAboutDlg)
	//}}AFX_DATA_INIT
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CAboutDlg)
	//}}AFX_DATA_MAP
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialog)
	//{{AFX_MSG_MAP(CAboutDlg)
		// メッセージ ハンドラがありません。
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CProject1Dlg ダイアログ

IMPLEMENT_DYNAMIC(CProject1Dlg, CDialog);

CProject1Dlg::CProject1Dlg(CWnd* pParent /*=NULL*/)
	: CDialog(CProject1Dlg::IDD, pParent)
{
	//{{AFX_DATA_INIT(CProject1Dlg)
	m_eCode = _T("");
	m_eValue = _T("");
	m_eNum = _T("");
	m_eOut = _T("");
	//}}AFX_DATA_INIT
	// メモ: LoadIcon は Win32 の DestroyIcon のサブシーケンスを要求しません。
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
	m_pAutoProxy = NULL;
}

CProject1Dlg::~CProject1Dlg()
{
	// このダイアログ用のオートメーション プロキシがある場合は、このダイアログ
	// へのポインタを NULL に戻します、それによってダイアログが削除されたこと
	// がわかります。
	if (m_pAutoProxy != NULL)
		m_pAutoProxy->m_pDialog = NULL;
}

void CProject1Dlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CProject1Dlg)
	DDX_Text(pDX, IDC_EDIT1, m_eCode);
	DDX_Text(pDX, IDC_EDIT2, m_eValue);
	DDX_Text(pDX, IDC_EDIT3, m_eNum);
	DDX_Text(pDX, IDC_EDIT4, m_eOut);
	//}}AFX_DATA_MAP
}

BEGIN_MESSAGE_MAP(CProject1Dlg, CDialog)
	//{{AFX_MSG_MAP(CProject1Dlg)
	ON_WM_SYSCOMMAND()
	ON_WM_CLOSE()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CProject1Dlg メッセージ ハンドラ

BOOL CProject1Dlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	// "バージョン情報..." メニュー項目をシステム メニューへ追加します。

	// IDM_ABOUTBOX はコマンド メニューの範囲でなければなりません。
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		CString strAboutMenu;
		strAboutMenu.LoadString(IDS_ABOUTBOX);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// このダイアログ用のアイコンを設定します。フレームワークはアプリケーションのメイン
	// ウィンドウがダイアログでない時は自動的に設定しません。
	SetIcon(m_hIcon, TRUE);			// 大きいアイコンを設定
	SetIcon(m_hIcon, FALSE);		// 小さいアイコンを設定
	
	// TODO: 特別な初期化を行う時はこの場所に追加してください。
	
	return TRUE;  // TRUE を返すとコントロールに設定したフォーカスは失われません。
}

void CProject1Dlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialog::OnSysCommand(nID, lParam);
	}
}

// もしダイアログボックスに最小化ボタンを追加するならば、アイコンを描画する
// コードを以下に記述する必要があります。MFC アプリケーションは document/view
// モデルを使っているので、この処理はフレームワークにより自動的に処理されます。

void CProject1Dlg::OnPaint() 
{
	if (IsIconic())
	{
		CPaintDC dc(this); // 描画用のデバイス コンテキスト

		SendMessage(WM_ICONERASEBKGND, (WPARAM) dc.GetSafeHdc(), 0);

		// クライアントの矩形領域内の中央
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// アイコンを描画します。
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialog::OnPaint();
	}
}

// システムは、ユーザーが最小化ウィンドウをドラッグしている間、
// カーソルを表示するためにここを呼び出します。
HCURSOR CProject1Dlg::OnQueryDragIcon()
{
	return (HCURSOR) m_hIcon;
}

// コントローラがオブジェクトの 1 つををまだ保持している場合、
// オートメーションサーバーはユーザーが UI を閉じる際に終了で
// きません。これらのメッセージ ハンドラはプロキシがまだ使用中
// かどうかを確認し、それから UI が非表示になりますがダイアロ
// グはそれが消された場合その場所に残ります。

void CProject1Dlg::OnClose() 
{
	if (CanExit())
		CDialog::OnClose();
}

void CProject1Dlg::OnOK() 
{
	char currentpath[MAX_PATH];
	CString strCSVPath;

	//インスタンス生成
	CCodeCnv *objCodeCnv = new CCodeCnv;
	
	//コード表のパスを取得
	GetCurrentDirectory(MAX_PATH,currentpath);
	strCSVPath = currentpath;
	strCSVPath += "\\CodeTable.csv";

	objCodeCnv->FileName(strCSVPath);

	UpdateData(TRUE);//コントロール>>メンバ変数
	m_eOut=objCodeCnv->GetCodeName(m_eCode,m_eValue,atoi(m_eNum));
	UpdateData(FALSE);//メンバ変数>>コントロール

	//破棄
	delete objCodeCnv;

	
}

void CProject1Dlg::OnCancel() 
{
	if (CanExit())
		CDialog::OnCancel();
}

BOOL CProject1Dlg::CanExit()
{
	// プロキシ オブジェクトがまだ残っている場合、オートメーション	
	// コントローラはこのアプリケーションをまだ保持しています。 
	// ダイアログの周囲は残しますが UI は非表示になります。
	if (m_pAutoProxy != NULL)
	{
		ShowWindow(SW_HIDE);
		return FALSE;
	}

	return TRUE;
}
