/*=======================================================================
  JRA-VAN Data Lab.プログラミングパーツ「データ登録パーツ　ダイアログ」

	   作成: JRA-VAN ソフトウェア工房  2003年6月26日

========================================================================
   (C) Copyright JRA SYSTEM SERVICE CO.,LTD. 2007 All rights reserved
========================================================================*/


// DataImportDlg.cpp : 実装ファイル
//

#include "stdafx.h"
#include "DataImport.h"
#include "DataImportDlg.h"
#include "clsDBImport.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CDataImportDlg ダイアログ

	//キャンセルフラグ
	bool DialogCancel;
	//JVOpen:総読込みファイル数
    long ReadCount;                     
	//JVOpen:総ダウンロードファイル数
    long DownloadCount;  

	//JVOpen:タイムスタンプ
	CString strLastFile;
	BSTR bstrLastFile;


CDataImportDlg::CDataImportDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CDataImportDlg::IDD, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
	m_iRadio = 0;
}

void CDataImportDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_JVLINK1, m_jvlink1);
	DDX_Control(pDX, IDC_EDIT3, m_txtCount);
	DDX_Control(pDX, IDC_EDIT2, m_txtFromDate);
	DDX_Control(pDX, IDC_EDIT1, m_txtDataSpec);

}

BEGIN_MESSAGE_MAP(CDataImportDlg, CDialog)
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	//}}AFX_MSG_MAP
	ON_BN_CLICKED(IDC_BUTTON1, OnBnClickedButton1)
	ON_BN_CLICKED(IDC_BUTTON2, OnBnClickedButton2)
	ON_BN_CLICKED(IDOK, OnBnClickedOk)
	ON_EN_CHANGE(IDC_EDIT1, OnEnChangeEdit1)
	ON_BN_CLICKED(IDC_RADIOTUJO, OnBnClickedRadiotujo)
	ON_BN_CLICKED(IDC_RADIOKONSHU, OnBnClickedRadiokonshu)
	ON_BN_CLICKED(IDC_RADIOSETUP, OnBnClickedRadiosetup)
	ON_BN_CLICKED(IDC_RADIOREAL, OnBnClickedRadioreal)
END_MESSAGE_MAP()

// CDataImportDlg メッセージ ハンドラ

BOOL CDataImportDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	// このダイアログのアイコンを設定します。アプリケーションのメイン ウィンドウがダイアログでない場合、
	//  Framework は、この設定を自動的に行います。
	SetIcon(m_hIcon, TRUE);			// 大きいアイコンの設定
	SetIcon(m_hIcon, FALSE);		// 小さいアイコンの設定

	// TODO: 初期化をここに追加します。
		long ReturnCode;                //JVLink戻り値
        CString sid;
        sid = "UNKNOWN";               //引数 JVInit:ソフトウェアID

        //**********************
        //JVLink初期化
        //**********************
        //※※※ JVInitは JVLinkメソッド使用前（但し、JVSetUIProPertiesを除く）に呼出す
        ReturnCode = m_jvlink1.JVInit(sid);

		//文字列に変換
		CString strReturnCode;
		strReturnCode.Format("%d", ReturnCode);
	
	return TRUE;  // フォーカスをコントロールに設定した場合を除き、TRUE を返します。
}

// ダイアログに最小化ボタンを追加する場合、アイコンを描画するための
//  下のコードが必要です。ドキュメント/ビュー モデルを使う MFC アプリケーションの場合、
//  これは、Framework によって自動的に設定されます。

void CDataImportDlg::OnPaint() 
{
	if (IsIconic())
	{
		CPaintDC dc(this); // 描画のデバイス コンテキスト

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// クライアントの四角形領域内の中央
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// アイコンの描画
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialog::OnPaint();
	}
}

//ユーザーが最小化したウィンドウをドラッグしているときに表示するカーソルを取得するために、
//  システムがこの関数を呼び出します。
HCURSOR CDataImportDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

void CDataImportDlg::OnBnClickedButton1()
{
	// TODO : ここにコントロール通知ハンドラ コードを追加します。

		long ReturnCode;					//JVLink戻り値
		CString DataSpec;
		CString FromDate;
		int DataOption;
	
		//初期値設定
		DialogCancel=false;					//キャンセルフラグ初期化	
		
		m_jvlink1.JVInit("UNKNOWN");


		UpdateData(true);
		m_txtDataSpec.GetWindowText(DataSpec);
		m_txtFromDate.GetWindowText(FromDate);

		clsDBImport mDBCon;
		
		if(MessageBox("データをクリアしますか？",NULL,MB_YESNO)==IDYES){
			mDBCon.ClearData("");
		}


		DataOption = m_iRadio;

		
		//**********************
		//JVLinkダウンロード処理
		//**********************
		if(DataOption==100){
			ReturnCode = m_jvlink1.JVRTOpen((LPCTSTR)DataSpec,(LPCTSTR)FromDate);
			DownloadCount=0;
		}else{
			ReturnCode = m_jvlink1.JVOpen((LPCTSTR)DataSpec,
													(LPCTSTR)FromDate,
													DataOption,
													&ReadCount,
													&DownloadCount,
													&bstrLastFile);
		}	
		//エラー判定
		if (ReturnCode != 0) {		   //エラー
			//文字列に変換
			CString strReturnCode;
			strReturnCode.Format("%d", ReturnCode);
			//終了処理
			JVClosing();

		}else{							//正常

			//文字列に変換
			CString strReturnCode;
			CString strDownloadCount;
			CString strReadCount;
			strReturnCode.Format("%d", ReturnCode);
			strDownloadCount.Format("%d", DownloadCount);
			strReadCount.Format("%d", ReadCount);
			

			//初期設定
			SetWindowText("ダウンロード中・・・");
			while(m_jvlink1.JVStatus()!=DownloadCount || m_jvlink1.JVStatus()<0){
				Sleep(1000);
			}
			//読込み処理
			JVReading();
			//終了処理
			JVClosing();
			return;
		}
	
}

void CDataImportDlg::OnBnClickedButton2()
{
	// TODO : ここにコントロール通知ハンドラ コードを追加します。
	long ReturnCode;

	//**********************
    //JVLink設定画面表示
    //**********************
	ReturnCode=m_jvlink1.JVSetUIProperties();

	CString strReturnCode;
	strReturnCode.Format("%d", ReturnCode);
}

void CDataImportDlg::OnBnClickedOk()
{
	// TODO : ここにコントロール通知ハンドラ コードを追加します。
	OnOK();
}

void CDataImportDlg::OnEnChangeEdit1()
{
	// TODO :  これが RICHEDIT コントロールの場合、まず、CDialog::OnInitDialog() 関数をオーバーライドして、
	// OR 状態の ENM_CORRECTTEXT フラグをマスクに入れて、
	// CRichEditCtrl().SetEventMask() を呼び出さない限り、
	// コントロールは、この通知を送信しません。

	// TODO :  ここにコントロール通知ハンドラ コードを追加してください。
}
//------------------------------------------------------------------------------
//		終了処理
//------------------------------------------------------------------------------
void CDataImportDlg::JVClosing()
{
		long ReturnCode;		//JVLink戻り値

		KillTimer(1);
		::SysFreeString(bstrLastFile);

		//***************
		//JVLink終了処理
		//***************
		ReturnCode = m_jvlink1.JVClose();

		//文字列に変換
		CString strReturnCode;
		strReturnCode.Format("%d", ReturnCode);

}
//------------------------------------------------------------------------------
//		読込み処理
//------------------------------------------------------------------------------
void CDataImportDlg::JVReading()
{
		long	ReturnCode; 					//JVLink戻り値
		long	BuffSize;						//バッファサイズ
		int ReturnCodeDB;
		BuffSize = 110000;						//バッファサイズ指定

		CString sBuff;							//バッファ
		BSTR bBuff;
		CString sBuffName;						//バッファ名
		BSTR bBuffName;
		CString sLineCount;

		//バッファ領域確保
		sBuff.GetBufferSetLength(BuffSize);
		bBuff=sBuff.AllocSysString();
		sBuffName.GetBuffer(32);
		bBuffName=sBuffName.AllocSysString();
		
		int 	JVReadingCount; 				//読込みファイル数
		int		LineCount;
		LineCount=0;

		//文字列に変換
		CString strReadCount;
		strReadCount.Format("%d", ReadCount);


		clsDBImport m_DBCon;
		m_DBCon.BeginTrans();
		m_txtCount.SetDlgItemText(IDC_EDIT3,"");
		//初期値
		ReturnCode=0;
		JVReadingCount=1;
		SetWindowText("データ読込み中．．．(0/" + strReadCount + ")");
		CString strReturnCode;
		CString strJVReadingCount;
		do {

			PumpMessages();

				//キャンセルが押されたら処理を抜ける
				if (DialogCancel==true) return;

				//***************
				//JVLink読込み処理
				//***************


				ReturnCode =  m_jvlink1.JVRead(&bBuff,&BuffSize,&bBuffName);


				//文字列に変換
				strReturnCode.Format("%d", ReturnCode);
				strJVReadingCount.Format("%d", JVReadingCount);
				
				//エラー判定
				if (ReturnCode > 0){		   //正常終了

					SetWindowText("データ読込み中．．．(" + strJVReadingCount + "/" + strReadCount + ": " + bBuffName + ")");
					sBuff.GetBufferSetLength(ReturnCode);
					sBuff = bBuff;
					//クリア
					ReturnCodeDB=m_DBCon.SetData(sBuff,sBuff.GetLength());
					if(ReturnCodeDB < 0) break;
					sBuff.Empty();
				}else if (ReturnCode == -1){   //ファイルの切れ目
					//ファイル名表示
					sBuffName.GetBufferSetLength(32);
					sBuffName = bBuffName;
					LineCount = 0;
					//プログレスバー表示
					JVReadingCount++; //カウントアップ
					SetWindowText("データ読込み中．．．(" + strJVReadingCount + "/" + strReadCount + ")");
					//クリア
					sBuff.Empty();					
				}else if (ReturnCode == 0){    //全レコード読込み終了(EOF)
				strJVReadingCount.Format("%d", JVReadingCount-1);
					SetWindowText("データ読込み完了(" + strJVReadingCount + "/" + strReadCount + ")");
					//Repeatを抜ける
					break;
				}else if (ReturnCode < -3 ){	//読込みエラー
					//Repeatを抜ける
				}else if (ReturnCode < -1 ){	//読込みエラー
					//Repeatを抜ける
					break;
				}
				LineCount++;
				sLineCount.Format("%d",LineCount);
				SetDlgItemText(IDC_EDIT3,sLineCount);
		} while (1);
		if(ReturnCodeDB==0){
			m_DBCon.CommitTrans();
		}else if(ReturnCodeDB == -1){
			m_DBCon.RollbackTrans();
		}

		::SysFreeString(bBuff);
		::SysFreeString(bBuffName);
		sBuff.Empty();
		sBuffName.Empty();
}

//------------------------------------------------------------------------------
//		バックグラウンド処理
//------------------------------------------------------------------------------
void CDataImportDlg::PumpMessages()
{
		MSG msg;
		while (::PeekMessage(&msg, NULL, 0, 0, PM_REMOVE)) {
			   ::TranslateMessage(&msg);
			   ::DispatchMessage (&msg);
		}
}

void CDataImportDlg::OnBnClickedRadiotujo()
{
	// TODO : ここにコントロール通知ハンドラ コードを追加します。
	m_iRadio = 1;
}

void CDataImportDlg::OnBnClickedRadiokonshu()
{
	// TODO : ここにコントロール通知ハンドラ コードを追加します。
	m_iRadio = 2;
}

void CDataImportDlg::OnBnClickedRadiosetup()
{
	// TODO : ここにコントロール通知ハンドラ コードを追加します。
	m_iRadio = 3;
}

void CDataImportDlg::OnBnClickedRadioreal()
{
	// TODO : ここにコントロール通知ハンドラ コードを追加します。
	m_iRadio = 100;
}
