VERSION 5.00
Object = "{2AB17740-0C41-11D7-916F-0003479BEB3F}#1.18#0"; "JVDTLab.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDBUpdate 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "データ更新"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'ｵｰﾅｰ ﾌｫｰﾑの中央
   Begin VB.Timer tmrRefresh 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   4440
      Top             =   1560
   End
   Begin VB.Timer tmrStartTrigger 
      Enabled         =   0   'False
      Left            =   4860
      Top             =   1560
   End
   Begin VB.Frame fraInfo 
      Height          =   1065
      Left            =   90
      TabIndex        =   7
      Top             =   1440
      Width           =   3735
      Begin VB.Label lblType 
         AutoSize        =   -1  'True
         Height          =   180
         Left            =   1470
         TabIndex        =   13
         Top             =   780
         Width           =   60
      End
      Begin VB.Label lblFix 
         AutoSize        =   -1  'True
         Caption         =   "取得中情報"
         Height          =   180
         Index           =   5
         Left            =   120
         TabIndex        =   12
         Top             =   780
         Width           =   900
      End
      Begin VB.Label lblFinish 
         AutoSize        =   -1  'True
         Height          =   180
         Left            =   1470
         TabIndex        =   11
         Top             =   510
         Width           =   60
      End
      Begin VB.Label lblPass 
         AutoSize        =   -1  'True
         Height          =   180
         Left            =   1470
         TabIndex        =   10
         Top             =   240
         Width           =   60
      End
      Begin VB.Label lblFix 
         AutoSize        =   -1  'True
         Caption         =   "経過時間"
         Height          =   180
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   720
      End
      Begin VB.Label lblFix 
         AutoSize        =   -1  'True
         Caption         =   "予想残り時間"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   510
         Width           =   1035
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "キャンセル"
      Height          =   285
      Left            =   3960
      TabIndex        =   1
      Top             =   2130
      Width           =   1365
   End
   Begin MSComctlLib.ProgressBar prgPart 
      Height          =   195
      Left            =   90
      TabIndex        =   0
      Top             =   1200
      Width           =   5205
      _ExtentX        =   9181
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.ProgressBar prgAll 
      Height          =   195
      Left            =   90
      TabIndex        =   2
      Top             =   690
      Width           =   5205
      _ExtentX        =   9181
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.ProgressBar prgDown 
      Height          =   195
      Left            =   90
      TabIndex        =   14
      Top             =   210
      Width           =   5205
      _ExtentX        =   9181
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Timer tmrOptimize 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   4020
      Top             =   1560
   End
   Begin VB.Label lblFix 
      AutoSize        =   -1  'True
      Caption         =   "ダウンロード進行状況"
      Height          =   180
      Index           =   6
      Left            =   90
      TabIndex        =   16
      Top             =   0
      Width           =   1650
   End
   Begin VB.Label lblPercentDown 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   5040
      TabIndex        =   15
      Top             =   0
      Width           =   60
   End
   Begin JVDTLabLibCtl.JVLink axJVLink 
      Left            =   4320
      OleObjectBlob   =   "frmDBUpdate.frx":0000
      Top             =   2100
   End
   Begin VB.Label lblPercentPart 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   5040
      TabIndex        =   6
      Top             =   990
      Width           =   60
   End
   Begin VB.Label lblPercentAll 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   5040
      TabIndex        =   5
      Top             =   480
      Width           =   60
   End
   Begin VB.Label lblFix 
      AutoSize        =   -1  'True
      Caption         =   "読み込み個別進行状況"
      Height          =   180
      Index           =   4
      Left            =   90
      TabIndex        =   4
      Top             =   990
      Width           =   1800
   End
   Begin VB.Label lblFix 
      AutoSize        =   -1  'True
      Caption         =   "読み込み進行状況"
      Height          =   180
      Index           =   3
      Left            =   90
      TabIndex        =   3
      Top             =   480
      Width           =   1440
   End
End
Attribute VB_Name = "frmDBUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'   データ更新処理
'
'   Other, SLOP, BLOD ,O6H6の四種類のモードがある。
'
Option Explicit

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部変数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private mlngGettingMode As Long                  '' BLOD SLOP O6H6単独モード

Private mblnFinishFlag  As Boolean              '' 終了フラグ 正常終了時に真となる
Private mblnCancelFlag  As Boolean              '' 中断フラグ 中断を取得ループに伝える
Private mblnAfterJVOpen As Boolean              '' JVOpen正常終了後に真（再セットアップキャンセル不可）

' 最適化処理関連
Private WithEvents mAsyncCN As ADODB.Connection '' 最適化用非同期コネクション
Attribute mAsyncCN.VB_VarHelpID = -1
Private mblnOptimizeFinish  As Boolean
Private mstrJobList()       As String
Private mlngJobCount        As Long
Private mstrTargetMDB       As String
Private mstrWorkMDB         As String
Private mstrTableName       As String

' 画面表示ステータス用
Private mstrFree        As String
Private mstrType        As String
Private mlngPercentPart As Long                 '' パーセント（個別）
Private mvarStartTime   As Variant              '' 読み込み開始日時

Private mlngReadCount       As Long             '' 読み込むファイル数
Private mlngFileCount       As Long             '' 処理ファイル数
Private mlngDownloadCount   As Long             '' ダウンロードが必要なファイル数
Private mdblReadedByte      As Double           '' 全体の読み込み済みバイト数
Private mlngReadedBytePart  As Long             '' 個別の読み込み済みバイト数
Private mlngCountParSpec    As Long             '' レコード種別毎カウンタ

Private mAV As clsImportAV
Private mBN As clsImportBN
Private mBR As clsImportBR
Private mCH As clsImportCH
Private mDM As clsImportDM
Private mHC As clsImportHC
Private mHN As clsImportHN
Private mHR As clsImportHR
Private mJC As clsImportJC
Private mKS As clsImportKS
Private mO1 As clsImportO1
Private mO2 As clsImportO2
Private mO3 As clsImportO3
Private mO4 As clsImportO4
Private mO5 As clsImportO5
Private mRA As clsImportRA
Private mRC As clsImportRC
Private mSE As clsImportSE
Private mSK As clsImportSK
Private mTK As clsImportTK
Private mUM As clsImportUM
Private mWE As clsImportWE
Private mWH As clsImportWH
Private mYS As clsImportYS
Private mODDS As clsImportODDS
Private mTC As clsImportTC
Private mCC As clsImportCC

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   プロパティ
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: データ更新モードを設定する
'
'   備考: なし
'
Public Property Let GettingMode(RHS As Long)
    Select Case RHS
    Case 0
        gApp.Log "Set DBUpdateMode: Other"
    Case 1
        gApp.Log "Set DBUpdateMode: SLOP"
    Case 2
        gApp.Log "Set DBUpdateMode: BLOD"
    Case 3
        gApp.Log "Set DBUpdateMode: O6H6"
    Case Else
        gApp.Log "Set DBUpdateMode: Error"
    End Select
    mlngGettingMode = RHS
End Property


'
'   機能: データ取得結果フラグを返す
'
'   備考: 正常終了 True, 異常終了 False
'
Public Property Get Finish() As Boolean
    Finish = mblnFinishFlag
End Property


'
'   機能: JVOpen正常終了後に真
'
'   備考: なし
'
Public Property Get AfterJVOpen() As Boolean
    AfterJVOpen = mblnAfterJVOpen
End Property


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   外部関数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: 取得メインループ
'
'   備考: なし
'
Public Function FetchJVData() As Boolean
On Error GoTo Errorhandler
    
    Const lngBuffSize         As Long = "102901"   '' JVRead用バッファ
    Const lngFileNameSize     As Long = "256"     '' ファイル名バッファのサイズ
    Const strDataSpecSetup    As String = "TOKURACEDIFFYSCH"            '' セットアップ
    Const strDataSpecUsual    As String = "TOKURACEDIFFYSCH"            '' 通常取得
    Const strDataSpecThisWeek As String = "TOKURACETCOVRCOV"            '' 今週モード
    
    Dim lngReturnCode       As Long                  '' JVLinkからの戻り値
    Dim lngPrevReturnCode   As Long                  '' 前回のJVLinkからの戻り値
    Dim strDataSpec         As String                '' JVOpen データ種別
    Dim strFromTime         As String                '' 取得開始日時
    Dim lngOptionFlag       As Long                  '' JVLink取得オプション

    Dim strLastTime         As String                '' JVOpenが返す取得時刻
    Dim strFileName         As String                '' JVGetsが返すファイル名用バッファ
    Dim bytBuff()           As Byte                  '' JVGets用バッファ
    Dim ImportObj           As clsIImport            '' Import Object Interface
    Dim strRecordIDOld      As String                '' レコード種別ID
    Dim strRecordIDNew      As String                '' レコード種別IDを保存する
    Dim lngCount            As Long                  '' 全体カウンタ
    Dim sngTimerStart       As Single                '' 全体タイマー
    Dim sngTimerEnd         As Single                '' 全体タイマー終了
    Dim sngSubTimerStart    As Single                '' レコード種別毎タイマー
    Dim sngSubTimerEnd      As Single                '' レコード種別毎タイマー終了
    Dim lngRecLen           As Long                  '' レコード長
    Dim strRegSetupCancelLastTime As String
    Dim strCurrentTimeStamp As String                '' 現在のファイルのタイムスタンプ

    Dim objMP As clsPointer '' マウスポインタ制御クラス
    
    Dim objDBRAKaiSel As frmDBRAKaiSel

    Set objMP = New clsPointer
    Call objMP.SetBusyPointer(Me)


    
    '-------
    ' JVInit
    
    lngReturnCode = axJVLink.JVInit(gJVLinkSID)
    If lngReturnCode <> 0 Then
        gApp.Log "JVLink - JVInitエラー"
        MsgBox "JVLink - JVInitエラー", vbExclamation, "馬吉：JVLinkエラー"
        ' エラー終了
        FetchJVData = False
        Exit Function
    End If
    
        
    '--------
    ' JVOpen
    
    ' 一度も取得されていない場合はセットアップを行う
    ' そうでなければ、通常データ取得をする
    If gApp.R_JVLMode = ukjThisWeek Then
        lngOptionFlag = 2 ' 今週データモード
        strDataSpec = strDataSpecThisWeek
        strFromTime = gApp.R_JVDLastTimeThisWeek ' 最終取得時間を今回の取得開始時間とする
        Me.Caption = "今週モードデータ取得"
    Else
        Select Case mlngGettingMode
        Case 0 ' Otherモード
            If val(Left$(gApp.R_JVDLastTime, 2)) = 0 Then
                lngOptionFlag = 3  'セットアップ用データモード
                strDataSpec = strDataSpecSetup
                Me.Caption = "データセットアップ"
                If gApp.R_SetupCancelLastTime <> "" Then
                    MsgBox "セットアップを再開します。" & vbCrLf & strFromTime, vbInformation, "馬吉：セットアップ再開"
                    Me.Caption = "データセットアップ再開"
                End If
                strFromTime = Format$(gApp.R_SetupYear, "0000") & String$(10, "0")
            Else
                lngOptionFlag = 1 ' 通常データモード
                strDataSpec = strDataSpecUsual
                strFromTime = gApp.R_JVDLastTime ' 最終取得時間を今回の取得開始時間とする
                Me.Caption = "更新データ取得"
            End If
            strDataSpec = strDataSpec & _
                            IIf(gApp.R_JVLGetSLOP, "SLOP", "") & _
                            IIf(gApp.R_JVLGetBLOD, "BLOD", "")
        Case 1 ' SLOP 単独モード
            If val(Left$(gApp.R_JVDLastTimeSLOP, 2)) = 0 Then
                lngOptionFlag = 3  'セットアップ用データモード
                strDataSpec = "SLOP"
                strFromTime = Format$(gApp.R_SetupYear, "0000") & String$(10, "0")
                Me.Caption = "坂路調教データセットアップ"
            Else
                lngOptionFlag = 1 ' 通常データモード
                strDataSpec = "SLOP"
                strFromTime = gApp.R_JVDLastTimeSLOP ' 最終取得時間を今回の取得開始時間とする
                Me.Caption = "坂路調教データ更新"
            End If
        Case 2 ' BLOD 単独モード
            If val(Left$(gApp.R_JVDLastTimeBLOD, 2)) = 0 Then
                lngOptionFlag = 3  'セットアップ用データモード
                strDataSpec = "BLOD"
                strFromTime = Format$(gApp.R_SetupYear, "0000") & String$(10, "0")
                Me.Caption = "血統・繁殖馬データセットアップ"
            Else
                lngOptionFlag = 1 ' 通常データモード
                strDataSpec = "BLOD"
                strFromTime = gApp.R_JVDLastTimeBLOD ' 最終取得時間を今回の取得開始時間とする
                Me.Caption = "血統・繁殖馬データ更新"
            End If
        End Select
    End If
    
    
    
    lngReturnCode = axJVLink.JVOpen(strDataSpec, _
                                    strFromTime, _
                                    lngOptionFlag, _
                                    mlngReadCount, _
                                    mlngDownloadCount, _
                                    strLastTime)
    
    Select Case lngReturnCode
        Case 0
            ' 続行
        Case -1
            MsgBox "データは最新です。", vbInformation, "馬吉：データ最新"
            FetchJVData = True  ' 成功で終了
            axJVLink.JVClose
            Exit Function
        Case -504 ' メンテナンス中
            ' メンテナンス中のメッセージはJVLinkが表示するので
            ' 馬吉側では何も表示しないで終了する。
            FetchJVData = False ' 失敗で終了
            axJVLink.JVClose
            Exit Function
        Case Else
            MsgBox "JV-Linkへ接続失敗しました。" & vbCrLf _
                & "JV-Linkからのエラーメッセージ: " & vbCrLf _
                & ErrMsgJVOpen(lngReturnCode), vbInformation, "馬吉：JV-Linkエラー"
            FetchJVData = False ' 失敗で終了
            axJVLink.JVClose
            Exit Function
    End Select
    
    ' ダウンロード進行状況プログレスバーにファイル数を設定する
    With prgDown
        .Min = 0
        If mlngDownloadCount > 0 Then '新規ダウンロードがあるなら
            .max = mlngDownloadCount
        End If
    End With
    
    '-------------------
    ' JVReadのループ処理
    FetchJVData = True
    
    If mlngReadCount > 0 Then
        
        If axJVLink.m_TotalReadFilesize > 0 Then
            prgAll.max = axJVLink.m_TotalReadFilesize
        Else
            MsgBox "m_TotalReadFilesizeが0以下です。"
            Exit Function
        End If
        
        ' ステータス描画開始
        tmrRefresh.Enabled = True

        
        mlngFileCount = 0
        mvarStartTime = Now
        
        ' 全体カウンタ初期化
        lngCount = 0
        
        ' 処理時間計測用　開始時間設定
        sngTimerStart = Timer
        
        'バッファ作成
        strFileName = String$(lngFileNameSize, vbNullChar)
        
        strRegSetupCancelLastTime = gApp.R_SetupCancelLastTime
        
        Do
            DoEvents    ' バックグラウンド処理
            
            If mblnCancelFlag Then
                ' OtherModeのセットアップ時なら、中断再開用タイムスタンプを記録する
                If mlngGettingMode = 0 And gApp.R_JVLMode = ukjUsual And lngOptionFlag = 3 Then
                    gApp.R_SetupCancelLastTime = strCurrentTimeStamp
                End If
                mblnFinishFlag = False ' 異常終了
                Exit Do ' 取得中止ボタンが押されたら終了
            End If
            
            ' JVGetsで1行読み込み
            lngReturnCode = axJVLink.JVGets(bytBuff, lngBuffSize, strFileName)
            
            'リターンコードにより処理を分岐
            Select Case lngReturnCode
            Case 0      ' 全ファイル読み込み終了
                ' キャンセル状態でなくする
                If lngOptionFlag = 3 Then
                    gApp.R_SetupCancelLastTime = ""
                End If
                
                ' 取得日時を保存する｡
                If mlngGettingMode = 0 Then
                    If gApp.R_JVLMode = ukjUsual Then
                        gApp.R_JVDLastTime = strLastTime
                        If gApp.R_JVLGetSLOP = True Then
                            gApp.R_JVDLastTimeSLOP = strLastTime
                        End If
                        If gApp.R_JVLGetBLOD = True Then
                            gApp.R_JVDLastTimeBLOD = strLastTime
                        End If
                    ElseIf gApp.R_JVLMode = ukjThisWeek Then
                        gApp.R_JVDLastTimeThisWeek = strLastTime
                    End If
                ElseIf mlngGettingMode = 1 Then ' SLOP
                    gApp.R_JVDLastTimeSLOP = strLastTime
                    gApp.R_SetupCancelLastTime = "" ' SLOP モードは中断をサポートしない
                ElseIf mlngGettingMode = 2 Then ' BLOD
                    gApp.R_JVDLastTimeBLOD = strLastTime
                    gApp.R_SetupCancelLastTime = "" ' BLOD モードは中断をサポートしない
                End If
                ' ループから脱出する
                mblnFinishFlag = True ' 正常終了
                Exit Do
            Case -1     ' ファイル切り替わり
                mlngReadedBytePart = prgPart.max
                Call tmrRefresh_Timer
                
                gApp.Log "ファイル切り替わり" & strFileName
                
                mlngFileCount = mlngFileCount + 1
                mlngReadedBytePart = 0
                prgPart.max = axJVLink.m_CurrentReadFilesize
                
                If Not ImportObj Is Nothing Then
                    'インポートオブジェクトに終了処理をさせる
                    gApp.Log ">Close DB"
                    Call ImportObj.CloseDB
                    gApp.Log "<Close DB"
                    ' DB最適化が必要なとき最適化を行う
                    Call DBOptimize(strRecordIDOld)
                    'インポートオブジェクトに開始処理をさせる
                    gApp.Log ">Open DB"
                    Call ImportObj.OpenDB
                    gApp.Log "<Open DB"
                End If
                
                        
                       
            Case -3     ' ダウンロード中
                If lngReturnCode <> lngPrevReturnCode Then gApp.Log "ダウンロード中"
                mstrType = "ダウンロード待機"
            Case -201   ' Initされてない
                   MsgBox "JVInitが行われていません。", vbExclamation, "馬吉：エラー"
                Exit Do
            Case -203   ' Openされてない
                MsgBox "JVOpenが行われていません。", vbExclamation, "馬吉：エラー"
                Exit Do
            Case -402, -403   ' ダウンロードしたファイルが異常
                gApp.Log lngReturnCode & "ダウンロードしたファイルが異常" & strFileName
                Do While axJVLink.JVFiledelete(strFileName) <> 0
                    Select Case MsgBox("ダウンロードしたファイルが異常な為、JVFileDeleteを試みましたが失敗しました。", vbCritical + vbAbortRetryIgnore)
                    Case VbMsgBoxResult.vbAbort     ' 中止
                        axJVLink.JVClose
                        Exit Function
                    Case VbMsgBoxResult.vbRetry     ' 再施行
                        gApp.Log "JVFiledeleteの再施行"
                    Case VbMsgBoxResult.vbIgnore    ' 無視
                        Exit Do
                    End Select
                Loop
                gApp.Log "JVFileDelete " & strFileName
                axJVLink.JVClose
                If mlngGettingMode = 0 And val(Left$(gApp.R_JVDLastTime, 2)) = 0 Then
                    lngOptionFlag = 4 ' セットアップの場合、ダイアログを出さないセットアップモードにする
                End If
                Select Case axJVLink.JVOpen(strDataSpec, _
                                            strFromTime, _
                                            lngOptionFlag, _
                                            mlngReadCount, _
                                            mlngDownloadCount, _
                                            strLastTime)
                Case 0
                    ' 続行
                Case -1
                    MsgBox "データは最新です。", vbInformation, "馬吉：データ最新"
                    FetchJVData = True  ' 成功で終了
                    axJVLink.JVClose
                    Exit Function
                Case Else
                    MsgBox "JV-Linkへ接続失敗しました。" & vbCrLf _
                        & "JV-Linkからのエラーメッセージ: " & vbCrLf _
                        & ErrMsgJVOpen(lngReturnCode), vbInformation, "馬吉：JV-Linkエラー"
                    FetchJVData = False ' 失敗で終了
                    axJVLink.JVClose
                    Exit Function
                End Select
                
                ' ダウンロード進行状況プログレスバーにファイル数を設定する
                With prgDown
                    .Min = 0
                    If mlngDownloadCount > 0 Then '新規ダウンロードがあるなら
                        .max = mlngDownloadCount
                    End If
                End With
                
                ' 中断再開用タイムスタンプを記録する
                gApp.R_SetupCancelLastTime = strCurrentTimeStamp
                strRegSetupCancelLastTime = gApp.R_SetupCancelLastTime
                
                mlngFileCount = 0
                mvarStartTime = Now
                
                ' 全体カウンタ初期化
                lngCount = 0
                
                ' 処理時間計測用　開始時間設定
                sngTimerStart = Timer

            Case -502   ' ダウンロード失敗
                MsgBox "ダウンロード中にエラーが発生しました。", vbExclamation, "馬吉：エラー"
                Exit Do
            Case -503   ' ファイルがない
                MsgBox "ファイルがありません。", vbExclamation, "馬吉：エラー"
                Exit Do
            Case Is > 0 ' 正常読み込み
            

                prgPart.max = axJVLink.m_CurrentReadFilesize
                ' 再開処理か、通常取得か
                strCurrentTimeStamp = axJVLink.m_CurrentFileTimeStamp
                If strRegSetupCancelLastTime > strCurrentTimeStamp Then
                    ' キャンセルタイムスタンプより古いものはスキップする
                    mdblReadedByte = mdblReadedByte + axJVLink.m_CurrentReadFilesize
                    mlngReadedBytePart = 0
                    mlngFileCount = mlngFileCount + 1
                    mstrType = "再開処理:" & strCurrentTimeStamp
                    gApp.Log "JVSkip : " & strFileName
                    Call tmrRefresh_Timer
                    axJVLink.JVSkip
                Else
                    'レコード種別IDを取得
                    strRecordIDNew = StrConv(LeftB(bytBuff, 2), vbUnicode)
                    mstrType = StrConv(LeftB(bytBuff, 20), vbUnicode)
                    
                        'レコード種別IDが変更された場合（または初処理時）
                        If strRecordIDNew <> strRecordIDOld Then
                        
                            prgPart.max = axJVLink.m_CurrentReadFilesize
        
                            mstrType = strRecordIDNew
                            
                            ' 測定結果を表示
                            If mlngCountParSpec <> 0 Then
                                sngSubTimerEnd = Timer
                                gApp.Log "レコード種別が " & strRecordIDOld & " から " & strRecordIDNew & "に変わりました"
                                gApp.Log vbTab & strRecordIDOld & "処理行数:" & CStr(mlngCountParSpec)
                                gApp.Log vbTab & strRecordIDOld & "読み込み時間:" & CStr(sngSubTimerEnd - sngSubTimerStart) & "秒"
                                If (sngSubTimerEnd - sngSubTimerStart) > 0 Then
                                    gApp.Log vbTab & strRecordIDOld & "処理速度:" & CStr(mlngCountParSpec / (sngSubTimerEnd - sngSubTimerStart)) & "行/秒"
                                End If
                            End If
                            
                            If Not ImportObj Is Nothing Then
                                'インポートオブジェクトに終了処理をさせる
                                gApp.Log "Close DB"
                                Call ImportObj.CloseDB
                            End If
                            
                            'インポートオブジェクトを変更する
                            Set ImportObj = SelectImportObj(strRecordIDNew)

                            If ImportObj Is Nothing Then
                                ' 未知のレコード種別の場合DBをOpenしない
                                mdblReadedByte = mdblReadedByte + axJVLink.m_CurrentReadFilesize
                                mlngReadedBytePart = 0
                                mlngFileCount = mlngFileCount + 1
                                Call tmrRefresh_Timer
                                gApp.Log strRecordIDNew
                                axJVLink.JVSkip
                            Else
                                gApp.Log "Open DB"
                                Call ImportObj.OpenDB
                            End If

                            'レコード種別IDの保存
                            strRecordIDOld = strRecordIDNew
                            
                            '開始時刻を保存
                            sngSubTimerStart = Timer
        
                            '種別毎カウンタ初期化
                            mlngCountParSpec = 0
                        
                        End If
                        
                        If Not ImportObj Is Nothing Then
                            
                                'DB追加処理
                                Do
                                    If Not ImportObj.Add(bytBuff) Then
                                        gApp.Log "レコード登録の失敗" & StrConv(bytBuff, vbUnicode)
                                        Select Case MsgBox("レコードの登録に失敗しました。" & vbCrLf & StrConv(bytBuff, vbUnicode), vbAbortRetryIgnore + vbQuestion, "馬吉：エラー")
                                        Case vbAbort
                                            gApp.Log "中止"
                                            mblnCancelFlag = True
                                            Exit Do
                                        Case vbRetry
                                            gApp.Log "再試行"
                
                                        Case vbIgnore
                                            gApp.Log "無視"
                                            Exit Do
                                        End Select
                                    Else
                                        Exit Do
                                    End If
                                Loop
                            
                            'カウントアップ
                            mlngCountParSpec = mlngCountParSpec + 1
                            lngCount = lngCount + 1
                            '読み込んだバイト数合計値
                            mdblReadedByte = mdblReadedByte + lngReturnCode - 1
                            mlngReadedBytePart = mlngReadedBytePart + lngReturnCode - 1
                        End If
    
                End If ' 再開処理か通常取得か
                
                ' バッファの開放
                Erase bytBuff
            
            Case Else
                gApp.Log "不明なリターンコード" & lngReturnCode
            End Select
            
            'リターンコードを保存
            lngPrevReturnCode = lngReturnCode
            

        Loop
        
        sngSubTimerEnd = Timer
        sngTimerEnd = Timer
        gApp.Log "レコード種別 " & strRecordIDOld & "の読み込み処理を終了しました"
        gApp.Log vbTab & strRecordIDOld & "処理行数:" & CStr(mlngCountParSpec)
        gApp.Log vbTab & strRecordIDOld & "読み込み時間:" & CStr(sngSubTimerEnd - sngSubTimerStart) & "秒"
        If (sngSubTimerEnd - sngSubTimerStart) > 0 Then
            gApp.Log vbTab & strRecordIDOld & "処理速度:" & CStr(mlngCountParSpec / (sngSubTimerEnd - sngSubTimerStart)) & "行/秒"
        End If
        
        gApp.Log "すべてのデータ読み込み処理を終了しました"
        gApp.Log vbTab & "全体処理行数:" & CStr(lngCount)
        gApp.Log vbTab & "全体読み込み時間:" & CStr(sngTimerEnd - sngTimerStart) & "秒"
        If (sngTimerEnd - sngTimerStart) > 0 Then
            gApp.Log "全体処理速度:" & CStr(lngCount / (sngTimerEnd - sngTimerStart)) & "行/秒"
        End If
    End If
    
    ' ステータス描画終了
    mdblReadedByte = prgAll.max
    Call tmrRefresh_Timer
    tmrRefresh.Enabled = False
    
    '---------
    ' JVClose

    axJVLink.JVClose
    
    
    Set objMP = Nothing

    ' 開催選択画面用データ作成
    If mlngGettingMode = 0 Then
        Call gApp.DeleteAllRAKaiSelCacheFlags
        Set objDBRAKaiSel = New frmDBRAKaiSel
        objDBRAKaiSel.TargetYear = CStr(Year(Now))
        objDBRAKaiSel.Show vbModal, Me
    End If
    
    MsgBox "取得が終了しました｡", vbOKOnly + vbInformation, "馬吉：取得終了"

    Exit Function
Errorhandler:
    gApp.ErrLog
    axJVLink.JVClose
    FetchJVData = False
End Function


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部関数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: キャンセルボタンイベント
'
'   備考: なし
'
Private Sub cmdCancel_Click()
On Error GoTo Errorhandler
    If (Not mblnFinishFlag) Then
        If MsgBox("データの取得を中止しますか？", vbYesNo + vbQuestion, "馬吉：データ取得中止の確認") = vbYes Then
            gApp.Log "データ取得のキャンセル"
            mblnCancelFlag = True
        End If
    End If
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   機能: フォームロードイベント
'
'   備考: なし
'
Private Sub Form_Load()
On Error GoTo Errorhandler
    Me.Icon = LoadResPicture(100, vbResIcon)

    mblnFinishFlag = False
    
    prgDown.max = 1
    prgAll.max = 1
    prgPart.max = 1
    lblPass.Caption = ""
    lblFinish.Caption = ""
    lblType.Caption = ""
    
    tmrStartTrigger.Interval = 1
    tmrStartTrigger.Enabled = True

    Set mAV = New clsImportAV
    Set mBN = New clsImportBN
    Set mBR = New clsImportBR
    Set mCH = New clsImportCH
    Set mDM = New clsImportDM
    'Set mH1 = New clsImportH1
    Set mHC = New clsImportHC
    Set mHN = New clsImportHN
    Set mHR = New clsImportHR
    Set mJC = New clsImportJC
    Set mKS = New clsImportKS
    Set mO1 = New clsImportO1
    Set mO2 = New clsImportO2
    Set mO3 = New clsImportO3
    Set mO4 = New clsImportO4
    Set mO5 = New clsImportO5
    Set mRA = New clsImportRA
    Set mRC = New clsImportRC
    Set mSE = New clsImportSE
    Set mSK = New clsImportSK
    Set mTK = New clsImportTK
    Set mUM = New clsImportUM
    Set mWE = New clsImportWE
    Set mWH = New clsImportWH
    Set mYS = New clsImportYS
    Set mODDS = New clsImportODDS
    Set mTC = New clsImportTC
    Set mCC = New clsImportCC
            
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   機能: フォームアンロード確認イベント
'
'   備考: なし
'
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo Errorhandler
    If UnloadMode <> vbFormCode Then
        If MsgBox("データの取得を中止しますか？", vbYesNo + vbQuestion, "馬吉：データ取得中止の確認") = vbYes Then
            gApp.Log "データ取得のキャンセル"
            Cancel = True
            mblnCancelFlag = True
        Else
            Cancel = True
        End If
    End If
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   機能: リフレッシュタイマーイベント
'
'   備考: なし
'
Private Sub tmrRefresh_Timer()
On Error GoTo Errorhandler
    Dim sngPersentAll   As Single  ''
    Dim varPass         As Variant '' 経過時間
    Dim lngJVStatus     As Long
    
    
    ' 経過時間
    varPass = Now - mvarStartTime
    
    ' 進行度合い
    If prgAll.max <> 0 Then
        sngPersentAll = mdblReadedByte / prgAll.max
    End If
    
    ' ダウンロード状況
    lngJVStatus = axJVLink.JVStatus
    If lngJVStatus = -502 Then
        Exit Sub
    ElseIf lngJVStatus > 0 Then
        lblPercentDown.Caption = axJVLink.JVStatus & " / " & mlngDownloadCount
        lblPercentDown.Left = ScaleWidth - lblPercentDown.width - 120
        If prgDown.value < mlngDownloadCount Then
            prgDown.value = axJVLink.JVStatus
        End If
    End If
    

    ' 全体読み込み進行度
    If axJVLink.m_TotalReadFilesize > 0 Then
        lblPercentAll.Caption = Format$(mdblReadedByte / 1024 / prgAll.max * 100, "##0.0") & " %  " & mlngFileCount & "/" & mlngReadCount
        lblPercentAll.Left = ScaleWidth - lblPercentAll.width - 120
        prgAll.value = Smaller(CLng(mdblReadedByte / 1024), prgAll.max)
    Else
        MsgBox "m_TotalReadFilesizeが0以下です。"
    End If
    
    ' 個別読み込み進行度
    
    lblPercentPart.Caption = mlngCountParSpec & " Rec  " & Format$(mlngReadedBytePart / prgPart.max * 100, "##0.0") & " %  "
    If prgPart.max >= mlngReadedBytePart Then
        prgPart.value = Smaller(mlngReadedBytePart, prgPart.max)
    Else
        lblPercentPart.Caption = "99.9 %"
        gApp.Log prgPart.value & "," & mlngReadedBytePart
    End If
    lblPercentPart.Left = ScaleWidth - lblPercentPart.width - 120
    
    '経過時間
    lblPass.Caption = Format$(varPass, "Long Time")
    '残り時間
    If mlngFileCount > 0 Then
        lblFinish.Caption = Format$((varPass * mlngReadCount / mlngFileCount) - varPass, "Long Time")
    Else
        lblFinish.Caption = ""
    End If
    lblType.Caption = mstrType
    
    Me.Refresh
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   機能: スタートタイマーイベント
'
'   備考: なし
'
Private Sub tmrStartTrigger_Timer()
On Error GoTo Errorhandler
    gApp.Log "データ取得の開始"
    tmrStartTrigger.Enabled = False
    If Not FetchJVData() Then
        gApp.Log "データ取得の失敗"
    Else
        gApp.Log "データ取得の正常終了"
    End If

    Unload Me
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   機能: 開催年月を取得
'
'   備考: なし
'
Private Sub NoteRTRace()
    Dim cn As ADODB.Connection
    Dim RA As ADODB.Recordset
    Dim strOut As String
    Set RA = New ADODB.Recordset
    
    RA.Open "SELECT [Year]&[MonthDay] AS YMD FROM RACE WHERE [DataKubun] <= '6' GROUP BY [Year]&[MonthDay]", gApp.GetCN_RACE, adOpenForwardOnly, adLockReadOnly
    Do While Not RA.EOF
        strOut = strOut & IIf(strOut = "", "", ",") & RA("YMD")
        RA.MoveNext
    Loop
    RA.Close
    
    gApp.R_RTDates = strOut
End Sub


'
'   機能: インポートオブジェクト選択
'
'   備考: なし
'
Private Function SelectImportObj(strRecordID As String) As clsIImport
    Select Case strRecordID
    Case "AV"
        Set SelectImportObj = mAV
    Case "BN"
        Set SelectImportObj = mBN
    Case "BR"
        Set SelectImportObj = mBR
    Case "CH"
        Set SelectImportObj = mCH
    Case "DM"
        Set SelectImportObj = mDM
    Case "HC"
        Set SelectImportObj = mHC
    Case "HN"
        Set SelectImportObj = mHN
    Case "HR"
        Set SelectImportObj = mHR
    Case "JC"
        Set SelectImportObj = mJC
    Case "KS"
        Set SelectImportObj = mKS
    Case "O1"
        Set SelectImportObj = mO1
    Case "O2"
        Set SelectImportObj = mO2
    Case "O3"
        Set SelectImportObj = mO3
    Case "O4"
        Set SelectImportObj = mO4
    Case "O5"
        Set SelectImportObj = mO5
    Case "RA"
        Set SelectImportObj = mRA
    Case "RC"
        Set SelectImportObj = mRC
    Case "SE"
        Set SelectImportObj = mSE
    Case "SK"
        Set SelectImportObj = mSK
    Case "TK"
        Set SelectImportObj = mTK
    Case "UM"
        Set SelectImportObj = mUM
    Case "WE"
        Set SelectImportObj = mWE
    Case "WH"
        Set SelectImportObj = mWH
    Case "YS"
        Set SelectImportObj = mYS
    Case "TC"
        Set SelectImportObj = mTC
    Case "CC"
        Set SelectImportObj = mCC
    Case "H1", "H6", "O6"
        Set SelectImportObj = mODDS
    Case Else
        Set SelectImportObj = Nothing
    End Select
End Function


'
'   機能: ＤＢ最適化の要／不要を判定
'
'   備考: なし
'
Private Sub DBOptimize(strRecordID As String)
On Error GoTo Errorhandler
    Dim fso As FileSystemObject
    Dim i   As Long

    gApp.Log "最適化：確認"

    Set fso = New FileSystemObject
    
    ReDim mstrJobList(0)
    
    Select Case strRecordID
    Case "HC"
        mstrJobList = Split("subHANRO.mdb", ",")
    Case "O5"
        mstrJobList = Split("subODDS_SANREN0.mdb,subODDS_SANREN1.mdb,subODDS_SANREN2.mdb,subODDS_SANREN3.mdb,subODDS_SANREN4.mdb,subODDS_SANREN5.mdb,subODDS_SANREN6.mdb,subODDS_SANREN7.mdb,subODDS_SANREN8.mdb,subODDS_SANREN9.mdb", ",")
    Case "RA"
        mstrJobList = Split("subRACE.mdb", ",")
    Case "SE"
        mstrJobList = Split("subUMA_RACE_A.mdb,subUMA_RACE_B.mdb", ",")
    Case "UM"
        mstrJobList = Split("subUMA.mdb", ",")
    Case Else
        gApp.Log "最適化：不要"
        Exit Sub
    End Select
        
    mlngJobCount = 0
    tmrOptimize.Interval = 100
    tmrOptimize.Enabled = False
    mblnOptimizeFinish = False
    On Error Resume Next
    For i = 0 To UBound(mstrJobList)
        ' ジョブリスト内のファイルサイズがどれか1.0GB以上であれば、Optimizeタイマーが始動する。
        tmrOptimize.Enabled = tmrOptimize.Enabled Or _
            (fso.GetFile(gApp.R_DBPath & "\" & mstrJobList(i)).Size > CLng(1000) * 1000 * 1000)
    Next i
    On Error GoTo Errorhandler
    
    If tmrOptimize.Enabled = False Then
        gApp.Log "最適化：不要"
        Exit Sub
    End If
    
    ' 最適化終了までループ
    gApp.Log "最適化：ジョブループ開始"
    Do While Not mblnOptimizeFinish
        DoEvents

        ' タイマーが停止していた場合終了
        If tmrOptimize.Enabled = False Then
            Exit Do
        End If
    Loop
    tmrOptimize.Enabled = False
    gApp.Log "最適化：ジョブループ終了"
    
    If mlngJobCount > 0 Then
        gApp.Log "最適化：終了"
        mstrType = "最適化終了"
        Call tmrRefresh_Timer
    Else
        gApp.Log "最適化：不要"
    End If
    
    Set mAsyncCN = Nothing
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   機能: 最適化タイマーイベント
'
'   備考: なし
'
Private Sub tmrOptimize_Timer()
On Error GoTo Errorhandler
    
    Dim cat             As ADOX.Catalog
    Dim strSQL          As String
    Dim i               As Long
    
    If mlngJobCount > UBound(mstrJobList) Then
        ' ジョブカウンタが最大値を超えていたら終了する
        mblnOptimizeFinish = True
        gApp.Log "最適化：全ジョブ終了"
    ElseIf mAsyncCN Is Nothing Then
        ' 非同期コネクションオブジェクトがなければ生成する
        Set mAsyncCN = New ADODB.Connection
        gApp.Log "最適化：コネクションオブジェクト生成"
    ElseIf mAsyncCN.State = adStateClosed Then
        ' コネクションが閉じていれば、最適化を開始する
        Set cat = New ADOX.Catalog
        mstrTargetMDB = gApp.R_DBPath & "\" & mstrJobList(mlngJobCount)
        mstrWorkMDB = gApp.R_DBPath & "\" & "Optimizing" & Timer & mstrJobList(mlngJobCount)
        
        ' 作業MDB作成
        cat.Create "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & mstrWorkMDB
        
        ' コネクション接続
        mAsyncCN.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                        "Data Source=" & mstrTargetMDB
        
        ' テーブル名取得
        Set cat.ActiveConnection = mAsyncCN
        For i = 0 To cat.Tables.count - 1
            If cat.Tables(i).Type = "TABLE" Then
                mstrTableName = cat.Tables(i).Name
            End If
        Next i
        
        ' テーブルコピーSQL文
        strSQL = "SELECT " & mstrTableName & ".* INTO " & mstrTableName & _
                " IN '" & mstrWorkMDB & "'" & _
                " FROM " & mstrTableName
                
        mAsyncCN.Execute strSQL, , adAsyncExecute
        mstrType = "最適化中:" & mstrTargetMDB
        Call tmrRefresh_Timer
        gApp.Log "最適化：" & mstrTargetMDB
    ElseIf mblnCancelFlag = True Then
        ' キャンセルボタンが押された場合
        If mAsyncCN.State And adStateExecuting <> 0 Then
            ' 実行中ならキャンセルする
            mAsyncCN.Cancel
        End If
        mAsyncCN.Close
    End If
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   機能: 非同期コネクション実行完了イベント
'
'   備考: なし
'
Private Sub mAsyncCN_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
On Error GoTo Errorhandler
    Dim IndexCreator    As clsCreateMDB
    Dim fso             As FileSystemObject
    Dim MsgResult       As VbMsgBoxResult
    
    gApp.Log "最適化：mAsyncCN_ExecuteComplete " & mstrTargetMDB
    
    Set IndexCreator = New clsCreateMDB
    Set fso = New FileSystemObject
    
    ' コネクションを閉じる
    mAsyncCN.Close
    
    If Not pError Is Nothing Then
        MsgBox "最適化失敗"
    Else
        ' インデックス作成
        gApp.Log "最適化：インデックス作成開始"
        mstrType = "インデックス作成中:" & mstrWorkMDB
        Call tmrRefresh_Timer

        With IndexCreator
            Set .mConnection = New ADODB.Connection
            Call .mConnection.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & mstrWorkMDB)
            Call CallByName(IndexCreator, "createIndex_" & mstrTableName, VbMethod)
            Call .mConnection.Close
        End With
        gApp.Log "最適化：インデックス作成終了"
        

        ' オリジナルファイル削除＆リネーム
        Do
            On Error Resume Next
            fso.DeleteFile mstrTargetMDB
            If Err.Number = 0 Then
                ' 削除成功ならリネームして終了
                gApp.Log "最適化：リネーム成功"
                fso.MoveFile mstrWorkMDB, mstrTargetMDB
                Exit Do
            Else
                MsgResult = MsgBox("ファイルを削除できません。", vbAbortRetryIgnore + vbDefaultButton2 + vbExclamation, "馬吉：最適化中のエラー")
                If MsgResult = vbAbort Then
                    ' 中止なら、最適化と取得全部を終了する
                    gApp.Log "最適化：リネーム失敗　中止"
                    fso.DeleteFile mstrWorkMDB
                    mblnFinishFlag = True
                    mblnCancelFlag = True
                    Exit Do
                ElseIf MsgResult = vbIgnore Then
                    gApp.Log "最適化：リネーム失敗　無視"
                    ' 無視なら、次の最適化へ
                    fso.DeleteFile mstrWorkMDB
                    Exit Do
                End If
                gApp.Log "最適化：リネーム失敗　リトライ"
            End If
            On Error GoTo Errorhandler
        Loop
    End If
    
    
    ' ジョブカウンタを進める
    mlngJobCount = mlngJobCount + 1
    Exit Sub
Errorhandler:
    gApp.ErrLog
    mlngJobCount = mlngJobCount + 1
End Sub

