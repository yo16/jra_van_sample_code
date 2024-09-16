VERSION 5.00
Object = "{2AB17740-0C41-11D7-916F-0003479BEB3F}#1.18#0"; "JVDTLab.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDBPrompt 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "速報データ取得中"
   ClientHeight    =   1395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3825
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   3825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows の既定値
   Begin MSComctlLib.StatusBar stbBar 
      Align           =   2  '下揃え
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   1110
      Width           =   3825
      _ExtentX        =   6747
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6694
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmrRefresh 
      Interval        =   100
      Left            =   540
      Top             =   30
   End
   Begin VB.Timer tmrStartTrigger 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   30
      Top             =   30
   End
   Begin JVDTLabLibCtl.JVLink axJVLink 
      Left            =   2730
      OleObjectBlob   =   "frmDBPrompt.frx":0000
      Top             =   150
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "速報データを取得しています。しばらくお待ちください。"
      Height          =   405
      Left            =   780
      TabIndex        =   0
      Top             =   510
      Width           =   2400
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmDBPrompt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'   速報データ取得 ダイアログ
'
Option Explicit

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部変数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private mblnFinishFlag As Boolean     '' 終了フラグ 正常終了時に真
Private mblnCancelFlag As Boolean     '' 中断フラグ 中断を取得ループに伝える

Private mMode    As ukPromptMode    '' 速報取得モード
Private mstrKey  As String          '' JVDataキー
Private mstrType As String

Private mlngReadByte As Long

Private objMP As clsPointer '' マウスポインタ制御クラス


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   プロパティ
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: 速報取得モードを設定する
'
'   備考: なし
'
Public Property Let Mode(RHS As ukPromptMode)
    mMode = RHS
End Property

'
'   機能: JVLinkから取得するデータのキーを設定する
'
'   備考: なし
'
Public Property Let key(RHS As String)
    mstrKey = RHS
End Property


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   外部関数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8


'
'   機能: 取得メインループ
'
'   備考: なし
'
Public Function FetchJVData(DataSpec As String) As Boolean
On Error GoTo ErrorHandler

    Const lngBuffSize         As Long = "102901"   '' JVRead用バッファ
    Const lngFileNameSize     As Long = "256"     '' ファイル名バッファのサイズ

    Dim lngReturnCode       As Long                  '' JVLinkからの戻り値
    Dim lngPrevReturnCode   As Long                  '' 前回のJVLinkからの戻り値
    Dim strDataSpec         As String                '' JVOpen データ種別
    Dim strFromTime         As String                '' 取得開始日時
    Dim lngOptionFlag       As Long                  '' JVLink取得オプション

    Dim strLastTime         As String                '' JVOpenが返す取得時刻
    Dim strFileName         As String                '' JVReadが返すファイル名用バッファ
    Dim strBuff             As String                '' JVRead用バッファ
    Dim ImportObj           As clsIImport            '' Import Object Interface
    Dim strRecordIDOld      As String                '' レコード種別ID
    Dim strRecordIDNew      As String                '' レコード種別IDを保存する
    Dim lngFileCount        As Long                  '' 処理ファイル数
    Dim varStartTime        As Variant               '' 読み込み開始日時
    Dim lngCount            As Long                  '' 全体カウンタ
    Dim sngTimerStart       As Single                '' 全体タイマー
    Dim sngTimerEnd         As Single                '' 全体タイマー終了
    Dim lngSubCount         As Long                  '' レコード種別毎カウンタ
    Dim sngSubTimerStart    As Single                '' レコード種別毎タイマー
    Dim sngSubTimerEnd      As Single                '' レコード種別毎タイマー終了

    Me.Icon = LoadResPicture(100, vbResIcon)
    Set objMP = New clsPointer
    Call objMP.SetBusyPointer(Me)

    '-------
    ' JVInit

    lngReturnCode = axJVLink.JVInit(gJVLinkSID)
    If lngReturnCode <> 0 Then
        gApp.Log "JVLink - JVInitエラー"
        MsgBox "JVLink - JVInitエラー", vbExclamation, "馬吉：エラー"
        ' エラー終了
        FetchJVData = False
        Exit Function
    End If


    '--------
    ' JVRTOpen

    lngReturnCode = axJVLink.JVRTOpen(DataSpec, mstrKey)

    Select Case lngReturnCode
        Case 0
            ' 続行
        Case -1
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


    '-------------------
    ' JVReadのループ処理

    lngFileCount = 0
    varStartTime = Now

    ' 全体カウンタ初期化
    lngCount = 0

    ' 処理時間計測用　開始時間設定
    sngTimerStart = Timer
    
    mlngReadByte = 0

    Do
        DoEvents    ' バックグラウンド処理

        If mblnCancelFlag Then
            
            FetchJVData = False
            Exit Function ' 取得中止ボタンが押されたら終了
        End If

        
        'バッファ作成
        strBuff = String$(lngBuffSize, vbNullChar)
        strFileName = String$(lngFileNameSize, vbNullChar)

        lngReturnCode = axJVLink.JVRead(strBuff, lngBuffSize, strFileName)
        
        'リターンコードにより処理を分岐
        Select Case lngReturnCode
        Case 0      ' 全ファイル読み込み終了
            ' ループから脱出する
            Exit Do
        Case -1     ' ファイル切り替わり
            If lngReturnCode <> lngPrevReturnCode Then
                gApp.Log "ファイル切り替わり" & strFileName
            End If
            lngFileCount = lngFileCount + 1
        Case -3     ' ダウンロード中
            If lngReturnCode <> lngPrevReturnCode Then gApp.Log "ダウンロード中"
            mstrType = "ダウンロード待機"
        Case -201   ' Initされてない
            MsgBox "JVInitが行われていません。", vbExclamation, "馬吉：エラー"
            Exit Do
        Case -203   ' Openされてない
            MsgBox "JVOpenが行われていません。", vbExclamation, "馬吉：エラー"
            Exit Do
        Case -502   ' ダウンロード失敗
            MsgBox "ダウンロード中にエラーが発生しました。", vbExclamation, "馬吉：エラー"
            Exit Do
        Case -503   ' ファイルがない
            MsgBox "ファイルがありません。", vbExclamation, "馬吉：エラー"
            Exit Do
        Case Is > 0 ' 正常読み込み
            mlngReadByte = mlngReadByte + lngReturnCode
            
            'レコード種別IDを取得
            strRecordIDNew = Left$(strBuff, 2)
            mstrType = Left$(strBuff, 20)


            'レコード種別IDが変更された場合（または初処理時）
            If strRecordIDNew <> strRecordIDOld Then
                mstrType = strRecordIDNew

                ' 測定結果を表示
                If lngSubCount <> 0 Then
                    sngSubTimerEnd = Timer
                    gApp.Log "レコード種別が " & strRecordIDOld & " から " & strRecordIDNew & "に変わりました"
                    gApp.Log vbTab & strRecordIDOld & "処理行数:" & CStr(lngSubCount)
                    gApp.Log vbTab & strRecordIDOld & "読み込み時間:" & CStr(sngSubTimerEnd - sngSubTimerStart) & "秒"
                    If (sngSubTimerEnd - sngSubTimerStart) > 0 Then
                        gApp.Log vbTab & strRecordIDOld & "処理速度:" & CStr(lngSubCount / (sngSubTimerEnd - sngSubTimerStart)) & "行/秒"
                    End If
                End If
                
                If Not ImportObj Is Nothing Then
                    'インポートオブジェクトに終了処理をさせる
                    Call ImportObj.CloseDB
                End If
                
                'インポートオブジェクトを変更する
                Set ImportObj = SelectImportObj(strRecordIDNew)
                Call ImportObj.OpenDB

                'レコード種別IDの保存
                strRecordIDOld = strRecordIDNew

                '開始時刻を保存
                sngSubTimerStart = Timer

                '種別毎カウンタ初期化
                lngSubCount = 0

            End If

            If Not ImportObj Is Nothing Then

                'DB追加処理
                If Not ImportObj.Add(StrConv(strBuff, vbFromUnicode)) Then
                    gApp.Log "レコード登録の失敗" & strBuff

                End If

                'カウントアップ
                lngSubCount = lngSubCount + 1
                lngCount = lngCount + 1
            End If


        Case Else
            gApp.Log "不明なリターンコード" & lngReturnCode
        End Select

        'リターンコードを保存
        lngPrevReturnCode = lngReturnCode

    Loop

    sngSubTimerEnd = Timer
    sngTimerEnd = Timer
    gApp.Log "レコード種別 " & strRecordIDOld & "の読み込み処理を終了しました"
    gApp.Log vbTab & strRecordIDOld & "処理行数:" & CStr(lngSubCount)
    gApp.Log vbTab & strRecordIDOld & "読み込み時間:" & CStr(sngSubTimerEnd - sngSubTimerStart) & "秒"
    If (sngSubTimerEnd - sngSubTimerStart) > 0 Then
        gApp.Log vbTab & strRecordIDOld & "処理速度:" & CStr(lngSubCount / (sngSubTimerEnd - sngSubTimerStart)) & "行/秒"
    End If

    gApp.Log "すべてのデータ読み込み処理を終了しました"
    gApp.Log vbTab & "全体処理行数:" & CStr(lngCount)
    gApp.Log vbTab & "全体読み込み時間:" & CStr(sngTimerEnd - sngTimerStart) & "秒"
    If (sngTimerEnd - sngTimerStart) > 0 Then
        gApp.Log "全体処理速度:" & CStr(lngCount / (sngTimerEnd - sngTimerStart)) & "行/秒"
    End If

    '---------
    ' JVClose
    
    axJVLink.JVClose

    Set objMP = Nothing

    FetchJVData = True
    Exit Function
ErrorHandler:
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
On Error GoTo ErrorHandler
    Unload Me
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: フォーム初期化イベント
'
'   備考: なし
'
Private Sub Form_Initialize()
On Error GoTo ErrorHandler
    mMode = ukpRA
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: フォームロードイベント
'
'   備考: なし
'
Private Sub Form_Load()
On Error GoTo ErrorHandler


    If mstrKey = "" Then
        gApp.Log TypeName(Me) & "キー未設定"
        mblnFinishFlag = True
        Unload Me
    Else
        mblnFinishFlag = False
    
        tmrStartTrigger.Interval = 1
        tmrStartTrigger.Enabled = True
    End If

    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: フォームアンロード確認イベント
'
'   備考: なし
'
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrorHandler
    If UnloadMode <> vbFormCode Then
        If MsgBox("データの取得を中止しますか？", vbYesNo + vbQuestion, "馬吉：データ取得中止の確認") = vbYes Then
            gApp.Log "データ取得のキャンセル"
            axJVLink.JVClose
            Set objMP = Nothing
            Cancel = True
            mblnCancelFlag = True
        Else
            Cancel = True
        End If
    End If
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: リフレッシュタイマーイベント
'
'   備考: なし
'
Private Sub tmrRefresh_Timer()
On Error GoTo ErrorHandler
    Dim max         As Long
    Dim persent     As Long
    
    max = axJVLink.m_TotalReadFilesize
    If max <> 0 Then
        persent = mlngReadByte / max
    Else
        persent = 100
    End If
    
    
    ' プログレスバー変更
    If max > 0 Then
        '
    End If
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: 取得メインループ
'
'   備考: なし
'
Private Sub tmrStartTrigger_Timer()
On Error GoTo ErrorHandler
    Dim returnFlag As Boolean
    Dim AllKeys()  As String
    Dim i As Long

    Log "データ取得の開始"
    tmrStartTrigger.Enabled = False
    Select Case mMode
    
    Case ukpRA ' 速報レース
        Log "レース結果取得"
        returnFlag = returnFlag Or Not FetchJVData("0B12") ' レース結果
        If returnFlag = False Then
            Log "馬体重取得"
            returnFlag = returnFlag Or Not FetchJVData("0B11") ' 馬体重
        End If
        If returnFlag = False Then
            Log "データマイニング取得"
            returnFlag = returnFlag Or Not FetchJVData("0B13") ' データマイニング
        End If
        If returnFlag = False Then
            Log "開催情報取得"
            returnFlag = returnFlag Or Not FetchJVData("0B14") ' 開催情報
        End If
        

    Case ukpOD ' 速報オッズ
        Log "オッズ取得"
        returnFlag = returnFlag Or Not FetchJVData("0B30") ' 速報オッズ(全賭式)
        If returnFlag = False Then
            Log "票数取得"
            returnFlag = returnFlag Or Not FetchJVData("0B20") ' 速報票数(全賭式)
        End If
    Case ukpPALLET ' メニューパレット
        AllKeys = Split(mstrKey, ",")
                        
        For i = 0 To UBound(AllKeys)
            mstrKey = AllKeys(i)
            Log "RT取得:" & mstrKey
            If returnFlag = False Then
                Log "レース結果取得"
                returnFlag = returnFlag Or Not FetchJVData("0B12") ' レース結果
            End If
            If returnFlag = False Then
                Log "馬体重取得"
                returnFlag = returnFlag Or Not FetchJVData("0B11") ' 馬体重
            End If
            If returnFlag = False Then
                Log "データマイニング取得"
                returnFlag = returnFlag Or Not FetchJVData("0B13") ' データマイニング
            End If
            If returnFlag = False Then
                Log "開催情報取得"
                returnFlag = returnFlag Or Not FetchJVData("0B14") ' 開催情報
            End If
        Next i
    End Select
    
    If returnFlag Then
        gApp.Log "データ取得の失敗"
    Else
        gApp.Log "データ取得の正常終了"
        MsgBox "取得が終了しました｡", vbOKOnly + vbInformation, "馬吉：取得終了"
    End If
    
    mblnFinishFlag = True
    Unload Me
    Exit Sub
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   機能: インポートオブジェクト選択処理
'
'   備考: なし
'
Private Function SelectImportObj(strRecrodID As String) As clsIImport
    Select Case strRecrodID
    Case "AV"
        Set SelectImportObj = New clsImportAV
    Case "BN"
        Set SelectImportObj = New clsImportBN
    Case "BR"
        Set SelectImportObj = New clsImportBR
    Case "CH"
        Set SelectImportObj = New clsImportCH
    Case "DM"
        Set SelectImportObj = New clsImportDM
    Case "HC"
        Set SelectImportObj = New clsImportHC
    Case "HN"
        Set SelectImportObj = New clsImportHN
    Case "HR"
        Set SelectImportObj = New clsImportHR
    Case "JC"
        Set SelectImportObj = New clsImportJC
    Case "KS"
        Set SelectImportObj = New clsImportKS
    Case "O1"
        Set SelectImportObj = New clsImportO1
    Case "O2"
        Set SelectImportObj = New clsImportO2
    Case "O3"
        Set SelectImportObj = New clsImportO3
    Case "O4"
        Set SelectImportObj = New clsImportO4
    Case "O5"
        Set SelectImportObj = New clsImportO5
    Case "RA"
        Set SelectImportObj = New clsImportRA
    Case "RC"
        Set SelectImportObj = New clsImportRC
    Case "SE"
        Set SelectImportObj = New clsImportSE
    Case "SK"
        Set SelectImportObj = New clsImportSK
    Case "TK"
        Set SelectImportObj = New clsImportTK
    Case "UM"
        Set SelectImportObj = New clsImportUM
    Case "WE"
        Set SelectImportObj = New clsImportWE
    Case "WH"
        Set SelectImportObj = New clsImportWH
    Case "YS"
        Set SelectImportObj = New clsImportYS
    Case "TC"
        Set SelectImportObj = New clsImportTC
    Case "CC"
        Set SelectImportObj = New clsImportCC
    Case "H1", "H6", "O6"
        Set SelectImportObj = New clsImportODDS
    Case Else
        Set SelectImportObj = Nothing
    End Select
End Function


'
'   機能: ログ出力処理
'
'   備考: なし
'
Private Sub Log(strText As String)
    stbBar.Panels(1).Text = strText
    gApp.Log strText
End Sub
