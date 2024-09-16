VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "アプリケーションクラス"
'
'   アプリケーションクラス
'

Option Explicit

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   API関数宣言
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

' 表示要素の寸法とシステム構成の設定を取得する関数の宣言
Private Declare Function GetSystemMetrics Lib "user32.dll" _
                    (ByVal nIndex As Long) As Long
' 取得するシステムメトリックまたは現在の構成を示す定数の宣言
Private Const SM_CXVSCROLL As Long = 2  '' 垂直スクロールバー高さ
Private Const SM_CYHSCROLL As Long = 3  '' 水平スクロールバー幅
Private Const SM_CYVSCROLL As Long = 20 '' 垂直スクロールバー
Private Const SM_CXHSCROLL As Long = 21 '' 水平スクロールバー


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部変数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private mBM     As clsBrowserMgr
Private mDM     As clsDatabaseMgr
Private mPallet As frmMenu

Private mlngFetchingCounter As Long '' 取得中状態カウンター


' 不揮発メモリから取得済みかどうかを判定するフラグ
Private mblnDBPath              As Boolean      '' データベース保管フォルダパス
Private mblnKakoNum             As Boolean      '' 表示過去走数
Private mblnHomeViewer          As Boolean      '' ホーム画面Viewer名
Private mblnHomeKey             As Boolean      '' ホーム画面キー

' 不揮発メモリから取得した値を保持するモジュール変数
Private mstrDBPath              As String       '' データベース保管フォルダパス
Private mlngKakoNum             As Long         '' 表示過去走数
Private mstrHomeViewer          As String       '' ホーム画面Viewer名
Private mstrHomeKey             As String       '' ホーム画面キー

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   プロパティ
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: データベース保管フォルダパスの値を取り出す
'
'   備考: 初期値のままの場合、"0" を返す
'
Public Property Get R_DBVersion() As String
    R_DBVersion = val(GetIniData("DatabaseInfo", "Version", "0", gApp.R_DBPath & "\UmakichiDB.ini"))
End Property


'
'   機能: データベースバージョンの値を保存する
'
'   備考: なし
'
Public Property Let R_DBVersion(RHS As String)
    SetIniData "DatabaseInfo", "Version", RHS, gApp.R_DBPath & "\UmakichiDB.ini"    'DBのバージョン変更
End Property


'
'   機能: レジストリから、データベース保管フォルダパスの値を取り出す
'
'   備考: 初期値のままの場合、App.Path\DB を返す
'
Public Property Get R_DBPath() As String
    Dim strPath As String

    If mblnDBPath = False Then
        strPath = LoadRegSub("Database", "Path", App.Path & "\DB")
        mstrDBPath = strPath
        mblnDBPath = True
    Else
        strPath = mstrDBPath
    End If

    R_DBPath = strPath
End Property


'
'   機能: レジストリに、データベース保管フォルダパスの値を保存する
'
'   備考: なし
'
Public Property Let R_DBPath(RHS As String)
    WriteRegSub "Database", "Path", RHS
    mblnDBPath = False
End Property


'
'   機能: JVData最終取得日時の値を取り出す
'
'   備考: なし
'
Public Property Get R_SetupYear() As String
    R_SetupYear = val(GetIniData("DatabaseInfo", "SetupYear", "0000", gApp.R_DBPath & "\UmakichiDB.ini"))
End Property


'
'   機能: JVData最終取得日時の値を保存する
'
'   備考: なし
'
Public Property Let R_SetupYear(RHS As String)
     SetIniData "DatabaseInfo", "SetupYear", RHS, gApp.R_DBPath & "\UmakichiDB.ini"
End Property


'
'   機能: JVData最終取得日時の値を取り出す
'
'   備考: なし
'
Public Property Get R_SetupCancelLastTime() As String
    R_SetupCancelLastTime = GetIniData("DatabaseInfo", "SetupCancelLastTime", "", gApp.R_DBPath & "\UmakichiDB.ini")
End Property


'
'   機能: JVData最終取得日時の値を保存する
'
'   備考: なし
'
Public Property Let R_SetupCancelLastTime(RHS As String)
     SetIniData "DatabaseInfo", "SetupCancelLastTime", RHS, gApp.R_DBPath & "\UmakichiDB.ini"
End Property


'
'   機能: JVData最終取得日時の値を取り出す
'
'   備考: なし
'
Public Property Get R_JVDLastTime() As String
    R_JVDLastTime = GetIniData("DatabaseInfo", "JVDLastTime", String$(14, "0"), gApp.R_DBPath & "\UmakichiDB.ini")
End Property


'
'   機能: JVData最終取得日時の値を保存する
'
'   備考: なし
'
Public Property Let R_JVDLastTime(RHS As String)
     SetIniData "DatabaseInfo", "JVDLastTime", RHS, gApp.R_DBPath & "\UmakichiDB.ini"
End Property


'
'   機能: JVData SLOP 最終取得日時の値を取り出す
'
'   備考: なし
'
Public Property Get R_JVDLastTimeSLOP() As String
    R_JVDLastTimeSLOP = GetIniData("DatabaseInfo", "JVDLastTimeSLOP", String$(14, "0"), gApp.R_DBPath & "\UmakichiDB.ini")
End Property


'
'   機能: JVData SLOP 最終取得日時の値を保存する
'
'   備考: なし
'
Public Property Let R_JVDLastTimeSLOP(RHS As String)
     SetIniData "DatabaseInfo", "JVDLastTimeSLOP", RHS, gApp.R_DBPath & "\UmakichiDB.ini"
End Property


'
'   機能: JVData BLOD 最終取得日時の値を取り出す
'
'   備考: なし
'
Public Property Get R_JVDLastTimeBLOD() As String
    R_JVDLastTimeBLOD = GetIniData("DatabaseInfo", "JVDLastTimeBLOD", String$(14, "0"), gApp.R_DBPath & "\UmakichiDB.ini")
End Property


'
'   機能: JVData BLOD 最終取得日時の値を保存する
'
'   備考: なし
'
Public Property Let R_JVDLastTimeBLOD(RHS As String)
     SetIniData "DatabaseInfo", "JVDLastTimeBLOD", RHS, gApp.R_DBPath & "\UmakichiDB.ini"
End Property


'
'   機能: JVDataThisWeek最終取得日時の値を取り出す
'
'   備考: なし
'
Public Property Get R_JVDLastTimeThisWeek() As String
    R_JVDLastTimeThisWeek = GetIniData("DatabaseInfo", "JVDLastTimeThisWeek", "00000000000000", gApp.R_DBPath & "\UmakichiDB.ini")
End Property


'
'   機能: JVDataThisWeek最終取得日時の値を保存する
'
'   備考: なし
'
Public Property Let R_JVDLastTimeThisWeek(RHS As String)
    SetIniData "DatabaseInfo", "JVDLastTimeThisWeek", RHS, gApp.R_DBPath & "\UmakichiDB.ini"
End Property


'
'   機能: レジストリから、表示過去走数の値を取り出す
'
'   備考: なし
'
Public Property Get R_KakoNum() As Long
    Dim lngOut As Long
    
    If mblnKakoNum = False Then
        lngOut = CLng(LoadRegSub("ViewerState", "KakoNum", "5"))
        mlngKakoNum = lngOut
        mblnKakoNum = True
    Else
        lngOut = mlngKakoNum
    End If
    R_KakoNum = lngOut
End Property


'
'   機能: レジストリに、表示過去走数の値を保存する
'
'   備考: なし
'
Public Property Let R_KakoNum(RHS As Long)
    WriteRegSub "ViewerState", "KakoNum", CStr(RHS)
    mblnKakoNum = False
End Property


'
'   機能: レジストリから、ホーム画面Viewer名の値を取り出す
'
'   備考: なし
'
Public Property Get R_HomeViewer() As String
    Dim strOut As String
    If mblnHomeViewer = False Then
        strOut = LoadRegSub("Home", "Viewer", "Home")
        mstrHomeViewer = strOut
        mblnHomeViewer = True
    Else
        strOut = mstrHomeViewer
    End If
    R_HomeViewer = strOut
End Property


'
'   機能: レジストリに、ホーム画面Viewer名の値を保存する
'
'   備考: なし
'
Public Property Let R_HomeViewer(RHS As String)
    WriteRegSub "Home", "Viewer", RHS
    mblnHomeViewer = False
End Property


'
'   機能: レジストリから、ホーム画面キーの値を取り出す
'
'   備考: なし
'
Public Property Get R_HomeKey() As String
    Dim strOut As String
    If mblnHomeKey = False Then
        strOut = LoadRegSub("Home", "Key", "Empty")
        mstrHomeKey = strOut
        mblnHomeKey = True
    Else
        strOut = mstrHomeKey
    End If
    R_HomeKey = strOut
End Property


'
'   機能: レジストリに、ホーム画面キーの値を保存する
'
'   備考: なし
'
Public Property Let R_HomeKey(RHS As String)
    WriteRegSub "Home", "Key", RHS
    mblnHomeKey = False
End Property


'
'   機能: レジストリから、背景色明の値を取り出す
'
'   備考: なし
'
Public Property Get R_BackColorLight() As Long
    R_BackColorLight = CLng(LoadRegSub("Color", "BackColorLight", CStr(&HE0EEEE)))
End Property


'
'   機能: レジストリに、背景色明の値を保存する
'
'   備考: なし
'
Public Property Let R_BackColorLight(RHS As Long)
    WriteRegSub "Color", "BackColorLight", CStr(RHS)
End Property


'
'   機能: レジストリから、背景色暗の値を取り出す
'
'   備考: なし
'
Public Property Get R_BackColorDark() As Long
    R_BackColorDark = CLng(LoadRegSub("Color", "BackColorDark", CStr(&HC0CCCC)))
End Property


'
'   機能: レジストリに、背景色暗の値を保存する
'
'   備考: なし
'
Public Property Let R_BackColorDark(RHS As Long)
    WriteRegSub "Color", "BackColorDark", CStr(RHS)
End Property


'
'   機能: レジストリから、メニューの表示状態の値を取り出す
'
'   備考: なし
'
Public Property Get R_MenuVisible() As Boolean
    R_MenuVisible = CBool(LoadRegSub("MenuPalette", "Visible", CStr(True)))
End Property


'
'   機能: レジストリに、メニューの表示状態の値を保存する
'
'   備考: なし
'
Public Property Let R_MenuVisible(RHS As Boolean)
    WriteRegSub "MenuPalette", "Visible", CStr(RHS)
End Property


'
'   機能: レジストリから、メニューを常に手前に表示するの値を取り出す
'
'   備考: なし
'
Public Property Get R_MenuAlwaysTopFlag() As Boolean
    R_MenuAlwaysTopFlag = CBool(LoadRegSub("MenuPalette", "AlwaysTop", CStr(False)))
End Property


'
'   機能: レジストリに、メニューを常に手前に表示するの値を保存する
'
'   備考: なし
'
Public Property Let R_MenuAlwaysTopFlag(RHS As Boolean)
    WriteRegSub "MenuPalette", "AlwaysTop", CStr(RHS)
    Call mPallet.AlwaysTop(RHS)
End Property


'
'   機能: レジストリから、メニューを残すの値を取り出す
'
'   備考: なし
'
Public Property Get R_MenuCanAlone() As Boolean
    R_MenuCanAlone = CBool(LoadRegSub("MenuPalette", "CanAlone", CStr(False)))
End Property


'
'   機能: レジストリに、メニューを残すの値を保存する
'
'   備考: なし
'
Public Property Let R_MenuCanAlone(RHS As Boolean)
    WriteRegSub "MenuPalette", "CanAlone", CStr(RHS)
End Property


'
'   機能: JV-Link取得モード 通常<->今週　を取り出す
'
'   備考: なし
'
Public Property Get R_JVLMode() As ukJVLMode
    Dim tmp As ukJVLMode
    Select Case GetIniData("DatabaseInfo", "Mode", "Usual", gApp.R_DBPath & "\UmakichiDB.ini")
    Case "Usual"
        tmp = ukjUsual
    Case "ThisWeek"
        tmp = ukjThisWeek
    End Select
    R_JVLMode = tmp
End Property


'
'   機能: JV-Link取得モード 通常<->今週　を保存する
'
'   備考: なし
'
Public Property Let R_JVLMode(RHS As ukJVLMode)
    Dim strTmp As String
    If RHS = ukjUsual Then
        strTmp = "Usual"
    ElseIf RHS = ukjThisWeek Then
        strTmp = "ThisWeek"
    End If
    SetIniData "DatabaseInfo", "Mode", strTmp, gApp.R_DBPath & "\UmakichiDB.ini"
End Property


'
'   機能: JV-Link取得モード SLOP　を取り出す
'
'   備考: なし
'
Public Property Get R_JVLGetSLOP() As Boolean
    R_JVLGetSLOP = (GetIniData("DatabaseInfo", "GetSLOP", "True", gApp.R_DBPath & "\UmakichiDB.ini") = "True")
End Property


'
'   機能: JV-Link取得モード SLOP　を保存する
'
'   備考: なし
'
Public Property Let R_JVLGetSLOP(RHS As Boolean)
    SetIniData "DatabaseInfo", "GetSLOP", IIf(RHS, "True", "False"), gApp.R_DBPath & "\UmakichiDB.ini"
End Property


'
'   機能: JV-Link産駒・繁殖馬を含めるフラグを取り出す
'
'   備考: なし
'
Public Property Get R_JVLGetBLOD() As Boolean
    R_JVLGetBLOD = (GetIniData("DatabaseInfo", "GetBLOD", "True", gApp.R_DBPath & "\UmakichiDB.ini") = "True")
End Property


'
'   機能: JV-Link産駒・繁殖馬を含めるフラグを保存する
'
'   備考: なし
'
Public Property Let R_JVLGetBLOD(RHS As Boolean)
    SetIniData "DatabaseInfo", "GetBLOD", IIf(RHS, "True", "False"), gApp.R_DBPath & "\UmakichiDB.ini"
End Property


'
'   機能: 出馬表開催選択のキャッシュが有効である場合に真を返す
'
'   備考: 高速に表示するための先読みキャッシュが存在して最新である事の確認
'
Public Property Get R_RAKaiSelCacheExist(strYear As String) As Boolean
    R_RAKaiSelCacheExist = (GetIniData("RAKaiSelCache", strYear, "False", gApp.R_DBPath & "\UmakichiDB.ini") = "True")
End Property


'
'   機能: 出馬表開催選択のキャッシュの有効状態を設定する
'
'   備考: なし
'
Public Property Let R_RAKaiSelCacheExist(strYear As String, RHS As Boolean)
     SetIniData "RAKaiSelCache", strYear, IIf(RHS, "True", "False"), gApp.R_DBPath & "\UmakichiDB.ini"
End Property


'
'   機能: 速報取得用の日付を返す
'
'   備考: なし
'
Public Property Get R_RTDates() As String
    R_RTDates = GetIniData("DatabaseInfo", "RTDates", "", gApp.R_DBPath & "\UmakichiDB.ini")
End Property


'
'   機能: 速報取得用の日付を保存する
'
'   備考: なし
'
Public Property Let R_RTDates(RHS As String)
     SetIniData "DatabaseInfo", "RTDates", RHS, gApp.R_DBPath & "\UmakichiDB.ini"
End Property


'
'   機能: レジストリのサブキーを返す
'
'   備考: なし
'
Private Property Get RegSubKey() As String
On Error GoTo ErrH
    
    RegSubKey = "Software\" & cRegistrySubKey & "\"
    
    Exit Property
    
ErrH:
    Debug.Print Err.Description
End Property


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   外部関数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: 開始処理
'
'   備考: ブラウザを生成する
'
Public Sub start()
On Error GoTo ErrorHandler
    Call Log("Start")
    
    ' DB接続
    If Not mDM.Connect Then
        MsgBox "データベースに接続できません。終了します。", vbOKOnly + vbCritical, "馬吉：起動エラー"
        Call gApp.ExitApp
        End
    End If
    
    DoEvents
    
    ' 操作パレットの作成
    Me.R_MenuVisible = True
    mPallet.Show
    
    ' 新規ブラウザの作成（ホーム画面を表示）
    Call mBM.NewBrowser(gApp.R_HomeViewer, gApp.R_HomeKey)
    DoEvents
    

    Call mBM.AllShowMenuPalette(Me.R_MenuVisible)
    DoEvents
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
    mDM.Disconnect
End Sub


'
'   機能: コネクションを得る－BANUSI
'
'   備考: なし
'
Public Function GetCN_BANUSI() As ADODB.Connection
    Set GetCN_BANUSI = mDM.GetCN(0)
End Function

'
'   機能: コネクションを得る－BATAIJYU
'
'   備考: なし
'
Public Function GetCN_BATAIJYU() As ADODB.Connection
    Set GetCN_BATAIJYU = mDM.GetCN(1)
End Function

'
'   機能: コネクションを得る－CHOKYO
'
'   備考: なし
'
Public Function GetCN_CHOKYO() As ADODB.Connection
    Set GetCN_CHOKYO = mDM.GetCN(2)
End Function

'
'   機能: コネクションを得る－CHOKYO_SEISEKI
'
'   備考: なし
'
Public Function GetCN_CHOKYO_SEISEKI() As ADODB.Connection
    Set GetCN_CHOKYO_SEISEKI = mDM.GetCN(3)
End Function

'
'   機能: コネクションを得る－HANRO
'
'   備考: なし
'
Public Function GetCN_HANRO() As ADODB.Connection
    Set GetCN_HANRO = mDM.GetCN(4)
End Function

'
'   機能: コネクションを得る－HANSYOKU
'
'   備考: なし
'
Public Function GetCN_HANSYOKU() As ADODB.Connection
    Set GetCN_HANSYOKU = mDM.GetCN(5)
End Function

'
'   機能: コネクションを得る－HARAI
'
'   備考: なし
'
Public Function GetCN_HARAI() As ADODB.Connection
    Set GetCN_HARAI = mDM.GetCN(6)
End Function

'
'   機能: コネクションを得る－KISHU
'
'   備考: なし
'
Public Function GetCN_KISHU() As ADODB.Connection
    Set GetCN_KISHU = mDM.GetCN(7)
End Function

'
'   機能: コネクションを得る－KISHU_CHANGE
'
'   備考: なし
'
Public Function GetCN_KISHU_CHANGE() As ADODB.Connection
    Set GetCN_KISHU_CHANGE = mDM.GetCN(8)
End Function

'
'   機能: コネクションを得る－KISHU_SEISEKI
'
'   備考: なし
'
Public Function GetCN_KISHU_SEISEKI() As ADODB.Connection
    Set GetCN_KISHU_SEISEKI = mDM.GetCN(9)
End Function

'
'   機能: コネクションを得る－MINING
'
'   備考: なし
'
Public Function GetCN_MINING() As ADODB.Connection
    Set GetCN_MINING = mDM.GetCN(10)
End Function

'
'   機能: コネクションを得る－ODDS_SANREN
'
'   備考: indexを引き数にして各個接続
'
Public Function GetCN_ODDS_SANREN(Index As Long) As ADODB.Connection
    Set GetCN_ODDS_SANREN = mDM.GetCN(11 + Index)
End Function

'
'   機能: コネクションを得る－ODDS_SANREN0
'
'   備考: なし
'
Public Function GetCN_ODDS_SANREN0() As ADODB.Connection
    Set GetCN_ODDS_SANREN0 = mDM.GetCN(11)
End Function

'
'   機能: コネクションを得る－ODDS_SANREN1
'
'   備考: なし
'
Public Function GetCN_ODDS_SANREN1() As ADODB.Connection
    Set GetCN_ODDS_SANREN1 = mDM.GetCN(12)
End Function

'
'   機能: コネクションを得る－ODDS_SANREN2
'
'   備考: なし
'
Public Function GetCN_ODDS_SANREN2() As ADODB.Connection
    Set GetCN_ODDS_SANREN2 = mDM.GetCN(13)
End Function

'
'   機能: コネクションを得る－ODDS_SANREN3
'
'   備考: なし
'
Public Function GetCN_ODDS_SANREN3() As ADODB.Connection
    Set GetCN_ODDS_SANREN3 = mDM.GetCN(14)
End Function

'
'   機能: コネクションを得る－ODDS_SANREN4
'
'   備考: なし
'
Public Function GetCN_ODDS_SANREN4() As ADODB.Connection
    Set GetCN_ODDS_SANREN4 = mDM.GetCN(15)
End Function

'
'   機能: コネクションを得る－ODDS_SANREN5
'
'   備考: なし
'
Public Function GetCN_ODDS_SANREN5() As ADODB.Connection
    Set GetCN_ODDS_SANREN5 = mDM.GetCN(16)
End Function

'
'   機能: コネクションを得る－ODDS_SANREN6
'
'   備考: なし
'
Public Function GetCN_ODDS_SANREN6() As ADODB.Connection
    Set GetCN_ODDS_SANREN6 = mDM.GetCN(17)
End Function

'
'   機能: コネクションを得る－ODDS_SANREN7
'
'   備考: なし
'
Public Function GetCN_ODDS_SANREN7() As ADODB.Connection
    Set GetCN_ODDS_SANREN7 = mDM.GetCN(18)
End Function

'
'   機能: コネクションを得る－ODDS_SANREN8
'
'   備考: なし
'
Public Function GetCN_ODDS_SANREN8() As ADODB.Connection
    Set GetCN_ODDS_SANREN8 = mDM.GetCN(19)
End Function

'
'   機能: コネクションを得る－ODDS_SANREN9
'
'   備考: なし
'
Public Function GetCN_ODDS_SANREN9() As ADODB.Connection
    Set GetCN_ODDS_SANREN9 = mDM.GetCN(20)
End Function

'
'   機能: コネクションを得る－ODDS_TANPUKUWAKU
'
'   備考: なし
'
Public Function GetCN_ODDS_TANPUKUWAKU() As ADODB.Connection
    Set GetCN_ODDS_TANPUKUWAKU = mDM.GetCN(21)
End Function

'
'   機能: コネクションを得る－ODDS_UMAREN
'
'   備考: なし
'
Public Function GetCN_ODDS_UMAREN() As ADODB.Connection
    Set GetCN_ODDS_UMAREN = mDM.GetCN(22)
End Function

'
'   機能: コネクションを得る－ODDS_UMATAN
'
'   備考: なし
'
Public Function GetCN_ODDS_UMATAN(Index As Long) As ADODB.Connection
    Set GetCN_ODDS_UMATAN = mDM.GetCN(23 + Index)
End Function

'
'   機能: コネクションを得る－ODDS_UMATAN0
'
'   備考: なし
'
Public Function GetCN_ODDS_UMATAN0() As ADODB.Connection
    Set GetCN_ODDS_UMATAN0 = mDM.GetCN(23)
End Function

'
'   機能: コネクションを得る－ODDS_UMATAN1
'
'   備考: なし
'
Public Function GetCN_ODDS_UMATAN1() As ADODB.Connection
    Set GetCN_ODDS_UMATAN1 = mDM.GetCN(24)
End Function

'
'   機能: コネクションを得る－ODDS_UMATAN2
'
'   備考: なし
'
Public Function GetCN_ODDS_UMATAN2() As ADODB.Connection
    Set GetCN_ODDS_UMATAN2 = mDM.GetCN(25)
End Function

'
'   機能: コネクションを得る－ODDS_UMATAN3
'
'   備考: なし
'
Public Function GetCN_ODDS_UMATAN3() As ADODB.Connection
    Set GetCN_ODDS_UMATAN3 = mDM.GetCN(26)
End Function

'
'   機能: コネクションを得る－ODDS_UMATAN4
'
'   備考: なし
'
Public Function GetCN_ODDS_UMATAN4() As ADODB.Connection
    Set GetCN_ODDS_UMATAN4 = mDM.GetCN(27)
End Function

'
'   機能: コネクションを得る－ODDS_UMATAN5
'
'   備考: なし
'
Public Function GetCN_ODDS_UMATAN5() As ADODB.Connection
    Set GetCN_ODDS_UMATAN5 = mDM.GetCN(28)
End Function

'
'   機能: コネクションを得る－ODDS_UMATAN6
'
'   備考: なし
'
Public Function GetCN_ODDS_UMATAN6() As ADODB.Connection
    Set GetCN_ODDS_UMATAN6 = mDM.GetCN(29)
End Function

'
'   機能: コネクションを得る－ODDS_UMATAN7
'
'   備考: なし
'
Public Function GetCN_ODDS_UMATAN7() As ADODB.Connection
    Set GetCN_ODDS_UMATAN7 = mDM.GetCN(30)
End Function

'
'   機能: コネクションを得る－ODDS_UMATAN8
'
'   備考: なし
'
Public Function GetCN_ODDS_UMATAN8() As ADODB.Connection
    Set GetCN_ODDS_UMATAN8 = mDM.GetCN(31)
End Function

'
'   機能: コネクションを得る－ODDS_UMATAN9
'
'   備考: なし
'
Public Function GetCN_ODDS_UMATAN9() As ADODB.Connection
    Set GetCN_ODDS_UMATAN9 = mDM.GetCN(32)
End Function

'
'   機能: コネクションを得る－ODDS_WIDE
'
'   備考: なし
'
Public Function GetCN_ODDS_WIDE() As ADODB.Connection
    Set GetCN_ODDS_WIDE = mDM.GetCN(33)
End Function

'
'   機能: コネクションを得る－RACE
'
'   備考: なし
'
Public Function GetCN_RACE() As ADODB.Connection
    Set GetCN_RACE = mDM.GetCN(34)
End Function

'
'   機能: コネクションを得る－RECORD
'
'   備考: なし
'
Public Function GetCN_RECORD() As ADODB.Connection
    Set GetCN_RECORD = mDM.GetCN(35)
End Function

'
'   機能: コネクションを得る－SANKU
'
'   備考: なし
'
Public Function GetCN_SANKU() As ADODB.Connection
    Set GetCN_SANKU = mDM.GetCN(36)
End Function

'
'   機能: コネクションを得る－SCHEDULE
'
'   備考: なし
'
Public Function GetCN_SCHEDULE() As ADODB.Connection
    Set GetCN_SCHEDULE = mDM.GetCN(37)
End Function

'
'   機能: コネクションを得る－SEISAN
'
'   備考: なし
'
Public Function GetCN_SEISAN() As ADODB.Connection
    Set GetCN_SEISAN = mDM.GetCN(38)
End Function

'
'   機能コネクションを得る－TENKO_BABA
'
'   備考: なし
'
Public Function GetCN_TENKO_BABA() As ADODB.Connection
    Set GetCN_TENKO_BABA = mDM.GetCN(39)
End Function

'
'   機能コネクションを得る－TOKU
'
'   備考: なし
'
Public Function GetCN_TOKU() As ADODB.Connection
    Set GetCN_TOKU = mDM.GetCN(40)
End Function

'
'   機能コネクションを得る－TOKU_RACE
'
'   備考: なし
'
Public Function GetCN_TOKU_RACE() As ADODB.Connection
    Set GetCN_TOKU_RACE = mDM.GetCN(41)
End Function

'
'   機能: コネクションを得る－TORIKESI_JYOGAI
'
'   備考: なし
'
Public Function GetCN_TORIKESI_JYOGAI() As ADODB.Connection
    Set GetCN_TORIKESI_JYOGAI = mDM.GetCN(42)
End Function

'
'   機能: コネクションを得る－UMA
'
'   備考: なし
'
Public Function GetCN_UMA() As ADODB.Connection
    Set GetCN_UMA = mDM.GetCN(43)
End Function

'
'   機能: コネクションを得る－UMA_RACE_A
'
'   備考: なし
'
Public Function GetCN_UMA_RACE_A() As ADODB.Connection
    Set GetCN_UMA_RACE_A = mDM.GetCN(44)
End Function

'
'   機能: コネクションを得る－UMA_RACE_B
'
'   備考: なし
'
Public Function GetCN_UMA_RACE_B() As ADODB.Connection
    Set GetCN_UMA_RACE_B = mDM.GetCN(45)
End Function

'
'   機能: コネクションを得る－LinkTables
'
'   備考: なし
'
Public Function GetCN_LinkTables() As ADODB.Connection
    Set GetCN_LinkTables = mDM.GetCN(46)
End Function


'
'   機能: コネクションを得る－RAKaiSel
'
'   備考: なし
'
Public Function GetCN_RAKaiSel() As ADODB.Connection
    Set GetCN_RAKaiSel = mDM.GetCN(47)
End Function


'
'   機能: コネクションを得る－HASSOU_CHANGE
'
'   備考: なし
'
Public Function GetCN_HASSOU_CHANGE() As ADODB.Connection
    Set GetCN_HASSOU_CHANGE = mDM.GetCN(48)
End Function


'
'   機能: コネクションを得る－COURSE_CHANGE
'
'   備考: なし
'
Public Function GetCN_COURSE_CHANGE() As ADODB.Connection
    Set GetCN_COURSE_CHANGE = mDM.GetCN(49)
End Function


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部関数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: レジストリ情報読み取り処理
'
'   備考: 引き数 Key - レジストリキー(アプリケーションルート以下)
' 　　　         ValueName - 値の名前
' 　　　         Default - デフォルトの値のデータ
'
Private Function LoadRegSub( _
    ByVal key As String, _
    ByVal ValueName As String, _
    Optional ByVal Default As String = "") As String
    
On Error GoTo ErrH
    
    Dim strReturn       As String
    
    strReturn = GetRegData(HKEY_CURRENT_USER, RegSubKey & key, ValueName)
    If strReturn <> "" Then
        LoadRegSub = strReturn
    Else
        LoadRegSub = Default
    End If
    
    Exit Function
        
ErrH:
    gApp.ErrLog
End Function


'
'   機能: レジストリ情報書き込み処理
'
'   備考: 引き数 Key - レジストリキー(アプリケーションルート以下)
' 　　　         ValueName - 値の名前
' 　　　         Value - 値のデータ
'
Private Function WriteRegSub( _
    ByVal key As String, _
    ByVal ValueName As String, _
    ByVal value As String) As Boolean
    
On Error GoTo ErrH
    
    WriteRegSub = SetRegData(HKEY_CURRENT_USER, RegSubKey & key, ValueName, value)
    
    Exit Function
        
ErrH:
    gApp.ErrLog
End Function


'
'   機能: INIファイルを使う前のDBデータがあったらINIに取り込む
'
'   備考: なし
'
Private Sub DBVerUP()
On Error GoTo ErrorHandler
    Dim fso As FileSystemObject
    Dim ts  As TextStream
    Dim strFromTime As String
    Dim filename As String
    
    ' 旧バージョンのFromtimeがあるか調べる
    strFromTime = gApp.R_JVDLastTime
    If strFromTime = String$(14, "0") Then
        Set fso = New FileSystemObject
        filename = gApp.R_DBPath & "\" & cFromtimeFN
        If fso.FileExists(filename) Then
            Set ts = fso.OpenTextFile(filename, ForReading)
            strFromTime = ts.ReadLine
            gApp.R_JVDLastTime = strFromTime
            ts.Close
            fso.DeleteFile filename
        End If
    End If

    ' 旧バージョンのFromtimeThisWeekがあるか調べる
    strFromTime = gApp.R_JVDLastTimeThisWeek
    If strFromTime = String$(14, "0") Then
        Set fso = New FileSystemObject
        filename = gApp.R_DBPath & "\" & cFromtimeThisWeekFN
        If fso.FileExists(filename) Then
            Set ts = fso.OpenTextFile(filename, ForReading)
            strFromTime = ts.ReadLine
            gApp.R_JVDLastTimeThisWeek = strFromTime
        End If
    End If
    Exit Sub
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   機能: 初期化処理
'
'   備考: オブジェクト生成
'
Private Sub Class_Initialize()
    Set mBM = New clsBrowserMgr     '' ブラウザマネージャ
    Set mDM = New clsDatabaseMgr    '' データベースマネージャ
    Set mPallet = New frmMenu
End Sub


'
'   機能: 終了処理
'
'   備考: ログ記録、DB切断、オブジェクト参照削除
'
Private Sub Class_Terminate()
    Call Log("End")
    mDM.Disconnect
    Set mBM = Nothing
    Set mDM = Nothing
End Sub


'
'   機能: アプリケーションを終了する
'
'   備考: なし
'
Public Sub ExitApp()
    ' END
    Dim frm As Form
    ' すべてのフォームをアンロードする
    For Each frm In Forms
        gApp.Log "Unload " & TypeName(frm)
        Unload frm
    Next frm
End Sub


'
'   機能: メニューパレットの表示状態を設定する
'
'   備考: なし
'
Public Sub ShowMenuPalette(Mode As Boolean)
    If Mode = True Then
        mPallet.Show
        Me.R_MenuVisible = True
    Else
        Unload mPallet
        If mBM.BrowserExist Then
            Me.R_MenuVisible = False
        End If
    End If
    Call mBM.AllShowMenuPalette(Mode)
End Sub


'
'   機能: 新規ブラウザを作る「新しいウインドウで開く」
'
'   備考: 引き数 ViewerName - Viewer名
'                Key - キー
'
Public Sub NewWindow(ViewerName As String, key As String)
    Call mBM.NewBrowser(ViewerName, key)
End Sub


'
'   機能: ログにメッセージを出力する
'
'   備考: イミディエイトウインドウに出力
'
Public Sub Log(ByVal strMess As String)
On Error Resume Next
    Dim strFileName As String
    Dim FileNum As Integer
    Dim strOut As String
    
    
    strOut = Now & " " & Format$(Timer, "000000.00") & " " & strMess
    
    Debug.Print strOut
End Sub


'
'   機能: エラーメッセージ、エラー番号をログに出力する
'
'   備考: なし
'
Public Sub ErrLog()
    Log "ERROR No:" & Err.Number & " Description: " & Err.Description
    Debug.Assert (ASSERTMODE = 0)
End Sub


'
'   機能: オブジェクト初期化 をログに出力する
'
'   備考: なし
'
Public Sub InitLog(obj As Object, Optional ByVal mess As String)
    Log "≫ " & TypeName(obj) & "(" & ObjPtr(obj) & ") " & mess
End Sub


'
'   機能: オブジェクト終了 をログに出力する
'
'   備考: なし
'
Public Sub TermLog(obj As Object, Optional ByVal mess As String)
    Log "≪ " & TypeName(obj) & "(" & ObjPtr(obj) & ") " & mess
End Sub


'
'   機能: ブラウザ終了の為に、ブラウザマネージャからブラウザを登録抹消する
'
'   備考: ブラウザ終了時に必ずよぶ
'
Public Sub BrowserUnregist(obj As Form)
    Call mBM.Unregist(obj)
    If (Not gApp.R_MenuCanAlone) And (Not mBM.BrowserExist) Then
        gApp.ExitApp
    End If
End Sub


'
'   機能: DBからH1レコードを取得する
'
'   備考: H1のデータ取得は、この関数を用いて行う
'
Public Function GetH1(key As String) As String
On Error GoTo ErrorHandler
    Dim intFileNum As Integer
    
    Dim Year     As String
    Dim MonthDay As String
    Dim JyoRace  As String
    Dim Path     As String
    Dim filename As String
    Dim buf      As String
    
    Year = Mid$(key, 1, 4)
    MonthDay = Mid$(key, 5, 4)
    JyoRace = Mid$(key, 9, 2) & Mid$(key, 15, 2)
    
    Path = gApp.R_DBPath & "\ODDS\H1" & Year
    filename = Year & MonthDay & JyoRace
    
    intFileNum = FreeFile
    If Dir(Path & "\" & filename) <> "" Then
        Open Path & "\" & filename For Binary Access Read As intFileNum
        buf = String$(65535, " ")
        Get #intFileNum, , buf
        Close intFileNum
        
        GetH1 = buf
    End If
    Exit Function
ErrorHandler:
    gApp.ErrLog
End Function

'
'   機能: DBからH6レコードを取得する
'
'   備考: H6のデータ取得は、この関数を用いて行う
'
Public Function GetH6(key As String) As String
On Error GoTo ErrorHandler
    Dim intFileNum As Integer
    
    Dim Year     As String
    Dim MonthDay As String
    Dim JyoRace  As String
    Dim Path     As String
    Dim filename As String
    Dim buf      As String
    
    Year = Mid$(key, 1, 4)
    MonthDay = Mid$(key, 5, 4)
    JyoRace = Mid$(key, 9, 2) & Mid$(key, 15, 2)
    
    Path = gApp.R_DBPath & "\ODDS\H6" & Year
    filename = Year & MonthDay & JyoRace
    
    intFileNum = FreeFile
    If Dir(Path & "\" & filename) <> "" Then
        Open Path & "\" & filename For Binary Access Read As intFileNum
        buf = String$(102890, " ")
        Get #intFileNum, , buf
        Close intFileNum
        
        GetH6 = buf
    End If
    
    Exit Function
ErrorHandler:
    gApp.ErrLog
End Function

'
'   機能: DBからO6レコードを取得する
'
'   備考: O6のデータ取得は、この関数を用いて行う
'
Public Function GetO6(key As String) As String
On Error GoTo ErrorHandler
    Dim intFileNum As Integer
    
    Dim Year     As String
    Dim MonthDay As String
    Dim JyoRace  As String
    Dim Path     As String
    Dim filename As String
    Dim buf      As String
    
    Year = Mid$(key, 1, 4)
    MonthDay = Mid$(key, 5, 4)
    JyoRace = Mid$(key, 9, 2) & Mid$(key, 15, 2)
    
    Path = gApp.R_DBPath & "\ODDS\O6" & Year
    filename = Year & MonthDay & JyoRace
    
    intFileNum = FreeFile
    If Dir(Path & "\" & filename) <> "" Then
        Open Path & "\" & filename For Binary Access Read As intFileNum
        buf = String$(83285, " ")
        Get #intFileNum, , buf
        Close intFileNum
        
        GetO6 = buf
    End If

    Exit Function
ErrorHandler:
    gApp.ErrLog
End Function


'
'   機能: 全てのブラウザを再表示する
'
'   備考: なし
'
Public Sub AllReload()
On Error GoTo ErrorHandler
    mBM.AllReload
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: 環境設定モード
'
'   備考: モーダルダイアログとして環境設定画面を表示、DBのリンクテーブル修正など
'
Public Sub Configulation()
On Error Resume Next
    Dim frm As frmConfig
    Dim result As Long      '' MsgBoxの戻り値
    Dim BackupAlwaysTopFlag As Boolean
    
    BackupAlwaysTopFlag = gApp.R_MenuAlwaysTopFlag
    
    gApp.R_MenuAlwaysTopFlag = False
    
    Set frm = New frmConfig
    
    ' DBから切断
    mDM.Disconnect
    ' 設定画面表示（モーダル）
    frm.Show vbModal

    ' DBへ接続
    mDM.Connect

    mBM.AllReload
    
    Set frm = Nothing
    
    gApp.R_MenuAlwaysTopFlag = BackupAlwaysTopFlag

End Sub


'
'   機能: データ取得モード
'
'   備考: モーダルダイアログとしてデータ取得画面を表示、DBのリンクテーブル修正など
'
Public Sub DBUpdate()
On Error Resume Next
    Dim DBUpdateForm As frmDBUpdate
    Dim ConfigFirst As frmConfigFirst
    Dim BackupAlwaysTopFlag As Boolean
    
    BackupAlwaysTopFlag = gApp.R_MenuAlwaysTopFlag
    
    gApp.R_MenuAlwaysTopFlag = False

    If JVLinkCheck = False Then
        MsgBox "JV-Linkがインストールされていません。", vbOKOnly + vbExclamation, "馬吉：エラー"
        Exit Sub
    End If
    
    If Not JVLinkVersion2Up Then
        MsgBox "インストールされているJVLinkのバージョンでは、データの更新を行う事ができません。" _
            & vbCrLf & "データの更新を行う場合は、Ver2.1以降のJVLinkをインストールして下さい。", vbOKOnly + vbInformation, "馬吉：JVLinkバージョンチェック"
        Exit Sub
    End If
    
    If mDM.MissingDBNum <> 0 Then
        MsgBox "データベースが異常です。", vbOKOnly + vbCritical, "馬吉：エラー"
        Exit Sub
    End If
    
    If gApp.R_JVLMode = ukjUsual And _
        gApp.R_JVDLastTime = "00000000000000" And _
        gApp.R_SetupCancelLastTime = "" Then
        ' データセットアップの設定画面を出す
        Set ConfigFirst = New frmConfigFirst
        ConfigFirst.Show vbModal
        If ConfigFirst.ButtonType <> "OK" Then
            Exit Sub
        End If
    End If
    
    ' 通常取得を行う
    mBM.AllFreeViewer
    Set DBUpdateForm = New frmDBUpdate
    DBUpdateForm.Show vbModal
    mBM.AllReload
    gApp.R_MenuAlwaysTopFlag = BackupAlwaysTopFlag

End Sub


'
'   機能: 速報データ取得
'
'   備考: なし
'
Public Sub DBPrompt(Mode As ukPromptMode, key As String)
On Error Resume Next
    Dim frm As frmDBPrompt
    Dim BackupAlwaysTopFlag As Boolean
    
    BackupAlwaysTopFlag = gApp.R_MenuAlwaysTopFlag
    
    gApp.R_MenuAlwaysTopFlag = False
    
    If JVLinkCheck = False Then
        MsgBox "JV-Linkがインストールされていません。", vbOKOnly + vbExclamation, "馬吉：エラー"
        Exit Sub
    End If
    
    If Not JVLinkVersion2Up Then
        MsgBox "インストールされているJVLinkのバージョンでは、データの更新を行う事ができません。" _
            & vbCrLf & "データの更新を行う場合は、Ver2.1以降のJVLinkをインストールして下さい。", vbOKOnly + vbInformation, "馬吉：JVLinkバージョンチェック"
        Exit Sub
    End If
    
    If mlngFetchingCounter = 0 Then
        mDM.Disconnect
        mBM.AllFreeViewer
        Set frm = New frmDBPrompt
        frm.Mode = Mode
        frm.key = key
        frm.Show vbModal
        Set frm = Nothing
        mDM.Connect
    End If
    mBM.AllReload
    gApp.R_MenuAlwaysTopFlag = BackupAlwaysTopFlag
End Sub


'
'   機能: データベース最適化
'
'   備考: なし
'
Public Sub DBCompact()
On Error Resume Next
    Dim frm As frmDBMaintenance
    Dim fso As New FileSystemObject
    Dim BackupAlwaysTopFlag As Boolean
    
    BackupAlwaysTopFlag = gApp.R_MenuAlwaysTopFlag
    
    gApp.R_MenuAlwaysTopFlag = False

    If mDM.MissingDBNum <> 0 Then
        MsgBox "データベースが異常です。", vbOKOnly + vbCritical, "馬吉：エラー"
        Exit Sub
    End If

    
    Set frm = New frmDBMaintenance
    mBM.AllFreeViewer
    frm.Show vbModal
    Set frm = Nothing
    mBM.AllReload
    gApp.R_MenuAlwaysTopFlag = BackupAlwaysTopFlag
End Sub


'
'   機能: 枠番の色を返す
'
'   備考: なし
'
Public Function GetWakubanColor(Wakuban As Long) As Long
    Select Case Wakuban
    Case 1
        GetWakubanColor = RGB(255, 255, 255)
    Case 2
        GetWakubanColor = RGB(1, 1, 1)
    Case 3
        GetWakubanColor = RGB(255, 0, 0)
    Case 4
        GetWakubanColor = RGB(0, 0, 255)
    Case 5
        GetWakubanColor = RGB(255, 255, 0)
    Case 6
        GetWakubanColor = RGB(0, 255, 0)
    Case 7
        GetWakubanColor = RGB(255, 128, 0)
    Case 8
        GetWakubanColor = RGB(255, 128, 128)
    Case Else
        GetWakubanColor = RGB(255, 255, 255)
        
    End Select
    
End Function


'
'   機能: 着順色を返す
'
'   備考: なし
'
Public Function GetChakujyunColor(Chakujyun As Long) As Long
    Select Case Chakujyun
    Case 1
        GetChakujyunColor = RGB(255, 200, 200)
    Case 2
        GetChakujyunColor = RGB(255, 200, 128)
    Case 3
        GetChakujyunColor = RGB(200, 255, 255)
    Case Else
        GetChakujyunColor = RGB(255, 255, 255)
        
    End Select
    
End Function


'
'   機能: 背景色暗を返す
'
'   備考: なし
'
Public Function ColDarkBG() As Long
    ColDarkBG = gApp.R_BackColorDark
End Function


'
'   機能: 背景色を返す
'
'   備考: なし
'
Public Function ColBG() As Long
    ColBG = gApp.R_BackColorLight
End Function


'
'   機能: 垂直スクロールバーの幅
'
'   備考: なし
'
Public Function vsbWidth() As Long
    vsbWidth = GetSystemMetrics(SM_CXVSCROLL) * Screen.TwipsPerPixelX
End Function


'
'   機能: 水平スクロールバーの高さ
'
'   備考: なし
'
Public Function hsbHeight() As Long
    hsbHeight = GetSystemMetrics(SM_CYHSCROLL) * Screen.TwipsPerPixelY
End Function


'
'   機能: 切断
'
'   備考: なし
'
Public Sub Disconnect()
    mDM.Disconnect
End Sub



'
'   機能: 出馬表開催選択のキャッシュを全て削除
'
'   備考: なし
'
Public Sub DeleteAllRAKaiSelCacheFlags()
On Error GoTo ErrorHandler
    Dim y As String
    
    y = "1985"
    
    Do While y <= CStr(Year(Now))
        gApp.R_RAKaiSelCacheExist(y) = False
        y = y + 1
    Loop
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: JVLinkがインストールされているか調べる
'
'   備考: なし
'
Private Function JVLinkCheck() As Boolean
On Error GoTo ErrorHandler
    Dim JVlink As frmWrappedJVLink
    Set JVlink = New frmWrappedJVLink
    Load JVlink
    Unload JVlink
    JVLinkCheck = True
    Exit Function
ErrorHandler:
    JVLinkCheck = False
End Function


'
'   機能: JVLinkのバージョンを調べる
'
'   備考: なし
'
Public Function JVLinkVersion2Up() As Boolean
On Error GoTo ErrorHandler
    Dim JVlink As frmWrappedJVLink
    Dim lngVersion As Long
    Dim lngReturnCode As Long
    Dim strDummy As String
    Dim lngDummy1 As Long
    Dim lngDummy2 As Long
    
    Set JVlink = New frmWrappedJVLink
    Load JVlink
        
    Log "JVLink ver. " & JVlink.m_JVLinkVersion
    lngVersion = CLng(JVlink.m_JVLinkVersion)
    
    ' JVInit
    lngReturnCode = JVlink.axJVLink.JVInit(gJVLinkSID)
    If lngReturnCode = 0 Then
        ' JVOpen
        lngReturnCode = JVlink.axJVLink.JVOpen("RCOV", "99999999999999", 2, lngDummy1, lngDummy2, strDummy)
        If lngReturnCode < -1 Then
            gApp.Log "JVLink - JVOpenエラー"
            MsgBox "JVLink - JVOpenエラー", vbExclamation, "馬吉：JVLinkエラー" ' エラー終了
        End If
    
        ' JVClose
        JVlink.axJVLink.JVClose
    End If
    
    Unload JVlink
    If lngVersion < 210 Then
        JVLinkVersion2Up = False
    Else
        JVLinkVersion2Up = True
    End If
    
    Exit Function
ErrorHandler:
    JVLinkVersion2Up = False
End Function

