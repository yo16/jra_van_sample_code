VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDataOD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "オッズ・票数 データクラス"
'
'   オッズ・票数 データクラス
'

' GridData(n)
'  0 ~ 11 番号順グリッド
' 12 ~ 23 人気順グリッド
' 奇数 - オッズ表示
' 偶数 - 票数表示
'
'mGridData(0) 単複枠-オッズ（番号順）
'mGridData(1) 単複枠-票数　（番号順）
'mGridData(2) 馬連　-オッズ（番号順）
'mGridData(3) 馬連　-票数　（番号順）
'mGridData(4) ワイド-オッズ（番号順）
'mGridData(5) ワイド-票数　（番号順）
'mGridData(6) 馬単　-オッズ（番号順）
'mGridData(7) 馬単　-票数　（番号順）
'mGridData(8) 三連　-オッズ（番号順）
'mGridData(9) 三連　-票数　（番号順）
'mGridData(10)三連単-オッズ（番号順）
'mGridData(11)三連単-票数　（番号順）

'mGridData(12) 単複枠-オッズ（人気順）
'mGridData(13) 単複枠-票数　（人気順）
'mGridData(14) 馬連　-オッズ（人気順）
'mGridData(15) 馬連　-票数　（人気順）
'mGridData(16) ワイド-オッズ（人気順）
'mGridData(17) ワイド-票数　（人気順）
'mGridData(18) 馬単　-オッズ（人気順）
'mGridData(19) 馬単　-票数　（人気順）
'mGridData(20) 三連　-オッズ（人気順）
'mGridData(21) 三連　-票数　（人気順）
'mGridData(22) 三連単-オッズ（人気順）
'mGridData(23) 三連単-票数　（人気順）


Option Explicit

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   外部関数(イベント)
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

' イベント
Public Event NoData()
Public Event NoUMARACE()

Public Event Fetched(Index As Long)
Public Event FetchedH1()
Public Event FetchedAll()

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部変数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

' 変換ユーティリティー
Private mCC As clsCodeConverter           '' JV-Data のコード変換メソッド群
Private mSC As clsStringConverter         '' JV-Data のコード変換メソッド群
' Viewerに提供する為のデータ格納変数
Private mstrLabels(0 To 12)     As String       '' テキストデータ
Private mstrHyo(0 To 7)         As String       '' テキストデータ
Private mstrHappyo(0 To 7)      As String       '' 0:単複枠 1:馬連 2:ワイド 3:馬単 4:3連複 5:3連単 6:票数 7:3連単票数 (rs("DataKubun") & rs("HappyoTime"))
Private mGridData(0 To 23)      As clsGridData  '' グリッドデータ 'Sanrentan
Private mGridisExist(0 To 23)   As Boolean      '' グリッドデータの存在フラグ 'Sanrentan
Private mstrTitle               As String       '' タイトル
Private mlngTosu                As Long         '' 登録頭数
Private mblnIsG1                As Boolean      '' G1 レースかどうか
Private mblnIsPrompt            As Boolean      '' 速報を受け取れるレースかどうか
Private mblnO6Exist             As Boolean      '' 3連単オッズがあるかどうか
Private mblnH6Exist             As Boolean      '' 3連単票数があるかどうか

'モジュール内で使用する為のデータ格納変数
Private mbuf_H1         As JV_H1_HYOSU_ZENKAKE  '' 票数
Private mBuf_RA         As JV_RA_RACE
Private mKey            As clsKeyRA
Private mobj_Wakuban    As Scripting.Dictionary '馬番-枠番の対応を記憶
Private mblnOddsHyosuIsOdds As Boolean          'オッズ・票数どちらを取得するか

'コネクション
Private WithEvents mAsyncCN_UMA_RACE_A As ADODB.Connection
Attribute mAsyncCN_UMA_RACE_A.VB_VarHelpID = -1
' 変更情報
Private mRS_TENKO_BABA      As ADODB.Recordset ' 天候馬場
Private mRS_BATAIJYU        As ADODB.Recordset ' 馬体重
Private mRS_TORIKESI_JYOGAI As ADODB.Recordset ' 取り消し除外
Private mRS_KISHU_CHANGE    As ADODB.Recordset ' 騎手変更
Private mRS_MINING          As ADODB.Recordset ' マイニング
Private mRS_COURSE_CHANGE   As ADODB.Recordset ' コース変更
Private mRS_HASSOU_CHANGE   As ADODB.Recordset ' 発走時間

' 以下各グリッド用のデータ取得ディクショナリー
Private mobj_HyoTansyo As Scripting.Dictionary
Private mobj_HyoFukusyo As Scripting.Dictionary
Private mobj_HyoWakuren As Scripting.Dictionary
Private mobj_HyoUmaren As Scripting.Dictionary
Private mobj_HyoWide As Scripting.Dictionary
Private mobj_HyoUmatan As Scripting.Dictionary
Private mobj_HyoSanrenpuku As Scripting.Dictionary
Private mobj_OddsTansyo As Scripting.Dictionary
Private mobj_OddsFukusyo As Scripting.Dictionary
Private mobj_OddsWakuren As Scripting.Dictionary
Private mobj_OddsUmaren As Scripting.Dictionary
Private mobj_OddsWide As Scripting.Dictionary
Private mobj_OddsUmatan As Scripting.Dictionary
Private mobj_OddsSanrenpuku As Scripting.Dictionary
Private mobj_HyoSanrentan As Scripting.Dictionary
Private mobj_OddsSanrentan As Scripting.Dictionary

Private mblnCancelFetching As Boolean

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部定数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Const mlngNinkijunRow = 20                   '人気順の行数

'オッズ・票数　定数
Const ODDS_O1_NUM As Long = 0 'mGridData(0) 単複枠-オッズ（番号順）
Const ODDS_H1_NUM As Long = 1 'mGridData(1) 単複枠-票数　（番号順）
Const ODDS_O2_NUM As Long = 2 'mGridData(2) 馬連　-オッズ（番号順）
Const ODDS_H2_NUM As Long = 3 'mGridData(3) 馬連　-票数　（番号順）
Const ODDS_O3_NUM As Long = 4 'mGridData(4) ワイド-オッズ（番号順）
Const ODDS_H3_NUM As Long = 5 'mGridData(5) ワイド-票数　（番号順）
Const ODDS_O4_NUM As Long = 6 'mGridData(6) 馬単　-オッズ（番号順）
Const ODDS_H4_NUM As Long = 7 'mGridData(7) 馬単　-票数　（番号順）
Const ODDS_O5_NUM As Long = 8 'mGridData(8) 三連　-オッズ（番号順）
Const ODDS_H5_NUM As Long = 9 'mGridData(9) 三連　-票数　（番号順）
Const ODDS_O6_NUM As Long = 10 'mGridData(10) 三連単　-オッズ（番号順）
Const ODDS_H6_NUM As Long = 11 'mGridData(11) 三連単　-オッズ（番号順）

Const ODDS_O1_NIN As Long = 12 'mGridData(12) 単複枠-オッズ（人気順）
Const ODDS_H1_NIN As Long = 13 'mGridData(13) 単複枠-票数　（人気順）
Const ODDS_O2_NIN As Long = 14 'mGridData(14) 馬連　-オッズ（人気順）
Const ODDS_H2_NIN As Long = 15 'mGridData(15) 馬連　-票数　（人気順）
Const ODDS_O3_NIN As Long = 16 'mGridData(16) ワイド-オッズ（人気順）
Const ODDS_H3_NIN As Long = 17 'mGridData(17) ワイド-票数　（人気順）
Const ODDS_O4_NIN As Long = 18 'mGridData(18) 馬単　-オッズ（人気順）
Const ODDS_H4_NIN As Long = 19 'mGridData(19) 馬単　-票数　（人気順）
Const ODDS_O5_NIN As Long = 20 'mGridData(20) 三連　-オッズ（人気順）
Const ODDS_H5_NIN As Long = 21 'mGridData(21) 三連　-票数　（人気順）
Const ODDS_O6_NIN As Long = 22 'mGridData(22) 三連単　-オッズ（人気順）
Const ODDS_H6_NIN As Long = 23 'mGridData(23) 三連単　-票数　（人気順）


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   プロパティ
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'
'   機能: キー設定プロパティ
'
'   備考: Viewer必須プロパティ
'
Public Property Let OddsHyosuIsOdds(RHS As Boolean)
    mblnOddsHyosuIsOdds = RHS
End Property


'
'   機能: ラベル情報を取得するプロパティ
'
'   備考: なし
'
Public Property Get Labels(Index As Long) As String
    Labels = mstrLabels(Index)
End Property


'
'   機能: 票数情報を取得するプロパティ
'
'   備考: なし
'
Public Property Get Hyo(Index As Long) As String
    Hyo = mstrHyo(Index)
End Property


'
'   機能: 発表月日時分を取得するプロパティ
'
'   備考: なし
'
Public Property Get Happyo(Index As Long) As String
    Happyo = mstrHappyo(Index)
End Property


'
'   機能: グリッドデータを取得するプロパティ
'
'   備考: なし
'
Public Property Get GridDatas(Index As Long) As clsGridData
    Set GridDatas = mGridData(Index)
End Property


'
'   機能: グリッドデータを取得するプロパティ
'
'   備考: なし
'
Public Property Get GridisExist(Index As Long) As Boolean
    GridisExist = mGridisExist(Index)
End Property


'
'   機能: タイトル文字列
'
'   備考: なし
'
Public Property Get Title() As String
    Title = mstrTitle
End Property


'
'   機能: 出走頭数
'
'   備考: なし
'
Public Property Get TOSU() As Long
    TOSU = mlngTosu
End Property


'
'   機能: G1レースなら真
'
'   備考: なし
'
Public Property Get IsG1() As Boolean
    IsG1 = mblnIsG1
End Property

'
'   機能: 速報が受け取り可能なレースなら真
'
'   備考: なし
'
Public Property Get IsPrompt() As Boolean
    IsPrompt = mblnIsPrompt
End Property

'
'   機能: 3連単オッズがあれば真
'
'   備考: なし
'
Public Property Get O6Exist() As Boolean
    O6Exist = mblnO6Exist
End Property

'
'   機能: 3連単票数があれば真
'
'   備考: なし
'
Public Property Get H6Exist() As Boolean
    H6Exist = mblnH6Exist
End Property


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   外部関数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

'
'   機能: オッズ、票数データを取得する
'
'   備考: 初期表示用データの取得
'
Public Function Fetch(ByRef key As clsKeyRA) As Boolean
On Error GoTo ErrorHandler
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim strSQL As String
    Dim i      As Long              '' ループカウンタ
    Dim j      As Long              '' ループカウンタ
    Dim strTmp As String            '' 一時利用用文字列
    
    Set mKey = key
    
    For i = 0 To 19
        mGridisExist(i) = False
    Next i
    
    '始めにRACEからヘッダ情報を取得
    ' 太字レース名等グレーバー等文字列取得生成
    Set cn = gApp.GetCN_RACE
    strSQL = "SELECT * FROM RACE " & key.SQLWHEREString
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenKeyset, adLockReadOnly, adCmdText

    ' データが無かったらmsgを返す
    If rs.EOF Or rs.BOF Then
        rs.Close
        Set rs = Nothing
        Fetch = False
        RaiseEvent NoData
        Exit Function
    End If
    
    ' RACEレコードをすべて構造体で保持する
    Call SetDataFromRS_RA(rs, mBuf_RA)
    ' 該当のCC,TCがあるとき構造体に上書きする
    Call Set_CCTC
        
    ' タイトル文字列を作る
    mstrTitle = key.ReadableString _
                & mCC.YOBI7(rs("YoubiCD")) & " " _
                & mCC.KIBJ3(rs("JyoCD")) _
                & mSC.KN1(rs("Kaiji") & rs("Nichiji"))
                
    '頭数データの登録
    mlngTosu = mBuf_RA.TorokuTosu
    
    ' G1 J.G1 の場合、真
    mblnIsG1 = (mBuf_RA.GradeCD = "A" Or mBuf_RA.GradeCD = "F")
    
    ' 速報が受け取り可能な場合、真
    If (mBuf_RA.head.DataKubun >= "2" And mBuf_RA.head.DataKubun <= "6") Then
        mblnIsPrompt = True
    End If
    Call OpenHenkoRecordset
    
    ' レース情報を作る
    Call MakeData_Race
    
    freers mRS_TENKO_BABA
    freers mRS_BATAIJYU
    freers mRS_TORIKESI_JYOGAI
    freers mRS_KISHU_CHANGE
    freers mRS_MINING
    freers mRS_COURSE_CHANGE
    freers mRS_HASSOU_CHANGE

    rs.Close
    Set rs = Nothing
    
    '3連単データのあるなし、発売のあるなしを判定
    mblnO6Exist = Not (gApp.GetO6(key.str) = "" Or Mid$(gApp.GetO6(key.str), 40, 1) = "0")
    mblnH6Exist = Not (gApp.GetH6(key.str) = "" Or Mid$(gApp.GetH6(key.str), 32, 1) = "0")

    gApp.Log "OD Fetch Complete"
    
    Fetch = True
    Exit Function
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Function


'
'   機能: 非同期取得のトリガーを呼び出す
'
'   備考: なし
'
Public Sub AsyncFetch()
On Error GoTo ErrorHandler
    Dim strSQL As String
    
    '-----------------
    ' 非同期取得を行う
    '-----------------
    
    ' 非同期用コネクション生成
    Set mAsyncCN_UMA_RACE_A = gApp.GetCN_UMA_RACE_A
    
    'UMA_RACE_Aを開く
    strSQL = "SELECT * FROM UMA_RACE_A " & mKey.SQLWHEREString
    strSQL = strSQL & " ORDER BY Umaban"

    mAsyncCN_UMA_RACE_A.Execute strSQL, , adAsyncExecute

    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub
     

'
'   機能: キャンセル
'
'   備考: なし
'
Public Sub CancelFetching()
On Error GoTo ErrorHandler
    gApp.Log "DataOD Canceling"
    mblnCancelFetching = True

    freers mRS_TENKO_BABA
    freers mRS_BATAIJYU
    freers mRS_TORIKESI_JYOGAI
    freers mRS_KISHU_CHANGE
    freers mRS_MINING
    freers mRS_COURSE_CHANGE
    freers mRS_HASSOU_CHANGE
    
    freecn mAsyncCN_UMA_RACE_A
    
    Erase mGridData
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Sub
   
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部関数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: クラス初期化イベント
'
'   備考: なし
'
Private Sub Class_Initialize()
    Set mCC = New clsCodeConverter
    Set mSC = New clsStringConverter

End Sub

'
'   機能: クラス終了イベント
'
'   備考: なし
'
Private Sub Class_Terminate()
    Set mCC = Nothing
    Set mSC = Nothing
    
    Set mobj_HyoSanrenpuku = Nothing
    Set mobj_OddsSanrenpuku = Nothing
    Set mobj_HyoSanrentan = Nothing 'Sanrentan
    Set mobj_OddsSanrentan = Nothing 'Sanrentan
    Set mobj_Wakuban = Nothing
    
End Sub

'
'   機能: MakeData_Race前に各種変更情報テーブルを開く
'
'   備考: なし
'
Private Sub OpenHenkoRecordset()
    Set mRS_TENKO_BABA = New ADODB.Recordset
    Set mRS_BATAIJYU = New ADODB.Recordset
    Set mRS_TORIKESI_JYOGAI = New ADODB.Recordset
    Set mRS_KISHU_CHANGE = New ADODB.Recordset
    Set mRS_MINING = New ADODB.Recordset
    Set mRS_COURSE_CHANGE = New ADODB.Recordset
    Set mRS_HASSOU_CHANGE = New ADODB.Recordset

    Call OpenTableDirect(mRS_TENKO_BABA, gApp.GetCN_TENKO_BABA, "TENKO_BABA")
    Call OpenTableDirect(mRS_BATAIJYU, gApp.GetCN_BATAIJYU, "BATAIJYU")
    Call OpenTableDirect(mRS_TORIKESI_JYOGAI, gApp.GetCN_TORIKESI_JYOGAI, "TORIKESI_JYOGAI")
    Call OpenTableDirect(mRS_KISHU_CHANGE, gApp.GetCN_KISHU_CHANGE, "KISHU_CHANGE")
    Call OpenTableDirect(mRS_MINING, gApp.GetCN_MINING, "MINING")
    Call OpenTableDirect(mRS_COURSE_CHANGE, gApp.GetCN_COURSE_CHANGE, "COURSE_CHANGE")
    Call OpenTableDirect(mRS_HASSOU_CHANGE, gApp.GetCN_HASSOU_CHANGE, "HASSOU_CHANGE")
    
End Sub


'
'   機能: 共通部分、オッズ、票数の処理を呼び出す
'
'   備考: なし
'
Private Sub mAsyncCN_UMA_RACE_A_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
On Error GoTo ErrorHandler
    Dim i As Long
    
    gApp.Log "<mAsyncCN_UMA_RACE_A_ExecuteComplete"
    DoEvents
    If mblnCancelFetching Then
        Exit Sub
    End If
    
    If pError Is Nothing Then
        
        If pRecordset.EOF Then
            RaiseEvent NoUMARACE
        Else
            DoEvents
            If mblnCancelFetching Then Exit Sub
            
            ' テーブルを開き共通項目の設定(UMA_RACE_B,UMA,KISHU,TORIKESI_JYOGAI,KISHU_CHANGE)
            Call set_common(pRecordset) '馬番なしの評価必要
            If mblnCancelFetching Then Exit Sub
            
            If mblnOddsHyosuIsOdds Then
                'オッズを設定
                Call set_Odds
            Else
                '票数を設定
                Call set_Hyo
            End If
            If mblnCancelFetching Then Exit Sub
            
            RaiseEvent FetchedAll
        End If
    Else
        With pError
            gApp.Log .Description & .SQLState & .Source & .Number
        End With
    End If

    freers mRS_TENKO_BABA
    freers mRS_BATAIJYU
    freers mRS_TORIKESI_JYOGAI
    freers mRS_KISHU_CHANGE
    freers mRS_MINING
    freers mRS_COURSE_CHANGE
    freers mRS_HASSOU_CHANGE
    
    freecn mAsyncCN_UMA_RACE_A

    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub

    
'
'   機能: レース情報を作る
'
'   備考: なし
'
Private Sub MakeData_Race()
On Error GoTo ErrorHandler
    Dim str As String               '' 作業用文字列変数
    Dim i   As Long                 '' ループカウンタ
    Dim rs  As ADODB.Recordset
    Dim blnHenkouExist As Boolean   '' 変更情報の有無フラグ
    Dim strTencoCD As String
    Dim strSibaBabaCD As String
    Dim strDirtBabaCD As String
    
    Dim varField1 As Variant
    Dim varKey1 As Variant
    Dim varField2 As Variant
    Dim varKey2 As Variant

    varField1 = Array("Year", "MonthDay", "JyoCD", "Kaiji", "Nichiji")
    With mKey
        varKey1 = Array(.Year, .MonthDay, .JyoCD, .Kaiji, .Nichiji)
    End With
    varField2 = Array("Year", "MonthDay", "JyoCD", "Kaiji", "Nichiji", "RaceNum")
    With mKey
        varKey2 = Array(.Year, .MonthDay, .JyoCD, .Kaiji, .Nichiji, .RaceNum)
    End With
    
    With mBuf_RA
        
        ' 太字レース名等グレーバー　文字列生成
        str = ""
        str = str & IIf(.id.JyoCD > "10", mCC.KIBJ1(.id.JyoCD), mCC.KIBJ3(.id.JyoCD))           '' 競馬場名
        str = str & IIf(.id.RaceNum = "00", "", val(.id.RaceNum) & "R")         '' nR
        str = str & " "                                 '' 空白
        With .RaceInfo
            If .Nkai <> "000" Then
                str = str & "第" & val(.Nkai) & "回"    '' 第n回
            End If
        End With
        str = str & " "                                 '' 空白
        str = str & Trim$(.RaceInfo.Hondai)             '' 本題
        str = str & mCC.GRAD2(.GradeCD)                 '' グレードコード
        mstrLabels(0) = str
    
        ' データ区分
        mstrLabels(1) = mSC.RA_DKBN(.head.DataKubun)
        ' 変更情報の有無の表示
        blnHenkouExist = False
        If .head.DataKubun >= "2" And .head.DataKubun <= "6" Then
        
            '' 天候変更は条件に含まない
            
            '' 馬体重変更は条件に含まない
            
            ' 取り消し除外情報の有無確認
            Call SafeSeek(mRS_TORIKESI_JYOGAI, varField2, varKey2)
            blnHenkouExist = blnHenkouExist Or Not mRS_TORIKESI_JYOGAI.EOF
            
            ' 騎手変更情報の有無確認
            Call SafeSeek(mRS_KISHU_CHANGE, varField2, varKey2)
            blnHenkouExist = blnHenkouExist Or Not mRS_KISHU_CHANGE.EOF
            
            '' マイニングは条件に含まない
            
            ' コース変更情報の有無確認
            Call SafeSeek(mRS_COURSE_CHANGE, varField2, varKey2)
            blnHenkouExist = blnHenkouExist Or Not mRS_COURSE_CHANGE.EOF
            
            ' 発走時間変更情報の有無確認
            Call SafeSeek(mRS_HASSOU_CHANGE, varField2, varKey2)
            blnHenkouExist = blnHenkouExist Or Not mRS_HASSOU_CHANGE.EOF
            
            If blnHenkouExist Then
                mstrLabels(1) = mstrLabels(1) & " 変更情報あり"
            End If
        End If
        
        ' ヘッダ情報　文字列生成
        
        ' 年月日曜
        str = ""
        str = str & mSC.YMD2(.id.Year & .id.MonthDay) & mCC.YOBI7(.RaceInfo.YoubiCD) & "  "
        mstrLabels(2) = str
        
        ' 発走時間
        str = ""
        If .HassoTime <> "0000" Then
            str = str & "発走 " & mSC.HHNN2(.HassoTime) & "  "
        Else
            str = str & "            "
        End If
        
        ' 競走条件
        With .JyokenInfo
            str = str & mCC.KSSB4(.SyubetuCD) & " " & mCC.KSJK1(.JyokenCD(4))
            str = str & "　　"
            str = str & mCC.KSKG1(.KigoCD, mBuf_RA.id.Year)
            str = str & "　　"
            str = str & mCC.JRSB1(.JyuryoCD)
        End With
        mstrLabels(3) = str
        
        ' コース区分
        str = ""
        If .CourseKubunCD <> "  " Then
            str = str & Trim$(.CourseKubunCD) & " コース"
        Else
            str = str & "        "
        End If
        str = str & " "
        
        ' トラック＆距離
        str = str & mCC.TRCK2(.TrackCD)
        str = str & mSC.FitSpaceNum(.KYORI, 4) & "m  "
        mstrLabels(4) = str
        
        ' 天候馬場状態
        str = ""
        ' まずRAから取得
        With .TenkoBaba
            strTencoCD = .TenkoCD
            strSibaBabaCD = .SibaBabaCD
            strDirtBabaCD = .DirtBabaCD
        End With
        If mBuf_RA.head.DataKubun <= "2" Then ' 3 : 速報成績が出た時点で確定する為
            ' 変更情報を配慮した表示
            Call SafeSeek(mRS_TENKO_BABA, varField1, varKey1)
            Set rs = mRS_TENKO_BABA
            If Not rs.EOF Then
                ' WEで上書き
                Do While Not rs.EOF
                    If rs("Year") <> mKey.Year _
                    Or rs("MonthDay") <> mKey.MonthDay _
                    Or rs("JyoCD") <> mKey.JyoCD _
                    Or rs("Kaiji") <> mKey.Kaiji _
                    Or rs("Nichiji") <> mKey.Nichiji Then
                        ' 違う日付になったら終了
                        Exit Do
                    End If
                    Select Case rs("HenkoID")
                    '天候馬場初期状態
                    Case "1"
                        strTencoCD = rs("AtoTenkoCD")
                        strSibaBabaCD = rs("AtoSibaBabaCD")
                        strDirtBabaCD = rs("AtoDirtBabaCD")
                                        
                    Case "2"
                        strTencoCD = rs("AtoTenkoCD")
                    Case "3"
                        strSibaBabaCD = rs("AtoSibaBabaCD")
                        strDirtBabaCD = rs("AtoDirtBabaCD")
                    End Select
                    rs.MoveNext
                Loop
            End If
        End If
                
        str = str & mCC.TNKO1(strTencoCD)
        If strSibaBabaCD <> "0" Then
            str = str & " 芝:" & mCC.BBJT1(strSibaBabaCD)
        End If
        If strDirtBabaCD <> "0" Then
            str = str & " ダート:" & mCC.BBJT1(strDirtBabaCD)
        End If
        str = str & " "
        mstrLabels(5) = str
        
        ' 出走頭数
        str = ""
        If .SyussoTosu <> "00" Then
            str = str & "出走頭数 " & Right$("  " & val(.SyussoTosu), 2) & "頭"
        Else
            str = str & "             "
        End If
        str = str & "  "

        ' 本賞金
        If .Honsyokin(0) <> "00000000" Then
            str = str & "本賞金  "
            For i = 0 To 6
                If .Honsyokin(i) <> "00000000" Then
                    str = str & "  " & Format$(Int(.Honsyokin(i) / 10) / 10, "0.0")
                    If Right$(str, 2) = ".0" Then
                        str = Left$(str, Len(str) - 2)
                    End If
                End If
            Next i
            str = str & " 万円"
        End If
        str = str & "  "
        
'        ' レコード

        mstrLabels(6) = str

        ' レコード
        mstrLabels(9) = mSC.RA_RKBN(.RecordUpKubun)

        
        ' 入線頭数
        str = ""
        If .NyusenTosu <> "00" Then
            str = str & "入線頭数 " & Right$("  " & val(.NyusenTosu), 2) & "頭"
        Else
            str = str & "             "
        End If
        str = str & "  "

        ' 付加賞金
        If .Fukasyokin(0) <> "00000000" Then
            str = str & "付加賞金"
            For i = 0 To 4
                If .Fukasyokin(i) <> "00000000" Then
                    str = str & "  " & Format$(Int(.Fukasyokin(i) / 10) / 10, "0.0")
                    If Right$(str, 2) = ".0" Then
                        str = Left$(str, Len(str) - 2)
                    End If
                End If
            Next i
            str = str & " 万円"
        End If
        str = str & "  "
        mstrLabels(7) = str
        
        ' データ作成年月日
        str = ""
        With .head.MakeDate
            str = str & mSC.YMD1(.Year & .Month & .Day) & "作成データ"
        End With
        mstrLabels(8) = str
    End With
    Exit Sub
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Sub
    

'
'   機能: 共通項目設定
'
'   備考: なし
'
Private Function set_common(rs As ADODB.Recordset) As Boolean
On Error GoTo ErrorHandler
    Dim i As Long
    Dim j As Long
    Dim lngGrid As Long            '' セル入力用
    Dim lngCP        As Long            '' セル入力用
    Dim lngRP        As Long            '' セル入力用
    Dim strSQL As String
    
    Dim UB As ADODB.Recordset
    Dim UM As ADODB.Recordset
    Dim AV As ADODB.Recordset
    Dim KS As ADODB.Recordset
    Dim KC As ADODB.Recordset
    Dim dicKC As Scripting.Dictionary '騎手変更の発表タイム保存用
    
    set_common = False
    
    'UMA_RACE_Bを開く
    Set UB = New ADODB.Recordset
    Call OpenTableDirect(UB, gApp.GetCN_UMA_RACE_B, "UMA_RACE_B")

    'UMAを開く
    Set UM = New ADODB.Recordset
    Call OpenTableDirect(UM, gApp.GetCN_UMA, "UMA")
    
    'KISHUを開く
    Set KS = New ADODB.Recordset
    Call OpenTableDirect(KS, gApp.GetCN_KISHU, "KISHU")
    
    'TORIKESI_JYOGAI (出走取消・競走除外)を開く
    Set AV = New ADODB.Recordset
    Call OpenTableDirect(AV, gApp.GetCN_TORIKESI_JYOGAI, "TORIKESI_JYOGAI")

    'KISHU_CHANGE (騎手変更)を開く
    strSQL = "SELECT * FROM KISHU_CHANGE " & mKey.SQLWHEREString & " ORDER BY HappyoTime ;"
    Set KC = New ADODB.Recordset
    KC.Open strSQL, gApp.GetCN_KISHU_CHANGE, adOpenKeyset, adLockReadOnly, adCmdText
    
    ' 騎手変更情報用
    Set dicKC = New Scripting.Dictionary
    If Not KC.EOF Then
        KC.MoveFirst
        Do While Not KC.EOF
            If Not (KC("Year") = mKey.Year _
            And KC("MonthDay") = mKey.MonthDay _
            And KC("JyoCD") = mKey.JyoCD _
            And KC("Kaiji") = mKey.Kaiji _
            And KC("Nichiji") = mKey.Nichiji _
            And KC("RaceNum") = mKey.RaceNum) Then
                Exit Do
            End If
            If dicKC.Exists(KC("Umaban").value) Then
                Call dicKC.Remove(KC("Umaban").value)
            End If
            dicKC.Add KC("Umaban").value, KC("HappyoTime").value
            KC.MoveNext
        Loop
        KC.MoveFirst
    End If
    KC.Close
    Call OpenTableDirect(KC, gApp.GetCN_KISHU_CHANGE, "KISHU_CHANGE")
    
    Set mobj_Wakuban = New Scripting.Dictionary
    ' グリッドデータオブジェクトを生成する
    Set mGridData(0) = New clsGridData
    
    With mGridData(0)
        .Cols = 16
        .Rows = IIf(mlngTosu < 13, 13, mlngTosu + 1)
        'カラムテキスト
        lngRP = 0
        .SetItemMatrix lngRP, 0, "枠", , , , , ColorODBack1, ColorODFore0
        .SetItemMatrix lngRP, 1, "番", , , , , ColorODBack1, ColorODFore0
        .SetItemMatrix lngRP, 2, "馬名", , , , , ColorODBack1, ColorODFore0
        .SetItemMatrix lngRP, 3, "性齢", , , , , ColorODBack1, ColorODFore0
        .SetItemMatrix lngRP, 4, "騎手名", , , , , ColorODBack1, ColorODFore0
        .SetItemMatrix lngRP, 5, "単勝オッズ", , , , , ColorODBack1, ColorODFore0
        .SetItemMatrix lngRP, 6, "複勝オッズ", , , , , ColorODBack1, ColorODFore0
        .SetItemMatrix lngRP, 7, , , , , , ColorODBack0
        For j = 0 To 7
            .SetItemMatrix lngRP, 8 + j, "枠連オッズ", , , , , ColorODBack1, ColorODFore0
        Next j
        
        rs.MoveFirst
        j = 0
        Do While Not rs.EOF
            lngRP = j + 1
            .SetItemMatrix lngRP, 0, mSC.FitSpaceNum(rs("Wakuban"), 2), , "^-", , , gApp.GetWakubanColor(rs("Wakuban")), Contrast(gApp.GetWakubanColor(rs("Wakuban")))
            .SetItemMatrix lngRP, 1, mSC.FitSpaceNum(rs("Umaban"), 2), , ">-"
            Call SafeSeek(UM, Array("KettoNum"), Array(rs("KettoNum").value))
            Call SafeSeek(UB, Array("B_Year", "B_MonthDay", "B_JyoCD", "B_RaceNum", "B_Umaban", "B_KettoNum"), _
                        Array(rs("Year").value, rs("MonthDay").value, rs("JyoCD").value, rs("RaceNum").value, rs("Umaban").value, rs("KettoNum").value))
            Call SafeSeek(AV, Array("Year", "MonthDay", "JyoCD", "Kaiji", "Nichiji", "RaceNum", "Umaban"), _
                Array(rs("Year").value, rs("MonthDay").value, rs("JyoCD").value, rs("Kaiji").value, rs("Nichiji").value, rs("RaceNum").value, rs("Umaban").value))
            If rs("DataKubun") = 2 And Not AV.EOF Then
                .SetItemMatrix lngRP, 2, Trim$(rs("Bamei")), mSC.AV_DKBN(AV("DataKubun")) & "(" & mSC.AV_JKBN(AV("JiyuKubun")) & ")", "<-", "UM", IIf(UM.EOF, "", rs("KettoNum")), , &HFF, True
            Else
                .SetItemMatrix lngRP, 2, Trim$(rs("Bamei")), mCC.IJKB1(UB("IJyoCD")), "<-", "UM", IIf(UM.EOF, "", rs("KettoNum")), , , IIf(UB("IjyoCD") >= "1" And UB("IjyoCD") <= "3", True, False)
            End If
            .SetItemMatrix lngRP, 3, mCC.SEIB4(rs("SexCD")) & val(rs("Barei")), , ">-"
            If Not UB.EOF Then
                If rs("DataKubun") = "2" Then
                    ' 騎手変更情報
                    If KC.EOF Or Not dicKC.Exists(rs("Umaban").value) Then
                        .SetItemMatrix lngRP, 4, Trim$(UB("KisyuRyakusyo")), "", "<-", "KS", UB("KisyuCode")
                    Else
                        SafeSeek KC, Array("Year", "MonthDay", "JyoCD", "Kaiji", "Nichiji", "RaceNum", "HappyoTime", "Umaban"), _
                                    Array(rs("Year").value, rs("MonthDay").value, rs("JyoCD").value, rs("Kaiji").value, rs("Nichiji").value, rs("RaceNum").value, dicKC.item(rs("Umaban").value), rs("Umaban").value)
                        SafeSeek KS, Array("KisyuCode"), _
                                    Array(KC("AtoKisyuCode").value)
                        If Trim$(IfExist(KS, "KisyuRyakusyo")) <> "" Then '騎手変更名を略称で表示(04-033-4)
                            .SetItemMatrix lngRP, 4, Trim$(KS("KisyuRyakusyo")), Trim$(KC("MaeKisyuName")) & "→" & Trim$(KC("AtoKisyuName")) _
                                & "", "<-", "KS", IfExist(KS, "KisyuCode"), , &HFF
                        Else
                            .SetItemMatrix lngRP, 4, Trim$(KC("AtoKisyuName")), Trim$(KC("MaeKisyuName")) & "→" & Trim$(KC("AtoKisyuName")) _
                                & "", "<-", "KS", IfExist(KS, "KisyuCode"), , &HFF
                        End If
                    End If
                Else
                    If val(UB("KisyuCodeBefore")) = 0 Then
                        .SetItemMatrix lngRP, 4, Trim$(UB("KisyuRyakusyo")), "", "<-", "KS", UB("KisyuCode")
                    Else
                        .SetItemMatrix lngRP, 4, Trim$(UB("KisyuRyakusyo")), Trim$(UB("KisyuRyakusyoBefore")) & "→" & Trim$(UB("KisyuRyakusyo")) _
                            & "", "<-", "KS", UB("KisyuCode"), , &HFF
                    End If
                End If
            End If
            If (val(rs("Umaban").value) = 0) Then
                RaiseEvent NoUMARACE
                Call CancelFetching
                Exit Function
            End If
            Call mobj_Wakuban.Add(val(rs("Umaban").value), rs("Wakuban").value)
            rs.MoveNext
            j = j + 1
        Loop
    End With
    
    '単複枠-票数　（番号順）グリッドに共通部分をコピーする
    Set mGridData(1) = mGridData(0).Clone
    With mGridData(1) 'ラベル上書き（票数）
        lngRP = 0
        .SetItemMatrix lngRP, 5, "単勝票数", , , , , ColorODBack1, ColorODFore0
        .SetItemMatrix lngRP, 6, "複勝票数", , , , , ColorODBack1, ColorODFore0
        .SetItemMatrix lngRP, 7, , , , , , ColorODBack0
        For j = 0 To 7
            .SetItemMatrix lngRP, 8 + j, "枠連票数", , , , , ColorODBack1, ColorODFore0
        Next j
    End With
    
    AV.Close
    KC.Close
    rs.Close
    UB.Close
    UM.Close
    KS.Close
    Set dicKC = Nothing
    
    set_common = True
    
    Exit Function
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Function
    
    
'
'   機能: オッズ設定
'
'   備考: なし
'
Private Sub set_Odds()
On Error GoTo ErrorHandler

    '単複枠
    Set mobj_OddsTansyo = New Scripting.Dictionary
    Set mobj_OddsFukusyo = New Scripting.Dictionary
    Set mobj_OddsWakuren = New Scripting.Dictionary
    Call set_O1(mKey)
    Set mobj_OddsTansyo = Nothing
    Set mobj_OddsFukusyo = Nothing
    Set mobj_OddsWakuren = Nothing
    DoEvents
    If mblnCancelFetching Then Exit Sub
    
    '馬連
    Set mobj_OddsUmaren = New Scripting.Dictionary
    Call set_O2(mKey)
    Set mobj_OddsUmaren = Nothing
    DoEvents
    If mblnCancelFetching Then Exit Sub
    
    'ワイド
    Set mobj_OddsWide = New Scripting.Dictionary
    Call set_O3(mKey)
    Set mobj_OddsWide = Nothing
    DoEvents
    If mblnCancelFetching Then Exit Sub
    
    '馬単
    Set mobj_OddsUmatan = New Scripting.Dictionary
    Call set_O4(mKey)
    Set mobj_OddsUmatan = Nothing
    DoEvents
    If mblnCancelFetching Then Exit Sub
    
    '3連複
    Set mobj_OddsSanrenpuku = New Scripting.Dictionary 'Class_TerminateでNothing
    Call set_O5(mKey, 1) 'Axis
    DoEvents
    If mblnCancelFetching Then Exit Sub

    '3連単
    Set mobj_OddsSanrentan = New Scripting.Dictionary 'Class_TerminateでNothing
    Call set_O6(mKey.str, 1) 'Axis
    DoEvents
    If mblnCancelFetching Then Exit Sub
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Sub
        
'
'   機能: データ区分と発表タイムから、ヘッダの区分表示文字列を生成する
'
'   備考: なし
'
Private Function KubunAndHappyoStr(Kubun As String, HappyoTime As String) As String
On Error GoTo ErrorHandler
    Dim out As String
    
    Select Case Kubun
    Case 1
        out = "データ: " & mSC.MDHM(HappyoTime) & " 現在"
    Case 2
        out = "前売り最終"
    Case 3
        out = "最終"
    Case 4, 5
        out = "確定"
    Case 9
        out = "中止"
    Case Else
        out = ""
    End Select

    KubunAndHappyoStr = out
    Exit Function
ErrorHandler:
    gApp.ErrLog
End Function


'
'   機能: オッズ設定
'
'   備考: なし
'
Private Function set_Hyo() As Boolean
On Error GoTo ErrorHandler
    Dim i As Long
    Dim strItem As String            '' 一時利用用文字列
    
    DoEvents
    If mblnCancelFetching Then Exit Function

    strItem = gApp.GetH1(mKey.str)

    ' データが無かったらFalseを返して終了
    If strItem = "" Then
        set_Hyo = False
        Exit Function
    End If
    DoEvents
    
    Call SetData_H1_(strItem, mbuf_H1)
    DoEvents
    If mblnCancelFetching Then Exit Function
        
    ' 各tab 総票数設定
    mstrHyo(0) = IIf(val(mbuf_H1.HyoTotal(0)) = 0, "", "単勝：" & val(mbuf_H1.HyoTotal(0)) & "票")
    mstrHyo(1) = IIf(val(mbuf_H1.HyoTotal(1)) = 0, "", "複勝：" & val(mbuf_H1.HyoTotal(1)) & "票")
    mstrHyo(2) = IIf(val(mbuf_H1.HyoTotal(2)) = 0, "", "枠連：" & val(mbuf_H1.HyoTotal(2)) & "票")
    mstrHyo(3) = IIf(val(mbuf_H1.HyoTotal(3)) = 0, "", "馬連：" & val(mbuf_H1.HyoTotal(3)) & "票")
    mstrHyo(4) = IIf(val(mbuf_H1.HyoTotal(4)) = 0, "", "ワイド：" & val(mbuf_H1.HyoTotal(4)) & "票")
    mstrHyo(5) = IIf(val(mbuf_H1.HyoTotal(5)) = 0, "", "馬単：" & val(mbuf_H1.HyoTotal(5)) & "票")
    mstrHyo(6) = IIf(val(mbuf_H1.HyoTotal(6)) = 0, "", "３連複：" & val(mbuf_H1.HyoTotal(6)) & "票")
    
    '票数データ区分
    mstrHappyo(6) = KubunAndHappyoStr(mbuf_H1.head.DataKubun, "") '
    
    
    RaiseEvent FetchedH1
    
    With mbuf_H1
        '単複枠
        If .HatubaiFlag(0) <> 0 Or .HatubaiFlag(1) <> 0 Or .HatubaiFlag(2) <> 0 Then
            Set mobj_HyoTansyo = New Scripting.Dictionary
            Set mobj_HyoFukusyo = New Scripting.Dictionary
            Set mobj_HyoWakuren = New Scripting.Dictionary
            Call set_H1(Mid$(strItem, 1, 1464))
            Set mobj_HyoTansyo = Nothing
            Set mobj_HyoFukusyo = Nothing
            Set mobj_HyoWakuren = Nothing
        End If
        
        DoEvents
        If mblnCancelFetching Then Exit Function
        '馬連
        If .HatubaiFlag(3) <> 0 Then
            Set mobj_HyoUmaren = New Scripting.Dictionary
            Call set_H2(Mid$(strItem, 1464, 2754))
            Set mobj_HyoUmaren = Nothing
        End If
        
        DoEvents
        If mblnCancelFetching Then Exit Function
        'ワイド
        If .HatubaiFlag(4) <> 0 Then
            Set mobj_HyoWide = New Scripting.Dictionary
            Call set_H3(Mid$(strItem, 4218, 2754))
            Set mobj_HyoWide = Nothing
        End If
        
        DoEvents
        If mblnCancelFetching Then Exit Function
        '馬単
        If .HatubaiFlag(5) <> 0 Then
            Set mobj_HyoUmatan = New Scripting.Dictionary
            Call set_H4(Mid$(strItem, 6972, 5508))
            Set mobj_HyoUmatan = Nothing
        End If
    
        DoEvents
        If mblnCancelFetching Then Exit Function
        '3連複
        If .HatubaiFlag(6) <> 0 Then
            Set mobj_HyoSanrenpuku = New Scripting.Dictionary 'Class_TerminateでNothing
            Call set_H5(Mid$(strItem, 12480, 16320), 1)  '
        End If
        
        DoEvents
        If mblnCancelFetching Then Exit Function
        '3連単
        Set mobj_HyoSanrentan = New Scripting.Dictionary 'Class_TerminateでNothing
        Call set_H6(mKey.str, 1) '
    End With

    Exit Function
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Function
    
    
'
'   機能: 番号順 ３連複オッズデータを再取得する（軸ボタン用）
'
'   備考: なし
'
Public Function FetchSanrenOddsNum(ByVal lngAxis As Long) As Boolean
On Error GoTo ErrorHandler
    Dim strItem As String
    Dim i As Long
    Dim j As Long
    Dim k(1 To 18) As Long
    Dim colOffset As Long
    Dim rowOffset As Long
    Dim lngCP        As Long            '' セル入力用
    Dim lngRP        As Long            '' セル入力用
    Dim lngGrid As Long
        
    rowOffset = 0
    colOffset = 0
        
    For i = 1 To 18
        k(i) = IIf(i < lngAxis, i, i + 1)
    Next i
    
    lngGrid = ODDS_O5_NUM 'mGridData(8) 三連　-オッズ（番号順）
    With mGridData(lngGrid)
        For i = 1 To 16
            lngCP = ColPosition10(rowOffset, colOffset, i, j) - 1 '列index番号呼び出し
            'カラムインデックス
            lngRP = RowPosition10(rowOffset, colOffset, i, i)
            .SetItemMatrix lngRP, lngCP + 0, "", , "^-", , , ColorODBack0, ColorODFore0
            .SetItemMatrix lngRP, lngCP + 1, k(i) + 0, , "^-", , , gApp.GetWakubanColor(mobj_Wakuban(k(i))), Contrast(gApp.GetWakubanColor(mobj_Wakuban(k(i)))) '
            For j = i + 1 To 17
                lngRP = RowPosition10(rowOffset, colOffset, i, j)
                .SetItemMatrix lngRP, lngCP + 0, k(j) + 0, , ">-", , , ColorODBack1, ColorODFore0
                strItem = OddsConv(mobj_OddsSanrenpuku.item(SanrenDicKey(lngAxis, k(i), k(j))))
                .SetItemMatrix lngRP, lngCP + 1, strItem, , ">-", , , , ColorHML(strItem)
            Next j
        Next i
    End With
    
    Exit Function
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Function
    
    
'
'   機能: 番号順 ３連単オッズデータを再取得する（軸ボタン用）
'
'   備考: なし
'
Public Function FetchSanrentanOddsNum(ByVal lngAxis As Long) As Boolean
On Error GoTo ErrorHandler
    Dim strItem As String
    Dim i As Long
    Dim j As Long
    Dim k(1 To 18) As Long
    Dim colOffset As Long
    Dim rowOffset As Long
    Dim lngCP        As Long            '' セル入力用
    Dim lngRP        As Long            '' セル入力用
    Dim lngGrid As Long
        
    rowOffset = 0
    colOffset = 0
        
    For i = 1 To 18
        k(i) = IIf(i <= lngAxis, i, i - 1)
    Next i
    
    lngGrid = ODDS_O6_NUM  'mGridData(10) 三連単　-オッズ（番号順）
    ' グリッドデータオブジェクトを生成する
    Set mGridData(lngGrid) = New clsGridData
    With mGridData(lngGrid)
        .Cols = (18 - 1) * 2
        .Rows = 18
        For i = 1 To 18
            DoEvents
            lngCP = (k(i) - 1) * 2
            If lngAxis <> i Then '軸ボタンと同じ数の列は処理しない
                'カラムインデックス
                lngRP = rowOffset
                .SetItemMatrix lngRP, lngCP + 0, "", , ">-", , , ColorODBack0, ColorODFore0
                .SetItemMatrix lngRP, lngCP + 1, i + 0, , "^-", , , gApp.GetWakubanColor(mobj_Wakuban(i)), Contrast(gApp.GetWakubanColor(mobj_Wakuban(i))) '
                
                For j = 1 To 18
                    lngRP = k(j)
                    If lngAxis <> j Then '軸ボタンと同じ数の行は処理しない
                        'ロウインデックス
                        .SetItemMatrix lngRP, lngCP + 0, j + 0, , ">-", , , ColorODBack1, ColorODFore0
                        If i <> j Then  '2着3着が同じ数字の場合値を入れない
                            strItem = OddsConv(mobj_OddsSanrentan.item("B" & Format$(lngAxis * 10000 + i * 100 + j, "000000")))
                            .SetItemMatrix lngRP, lngCP + 1, strItem, , ">-", , , , ColorHML(strItem)
                        End If
                    End If
                Next j
            End If
        Next i
    End With
    
    Exit Function
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Function
    
    
'
'   機能: 番号順 ３連複票数データを再取得する（軸ボタン用）
'
'   備考: なし
'
Public Function FetchSanrenHyoNum(ByVal lngAxis As Long) As Boolean
On Error GoTo ErrorHandler
    Dim strItem As String
    Dim i As Long
    Dim j As Long
    Dim k(1 To 18) As Long
    Dim colOffset As Long
    Dim rowOffset As Long
    Dim lngCP        As Long            '' セル入力用
    Dim lngRP        As Long            '' セル入力用
    Dim lngGrid As Long
        
    rowOffset = 0
    colOffset = 0
        
    For i = 1 To 18
        k(i) = IIf(i < lngAxis, i, i + 1)
    Next i
    
    lngGrid = ODDS_H5_NUM 'mGridData(9) 三連　-票数　（番号順）
    With mGridData(lngGrid)
        For i = 1 To 16
            lngCP = ColPosition10(rowOffset, colOffset, i, j) - 1 '列index番号呼び出し
            'カラムインデックス
            lngRP = RowPosition10(rowOffset, colOffset, i, i)
            .SetItemMatrix lngRP, lngCP + 0, "", , "^-", , , ColorODBack0, ColorODFore0
            .SetItemMatrix lngRP, lngCP + 1, k(i) + 0, , "^-", , , gApp.GetWakubanColor(mobj_Wakuban(k(i))), Contrast(gApp.GetWakubanColor(mobj_Wakuban(k(i)))) '
            For j = i + 1 To 17
                lngRP = RowPosition10(rowOffset, colOffset, i, j)
                .SetItemMatrix lngRP, lngCP + 0, k(j) + 0, , ">-", , , ColorODBack1, ColorODFore0
                .SetItemMatrix lngRP, lngCP + 1, HyoConv(mobj_HyoSanrenpuku.item(SanrenDicKey(lngAxis, k(i), k(j)))), , ">-", , , ColorODBack0, ColorODForeL
            Next j
        Next i
    End With
    
    Exit Function
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Function
    
    
'
'   機能: 番号順 ３連単票数データを再取得する（軸ボタン用）
'
'   備考: なし
'
Public Function FetchSanrentanHyoNum(ByVal lngAxis As Long) As Boolean
On Error GoTo ErrorHandler
    Dim strItem As String
    Dim i As Long
    Dim j As Long
    Dim k(1 To 18) As Long
    Dim colOffset As Long
    Dim rowOffset As Long
    Dim lngCP        As Long            '' セル入力用
    Dim lngRP        As Long            '' セル入力用
    Dim lngGrid As Long
        
    rowOffset = 0
    colOffset = 0
        
    For i = 1 To 18
        k(i) = IIf(i <= lngAxis, i, i - 1)
    Next i
    
    lngGrid = ODDS_H6_NUM 'mGridData(11) 三連単　-票数　（番号順）
    ' グリッドデータオブジェクトを生成する
    Set mGridData(lngGrid) = New clsGridData
    With mGridData(lngGrid)
        .Cols = (18 - 1) * 2
        .Rows = 18
        For i = 1 To 18
            DoEvents
            lngCP = (k(i) - 1) * 2
            If lngAxis <> i Then '軸ボタンと同じ数の列は処理しない
                'カラムインデックス
                lngRP = rowOffset
                .SetItemMatrix lngRP, lngCP + 0, "", , ">-", , , ColorODBack0, ColorODFore0
                .SetItemMatrix lngRP, lngCP + 1, i + 0, , "^-", , , gApp.GetWakubanColor(mobj_Wakuban(i)), Contrast(gApp.GetWakubanColor(mobj_Wakuban(i))) '
                
                For j = 1 To 18
                    lngRP = k(j)
                    If lngAxis <> j Then '軸ボタンと同じ数の行は処理しない
                        .SetItemMatrix lngRP, lngCP + 0, j + 0, , ">-", , , ColorODBack1, ColorODFore0
                        If i <> j Then  '2着3着が同じ数字の場合値を入れない
                            strItem = HyoConv(mobj_HyoSanrentan.item("B" & Format$(lngAxis * 10000 + i * 100 + j, "000000")))
                            .SetItemMatrix lngRP, lngCP + 1, strItem, , ">-", , , ColorODBack0, ColorODForeL
                        End If
                    End If
                Next j
            End If
        Next i
    End With
    
    Exit Function
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Function


'
'   機能: 票数-単複枠用グリッドを作る
'
'   備考: 文字列をディクショナリに読み込み、ディクショナリからグリッドを作る
'
Private Function set_H1(buf As String) As Boolean
On Error GoTo ErrorHandler
    Dim strItem As String
    Dim strValue As String
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim colOffset As Long
    Dim rowOffset As Long
    Dim lngCP        As Long            '' セル入力用
    Dim lngRP        As Long            '' セル入力用
    Dim lngGrid As Long
    
    rowOffset = 0
    colOffset = 0
    
    '発売フラグを参照し、発売があるならデータをディクショナリ登録
    With mbuf_H1
        If .HatubaiFlag(0) <> "0" Then
            Call DataToDic(mobj_HyoTansyo, Mid$(buf, 84, 504 - 84), 2, 11, 2)   '単勝
        End If
        DoEvents
        If .HatubaiFlag(1) <> "0" Then
            Call DataToDic(mobj_HyoFukusyo, Mid$(buf, 504, 924 - 504), 2, 11, 2)   '複勝
        End If
        DoEvents
        If .HatubaiFlag(2) <> "0" Then
            Call DataToDic(mobj_HyoWakuren, Mid$(buf, 924, 1464 - 924), 2, 11, 2)  '枠連
        End If
    End With

    lngGrid = ODDS_H1_NUM 'mGridData(1) 単複枠-票数　（番号順）
    With mGridData(lngGrid)
        For i = 0 To mlngTosu - 1
            DoEvents
            '単勝
            strItem = HyoConv(mobj_HyoTansyo.item("B" & Right$(("0" & i + 1), 2)))
            .SetItemMatrix i + 1, 5, strItem, , ">-"
            '複勝
            strItem = HyoConv(mobj_HyoFukusyo.item("B" & Right$(("0" & i + 1), 2)))
            .SetItemMatrix i + 1, 6, strItem, , ">-"
        Next i
        
        '枠連
        rowOffset = 1
        colOffset = 8

        For i = 1 To 8
            DoEvents
            lngCP = ColPosition4(rowOffset, colOffset, i, j) - 1
            'カラムインデックス
            lngRP = RowPosition4(rowOffset, colOffset, i, i - 1)
            .SetItemMatrix lngRP, lngCP + 0, "", , ">-", , , ColorODBack0, ColorODFore0
            .SetItemMatrix lngRP, lngCP + 1, i & "枠", , "^-", , , gApp.GetWakubanColor(i), Contrast(gApp.GetWakubanColor(i))
            For j = i To 8 '枠番、票数書き込み
                lngRP = RowPosition4(rowOffset, colOffset, i, j)
                .SetItemMatrix lngRP, lngCP + 0, i & "-" & j, , ">-", , , ColorODBack1, ColorODFore0
                .SetItemMatrix lngRP, lngCP + 1, HyoConv(mobj_HyoWakuren.item("B" & i & j)), , ">-", , , ColorODBack0, ColorODForeL
            Next j
        Next i
    End With
    mGridisExist(lngGrid) = True
    RaiseEvent Fetched(lngGrid)
        
    lngGrid = ODDS_H1_NIN 'mGridData(13) 単複枠-票数　（人気順）
    ' グリッドデータオブジェクトを生成する
    Set mGridData(lngGrid) = New clsGridData
    
    With mGridData(lngGrid)
        .Cols = 17
        .Rows = IIf(mlngTosu < 13, 13, mlngTosu + 1)
        'カラムテキスト
        lngRP = 0
        .SetItemMatrix lngRP, 0, "単勝票数", , , , , ColorODBack1, ColorODFore0
        .SetItemMatrix lngRP, 1, "単勝票数", , , , , ColorODBack1, ColorODFore0
        .SetItemMatrix lngRP, 2, "単勝票数", , , , , ColorODBack1, ColorODFore0
        .SetItemMatrix lngRP, 3, , , , , , ColorODBack0
        .SetItemMatrix lngRP, 4, "複勝票数", , , , , ColorODBack1, ColorODFore0
        .SetItemMatrix lngRP, 5, "複勝票数", , , , , ColorODBack1, ColorODFore0
        .SetItemMatrix lngRP, 6, "複勝票数", , , , , ColorODBack1, ColorODFore0
        .SetItemMatrix lngRP, 7, , , , , , ColorODBack0
        For i = 0 To 8
            .SetItemMatrix 0, 8 + i, "枠連票数", , , , , ColorODBack1, ColorODFore0
        Next i
    End With
    
    Call setNin(mobj_HyoTansyo, lngGrid, "H1", 2, 2, 1, 0, mlngTosu)
    Call setNin(mobj_HyoFukusyo, lngGrid, "H1", 2, 2, 1, 4, mlngTosu)
    Call setNin(mobj_HyoWakuren, lngGrid, "H1", 2, 2, 1, 8, 12)
    mGridisExist(lngGrid) = True
    RaiseEvent Fetched(lngGrid)

    Exit Function
ErrorHandler:
    gApp.ErrLog
End Function

'
'   機能: 票数-馬連用グリッドを作る
'
'   備考: 文字列をディクショナリに読み込み、ディクショナリからグリッドを作る
'
Private Function set_H2(buf As String) As Boolean
On Error GoTo ErrorHandler
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim colOffset As Long
    Dim rowOffset As Long
    Dim lngCP        As Long            '' セル入力用
    Dim lngRP        As Long            '' セル入力用
    Dim lngGrid As Long

    rowOffset = 0
    colOffset = 0
    
    Call DataToDic(mobj_HyoUmaren, buf, 4, 11, 3)

    lngGrid = ODDS_H2_NUM 'mGridData(3) 馬連　-票数　（番号順）
    ' グリッドデータオブジェクトを生成する
    Set mGridData(lngGrid) = New clsGridData
    
    With mGridData(lngGrid)
        .Cols = 20
        .Rows = 18
        For i = 1 To 17
            DoEvents
            lngCP = ColPosition10(rowOffset, colOffset, i, j) - 1
            'カラムインデックス
            lngRP = RowPosition10(rowOffset, colOffset, i, i)
            .SetItemMatrix lngRP, lngCP + 0, "", , ">-", , , ColorODBack0, ColorODFore0
            .SetItemMatrix lngRP, lngCP + 1, i + 0, , "^-", , , gApp.GetWakubanColor(mobj_Wakuban.item(i)), Contrast(gApp.GetWakubanColor(mobj_Wakuban.item(i))) '
            
            For j = i + 1 To 18 '枠番、票数書き込み
                lngRP = RowPosition10(rowOffset, colOffset, i, j)
                .SetItemMatrix lngRP, lngCP + 0, j + 0, , ">-", , , ColorODBack1, ColorODFore0
                .SetItemMatrix lngRP, lngCP + 1, HyoConv(mobj_HyoUmaren.item("B" & Format$(i * 100 + j, "0000"))), , ">-", , , ColorODBack0, ColorODForeL
            Next j
        Next i
    End With
    mGridisExist(lngGrid) = True
    RaiseEvent Fetched(lngGrid)

    lngGrid = ODDS_H2_NIN 'mGridData(15) 馬連　-票数　（人気順）
    ' グリッドデータオブジェクトを生成する
    Set mGridData(lngGrid) = New clsGridData
    Call setNin(mobj_HyoUmaren, lngGrid, "H2", 4, 3, 0, 0, mlngNinkijunRow)
    mGridisExist(lngGrid) = True
    RaiseEvent Fetched(lngGrid)

    Exit Function
ErrorHandler:
    gApp.ErrLog
End Function

'
'   機能: 票数-ワイド用グリッドを作る
'
'   備考: 文字列をディクショナリに読み込み、ディクショナリからグリッドを作る
'
Private Function set_H3(buf As String) As Boolean
On Error GoTo ErrorHandler
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim colOffset As Long
    Dim rowOffset As Long
    Dim lngCP        As Long            '' セル入力用
    Dim lngRP        As Long            '' セル入力用
    Dim lngGrid As Long

    rowOffset = 0
    colOffset = 0
    
    Call DataToDic(mobj_HyoWide, buf, 4, 11, 3)

    lngGrid = ODDS_H3_NUM 'mGridData(5) ワイド-票数　（番号順）
    ' グリッドデータオブジェクトを生成する
    Set mGridData(lngGrid) = New clsGridData
    
    With mGridData(lngGrid)
        .Cols = 20
        .Rows = 18
        For i = 1 To 17
            DoEvents
            lngCP = ColPosition10(rowOffset, colOffset, i, j) - 1
            'カラムインデックス
            lngRP = RowPosition10(rowOffset, colOffset, i, i)
            .SetItemMatrix lngRP, lngCP + 0, "", , ">-", , , ColorODBack0, ColorODFore0
            .SetItemMatrix lngRP, lngCP + 1, i + 0, , "^-", , , gApp.GetWakubanColor(mobj_Wakuban(i)), Contrast(gApp.GetWakubanColor(mobj_Wakuban(i))) '
            
            For j = i + 1 To 18 '枠番、票数書き込み
                lngRP = RowPosition10(rowOffset, colOffset, i, j)
                .SetItemMatrix lngRP, lngCP + 0, j + 0, , ">-", , , ColorODBack1, ColorODFore0
                .SetItemMatrix lngRP, lngCP + 1, HyoConv(mobj_HyoWide.item("B" & Format$(i * 100 + j, "0000"))), , ">-", , , ColorODBack0, ColorODForeL
            Next j
        Next i
    End With
    mGridisExist(lngGrid) = True
    RaiseEvent Fetched(lngGrid)


    lngGrid = ODDS_H3_NIN 'mGridData(17) ワイド-票数　（人気順）
    ' グリッドデータオブジェクトを生成する
    Set mGridData(lngGrid) = New clsGridData
    Call setNin(mobj_HyoWide, lngGrid, "H3", 4, 3, 0, 0, mlngNinkijunRow)
    mGridisExist(lngGrid) = True
    RaiseEvent Fetched(lngGrid)

    Exit Function
ErrorHandler:
    gApp.ErrLog
End Function


'
'   機能: 票数-馬単用グリッドを作る
'
'   備考: 文字列をディクショナリに読み込み、ディクショナリからグリッドを作る
'
Private Function set_H4(buf As String) As Boolean
On Error GoTo ErrorHandler
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim colOffset As Long
    Dim rowOffset As Long
    Dim lngCP        As Long            '' セル入力用
    Dim lngRP        As Long            '' セル入力用
    Dim lngGrid As Long

    rowOffset = 0
    colOffset = 0

    Call DataToDic(mobj_HyoUmatan, buf, 4, 11, 3)

    lngGrid = ODDS_H4_NUM 'mGridData(7) 馬単　-票数　（番号順）
    ' グリッドデータオブジェクトを生成する
    Set mGridData(lngGrid) = New clsGridData

    With mGridData(lngGrid)
        .Cols = 18 * 2
        .Rows = 18 + 1
        For i = 1 To 18
            DoEvents
            lngCP = (i - 1) * 2
            'カラムインデックス
            lngRP = rowOffset
            .SetItemMatrix lngRP, lngCP + 0, "", , ">-", , , ColorODBack0, ColorODFore0
            .SetItemMatrix lngRP, lngCP + 1, i + 0, , "^-", , , gApp.GetWakubanColor(mobj_Wakuban(i)), Contrast(gApp.GetWakubanColor(mobj_Wakuban(i))) '
            
            For j = 1 To 18  '枠番、票数書き込み
                lngRP = j
                .SetItemMatrix lngRP, lngCP + 0, j + 0, , ">-", , , ColorODBack1, ColorODFore0
                .SetItemMatrix lngRP, lngCP + 1, HyoConv(mobj_HyoUmatan.item("B" & Format$(i * 100 + j, "0000"))), , ">-", , , ColorODBack0, ColorODForeL
            Next j
        Next i
    End With
    mGridisExist(lngGrid) = True
    RaiseEvent Fetched(lngGrid)

    lngGrid = ODDS_H4_NIN 'mGridData(19) 馬単　-票数　（人気順）
    ' グリッドデータオブジェクトを生成する
    Set mGridData(lngGrid) = New clsGridData
    Call setNin(mobj_HyoUmatan, lngGrid, "H4", 4, 3, 0, 0, mlngNinkijunRow)
    mGridisExist(lngGrid) = True
    RaiseEvent Fetched(lngGrid)

    Exit Function
ErrorHandler:
    gApp.ErrLog
End Function

'
'   機能: 票数-3連複用グリッドを作る
'
'   備考: 文字列をディクショナリに読み込み、ディクショナリからグリッドを作る
'
Private Function set_H5(buf As String, lngAxis As Long) As Boolean
On Error GoTo ErrorHandler
    Dim i As Long
    Dim j As Long
    Dim k(1 To 18) As Long
    Dim colOffset As Long
    Dim rowOffset As Long
    Dim lngCP        As Long            '' セル入力用
    Dim lngRP        As Long            '' セル入力用
    Dim lngGrid As Long

    rowOffset = 0
    colOffset = 0

    For i = 1 To 18
        k(i) = IIf(i < lngAxis, i, i + 1)
    Next i
    
    Call DataToDic(mobj_HyoSanrenpuku, buf, 6, 11, 3)

    lngGrid = ODDS_H5_NUM 'mGridData(9) 三連　-票数　（番号順）
    ' グリッドデータオブジェクトを生成する
    Set mGridData(lngGrid) = New clsGridData

    With mGridData(lngGrid)
        .Rows = 17
        .Cols = 20
        For i = 1 To 16
            DoEvents
            lngCP = ColPosition10(rowOffset, colOffset, i, j) - 1 '列index番号呼び出し
            'カラムインデックス
            lngRP = RowPosition10(rowOffset, colOffset, i, i)
            .SetItemMatrix lngRP, lngCP + 0, "", , ">-", , , ColorODBack0, ColorODFore0
            .SetItemMatrix lngRP, lngCP + 1, k(i) + 0, , "^-", , , gApp.GetWakubanColor(mobj_Wakuban(k(i))), Contrast(gApp.GetWakubanColor(mobj_Wakuban(k(i)))) '
            For j = i + 1 To 17
                lngRP = RowPosition10(rowOffset, colOffset, i, j)
                .SetItemMatrix lngRP, lngCP + 0, k(j) + 0, , ">-", , , ColorODBack1, ColorODFore0
                .SetItemMatrix lngRP, lngCP + 1, HyoConv(mobj_HyoSanrenpuku.item(SanrenDicKey(lngAxis, k(i), k(j)))), , ">-", , , ColorODBack0, ColorODForeL
            Next j
        Next i
    End With
    mGridisExist(lngGrid) = True
    RaiseEvent Fetched(lngGrid)


    lngGrid = ODDS_H5_NIN 'mGridData(21) 三連　-票数　（人気順）
    ' グリッドデータオブジェクトを生成する
    Set mGridData(lngGrid) = New clsGridData
    Call setNin(mobj_HyoSanrenpuku, lngGrid, "H5", 6, 3, 0, 0, mlngNinkijunRow)
    mGridisExist(lngGrid) = True
    RaiseEvent Fetched(lngGrid)

    Exit Function
ErrorHandler:
    gApp.ErrLog
End Function


'
'   機能: 票数-3連単用グリッドを作る
'
'   備考: 文字列をディクショナリに読み込み、ディクショナリからグリッドを作る
'
Private Function set_H6(key As String, lngAxis As Long) As Boolean  'Sanren
On Error GoTo ErrorHandler
    Dim strItem As String
    Dim strHyo  As String            '' 3連単　票総数
    Dim i As Long
    Dim j As Long
    Dim k(1 To 18) As Long
    Dim colOffset As Long
    Dim rowOffset As Long
    Dim lngCP        As Long            '' セル入力用
    Dim lngRP        As Long            '' セル入力用
    Dim lngGrid As Long
    Dim lngMDBIndex As Long
    
    rowOffset = 0
    colOffset = 0
        
    gApp.Log "GetH6 start"
    strItem = gApp.GetH6(key)
    gApp.Log "GetH6 end"

    ' データが無かったらあるいは、発売がなかったらFalseを返して終了
    If strItem = "" Or Mid$(strItem, 32, 1) = "0" Then
        set_H6 = False
        Exit Function
    End If
    
    '3連単　票数データ区分
    mstrHappyo(7) = KubunAndHappyoStr(Mid$(strItem, 3, 1), "") '
    
    '3連単票総数
    strHyo = Mid$(strItem, 102867, 11)
    mstrHyo(7) = IIf(val(strHyo) = 0, "", "３連単：" & val(strHyo) & "票")
    
    Call DataToDic(mobj_HyoSanrentan, Mid$(strItem, 51, 102816), 6, 11, 4)  '6114

    For i = 1 To 18
        k(i) = IIf(i <= lngAxis, i, i - 1)
    Next i
    
    lngGrid = ODDS_H6_NUM 'mGridData(11) 三連単　-票数　（番号順）
    ' グリッドデータオブジェクトを生成する
    Set mGridData(lngGrid) = New clsGridData
    With mGridData(lngGrid)
        .Cols = (18 - 1) * 2
        .Rows = 18
        For i = 1 To 18
            DoEvents
            lngCP = (k(i) - 1) * 2
            If lngAxis <> i Then '軸ボタンと同じ数の列、は処理しない
                'カラムインデックス
                lngRP = rowOffset
                .SetItemMatrix lngRP, lngCP + 0, "", , ">-", , , ColorODBack0, ColorODFore0
                .SetItemMatrix lngRP, lngCP + 1, i + 0, , "^-", , , gApp.GetWakubanColor(mobj_Wakuban(i)), Contrast(gApp.GetWakubanColor(mobj_Wakuban(i))) '
                
                For j = 1 To 18
                    lngRP = k(j)
                    If lngAxis <> j Then '軸ボタンと同じ数の行、は処理しない
                        .SetItemMatrix lngRP, lngCP + 0, j + 0, , ">-", , , ColorODBack1, ColorODFore0
                        If i <> j Then  '2着3着が同じ数字の場合、値を入れない
                            strItem = HyoConv(mobj_HyoSanrentan.item("B" & Format$(lngAxis * 10000 + i * 100 + j, "000000")))
                            .SetItemMatrix lngRP, lngCP + 1, strItem, , ">-", , , ColorODBack0, ColorODForeL
                        End If
                    End If
                Next j
            End If
        Next i
    End With
    mGridisExist(lngGrid) = True
    RaiseEvent Fetched(lngGrid)
    
    lngGrid = ODDS_H6_NIN 'mGridData(23) 三連単　-票数　（人気順）
    ' グリッドデータオブジェクトを生成する
    Set mGridData(lngGrid) = New clsGridData
    Call setNin(mobj_HyoSanrentan, lngGrid, "H6", 6, 4, 0, 0, mlngNinkijunRow)
    mGridisExist(lngGrid) = True
    RaiseEvent Fetched(lngGrid)
    
    Exit Function
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Function


'
'   機能: オッズ-単複枠用グリッドを作る
'
'   備考: なし
'
Private Function set_O1(ByRef objKey As clsKeyRA) As Boolean
On Error GoTo ErrorHandler
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim buf As String
    Dim strSQL As String
    Dim strItem As String
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim colOffset As Long
    Dim rowOffset As Long
    Dim lngCP        As Long            '' セル入力用
    Dim lngRP        As Long            '' セル入力用
    Dim lngGrid As Long

    rowOffset = 0
    colOffset = 0
    
    ' ODDS_TANPUKUWAKU (オッズ_単複枠)
    strSQL = "SELECT * FROM ODDS_TANPUKUWAKU " & objKey.SQLWHEREString
    Set cn = gApp.GetCN_ODDS_TANPUKUWAKU
    Set rs = cn.Execute(strSQL)

    ' データが無かったらFalseを返して終了
    If rs.EOF Or rs.BOF Then
        rs.Close
        Set rs = Nothing
        set_O1 = False
        Exit Function
    End If

    '発売フラグを参照し、発売があるならデータをディクショナリ登録
    If IfExist(rs, "TansyoFlag") = "1" Or IfExist(rs, "TansyoFlag") = "3" Or IfExist(rs, "TansyoFlag") = "7" Then
        Call DataToDic(mobj_OddsTansyo, IfExist(rs, "TanDATA"), 2, 4, 2)
    End If
    DoEvents
    If IfExist(rs, "FukusyoFlag") = "1" Or IfExist(rs, "FukusyoFlag") = "3" Or IfExist(rs, "FukusyoFlag") = "7" Then
        Call DataToDic(mobj_OddsFukusyo, IfExist(rs, "FukuDATA"), 2, 8, 2) '2442
    End If
    DoEvents
    If IfExist(rs, "WakurenFlag") = "1" Or IfExist(rs, "WakurenFlag") = "3" Or IfExist(rs, "WakurenFlag") = "7" Then
        Call DataToDic(mobj_OddsWakuren, IfExist(rs, "WakuDATA"), 2, 5, 2)
    End If
    
    '単複枠発表タイム
    mstrHappyo(0) = KubunAndHappyoStr(rs("DataKubun"), rs("HappyoTime"))
    '単複枠、各票総数を設定
    mstrHyo(0) = IIf(val(rs("TotalHyosuTansyo")) = 0, "", "単勝：" & val(rs("TotalHyosuTansyo")) & "票")
    mstrHyo(1) = IIf(val(rs("TotalHyosuFukusyo")) = 0, "", "複勝：" & val(rs("TotalHyosuFukusyo")) & "票")
    mstrHyo(2) = IIf(val(rs("TotalHyosuWakuren")) = 0, "", "枠連：" & val(rs("TotalHyosuWakuren")) & "票")
    
    rs.Close

    Set rs = Nothing
    
    
    lngGrid = ODDS_O1_NUM 'mGridData(0) 単複枠-オッズ（番号順）
    With mGridData(lngGrid)
        For i = 0 To mlngTosu - 1
            DoEvents
            '単勝
            strItem = OddsConv(mobj_OddsTansyo.item("B" & Right$(("0" & i + 1), 2)))
            .SetItemMatrix i + 1, 5, strItem, , ">-", , , , ColorHML(strItem)
            '複勝
            strItem = OddsConv(mobj_OddsFukusyo.item("B" & Right$(("0" & i + 1), 2)))
            .SetItemMatrix i + 1, 6, strItem, , ">-", , , , ColorHML(strItem)
        Next i
        '枠連
        rowOffset = 1
        colOffset = 8

        For i = 1 To 8
            DoEvents
            lngCP = ColPosition4(rowOffset, colOffset, i, j) - 1
            'カラムインデックス
            lngRP = RowPosition4(rowOffset, colOffset, i, i - 1)
            .SetItemMatrix lngRP, lngCP + 0, "", , ">-", , , ColorODBack0, ColorODFore0
            .SetItemMatrix lngRP, lngCP + 1, i & "枠", , "^-", , , gApp.GetWakubanColor(i), Contrast(gApp.GetWakubanColor(i))
            For j = i To 8 '枠番、票数書き込み
                lngRP = RowPosition4(rowOffset, colOffset, i, j)
                .SetItemMatrix lngRP, lngCP + 0, i & "-" & j, , ">-", , , ColorODBack1, ColorODFore0
                strItem = OddsConv(mobj_OddsWakuren.item("B" & i & j))
                .SetItemMatrix lngRP, lngCP + 1, strItem, , ">-", , , , ColorHML(strItem)
            Next j
        Next i
    End With
    mGridisExist(lngGrid) = True
    RaiseEvent Fetched(lngGrid)
    

    lngGrid = ODDS_O1_NIN 'mGridData(12) 単複枠-オッズ（人気順）'現在　番号順で代用中
    ' グリッドデータオブジェクトを生成する
    Set mGridData(lngGrid) = New clsGridData

    With mGridData(lngGrid)
        .Cols = 17
        .Rows = IIf(mlngTosu < 13, 13, mlngTosu + 1)
        'カラムテキスト
        lngRP = 0
        .SetItemMatrix lngRP, 0, "単勝オッズ", , , , , ColorODBack1, ColorODFore0
        .SetItemMatrix lngRP, 1, "単勝オッズ", , , , , ColorODBack1, ColorODFore0
        .SetItemMatrix lngRP, 2, "単勝オッズ", , , , , ColorODBack1, ColorODFore0
        .SetItemMatrix lngRP, 3, , , , , , ColorODBack0
        .SetItemMatrix lngRP, 4, "複勝オッズ", , , , , ColorODBack1, ColorODFore0
        .SetItemMatrix lngRP, 5, "複勝オッズ", , , , , ColorODBack1, ColorODFore0
        .SetItemMatrix lngRP, 6, "複勝オッズ", , , , , ColorODBack1, ColorODFore0
        .SetItemMatrix lngRP, 7, , , , , , ColorODBack0
        For i = 0 To 8
            .SetItemMatrix 0, 8 + i, "枠連オッズ", , , , , ColorODBack1, ColorODFore0
        Next i
    End With
    
    Call setNin(mobj_OddsTansyo, lngGrid, "O1", 2, 2, 1, 0, mlngTosu)
    Call setNin(mobj_OddsFukusyo, lngGrid, "O1", 2, 2, 1, 4, mlngTosu)
    Call setNin(mobj_OddsWakuren, lngGrid, "O1", 2, 2, 1, 8, 12)
    
    mGridisExist(lngGrid) = True
    RaiseEvent Fetched(lngGrid)
    
    Exit Function
ErrorHandler:
    gApp.ErrLog
End Function

'
'   機能: オッズ-馬連用グリッドを作る
'
'   備考: なし
'
Private Function set_O2(ByRef objKey As clsKeyRA) As Boolean
On Error GoTo ErrorHandler
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim buf As String
    Dim strSQL As String
    Dim strItem As String
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim colOffset As Long
    Dim rowOffset As Long
    Dim lngCP        As Long            '' セル入力用
    Dim lngRP        As Long            '' セル入力用
    Dim lngGrid As Long

    rowOffset = 0
    colOffset = 0
    
    ' ODDS_UMAREN (オッズ_馬連)
    strSQL = "SELECT * FROM ODDS_UMAREN " & objKey.SQLWHEREString
    Set cn = gApp.GetCN_ODDS_UMAREN
    Set rs = cn.Execute(strSQL)

    ' データが無かったらあるいは、発売がなかったらFalseを返して終了
    If rs.EOF Or rs.BOF Or IfExist(rs, "UmarenFlag") = "0" Or IfExist(rs, "UmarenFlag") = "" Then
        rs.Close
        Set rs = Nothing
        set_O2 = False
        Exit Function
    End If

    Call DataToDic(mobj_OddsUmaren, IfExist(rs, "DATA"), 4, 6, 3) '463
    
    '馬連発表タイム
    mstrHappyo(1) = KubunAndHappyoStr(rs("DataKubun"), rs("HappyoTime"))
    '馬連票総数を設定
    mstrHyo(3) = IIf(val(rs("TotalHyosuUmaren")) = 0, "", "馬連：" & val(rs("TotalHyosuUmaren")) & "票")
    
    rs.Close

    Set rs = Nothing
    
    lngGrid = ODDS_O2_NUM 'mGridData(2) 馬連　-オッズ（番号順）
    ' グリッドデータオブジェクトを生成する
    Set mGridData(lngGrid) = New clsGridData
    With mGridData(lngGrid)
        .Cols = 20
        .Rows = 18
        For i = 1 To 17
            DoEvents
            lngCP = ColPosition10(rowOffset, colOffset, i, j) - 1
            'カラムインデックス
            lngRP = RowPosition10(rowOffset, colOffset, i, i)
            .SetItemMatrix lngRP, lngCP + 0, "", , ">-", , , ColorODBack0, ColorODFore0
            .SetItemMatrix lngRP, lngCP + 1, i + 0, , "^-", , , gApp.GetWakubanColor(mobj_Wakuban(i)), Contrast(gApp.GetWakubanColor(mobj_Wakuban(i))) '
            
            For j = i + 1 To 18 '枠番、オッズ書き込み
                lngRP = RowPosition10(rowOffset, colOffset, i, j)
                .SetItemMatrix lngRP, lngCP + 0, j + 0, , ">-", , , ColorODBack1, ColorODFore0
                strItem = OddsConv(mobj_OddsUmaren.item("B" & Format$(i * 100 + j, "0000")))
                .SetItemMatrix lngRP, lngCP + 1, strItem, , ">-", , , , ColorHML(strItem)
            Next j
        Next i
    End With
    mGridisExist(lngGrid) = True
    RaiseEvent Fetched(lngGrid)
    
    lngGrid = ODDS_O2_NIN 'mGridData(14) 馬連　-オッズ（人気順）
    ' グリッドデータオブジェクトを生成する
    Set mGridData(lngGrid) = New clsGridData
    Call setNin(mobj_OddsUmaren, lngGrid, "O2", 4, 3, 0, 0, mlngNinkijunRow)
    mGridisExist(lngGrid) = True
    RaiseEvent Fetched(lngGrid)
    
    Exit Function
ErrorHandler:
    gApp.ErrLog
End Function

'
'   機能: オッズ-ワイド用グリッドを作る
'
'   備考: なし
'
Private Function set_O3(ByRef objKey As clsKeyRA) As Boolean
On Error GoTo ErrorHandler
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim buf As String
    Dim strSQL As String
    Dim strItem As String
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim colOffset As Long
    Dim rowOffset As Long
    Dim lngCP        As Long            '' セル入力用
    Dim lngRP        As Long            '' セル入力用
    Dim lngGrid As Long

    rowOffset = 0
    colOffset = 0
    
    ' ODDS_WIDE (オッズ_ワイド)
    strSQL = "SELECT * FROM ODDS_WIDE " & objKey.SQLWHEREString
    Set cn = gApp.GetCN_ODDS_WIDE
    Set rs = cn.Execute(strSQL)

    ' データが無かったらあるいは、発売がなかったらFalseを返して終了
    If rs.EOF Or rs.BOF Or IfExist(rs, "WideFlag") = "0" Or IfExist(rs, "WideFlag") = "" Then
        rs.Close
        Set rs = Nothing
        set_O3 = False
        Exit Function
    End If

    Call DataToDic(mobj_OddsWide, IfExist(rs, "DATA"), 4, 10, 3) ' 4553
    
    'ワイド発表タイム
    mstrHappyo(2) = KubunAndHappyoStr(rs("DataKubun"), rs("HappyoTime"))
    'ワイド票総数を設定
    mstrHyo(4) = IIf(val(rs("TotalHyosuWide")) = 0, "", "ワイド：" & val(rs("TotalHyosuWide")) & "票")
    
    rs.Close

    Set rs = Nothing
    
    
    lngGrid = ODDS_O3_NUM 'mGridData(4) ワイド-オッズ（番号順）
    ' グリッドデータオブジェクトを生成する
    Set mGridData(lngGrid) = New clsGridData
    With mGridData(lngGrid)
        .Cols = 20
        .Rows = 18
        For i = 1 To 17
            DoEvents
            lngCP = ColPosition10(rowOffset, colOffset, i, j) - 1
            'カラムインデックス
            lngRP = RowPosition10(rowOffset, colOffset, i, i)
            .SetItemMatrix lngRP, lngCP + 0, "", , ">-", , , ColorODBack0, ColorODFore0
            .SetItemMatrix lngRP, lngCP + 1, i + 0, , "^-", , , gApp.GetWakubanColor(mobj_Wakuban(i)), Contrast(gApp.GetWakubanColor(mobj_Wakuban(i))) '
            
            For j = i + 1 To 18 '枠番、オッズ書き込み
                lngRP = RowPosition10(rowOffset, colOffset, i, j)
                .SetItemMatrix lngRP, lngCP + 0, j + 0, , ">-", , , ColorODBack1, ColorODFore0
                strItem = OddsConv(mobj_OddsWide.item("B" & Format$(i * 100 + j, "0000")))
                .SetItemMatrix lngRP, lngCP + 1, strItem, , ">-", , , , ColorHML(strItem)
            Next j
        Next i
    End With
    mGridisExist(lngGrid) = True
    RaiseEvent Fetched(lngGrid)

    lngGrid = ODDS_O3_NIN 'mGridData(16) ワイド-オッズ（人気順）
    ' グリッドデータオブジェクトを生成する
    Set mGridData(lngGrid) = New clsGridData
    Call setNin(mobj_OddsWide, lngGrid, "O3", 4, 3, 0, 0, mlngNinkijunRow)
    mGridisExist(lngGrid) = True
    RaiseEvent Fetched(lngGrid)

    Exit Function
ErrorHandler:
    gApp.ErrLog
End Function

'
'   機能: オッズ-馬単用グリッドを作る
'
'   備考: なし
'
Private Function set_O4(ByRef objKey As clsKeyRA) As Boolean
On Error GoTo ErrorHandler
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim buf As String
    Dim strSQL As String
    Dim strItem As String
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim colOffset As Long
    Dim rowOffset As Long
    Dim lngCP        As Long            '' セル入力用
    Dim lngRP        As Long            '' セル入力用
    Dim lngGrid As Long
    Dim lngMDBIndex  As Long

    rowOffset = 0
    colOffset = 0
    
    ' 分割MDBのインデックス番号を場所コードから得る
    If objKey.JyoCD >= 1 And objKey.JyoCD <= 10 Then
        lngMDBIndex = objKey.JyoCD - 1
    Else
        lngMDBIndex = 0
    End If
    ' ODDS_UMATAN (オッズ_馬単)
    strSQL = "SELECT * FROM ODDS_UMATAN " & objKey.SQLWHEREString
    Set cn = gApp.GetCN_ODDS_UMATAN(lngMDBIndex)
    Set rs = cn.Execute(strSQL)

    ' データが無かったらあるいは、発売がなかったらFalseを返して終了
    If rs.EOF Or rs.BOF Or IfExist(rs, "UmatanFlag") = "0" Or IfExist(rs, "UmatanFlag") = "" Then
        rs.Close
        Set rs = Nothing
        set_O4 = False
        Exit Function
    End If

    Call DataToDic(mobj_OddsUmatan, IfExist(rs, "DATA"), 4, 6, 3) ' 463
    
    '馬単発表タイム
    mstrHappyo(3) = KubunAndHappyoStr(rs("DataKubun"), rs("HappyoTime"))
    '馬単票総数を設定
    mstrHyo(5) = IIf(val(rs("TotalHyosuUmatan")) = 0, "", "馬単：" & val(rs("TotalHyosuUmatan")) & "票")
    
    rs.Close

    Set rs = Nothing
    
    
    lngGrid = ODDS_O4_NUM 'mGridData(6) 馬単　-オッズ（番号順）
    ' グリッドデータオブジェクトを生成する
    Set mGridData(lngGrid) = New clsGridData
    With mGridData(lngGrid)
        .Cols = 18 * 2
        .Rows = 18 + 1
        For i = 1 To 18
            DoEvents
            lngCP = (i - 1) * 2
            'カラムインデックス
            lngRP = rowOffset
            .SetItemMatrix lngRP, lngCP + 0, "", , ">-", , , ColorODBack0, ColorODFore0
            .SetItemMatrix lngRP, lngCP + 1, i + 0, , "^-", , , gApp.GetWakubanColor(mobj_Wakuban(i)), Contrast(gApp.GetWakubanColor(mobj_Wakuban(i))) '
            
            For j = 1 To 18  '枠番、オッズ書き込み
                lngRP = j
                .SetItemMatrix lngRP, lngCP + 0, j + 0, , ">-", , , ColorODBack1, ColorODFore0
                strItem = OddsConv(mobj_OddsUmatan.item("B" & Format$(i * 100 + j, "0000")))
                .SetItemMatrix lngRP, lngCP + 1, strItem, , ">-", , , , ColorHML(strItem)
            Next j
        Next i
    End With
    mGridisExist(lngGrid) = True
    RaiseEvent Fetched(lngGrid)

    lngGrid = ODDS_O4_NIN 'mGridData(16) 馬単　-オッズ（人気順）
    ' グリッドデータオブジェクトを生成する
    Set mGridData(lngGrid) = New clsGridData
    Call setNin(mobj_OddsUmatan, lngGrid, "O4", 4, 3, 0, 0, mlngNinkijunRow)
    mGridisExist(lngGrid) = True
    RaiseEvent Fetched(lngGrid)

    Exit Function
ErrorHandler:
    gApp.ErrLog
End Function

'
'   機能: オッズ-3連複用グリッドを作る
'
'   備考: なし
'
Private Function set_O5(ByRef objKey As clsKeyRA, lngAxis As Long) As Boolean
On Error GoTo ErrorHandler
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim buf As String
    Dim strSQL As String
    Dim strItem As String
    Dim i As Long
    Dim j As Long
    Dim k(1 To 18) As Long
    Dim colOffset As Long
    Dim rowOffset As Long
    Dim lngCP        As Long            '' セル入力用
    Dim lngRP        As Long            '' セル入力用
    Dim lngGrid As Long
    Dim lngMDBIndex As Long
    
    rowOffset = 0
    colOffset = 0
        
    ' 分割MDBのインデックス番号を場所コードから得る
    If objKey.JyoCD >= 1 And objKey.JyoCD <= 10 Then
        lngMDBIndex = objKey.JyoCD - 1
    Else
        lngMDBIndex = 0
    End If
    
    ' ODDS_SANREN (オッズ_3連複)
    strSQL = "SELECT * FROM ODDS_SANREN " & objKey.SQLWHEREString
    Set cn = gApp.GetCN_ODDS_SANREN(lngMDBIndex)
    Set rs = cn.Execute(strSQL)

    ' データが無かったらあるいは、発売がなかったらFalseを返して終了
    If rs.EOF Or rs.BOF Or IfExist(rs, "SanrenFlag") = "0" Or IfExist(rs, "SanrenFlag") = "" Then
        rs.Close
        Set rs = Nothing
        set_O5 = False
        Exit Function
    End If

    Call DataToDic(mobj_OddsSanrenpuku, IfExist(rs, "DATA"), 6, 6, 3) '663
    
    '3連複発表タイム
    mstrHappyo(4) = KubunAndHappyoStr(rs("DataKubun"), rs("HappyoTime"))
    '3連複票総数を設定
    mstrHyo(6) = IIf(val(rs("TotalHyosuSanren")) = 0, "", "３連複：" & val(rs("TotalHyosuSanren")) & "票")
    
    rs.Close

    Set rs = Nothing
    
    For i = 1 To 18
        k(i) = IIf(i < lngAxis, i, i + 1)
    Next i
    
    lngGrid = ODDS_O5_NUM 'mGridData(8) 三連　-オッズ（番号順）
    ' グリッドデータオブジェクトを生成する
    Set mGridData(lngGrid) = New clsGridData
    With mGridData(lngGrid)
        .Rows = 17
        .Cols = 20
        For i = 1 To 16
            DoEvents
            lngCP = ColPosition10(rowOffset, colOffset, i, j) - 1 '列index番号呼び出し
            'カラムインデックス
            lngRP = RowPosition10(rowOffset, colOffset, i, i)
            .SetItemMatrix lngRP, lngCP + 0, "", , "^-", , , ColorODBack0, ColorODFore0
            .SetItemMatrix lngRP, lngCP + 1, k(i) + 0, , "^-", , , gApp.GetWakubanColor(mobj_Wakuban(k(i))), Contrast(gApp.GetWakubanColor(mobj_Wakuban(k(i)))) '
            For j = i + 1 To 17
                lngRP = RowPosition10(rowOffset, colOffset, i, j)
                .SetItemMatrix lngRP, lngCP + 0, k(j) + 0, , ">-", , , ColorODBack1, ColorODFore0
                strItem = OddsConv(mobj_OddsSanrenpuku.item(SanrenDicKey(lngAxis, k(i), k(j))))
                .SetItemMatrix lngRP, lngCP + 1, strItem, , ">-", , , , ColorHML(strItem)
            Next j
        Next i
    End With
    mGridisExist(lngGrid) = True
    RaiseEvent Fetched(lngGrid)

    lngGrid = ODDS_O5_NIN 'mGridData(20) 三連　-オッズ（人気順）
    ' グリッドデータオブジェクトを生成する
    Set mGridData(lngGrid) = New clsGridData
    Call setNin(mobj_OddsSanrenpuku, lngGrid, "O5", 6, 3, 0, 0, mlngNinkijunRow)
    mGridisExist(lngGrid) = True
    RaiseEvent Fetched(lngGrid)

    Exit Function
ErrorHandler:
    gApp.ErrLog
End Function


'
'   機能: オッズ-3連単用グリッドを作る
'
'   備考: なし
'
Private Function set_O6(key As String, lngAxis As Long) As Boolean  'Sanren
On Error GoTo ErrorHandler
    Dim strBuff As String
    Dim strHyo  As String            '' 3連単　票総数
    Dim strItem As String
    Dim i As Long
    Dim j As Long
    Dim k(1 To 18) As Long
    Dim colOffset As Long
    Dim rowOffset As Long
    Dim lngCP        As Long            '' セル入力用
    Dim lngRP        As Long            '' セル入力用
    Dim lngGrid As Long
    Dim lngMDBIndex As Long
    
    rowOffset = 0
    colOffset = 0
        
    gApp.Log "GetO6 start"
    strBuff = gApp.GetO6(key)
    gApp.Log "GetO6 end"

    ' データが無かったらあるいは、発売がなかったらFalseを返して終了
    If strBuff = "" Or Mid$(strBuff, 40, 1) = "0" Then
        set_O6 = False
        Exit Function
    End If
    
    '3連単発表タイム
    mstrHappyo(5) = KubunAndHappyoStr(Mid$(strBuff, 3, 1), Mid$(strBuff, 28, 8))
    
    '3連単票総数
    strHyo = val(Mid$(strBuff, 83273, 11))
    mstrHyo(7) = IIf(val(strHyo) = 0, "", "３連単：" & val(strHyo) & "票")
    
    Call DataToDic(mobj_OddsSanrentan, Mid$(strBuff, 41, 83232), 6, 7, 4) '674
    
    For i = 1 To 18
        k(i) = IIf(i <= lngAxis, i, i - 1)
    Next i
    
    lngGrid = ODDS_O6_NUM  'mGridData(10) 三連単　-オッズ（番号順）
    ' グリッドデータオブジェクトを生成する
    Set mGridData(lngGrid) = New clsGridData
    With mGridData(lngGrid)
        .Cols = (18 - 1) * 2
        .Rows = 18
        For i = 1 To 18
            DoEvents
            lngCP = (k(i) - 1) * 2
            If lngAxis <> i Then '軸ボタンと同じ数の列、は処理しない
                'カラムインデックス
                lngRP = rowOffset
                .SetItemMatrix lngRP, lngCP + 0, "", , ">-", , , ColorODBack0, ColorODFore0
                .SetItemMatrix lngRP, lngCP + 1, i + 0, , "^-", , , gApp.GetWakubanColor(mobj_Wakuban(i)), Contrast(gApp.GetWakubanColor(mobj_Wakuban(i))) '
                
                For j = 1 To 18
                    lngRP = k(j)
                    If lngAxis <> j Then '軸ボタンと同じ数の行、は処理しない
                        'ロウインデックス
                        .SetItemMatrix lngRP, lngCP + 0, j + 0, , ">-", , , ColorODBack1, ColorODFore0
                        If i <> j Then  '2着3着が同じ数字の場合、値を入れない
                            strItem = OddsConv(mobj_OddsSanrentan.item("B" & Format$(lngAxis * 10000 + i * 100 + j, "000000")))
                            .SetItemMatrix lngRP, lngCP + 1, strItem, , ">-", , , , ColorHML(strItem)
                        End If
                    End If
                Next j
            End If
        Next i
    End With
    mGridisExist(lngGrid) = True
    RaiseEvent Fetched(lngGrid)
    
    
    lngGrid = ODDS_O6_NIN 'mGridData(22) 三連　-オッズ（人気順）
    ' グリッドデータオブジェクトを生成する
    Set mGridData(lngGrid) = New clsGridData
    Call setNin(mobj_OddsSanrentan, lngGrid, "O6", 6, 4, 0, 0, mlngNinkijunRow)
    mGridisExist(lngGrid) = True
    RaiseEvent Fetched(lngGrid)

    Exit Function
ErrorHandler:
    gApp.ErrLog
End Function


'
'   機能: Dictionary登録関数
'
'   備考: 引き数 dic  - 登録先Dictionary
'                buf  - データ文字列
'                k    - keyの文字数
'                o    - Odsの文字数
'                n    - Ninの文字数
'
Private Sub DataToDic(ByRef dic As Dictionary, ByVal buf As String, k As Long, o As Long, n As Long)
    Dim p As Long
    Dim tmp As String
    Dim key As String
    Dim Ods As String  'オッズor票数
    Dim Nin As String
    
    p = 0
    Do While p < Len(buf)
        key = Mid$(buf, 1 + p, k)
        Ods = Mid$(buf, 1 + p + k, o)
        Nin = "N" & Mid$(buf, 1 + p + k + o, n)
        
        If Trim$(key) <> "" Then
            Call dic.Add("B" & key, Mid$(Nin, 2, 1) & Ods)  '票数ALL0がきたとき人気順参照の要あり
            
            If dic.Exists(Nin) Then
                tmp = dic(Nin)
                Call dic.Remove(Nin)
                Call dic.Add(Nin, tmp & key)
            Else
                Call dic.Add(Nin, key)
            End If
        End If
        
        p = p + k + o + n
    Loop
End Sub


'
'   機能: グリッドに人気順を書き込む関数
'
'   備考: 引き数 ARG1 - 登録用Dictionary
'                ARG2 - グリッドナンバー
'                ARG3 - 1:単複枠 2:馬連 3:ワイド 4:馬単 5:3連複 5:3連単
'                ARG4 - Dictionary's Key文字数
'                ARG5 - Dictionary's Nin文字数
'                ARG6 - rowOffset
'                ARG7 - colOffset
'                ARG8 - 折り返し行数
'
Private Sub setNin(dic As Dictionary, lngGrid As Long, opt As String, lngKey As Long, lngNin As Long _
                    , rowOffset As Long, colOffset As Long, lngRow As Long)
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim lngColSet As Long
    Dim lngNum As Long                  '組み合わせ数
    Dim lngCounter As Long
    Dim lngTmp As Long
    Dim strTmp As String
    Dim strItem As String
    Dim lngCP        As Long            '' セル入力用
    Dim lngRP        As Long            '' セル入力用
    
    With mGridData(lngGrid)
        '出走頭数から組み合わせ数を計算
        Select Case Mid$(opt, 2, 1)
        Case "1"
            Select Case colOffset
            Case 0, 4  '単複
                lngNum = lngRow
            Case 8  '枠
                lngNum = 36
            End Select
        Case "2"  '馬連
            lngNum = mlngTosu * (mlngTosu - 1) / 2
        Case "3"  'ワイド
            lngNum = mlngTosu * (mlngTosu - 1) / 2
        Case "4"  '馬単
            lngNum = mlngTosu * (mlngTosu - 1)
        Case "5"  '3連複
            lngNum = mlngTosu * (mlngTosu - 1) * (mlngTosu - 2) / 6
        Case "6"  '3連単
            lngNum = mlngTosu * (mlngTosu - 1) * (mlngTosu - 2)
        End Select
        
        If Not Mid$(opt, 2, 1) = "1" Then
            '単複枠以外
            lngColSet = -Int(-lngNum / lngRow) '組み合わせの数と行数からカラムセット(3カラム)の数を計算
            .Cols = colOffset + (3 * lngColSet)
            .Rows = rowOffset + lngRow
        End If
        
        lngTmp = 0
        For i = 0 To lngNum - 1  '
            DoEvents
            If dic.Exists("N" & Right$("000" & i + 1, lngNin)) Then '人気順データがあるなら
                strItem = dic.item("N" & Right$("000" & i + 1, lngNin))
                k = 0
                Do While ((k + 1) * lngKey <= Len(strItem))
                    lngTmp = lngTmp + 1
                    lngCP = ColPositionNin(colOffset, colOffset, lngTmp, lngRow)
                    lngRP = RowPositionNin(rowOffset, colOffset, lngTmp, lngRow)
                    .SetItemMatrix lngRP, lngCP + 0, i + 1, , ">-", , , ColorODBack1, ColorODFore0
                    strTmp = Mid$(strItem, lngKey * k + 1, lngKey)  '馬番
                    '項目名（3種類に分類して処理："1"," 1-10"," 1-10-12"）
                    If Mid$(opt, 2, 1) = "5" Or Mid$(opt, 2, 1) = "6" Then  '3連複、3連単のとき（EX." 1-10-12"）
                        .SetItemMatrix lngRP, lngCP + 1, mSC.FitSpaceNum(Mid$(strTmp, 1, 2), 2) & "-" & mSC.FitSpaceNum(Mid$(strTmp, 3, 2), 2) & "-" & mSC.FitSpaceNum(Mid$(strTmp, 5, 2), 2), , ">-", , , ColorODBack0, ColorODForeL
                    ElseIf Mid$(opt, 2, 1) = "1" And Not colOffset = 8 Then  '単勝、複勝のとき（EX."1"）
                        .SetItemMatrix lngRP, lngCP + 1, val(strTmp), , ">-", , , ColorODBack0, ColorODForeL
                    Else                                                     '上記以外（枠連、馬連、ワイド、馬単）のとき（EX." 1-10"）
                        .SetItemMatrix lngRP, lngCP + 1, mSC.FitSpaceNum(Left$(strTmp, lngKey / 2), 2) & "-" & mSC.FitSpaceNum(Right$(strTmp, lngKey / 2), 2), , ">-", , , ColorODBack0, ColorODForeL
                    End If
                    '項目内容（オッズ・票数）
                    strTmp = dic.item("B" & strTmp)
                    If Mid$(opt, 1, 1) = "O" Then 'オッズ
                        strTmp = OddsConv(strTmp)
                        .SetItemMatrix lngRP, lngCP + 2, strTmp, , ">-", , , ColorODBack0, ColorHML(strTmp)
                    ElseIf Mid$(opt, 1, 1) = "H" Then  '票数
                        strTmp = HyoConv(strTmp)
                        .SetItemMatrix lngRP, lngCP + 2, strTmp, , ">-", , , ColorODBack0, ColorODForeL
                    End If
                    k = k + 1
                Loop
            End If
        Next i
    End With
End Sub


'
'   機能: オッズの値を適切な形に変換
'
'   備考: なし
'
Private Function OddsConv(ByVal str As String) As String
    Dim lngTmp As Long
    
    str = Mid$(str, 2)  '先頭1byte削除
    If val(str) <> 0 Then
        If Len(str) >= 8 Then
            lngTmp = Len(str) / 2
            OddsConv = Right$("   " & Format$(Mid$(str, 1, lngTmp) / 10, "0.0"), lngTmp + 1) & "～" & Right$("   " & Format$(Mid$(str, lngTmp + 1, lngTmp) / 10, "0.0"), lngTmp + 1)
        Else
            OddsConv = Format$(str / 10, "0.0")
        End If
    Else
        OddsConv = DataException(str)
    End If
End Function


'
'   機能: 票数の値を適切な形に変換
'
'   備考: なし
'
Private Function HyoConv(ByVal str As String) As String
    If val(Mid$(str, 2)) <> 0 Then
        HyoConv = val(Mid$(str, 2))  '票数
    Else
        If Left$(str, 1) = "-" Or Left$(str, 1) = "*" Then 'ALL0 発売前取消
            HyoConv = "ﾄﾘｹｼ"
        ElseIf Left$(str, 1) = " " Or str = "" Then  '登録無し
            HyoConv = Trim$(str)
        Else
            HyoConv = val(Mid$(str, 2)) 'ALL0 発売票数無し
        End If
    End If
End Function

'
'   機能: 例外文字列の処理
'
'   備考: なし
'
Private Function DataException(str As String) As String
    If mSC.Emp(str) Then                '登録無し、無投票 '無投票表示の変更可能性あり
        DataException = " "
    ElseIf Mid$(str, 1, 1) = "-" Then    '発売前取消
        DataException = "ﾄﾘｹｼ"
    ElseIf Mid$(str, 1, 1) = "*" Then    '発売後取消
        DataException = "ﾄﾘｹｼ"
    Else                                'オッズ
        DataException = str
    End If
End Function

'
'   機能: 値に対して色を与える
'
'   備考: なし
'
Private Function ColorHML(ByVal str As String) As String
    Dim lngTmp As Long
    
    If InStr(str, "～") Then str = Format$(val(Mid$(str, 1, InStr(str, "～") - 1)), "0.0")
    lngTmp = Len(str)
    Select Case lngTmp
    Case Is > 4
        ColorHML = ColorODForeL
    Case Is <= 3
        ColorHML = ColorODForeH
    Case Else
        If str = "ﾄﾘｹｼ" Then
            ColorHML = ColorODForeL
        Else
            ColorHML = ColorODForeM
        End If
    End Select
    
End Function


'
'   機能: ３連複ディクショナリーキー取得
'
'   備考: 引き数 ARG1 - 軸馬
'      　　      ARG2 - 行番
'      　　      ARG3 - 列番
'         大小で並べ替えるだけ
'
Private Function SanrenDicKey(i As Long, j As Long, k As Long) As String
    If i < j And j < k Then
        SanrenDicKey = "B" & Format$(i * 10000 + j * 100 + k, "000000")
    ElseIf i < k And k < j Then
        SanrenDicKey = "B" & Format$(i * 10000 + k * 100 + j, "000000")
    ElseIf j < i And i < k Then
        SanrenDicKey = "B" & Format$(j * 10000 + i * 100 + k, "000000")
    ElseIf j < k And k < i Then
        SanrenDicKey = "B" & Format$(j * 10000 + k * 100 + i, "000000")
    ElseIf k < i And i < j Then
        SanrenDicKey = "B" & Format$(k * 10000 + i * 100 + j, "000000")
    ElseIf k < j And j < i Then
        SanrenDicKey = "B" & Format$(k * 10000 + j * 100 + i, "000000")
    Else
        Debug.Print "Odds i,j,kの大小が選択できません"
    End If
End Function

'
'   機能: 枠連タブの枠に対するマトリクス位置のカラムを返す
'
'   備考: 引き数 ARG1 - ロウオフセット値
'                ARG2 - カラムオフセット値
'                ARG3 - 枠1 　１～８
'                ARG4 - 枠2  枠1～８
'
Private Function ColPosition4(oRow As Long, oCol As Long, i As Long, j As Long) As Long
    ColPosition4 = IIf(i <= 4, oCol + 1 + ((i - 1) * 2), oCol + 1 + ((7 - (i - 1)) * 2))
End Function


'
'   機能: 枠連タブの枠に対するマトリクス位置のロウを返す
'
'   備考: 引き数 ARG1 - ロウオフセット値
'                ARG2 - カラムオフセット値
'                ARG3 - 枠1 　１～８
'                ARG4 - 枠2  枠1～８
'
Private Function RowPosition4(oRow As Long, oCol As Long, i As Long, j As Long) As Long
    RowPosition4 = IIf(i <= 4, oRow + 1 + j - i, oRow + 1 + 2 + j)
End Function


'
'   機能: 馬連タブの枠に対するマトリクス位置のカラムを返す
'
'   備考: 引き数 ARG1 - ロウオフセット値
'                ARG2 - カラムオフセット値
'                ARG3 - １～１８
'                ARG4 - １～１８
'
Private Function ColPosition10(oRow As Long, oCol As Long, i As Long, j As Long) As Integer
    ColPosition10 = IIf(i <= 10, oCol + 1 + ((i - 1) * 2), oCol + 1 + ((19 - (i - 1)) * 2))
End Function

'
'   機能: 馬連タブの枠に対するマトリクス位置のカラムを返す
'
'   備考: 引き数 ARG1 - ロウオフセット値
'                ARG2 - カラムオフセット値
'                ARG3 - １～１８
'                ARG4 - １～１８
'
Private Function RowPosition10(oRow As Long, oCol As Long, i As Long, j As Long) As Integer
    RowPosition10 = IIf(i <= 10, oRow + j - i, oRow + j - 1)
End Function

'
'   機能: 人気順の順番に対するマトリクス位置のカラムを返す
'
'   備考: 引き数 ARG1 - ロウオフセット値
'                ARG2 - カラムオフセット値
'                ARG3 - 順番
'
Private Function ColPositionNin(oRow As Long, oCol As Long, i As Long, lngRow As Long) As Integer
    ColPositionNin = oCol + 3 * Fix((i - 1) / lngRow)
End Function

'
'   機能: 人気順の順番に対するマトリクス位置のロウを返す
'
'   備考: 引き数 ARG1 - ロウオフセット値
'                ARG2 - カラムオフセット値
'                ARG3 - 順番
'
Private Function RowPositionNin(oRow As Long, oCol As Long, i As Long, lngRow As Long) As Integer
    RowPositionNin = oRow + ((i - 1) Mod lngRow)
End Function


'
'   機能: 票数（全掛式）、必要な部分（票数データ以外）だけ構造体にデータを入れる
'
'   備考: 引き数 ARG1 - H1文字列データ
'                ARG2 - H1構造体
'
Private Sub SetData_H1_(lBuf As String, ByRef mBuf As JV_H1_HYOSU_ZENKAKE)
    Dim i As Integer                                '' ループカウンタ
    Dim j As Integer                                '' ループカウンタ
    Dim k As Integer                                '' ループカウンタ
    Dim p As Long                                   '' 切り分け開始位置
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMid(lBuf, p, 2)      '' レコード種別
            .DataKubun = IncMid(lBuf, p, 1)       '' データ区分
            With .MakeDate
                .Year = IncMid(lBuf, p, 4)        '' 年
                .Month = IncMid(lBuf, p, 2)       '' 月
                .Day = IncMid(lBuf, p, 2)         '' 日
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMid(lBuf, p, 4)            '' 開催年
            .MonthDay = IncMid(lBuf, p, 4)        '' 開催月日
            .JyoCD = IncMid(lBuf, p, 2)           '' 競馬場コード
            .Kaiji = IncMid(lBuf, p, 2)           '' 開催回[第N回]
            .Nichiji = IncMid(lBuf, p, 2)         '' 開催日目[N日目]
            .RaceNum = IncMid(lBuf, p, 2)         '' レース番号
        End With ' id
        .TorokuTosu = IncMid(lBuf, p, 2)          '' 登録頭数
        .SyussoTosu = IncMid(lBuf, p, 2)          '' 出走頭数
        For i = 0 To 6
            .HatubaiFlag(i) = IncMid(lBuf, p, 1)  '' 発売フラグ
        Next i
        .FukuChakuBaraiKey = IncMid(lBuf, p, 1)   '' 複勝着払キー
        For i = 0 To 27
            .HenkanUma(i) = IncMid(lBuf, p, 1)    '' 返還馬番情報(馬番01～28)
        Next i
        For i = 0 To 7
            .HenkanWaku(i) = IncMid(lBuf, p, 1)   '' 返還枠番情報(枠番1～8)
        Next i
        For i = 0 To 7
            .HenkanDoWaku(i) = IncMid(lBuf, p, 1) '' 返還同枠情報(枠番1～8)
        Next i
        p = 28800
        For i = 0 To 13
            .HyoTotal(i) = IncMid(lBuf, p, 11)    '' 票数合計
        Next i
        .CRLF = IncMid(lBuf, p, 2)                '' レコード区切り
    End With

End Sub


'
'   機能: DBVer1以上且つRA(DataKubun)<=2のとき、TC,CCに本レースの情報があればmBuf_RA構造体に上書きする
'
'   備考: 上書き対象はKyoriとTrackCD
'
Private Sub Set_CCTC()
On Error GoTo ErrorHandler
    Dim CC          As ADODB.Recordset 'コースチェンジ
    Dim TC          As ADODB.Recordset '発走チェンジ
    
    If Not val(gApp.R_DBVersion) > 0 Or mBuf_RA.head.DataKubun > "2" Then
    '古いDBの時はチェックしない
    '速報成績が出た時点で確定する為,DataKubun3以上はチェックしない
        Exit Sub
    End If
    
    Set CC = New ADODB.Recordset
    Set TC = New ADODB.Recordset
    Call OpenTableDirect(CC, gApp.GetCN_COURSE_CHANGE, "COURSE_CHANGE")
    Call OpenTableDirect(TC, gApp.GetCN_HASSOU_CHANGE, "HASSOU_CHANGE")
    With mKey
        SafeSeek CC, Array("Year", "MonthDay", "JyoCD", "Kaiji", "Nichiji", "RaceNum"), _
                    Array(.Year, .MonthDay, .JyoCD, .Kaiji, .Nichiji, .RaceNum)
        SafeSeek TC, Array("Year", "MonthDay", "JyoCD", "Kaiji", "Nichiji", "RaceNum"), _
                    Array(.Year, .MonthDay, .JyoCD, .Kaiji, .Nichiji, .RaceNum)
    End With
    
    ' CC,TC双方にデータが無かったらFalseを返して終了
    If (CC.EOF Or CC.BOF) And (TC.EOF Or TC.BOF) Then
        CC.Close
        TC.Close
        Set CC = Nothing
        Set TC = Nothing
        Exit Sub
    End If
        
    If Not (CC.EOF Or CC.BOF) Then
        'CCに変更情報があれば
        If CC("AtoKyori") <> "0000" Then
            mBuf_RA.KYORI = CC("AtoKyori")
        End If
        If CC("AtoTrackCD") <> "00" Then
            mBuf_RA.TrackCD = CC("AtoTrackCD")
        End If
    End If
    If Not (TC.EOF Or TC.BOF) Then
        'TCに変更情報があれば
        If TC("AtoHassoTime") <> "0000" Then
            mBuf_RA.HassoTime = TC("AtoHassoTime")
        End If
    End If
    
    CC.Close
    TC.Close
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub

