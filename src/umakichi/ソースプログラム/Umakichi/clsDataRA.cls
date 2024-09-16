VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDataRA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "RA - RACE データ取得オブジェクト"
'
'   RA - RACE データ取得オブジェクト
'

Option Explicit

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   外部関数(イベント)
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

' イベント
Public Event FetchedKihon(GridData As clsGridData)
Public Event FetchedKetto(GridData As clsGridData)
Public Event FetchedKako(GridData As clsGridData)
Public Event FetchedMining(GridData As clsGridData, DMKubun As String)
Public Event FetchedJokenBetu(GridData As clsGridData)
Public Event FetchedMotiTIme(GridData As clsGridData)
Public Event FetchedSeiseki(GridData As clsGridData, GridHarai As clsGridData, GridLap As clsGridData, flag As Long)
Public Event NoUMARACE()

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部変数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

' 変換ユーティリティー
Private mCC As clsCodeConverter          '' JV-Data のコード変換メソッド群
Private mSC As clsStringConverter        '' 文字列変換メソッド群

' Viewerに提供する為のデータ格納変数
Private mstrLabels(0 To 9)  As String    '' テキストデータ
Private mstrTitle           As String    '' タイトル
Private mlngTosu            As Long      '' 登録頭数
Private mlngHenkanRow       As Long      '' 変換情報の始まる行番号
Private mlngCornerRow       As Long      '' コーナー通過順情報の始まる行番号
Private mlngFirstTabNumber  As Long      '' 初期表示タブ番号
Private mstrRCKey           As String    '' レコード詳細画面のキー
Private mstrRCG1Key         As String    '' レコード詳細画面のキー
Private mblnIsG1            As Boolean   '' G1 レースかどうか
Private mblnIsPrompt        As Boolean   '' 速報を受け取れるレースかどうか
Private mblnIsShougai       As Boolean   '' 障害レースかどうか（障害なら持ちタイムタブが無い）

Private mCN_RACE        As ADODB.Connection
Attribute mCN_RACE.VB_VarHelpID = -1
Private WithEvents mAsyncCN_UMA_RACE_A As ADODB.Connection
Attribute mAsyncCN_UMA_RACE_A.VB_VarHelpID = -1
Private mCN_UMA_RACE_B  As ADODB.Connection
Attribute mCN_UMA_RACE_B.VB_VarHelpID = -1
Private mCN_UMA         As ADODB.Connection
Attribute mCN_UMA.VB_VarHelpID = -1
Private mCN_BANUSI      As ADODB.Connection
Attribute mCN_BANUSI.VB_VarHelpID = -1
Private mCN_CHOKYO      As ADODB.Connection
Attribute mCN_CHOKYO.VB_VarHelpID = -1
Private mCN_KISHU       As ADODB.Connection
Attribute mCN_KISHU.VB_VarHelpID = -1
Private mCN_SEISAN      As ADODB.Connection
Attribute mCN_SEISAN.VB_VarHelpID = -1
Private mCN_HANSYOKU    As ADODB.Connection
Attribute mCN_HANSYOKU.VB_VarHelpID = -1
Private mCN_HARAI       As ADODB.Connection
Attribute mCN_HARAI.VB_VarHelpID = -1

Private mRS_UMA_RACE_A As ADODB.Recordset
Attribute mRS_UMA_RACE_A.VB_VarHelpID = -1
Private mRS_UMA_RACE_B As ADODB.Recordset
Attribute mRS_UMA_RACE_B.VB_VarHelpID = -1
Private mRS_UMA        As ADODB.Recordset
Attribute mRS_UMA.VB_VarHelpID = -1
Private mRS_BANUSI     As ADODB.Recordset
Attribute mRS_BANUSI.VB_VarHelpID = -1
Private mRS_CHOKYO     As ADODB.Recordset
Attribute mRS_CHOKYO.VB_VarHelpID = -1
Private mRS_KISHU      As ADODB.Recordset
Attribute mRS_KISHU.VB_VarHelpID = -1
Private mRS_SEISAN     As ADODB.Recordset
Attribute mRS_SEISAN.VB_VarHelpID = -1
Private mRS_HANSYOKU   As ADODB.Recordset
Attribute mRS_HANSYOKU.VB_VarHelpID = -1
Private mRS_HARAI      As ADODB.Recordset
Private mRS_Time_K     As ADODB.Recordset
Private mRS_TimeJK     As ADODB.Recordset

' 変更情報
Private mRS_TENKO_BABA      As ADODB.Recordset ' 天候馬場
Private mRS_BATAIJYU        As ADODB.Recordset ' 馬体重
Private mRS_TORIKESI_JYOGAI As ADODB.Recordset ' 取り消し除外
Private mRS_KISHU_CHANGE    As ADODB.Recordset ' 騎手変更
Private mRS_MINING          As ADODB.Recordset ' マイニング
Private mRS_COURSE_CHANGE   As ADODB.Recordset ' コース変更
Private mRS_HASSOU_CHANGE   As ADODB.Recordset ' 発走時間

Private mDicAV              As Scripting.Dictionary ' 取り消し除外情報辞書
            
Private mBuf_RA       As JV_RA_RACE

Private mKey As clsKeyRA

Private mblnNowFetching         As Boolean  ' 取得中フラグ
Private mblnCancelFetching      As Boolean  ' キャンセルフラグ

Private mblnNowKakoFetching     As Boolean ' 過去走取得中フラグ
Private mblnCancelKakoFetching  As Boolean ' 過去走処理待ちフラグ


Private Type DM_INFO
    Umaban As String                    ''馬番
    DMTime As String                    ''予想走破タイム
    DMGosaP As String                   ''予想誤差(信頼度)＋
    DMGosaM As String                   ''予想誤差(信頼度)－
End Type
    
    
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   プロパティ
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: ラベル情報を取得するプロパティ
'
'   備考: なし
'
Public Property Get Labels(Index As Integer) As String
    Labels = mstrLabels(Index)
End Property

'
'   機能: タイトル文字列を取得する
'
'   備考: なし
'
Public Property Get Title() As String
    Title = mstrTitle
End Property


'
'   機能: 登録頭数を取得する
'
'   備考: なし
'
Public Property Get TOSU() As Long
    TOSU = mlngTosu
End Property

'
'   機能: 払い戻しグリッドの変換情報のはじまる行
'
'   備考: なし
'
Public Property Get HenkanRow() As Long
    HenkanRow = mlngHenkanRow
End Property


'
'   機能: ラップタイムグリッドのコーナー通過順情報のはじまる行
'
'   備考: なし
'
Public Property Get CornerRow() As Long
    CornerRow = mlngCornerRow
End Property


'
'   機能: 最初に開かれるタブの番号
'
'   備考: なし
'
Public Property Get FirstTabNumber() As Long
    FirstTabNumber = mlngFirstTabNumber
End Property


'
'   機能: 過去走データの取得状態
'
'   備考: なし
'
Public Property Get NowKakoFetching() As Boolean
    NowKakoFetching = mblnNowKakoFetching
End Property

'
'   機能: レコード詳細画面用のキーを返す
'
'   備考: なし
'
Public Property Get RCKey() As String
    RCKey = mstrRCKey
End Property


'
'   機能: G1レコード詳細画面用のキーを返す
'
'   備考: なし
'
Public Property Get RCG1Key() As String
    RCG1Key = mstrRCG1Key
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
'   機能: 障害レースなら真
'
'   備考: なし
'
Public Property Get IsShougai() As Boolean
    IsShougai = mblnIsShougai
End Property


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   外部関数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8


'
'   機能: キャンセル
'
'   備考: なし
'
Public Sub CancelFetching()
On Error GoTo Errorhandler
    gApp.Log "RA CancelFetching"
    If mblnNowFetching Then
        mblnCancelFetching = True
        
        freers mRS_UMA_RACE_A
        freers mRS_UMA_RACE_B
        freers mRS_UMA
        freers mRS_BANUSI
        freers mRS_CHOKYO
        freers mRS_KISHU
        freers mRS_SEISAN
        freers mRS_HANSYOKU
        freers mRS_HARAI
        freers mRS_Time_K
        freers mRS_TimeJK
        
        freers mRS_TENKO_BABA
        freers mRS_BATAIJYU
        freers mRS_TORIKESI_JYOGAI
        freers mRS_KISHU_CHANGE
        freers mRS_MINING
        freers mRS_COURSE_CHANGE
        freers mRS_HASSOU_CHANGE
        
        freecn mAsyncCN_UMA_RACE_A
        freecn mCN_RACE
        freecn mCN_UMA_RACE_B
        freecn mCN_UMA
        freecn mCN_BANUSI
        freecn mCN_CHOKYO
        freecn mCN_KISHU
        freecn mCN_SEISAN
        freecn mCN_HANSYOKU
        freecn mCN_HARAI
        
        mblnNowFetching = False
    End If
    Exit Sub
Errorhandler:
    gApp.ErrLog
    Resume Next
End Sub



'
'   機能: キャンセル(過去走)
'
'   備考: なし
'
Public Sub CancelKakoFetching()
    gApp.Log "RA CancelKakoFetching"

    mblnCancelKakoFetching = True
End Sub


'
'   機能: 過去走タブ用取得を開始する
'
'   備考: なし
'
Public Sub FetchKako()
    mblnCancelKakoFetching = False
    Call DirectMakeData_Kako
End Sub


'
'   機能: 取得を開始する
'
'   備考: 初期表示用データの取得
'
Public Function Fetch(ByRef objKey As clsKeyRA) As Boolean
On Error GoTo Errorhandler
    Dim rs           As ADODB.Recordset
    
    Dim lngNumRecord   As Long            '' レース数
    Dim strSQL As String          '' SQL文用
    Dim i              As Long            '' ループカウンタ
    Dim lngCP          As Long            '' セル入力用
    Dim lngRP          As Long            '' セル入力用
    Dim str            As String          '' 一次利用
    Dim intMDBIndex    As Integer         '' 分割MDBのインデクス番号

    Set mKey = objKey
    
    mblnNowFetching = True
   
    ' コネクションがオープンできていなければ False を返して終了
    If CheckCNOpened = False Then
        Fetch = False
        Exit Function
    End If
   
    '-------------
    ' RACE情報取得
    '-------------

    Set rs = New ADODB.Recordset
    Call OpenTableDirect(rs, mCN_RACE, "RACE")
    Call SafeSeek(rs, mKey.FieldArray, mKey.ValueArray)

    ' データが無かったらFalseを返して終了
    If rs.EOF Or rs.BOF Then
        rs.Close
        Set rs = Nothing
        Fetch = False
        Exit Function
    End If

    ' RACEレコードをすべて構造体で保持する
    Call SetDataFromRS_RA(rs, mBuf_RA)
    ' 該当のCC,TCがあるとき構造体に上書きする
    Call Set_CCTC
    
    ' タイトル文字列を作る
    mstrTitle = objKey.ReadableString _
                & mCC.YOBI7(rs("YoubiCD")) & " " _
                & mCC.KIBJ3(rs("JyoCD")) _
                & mSC.KN1(rs("Kaiji") & rs("Nichiji"))
    
    '頭数を取得、印刷ダイアログに必要
    mlngTosu = mBuf_RA.TorokuTosu
    
    ' G1 J.G1 の場合、真
    mblnIsG1 = (mBuf_RA.GradeCD = "A" Or mBuf_RA.GradeCD = "F")

    ' 障害レースなら真
    mblnIsShougai = (mBuf_RA.TrackCD = "27" Or mBuf_RA.TrackCD = "28" Or (mBuf_RA.TrackCD >= "51" And mBuf_RA.TrackCD <= "59"))

    ' 速報が受け取り可能な場合、真
    If (mBuf_RA.head.DataKubun >= "2" And mBuf_RA.head.DataKubun <= "6") Then
        mblnIsPrompt = True
    End If
    Call OpenHenkoRecordset
    
    ' レース情報を作る
    Call MakeData_Race
        
    rs.Close

    ' 初期表示タブ番号の選択
    ' データ区分に応じて最初に表示するタブの番号を返す
    Select Case mBuf_RA.head.DataKubun
    Case "3" To "7", "A", "B"
        mlngFirstTabNumber = 6
    Case Else
        mlngFirstTabNumber = 0
    End Select

    ' レコード詳細画面用のキー
    mstrRCKey = "1" & mKey.str & "0000" & mBuf_RA.JyokenInfo.SyubetuCD & mBuf_RA.KYORI & mBuf_RA.TrackCD
    mstrRCG1Key = "2" & mKey.str & mBuf_RA.RaceInfo.TokuNum & mBuf_RA.JyokenInfo.SyubetuCD & mBuf_RA.KYORI & mBuf_RA.TrackCD

    '-----------------
    ' 非同期取得を行う
    '-----------------
    
    strSQL = "SELECT * FROM UMA_RACE_A"
    strSQL = strSQL & " WHERE "
    strSQL = strSQL & "     Year     = '" & mKey.Year & "'"
    strSQL = strSQL & " AND MonthDay = '" & mKey.MonthDay & "'"
    strSQL = strSQL & " AND JyoCD    = '" & mKey.JyoCD & "'"
    strSQL = strSQL & " AND RaceNum  = '" & mKey.RaceNum & "'"
    strSQL = strSQL & " ORDER BY Umaban"
    
    gApp.Log strSQL
    mAsyncCN_UMA_RACE_A.Execute strSQL, , adAsyncExecute

    gApp.Log "RA Fetch Complete"
    Fetch = True
    Exit Function
Errorhandler:
    gApp.ErrLog
    Resume Next
End Function ' Fetch


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
    
    Set mCN_RACE = gApp.GetCN_RACE
    Set mAsyncCN_UMA_RACE_A = gApp.GetCN_UMA_RACE_A
    Set mCN_UMA_RACE_B = gApp.GetCN_UMA_RACE_B
    Set mCN_UMA = gApp.GetCN_UMA
    Set mCN_BANUSI = gApp.GetCN_BANUSI
    Set mCN_CHOKYO = gApp.GetCN_CHOKYO
    Set mCN_KISHU = gApp.GetCN_KISHU
    Set mCN_SEISAN = gApp.GetCN_SEISAN
    Set mCN_HANSYOKU = gApp.GetCN_HANSYOKU
    Set mCN_HARAI = gApp.GetCN_HARAI
 
End Sub

'
'   機能: クラス終了イベント
'
'   備考: なし
'
Private Sub Class_Terminate()
    Set mCC = Nothing
    Set mSC = Nothing
End Sub


'
'   機能: コネクションがオープンされているか調べる
'
'   備考: なし
'
Private Function CheckCNOpened() As Boolean
    Dim out As Boolean
    out = True
    out = out And (mAsyncCN_UMA_RACE_A.State And adStateOpen)
    out = out And (mCN_BANUSI.State And adStateOpen)
    out = out And (mCN_CHOKYO.State And adStateOpen)
    out = out And (mCN_HANSYOKU.State And adStateOpen)
    out = out And (mCN_HARAI.State And adStateOpen)
    out = out And (mCN_KISHU.State And adStateOpen)
    out = out And (mCN_RACE.State And adStateOpen)
    out = out And (mCN_SEISAN.State And adStateOpen)
    out = out And (mCN_UMA.State And adStateOpen)
    out = out And (mCN_UMA_RACE_B.State And adStateOpen)
    CheckCNOpened = out
End Function


'
'   機能: 変更情報のレコードセットを開く
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
'   機能: 競走馬レースレコードセットの取得完了イベントハンドラ
'
'   備考: なし
'
Private Sub mAsyncCN_UMA_RACE_A_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
On Error GoTo Errorhandler
    gApp.Log ">mAsyncCN_UMA_RACE_A_ExecuteComplete"
    
    Dim strSQL As String
    Dim i As Long
    Dim blnFirst As Boolean
    
    If pError Is Nothing Then
    
        Set mRS_UMA_RACE_A = pRecordset
        
        If mRS_UMA_RACE_A.EOF Then
            RaiseEvent NoUMARACE
        Else
            Set mRS_UMA_RACE_B = New ADODB.Recordset
            If OpenTableDirect(mRS_UMA_RACE_B, mCN_UMA_RACE_B, "UMA_RACE_B") = False Then
                Set mRS_UMA_RACE_B = Nothing
            End If
            
            Set mRS_UMA = New ADODB.Recordset
            If OpenTableDirect(mRS_UMA, mCN_UMA, "UMA") = False Then
                Set mRS_UMA = Nothing
            End If
            
            Set mRS_BANUSI = New ADODB.Recordset
            If OpenTableDirect(mRS_BANUSI, mCN_BANUSI, "BANUSI") = False Then
                Set mRS_BANUSI = Nothing
            End If
            
            Set mRS_CHOKYO = New ADODB.Recordset
            If OpenTableDirect(mRS_CHOKYO, mCN_CHOKYO, "CHOKYO") = False Then
                Set mRS_CHOKYO = Nothing
            End If
            
            Set mRS_KISHU = New ADODB.Recordset
            If OpenTableDirect(mRS_KISHU, mCN_KISHU, "KISHU") = False Then
                Set mRS_KISHU = Nothing
            End If
            
            Set mRS_SEISAN = New ADODB.Recordset
            If OpenTableDirect(mRS_SEISAN, mCN_SEISAN, "SEISAN") = False Then
                Set mRS_SEISAN = Nothing
            End If
        
            Set mRS_HANSYOKU = New ADODB.Recordset
            If OpenTableDirect(mRS_HANSYOKU, mCN_HANSYOKU, "HANSYOKU") = False Then
                Set mRS_HANSYOKU = Nothing
            End If
            
            Set mRS_HARAI = New ADODB.Recordset
            If OpenTableDirect(mRS_HARAI, mCN_HARAI, "HARAI") = False Then
                Set mRS_HARAI = Nothing
            End If
            
            Call DirectMakeData_Kihon
            If mblnCancelFetching Then
                mblnNowFetching = False
                Exit Sub
            End If
            
            Call DirectMakeData_Seiseki
            If mblnCancelFetching Then
                mblnNowFetching = False
                Exit Sub
            End If
            Call DirectMakeData_Ketto
            If mblnCancelFetching Then
                mblnNowFetching = False
                Exit Sub
            End If
            Call DirectMakeData_Mining
            If mblnCancelFetching Then
                mblnNowFetching = False
                Exit Sub
            End If
            Call DirectMakeData_JokenBetu
            If mblnCancelFetching Then
                mblnNowFetching = False
                Exit Sub
            End If
            If Not (mBuf_RA.TrackCD = "27" Or mBuf_RA.TrackCD = "28" Or (mBuf_RA.TrackCD >= "51" And mBuf_RA.TrackCD <= "59")) Then
                Call DirectMakeData_MochiTime
            End If
            If mblnCancelFetching Then
                mblnNowFetching = False
                Exit Sub
            End If
            Call DirectMakeData_Kako
        End If
    Else
        With pError
            gApp.Log .Description & .SQLState & .Source & .Number
        End With
        RaiseEvent NoUMARACE
    End If
    
    mblnNowFetching = False

    freers mRS_BANUSI
    freers mRS_CHOKYO
    freers mRS_KISHU
    freers mRS_SEISAN
    freers mRS_HANSYOKU
    freers mRS_HARAI
    freers mRS_Time_K
    freers mRS_TimeJK
    
    freers mRS_TENKO_BABA
    freers mRS_BATAIJYU
    freers mRS_KISHU_CHANGE
    freers mRS_MINING
    freers mRS_COURSE_CHANGE
    freers mRS_HASSOU_CHANGE


    freecn mCN_RACE
    freecn mCN_BANUSI
    freecn mCN_CHOKYO
    freecn mCN_KISHU
    freecn mCN_SEISAN
    freecn mCN_HANSYOKU
    freecn mCN_HARAI
        
        
    gApp.Log "<mAsyncCN_UMA_RACE_A_ExecuteComplete"
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   機能: レース情報を作る
'
'   備考: なし
'
Private Sub MakeData_Race()
On Error GoTo Errorhandler
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
                    gApp.Log "本賞金:" & .Honsyokin(i)
                    If Right$(str, 2) = ".0" Then
                        str = Left$(str, Len(str) - 2)
                    End If
                End If
            Next i
            str = str & " 万円"
        End If
        str = str & "  "
        
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
                    gApp.Log "付加賞金:" & .Fukasyokin(i)
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
Errorhandler:
    gApp.ErrLog
    Resume Next
End Sub

'
'   機能: 基本情報グリッドを作る
'
'   備考: なし
'
Private Sub DirectMakeData_Kihon()
On Error GoTo Errorhandler
    
    Const rowOffset As Long = 1 '' データ行オフセット
    
    Dim lngCP  As Long   '' カラムポインタ
    Dim i      As Long   '' ループカウンタ
    Dim strTmp As String '' 一時作業用文字列
    Dim strTmp2 As String '' 一時作業用文字列
    
    Dim gd As clsGridData
    
    Dim UA As ADODB.Recordset   '' 馬毎レース　前半
    Dim UB As ADODB.Recordset   '' 馬毎レース　後半
    Dim UM As ADODB.Recordset   '' 競走馬マスタ
    Dim KS As ADODB.Recordset   '' 騎手マスタ
    Dim CH As ADODB.Recordset   '' 調教師マスタ
    Dim BN As ADODB.Recordset   '' 馬主マスタ
    Dim BR As ADODB.Recordset   '' 生産者マスタ
    Dim AV As ADODB.Recordset   '' 出走取消・競走除外
    Dim KC As ADODB.Recordset   '' 騎手変更

    Dim bufBataiju As JV_WH_BATAIJYU
    Dim dicKC As New Scripting.Dictionary
    
    Set gd = New clsGridData
    
    Set UA = mRS_UMA_RACE_A
    Set UB = mRS_UMA_RACE_B
    Set UM = mRS_UMA
    Set KS = mRS_KISHU
    Set CH = mRS_CHOKYO
    Set BN = mRS_BANUSI
    Set BR = mRS_SEISAN
    Set AV = mRS_TORIKESI_JYOGAI
    Set KC = mRS_KISHU_CHANGE
    
    ' 馬体重変更情報
    Call SafeSeek(mRS_BATAIJYU, mKey.FieldArray, mKey.ValueArray)
    If Not mRS_BATAIJYU.EOF Then
        Call SetDataFromRS_WH(mRS_BATAIJYU, bufBataiju)
    End If
    
    ' 騎手変更情報
    KC.Close
    KC.Open "SELECT * FROM KISHU_CHANGE " & mKey.SQLWHEREString & " ORDER BY HappyoTime", gApp.GetCN_KISHU_CHANGE, adOpenForwardOnly, adLockReadOnly
    Do While Not KC.EOF
        If dicKC.Exists(KC("Umaban").value) Then
            Call dicKC.Remove(KC("Umaban").value)
        End If
        dicKC.Add KC("Umaban").value, KC("HappyoTime").value
        KC.MoveNext
    Loop
    KC.Close
    Call OpenTableDirect(KC, gApp.GetCN_KISHU_CHANGE, "KISHU_CHANGE")
    
    With gd
        .Rows = 10 + rowOffset
        .Cols = 22
        
        ' カラムヘッダ挿入
        lngCP = 0
        .SetItemMatrix 0, lngCP, "枠", "枠番", ">_"
        .SetItemMatrix 0, lngCP, "番", "馬番", ">_"
        .SetItemMatrix 0, lngCP, "B", "ブリンカー"
        .SetItemMatrix 0, lngCP, "馬記号"
        .SetItemMatrix 0, lngCP, "馬名"
        
        .SetItemMatrix 0, lngCP, "性齢"
        .SetItemMatrix 0, lngCP, "毛", "毛色"
        .SetItemMatrix 0, lngCP, "習"
        .SetItemMatrix 0, lngCP, "騎手"
        .SetItemMatrix 0, lngCP, "負担", "単位kg", ">_"
        
        .SetItemMatrix 0, lngCP, "馬体重", , ">_"
        .SetItemMatrix 0, lngCP, "増減", , ">_"
        .SetItemMatrix 0, lngCP, "本賞金累計", , ">_"
        .SetItemMatrix 0, lngCP, "収得賞金", , ">_"
        .SetItemMatrix 0, lngCP, "調教師"
        
        .SetItemMatrix 0, lngCP, "馬主"
        .SetItemMatrix 0, lngCP, "生産者"
        .SetItemMatrix 0, lngCP, "逃げ回数"
        .SetItemMatrix 0, lngCP, "先行回数"
        .SetItemMatrix 0, lngCP, "差し回数"
        
        .SetItemMatrix 0, lngCP, "追込回数"
        .SetItemMatrix 0, lngCP, "服色"
    End With
    
    ' グリッド挿入
    UA.MoveFirst
    
    i = 0
    Do While Not UA.EOF
        DoEvents
        If mblnCancelFetching Then
            Exit Sub
        End If
        
        ' 必要データを取得する
        
        SafeSeek UB, Array("B_Year", "B_MonthDay", "B_JyoCD", "B_RaceNum", "B_Umaban", "B_KettoNum"), _
                    Array(UA("Year").value, UA("MonthDay").value, UA("JyoCD").value, UA("RaceNum").value, UA("Umaban").value, UA("KettoNum").value)
        SafeSeek UM, Array("KettoNum"), _
                    Array(UA("KettoNum").value)
        SafeSeek KS, Array("KisyuCode"), _
                    Array(UB("KisyuCode").value)
        SafeSeek CH, Array("ChokyosiCode"), _
                    Array(UA("ChokyosiCode").value)
        SafeSeek BN, Array("BanusiCode"), _
                    Array(UA("BanusiCode").value)
        If Not UM.EOF Then
            SafeSeek BR, Array("BreederCode"), _
                        Array(UM("BreederCode").value)
        End If
        SafeSeek AV, Array("Year", "MonthDay", "JyoCD", "Kaiji", "Nichiji", "RaceNum", "Umaban"), _
                    Array(UA("Year").value, UA("MonthDay").value, UA("JyoCD").value, UA("Kaiji").value, UA("Nichiji").value, UA("RaceNum").value, UA("Umaban").value)
        If dicKC.Exists(UA("Umaban").value) Then
            SafeSeek KC, Array("Year", "MonthDay", "JyoCD", "Kaiji", "Nichiji", "RaceNum", "HappyoTime", "Umaban"), _
                        Array(UA("Year").value, UA("MonthDay").value, UA("JyoCD").value, UA("Kaiji").value, UA("Nichiji").value, UA("RaceNum").value, dicKC.item(UA("Umaban").value), UA("Umaban").value)
        End If
        
        
        lngCP = 0
        With gd
            ' 枠番
            .SetItemMatrix i + rowOffset, lngCP, mSC.FitSpaceNum(UA("Wakuban"), 2), , ">-", , , gApp.GetWakubanColor(val(UA("Wakuban"))), Contrast(gApp.GetWakubanColor(val(UA("Wakuban"))))
            ' 馬番
            .SetItemMatrix i + rowOffset, lngCP, mSC.FitSpaceNum(UA("Umaban"), 2), , ">-"
            ' ブリンカー
            .SetItemMatrix i + rowOffset, lngCP, IIf(UB("Blinker") = "0", "", "B")
            ' 馬記号
            .SetItemMatrix i + rowOffset, lngCP, mCC.UMKG1(UA("UmaKigoCD")), mCC.UMKG3(UA("UmaKigoCD"))
            ' 馬名
            Call BameiAV(AV, UA, UB, UM, gd, i + rowOffset, lngCP)
            ' 馬　性別年齢
            .SetItemMatrix i + rowOffset, lngCP, mCC.SEIB4(UA("SexCD")) & val(UA("Barei"))
            ' 毛色
            .SetItemMatrix i + rowOffset, lngCP, mCC.KEIR1(UA("KeiroCD"))
            ' 見習い区分、騎手名、負担重量
            ' 騎手変更情報を配慮
            If UA("DataKubun") = "2" And dicKC.Exists(UA("Umaban").value) Then
                SafeSeek KS, Array("KisyuCode"), _
                            Array(KC("AtoKisyuCode").value)
                If KC("MaeMinaraiCD") = KC("AtoMinaraiCD") Then
                    .SetItemMatrix i + rowOffset, lngCP, mCC.KSMN1(UB("MinaraiCD")), mCC.KSMN3(UB("MinaraiCD"))
                Else
                    .SetItemMatrix i + rowOffset, lngCP, mCC.KSMN1(KC("AtoMinaraiCD")), mCC.KSMN1(KC("MaeMinaraiCD")) & "→" & mCC.KSMN1(KC("AtoMinaraiCD")), , , , , RGB(255, 0, 0)
                End If
                If Trim$(IfExist(KS, "KisyuRyakusyo")) <> "" Then '騎手変更名を略称で表示(04-033-4)
                    .SetItemMatrix i + rowOffset, lngCP, Trim$(KS("KisyuRyakusyo")), Trim$(KC("MaeKisyuName")) & "→" & Trim$(KC("AtoKisyuName")), , "KS", IfExist(KS, "KisyuCode"), , RGB(255, 0, 0)
                Else
                    .SetItemMatrix i + rowOffset, lngCP, Trim$(KC("AtoKisyuName")), Trim$(KC("MaeKisyuName")) & "→" & Trim$(KC("AtoKisyuName")), , "KS", IfExist(KS, "KisyuCode"), , RGB(255, 0, 0)
                End If
                If KC("AtoFutan") = UB("Futan") Then
                    .SetItemMatrix i + rowOffset, lngCP, IIf(val(UB("Futan")) = 0, "", Format$(UB("Futan") / 10, "#0.0")), , ">-"
                Else
                    .SetItemMatrix i + rowOffset, lngCP, IIf(val(KC("AtoFutan")) = 0, "", Format$(KC("AtoFutan") / 10, "#0.0")), IIf(val(UB("Futan")) = 0, "", Format$(UB("Futan") / 10, "#0.0")) & "→" & IIf(val(KC("AtoFutan")) = 0, "", Format$(KC("AtoFutan") / 10, "#0.0")), ">-", , , , RGB(255, 0, 0)
                End If
            Else
                If val(UB("MinaraiCDBefore")) = 0 Then
                    .SetItemMatrix i + rowOffset, lngCP, mCC.KSMN1(UB("MinaraiCD")), mCC.KSMN3(UB("MinaraiCD"))
                Else
                    .SetItemMatrix i + rowOffset, lngCP, mCC.KSMN1(UB("MinaraiCD")), mCC.KSMN1(UB("MinaraiCDBefore")) & "→" & mCC.KSMN1(UB("MinaraiCD")), ">-", , , , RGB(255, 0, 0)
                End If
                If Trim$(UB("KisyuRyakusyoBefore")) = "" Then
                    .SetItemMatrix i + rowOffset, lngCP, Trim$(UB("KisyuRyakusyo")), , , "KS", IfExist(KS, "KisyuCode")
                Else
                    .SetItemMatrix i + rowOffset, lngCP, Trim$(UB("KisyuRyakusyo")), Trim$(UB("KisyuRyakusyoBefore")) & "→" & Trim$(UB("KisyuRyakusyo")), , "KS", IfExist(KS, "KisyuCode"), , &HFF
                End If
                If val(UB("FutanBefore")) = 0 Then
                    .SetItemMatrix i + rowOffset, lngCP, IIf(val(UB("Futan")) = 0, "", Format$(UB("Futan") / 10, "#0.0")), , ">-"
                Else
                    .SetItemMatrix i + rowOffset, lngCP, IIf(val(UB("Futan")) = 0, "", Format$(UB("Futan") / 10, "#0.0")), IIf(val(UB("FutanBefore")) = 0, "", Format$(UB("FutanBefore") / 10, "#0.0")) & "→" & IIf(val(UB("Futan")) = 0, "", Format$(UB("Futan") / 10, "#0.0")), ">-", , , , RGB(255, 0, 0)
                End If
            End If
            
            ' 馬体重変更情報を配慮
            If UA("DataKubun").value <= "2" And Not mRS_BATAIJYU.EOF Then
                ' 馬体重変更情報がある場合
                If val(bufBataiju.BataijyuInfo(i).BaTaijyu) < 2 Then
                    .SetItemMatrix i + rowOffset, lngCP, "     ", , ">-"
                ElseIf bufBataiju.BataijyuInfo(i).BaTaijyu = "999" Then
                    .SetItemMatrix i + rowOffset, lngCP, " --- ", , ">-"
                Else
                    .SetItemMatrix i + rowOffset, lngCP, val(bufBataiju.BataijyuInfo(i).BaTaijyu), , ">-"
                End If
                strTmp = "" ' 表示項目
                Select Case bufBataiju.BataijyuInfo(i).ZogenSa
                Case "999", "   "
                    strTmp = ""
                Case "000"
                    strTmp = "±0"
                Case Else
                    strTmp = bufBataiju.BataijyuInfo(i).ZogenFugo & mSC.ValStr(bufBataiju.BataijyuInfo(i).ZogenSa)
                End Select
                strTmp2 = ""  ' ツールチップ
                Select Case UB("ZogenSa")
                Case "999", "   "
                    strTmp2 = ""
                Case "000"
                    strTmp2 = "±0"
                Case Else
                    strTmp2 = UB("ZogenFugo") & mSC.ValStr(UB("ZogenSa"))
                End Select
                .SetItemMatrix i + rowOffset, lngCP, strTmp, strTmp2, ">-" ' 速報馬体重のあるときの増減差
            Else
                ' ない場合
                If val(UB("BaTaijyu")) < 2 Then
                    .SetItemMatrix i + rowOffset, lngCP, "     ", , ">-"
                ElseIf UB("BaTaijyu") = "999" Then
                    .SetItemMatrix i + rowOffset, lngCP, " --- ", , ">-"
                Else
                    .SetItemMatrix i + rowOffset, lngCP, val(UB("BaTaijyu")), , ">-"
                End If
                strTmp = ""
                Select Case UB("ZogenSa")
                Case "999", "   "
                    strTmp = ""
                Case "000"
                    strTmp = "±0"
                Case Else
                    strTmp = UB("ZogenFugo") & mSC.ValStr(UB("ZogenSa"))
                End Select
                .SetItemMatrix i + rowOffset, lngCP, strTmp, , ">-"
            End If
            ' 本賞金
            .SetItemMatrix i + rowOffset, lngCP, mSC.Money3(val(IfExist(UM, "RuikeiHonsyoHeiti")) + val(IfExist(UM, "RuikeiHonsyoSyogai"))), , ">-"
            ' 収得賞金
            If mBuf_RA.TrackCD >= "51" And mBuf_RA.TrackCD <= "59" Then
            ' 障害の場合
                .SetItemMatrix i + rowOffset, lngCP, mSC.Money3(val(IfExist(UM, "RuikeiSyutokuSyogai"))), , ">-"
            Else
            ' 障害ではない場合
                .SetItemMatrix i + rowOffset, lngCP, mSC.Money3(val(IfExist(UM, "RuikeiSyutokuHeichi"))), , ">-"
            End If
            ' 東西区分＋調教師名
            .SetItemMatrix i + rowOffset, lngCP, mCC.TZSZ4(UA("TozaiCD")) & Trim$(UA("ChokyosiRyakusyo")), , , "CH", IfExist(CH, "ChokyosiCode")
            ' 馬主名
            .SetItemMatrix i + rowOffset, lngCP, Trim$(UA("BanusiName")), , , "BN", IfExist(BN, "BanusiCode")
            ' 生産者名
            .SetItemMatrix i + rowOffset, lngCP, Trim$(IfExist(UM, "BreederName")), , , "BR", IfExist(BR, "BreederCode")
            ' 脚質 逃げ
            .SetItemMatrix i + rowOffset, lngCP, mSC.ValStr(IfExist(UM, "Kyakusitu1")), , ">-"
            ' 脚質 先行
            .SetItemMatrix i + rowOffset, lngCP, mSC.ValStr(IfExist(UM, "Kyakusitu2")), , ">-"
            ' 脚質 刺し
            .SetItemMatrix i + rowOffset, lngCP, mSC.ValStr(IfExist(UM, "Kyakusitu3")), , ">-"
            ' 脚質 追込
            .SetItemMatrix i + rowOffset, lngCP, mSC.ValStr(IfExist(UM, "Kyakusitu4")), , ">-"
            ' 服色
            .SetItemMatrix i + rowOffset, lngCP, Trim$(UA("Fukusyoku"))
        End With
        
        UA.MoveNext
        
        gd.ItemMatrix(i + rowOffset, 4).ToolTip = ""

        i = i + 1
        If gd.Rows = i + rowOffset Then
            gd.Rows = gd.Rows + 10
        End If
    Loop
    gd.Rows = i + rowOffset
    
    RaiseEvent FetchedKihon(gd)
    Exit Sub
    
Errorhandler:
    gApp.ErrLog
End Sub



'
'   機能: 血統グリッドを作る
'
'   備考: なし
'
Private Sub DirectMakeData_Ketto()
On Error GoTo Errorhandler
    
    Const rowOffset As Long = 1 '' データ行オフセット
    
    Dim lngCP As Long   '' カラムポインタ
    Dim i     As Long   '' ループカウンタ
    
    
    Dim gd As clsGridData
    
    Dim UA As ADODB.Recordset
    Dim UB As ADODB.Recordset
    Dim UM As ADODB.Recordset
    Dim HN As ADODB.Recordset
    Dim AV As ADODB.Recordset   '' 出走取消・競走除外

    Set gd = New clsGridData
    Set UA = mRS_UMA_RACE_A
    Set UB = mRS_UMA_RACE_B
    Set UM = mRS_UMA
    Set HN = mRS_HANSYOKU
    Set AV = mRS_TORIKESI_JYOGAI
    
    With gd
        .Rows = 8 + rowOffset
        .Cols = 6
    
        ' カラムヘッダ登録
        lngCP = 0
        .SetItemMatrix 0, lngCP, "枠", "枠番", ">_"
        .SetItemMatrix 0, lngCP, "番", "馬番", ">_"
        .SetItemMatrix 0, lngCP, "馬名"
        .SetItemMatrix 0, lngCP, "親(父/母)"
        .SetItemMatrix 0, lngCP, "二代"
        .SetItemMatrix 0, lngCP, "三代"
    
    End With
    
    ' グリッドデータ挿入
    UA.MoveFirst
    i = 0
    
    Do While Not UA.EOF
        DoEvents
        If mblnCancelFetching Then
            Exit Sub
        End If
        SafeSeek UB, Array("B_Year", "B_MonthDay", "B_JyoCD", "B_RaceNum", "B_Umaban", "B_KettoNum"), _
                    Array(UA("Year").value, UA("MonthDay").value, UA("JyoCD").value, UA("RaceNum").value, UA("Umaban").value, UA("KettoNum").value)
        SafeSeek UM, Array("KettoNum"), Array(UA("KettoNum").value)
        SafeSeek AV, Array("Year", "MonthDay", "JyoCD", "Kaiji", "Nichiji", "RaceNum", "Umaban"), _
                    Array(UA("Year").value, UA("MonthDay").value, UA("JyoCD").value, UA("Kaiji").value, UA("Nichiji").value, UA("RaceNum").value, UA("Umaban").value)
        
        Call KettoSub(UA, UB, UM, HN, AV, gd, i, 0, "1", "3", "7")
        Call KettoSub(UA, UB, UM, HN, AV, gd, i, 1, "1", "3", "8")
        Call KettoSub(UA, UB, UM, HN, AV, gd, i, 2, "1", "4", "9")
        Call KettoSub(UA, UB, UM, HN, AV, gd, i, 3, "1", "4", "10")
        Call KettoSub(UA, UB, UM, HN, AV, gd, i, 4, "2", "5", "11")
        Call KettoSub(UA, UB, UM, HN, AV, gd, i, 5, "2", "5", "12")
        Call KettoSub(UA, UB, UM, HN, AV, gd, i, 6, "2", "6", "13")
        Call KettoSub(UA, UB, UM, HN, AV, gd, i, 7, "2", "6", "14")
        UA.MoveNext
        i = i + 1
        If gd.Rows <= i * 8 + rowOffset Then
            gd.Rows = gd.Rows + 8 * 4
        End If
    Loop
    gd.Rows = i * 8 + rowOffset
    
    RaiseEvent FetchedKetto(gd)
    
    Exit Sub
Errorhandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   機能: 血統情報作成用サブルーチン
'
'   備考: 引き数 UA       - 馬毎情報を含む ADODB.Recordset
'       　　     UM       - 馬マスタ情報を含む ADODB.Recordset
'       　　     HN       - 繁殖馬情報を含む ADODB.Recordset
'                gd       - 書き込み先グリッドデータ
'                RotI     - 何頭目かのインデックス
'                RowP     - 一頭毎内の書き込み先Rowインデックス
'                Field1~3 - 父母等の組み合わせ
'
Private Sub KettoSub(UA As ADODB.Recordset, UB As ADODB.Recordset, UM As ADODB.Recordset, HN As ADODB.Recordset, AV As ADODB.Recordset _
                    , gd As clsGridData, RowI As Long, RowP As Long _
                    , Field1 As String, Field2 As String, Field3 As String)
    Dim lngCP As Long
    
    Const rowOffset As Long = 1 '' データ行オフセット
    
    lngCP = 0
    With gd
        .SetItemMatrix (RowI * 8) + RowP + rowOffset, lngCP, IIf(RowI Mod 2, " ", "") & mSC.FitSpaceNum(UA("Wakuban"), 2), , ">-", , , gApp.GetWakubanColor(val(UA("Wakuban"))), Contrast(gApp.GetWakubanColor(val(UA("Wakuban"))))
        .SetItemMatrix (RowI * 8) + RowP + rowOffset, lngCP, IIf(RowI Mod 2, " ", "") & mSC.FitSpaceNum(UA("Umaban"), 2), , ">-"
        Call BameiAV(AV, UA, UB, UM, gd, (RowI * 8) + RowP + rowOffset, lngCP)
        
        If Not UM.EOF Then
            SafeSeek HN, Array("HansyokuNum"), Array(UM("Ketto3InfoHansyokuNum" & Field1).value)
            .SetItemMatrix (RowI * 8) + RowP + rowOffset, lngCP, KettoGetName(UM, HN, Field1), IfExist(HN, "HansyokuNum"), , "HN", IfExist(HN, "HansyokuNum"), IIf(val(Field1) Mod 2 = 0, ColorMother, ColorFather), KettoFGColor(UM, HN, Field1)
            SafeSeek HN, Array("HansyokuNum"), Array(UM("Ketto3InfoHansyokuNum" & Field2).value)
            .SetItemMatrix (RowI * 8) + RowP + rowOffset, lngCP, KettoGetName(UM, HN, Field2), IfExist(HN, "HansyokuNum"), , "HN", IfExist(HN, "HansyokuNum"), IIf(val(Field2) Mod 2 = 0, ColorMother, ColorFather), KettoFGColor(UM, HN, Field2)
            SafeSeek HN, Array("HansyokuNum"), Array(UM("Ketto3InfoHansyokuNum" & Field3).value)
            .SetItemMatrix (RowI * 8) + RowP + rowOffset, lngCP, KettoGetName(UM, HN, Field3), IfExist(HN, "HansyokuNum"), , "HN", IfExist(HN, "HansyokuNum"), IIf(val(Field3) Mod 2 = 0, ColorMother, ColorFather), KettoFGColor(UM, HN, Field3)
        Else
            .SetItemMatrix (RowI * 8) + RowP + rowOffset, lngCP, String(CLng(Field1), "."), IfExist(HN, "HansyokuNum"), , , , IIf(val(Field1) Mod 2 = 0, ColorMother, ColorFather), IIf(val(Field1) Mod 2 = 0, ColorMother, ColorFather)
            .SetItemMatrix (RowI * 8) + RowP + rowOffset, lngCP, String(CLng(Field2), "."), IfExist(HN, "HansyokuNum"), , , , IIf(val(Field2) Mod 2 = 0, ColorMother, ColorFather), IIf(val(Field2) Mod 2 = 0, ColorMother, ColorFather)
            .SetItemMatrix (RowI * 8) + RowP + rowOffset, lngCP, String(CLng(Field3), "."), IfExist(HN, "HansyokuNum"), , , , IIf(val(Field3) Mod 2 = 0, ColorMother, ColorFather), IIf(val(Field3) Mod 2 = 0, ColorMother, ColorFather)
        End If
    End With
End Sub


'
'   機能: 血統情報から馬名を得る
'
'   備考:
'
Private Function KettoGetName(UM As ADODB.Recordset, HN As ADODB.Recordset, FieldNum As String) As String
    Dim str As String
    If Not HN.EOF Then
        str = Trim$(HN("Bamei"))
        If str <> "" Then
            KettoGetName = str
            Exit Function
        End If
        
        str = Trim$(HN("BameiEng"))
        If str <> "" Then
            KettoGetName = str
        Else
            KettoGetName = String(CLng(FieldNum), ".")
        End If
    Else
        str = Trim$(UM("Ketto3InfoBamei" & FieldNum))
        If str <> "" Then
            KettoGetName = str
        Else
            KettoGetName = String(CLng(FieldNum), ".")
        End If
    End If
End Function

'
'   機能: 馬名が無い場合はダミー文字が入るため、それは見えないように背景色と同化する
'
'   備考:
'
Private Function KettoFGColor(UM As ADODB.Recordset, HN As ADODB.Recordset, FieldNum As String) As Long
    Dim color As Long
    If Not HN.EOF Then
        If Trim$(HN("Bamei")) <> "" Then
            color = vbBlack
            Exit Function
        End If
        
        If Trim$(HN("BameiEng")) <> "" Then
            color = vbBlack
            Exit Function
        End If
    Else
        If Trim$(UM("Ketto3InfoBamei" & FieldNum)) <> "" Then
            color = vbBlack
            Exit Function
        End If
    End If
    KettoFGColor = IIf(val(FieldNum) Mod 2 = 0, ColorMother, ColorFather)
End Function


'
'   機能: マイニンググリッドを作る
'
'   備考: なし
'
Private Sub DirectMakeData_Mining()
On Error GoTo Errorhandler
    
    Const rowOffset As Long = 1 '' データ行オフセット
    
    Dim lngCP As Long        '' カラムポインタ
    Dim i     As Long        '' ループカウンタ
    Dim gd    As clsGridData '' グリッドデータ
    Dim strDMKubun As String '' マイニング区分
    
    Dim UA As ADODB.Recordset
    Dim UB As ADODB.Recordset
    Dim MN As ADODB.Recordset
    Dim UM As ADODB.Recordset
    Dim AV As ADODB.Recordset

    Set UA = mRS_UMA_RACE_A
    Set UB = mRS_UMA_RACE_B
    Set MN = mRS_MINING
    Set UM = mRS_UMA
    Set AV = mRS_TORIKESI_JYOGAI
    
    Set gd = New clsGridData
    
    With gd
        .Rows = 28 + rowOffset ' 適当
        .Cols = 7

        ' カラムヘッダ登録
        lngCP = 0
        .SetItemMatrix 0, lngCP, "枠", "枠番", ">_"
        .SetItemMatrix 0, lngCP, "番", "馬番", ">_"
        .SetItemMatrix 0, lngCP, "馬名"
        .SetItemMatrix 0, lngCP, "予想タイム", , ">_"
        .SetItemMatrix 0, lngCP, "予想誤差＋", , ">_"
        .SetItemMatrix 0, lngCP, "予想誤差－", , ">_"
        .SetItemMatrix 0, lngCP, "予想順位", , ">_"
    End With
    
    SafeSeek MN, Array("Year", "MonthDay", "JyoCD", "Kaiji", "Nichiji", "RaceNum"), _
                Array(mKey.Year, mKey.MonthDay, mKey.JyoCD, mKey.Kaiji, mKey.Nichiji, mKey.RaceNum)
    If mBuf_RA.head.DataKubun <= "6" And Not MN.EOF Then
        UA.MoveFirst
        strDMKubun = mSC.SE_DMKBN(MN("DataKubun"))
        i = 0
        Do While Not UA.EOF
            SafeSeek UB, Array("B_Year", "B_MonthDay", "B_JyoCD", "B_RaceNum", "B_Umaban", "B_KettoNum"), _
                        Array(UA("Year").value, UA("MonthDay").value, UA("JyoCD").value, UA("RaceNum").value, UA("Umaban").value, UA("KettoNum").value)
            SafeSeek UM, Array("KettoNum"), Array(UA("KettoNum").value)
            SafeSeek AV, Array("Year", "MonthDay", "JyoCD", "Kaiji", "Nichiji", "RaceNum", "Umaban"), _
                        Array(UA("Year").value, UA("MonthDay").value, UA("JyoCD").value, UA("Kaiji").value, UA("Nichiji").value, UA("RaceNum").value, UA("Umaban").value)
            
            DoEvents
            If mblnCancelFetching Then
                Exit Sub
            End If
            lngCP = 0
            With gd
                .SetItemMatrix i + rowOffset, lngCP, mSC.FitSpaceNum(UA("Wakuban"), 2), , ">^", , , gApp.GetWakubanColor(val(UA("Wakuban"))), Contrast(gApp.GetWakubanColor(val(UA("Wakuban"))))
                .SetItemMatrix i + rowOffset, lngCP, mSC.FitSpaceNum(UA("Umaban"), 2), , ">^"
                Call BameiAV(AV, UA, UB, UM, gd, i + rowOffset, lngCP)
                If UA("Umaban") >= "01" And UA("Umaban") <= "18" Then
                    .SetItemMatrix i + rowOffset, lngCP, IIf(val(MN("DMTime" & val(UA("Umaban")))) = 0, "", mSC.MSSss(Trim$(MN("DMTime" & val(UA("Umaban")))))), , ">^"
                    .SetItemMatrix i + rowOffset, lngCP, IIf(val(MN("DMTime" & val(UA("Umaban")))) = 0, "", Format$(val(MN("DMGosaP" & val(UA("Umaban"))) / 100), "#0.00")), , ">^"
                    .SetItemMatrix i + rowOffset, lngCP, IIf(val(MN("DMTime" & val(UA("Umaban")))) = 0, "", Format$(val(MN("DMGosaM" & val(UA("Umaban"))) / 100), "#0.00")), , ">^"
                    If .ItemMatrix(i + rowOffset, 2).Strikethru Then
                        .SetItemMatrix i + rowOffset, lngCP, .ItemMatrix(i + rowOffset, 2).ToolTip
                    Else
                        .SetItemMatrix i + rowOffset, lngCP, IIf(val(MN("DMTime" & val(UA("Umaban")))) = 0, "", mSC.MSSss(Trim$(MN("DMTime" & val(UA("Umaban")))))), , ">^"
                    End If
                End If
            End With
            UA.MoveNext
            i = i + 1
            If gd.Rows <= i + rowOffset Then
                gd.Rows = gd.Rows + 10
            End If
        Loop
    Else
        ' グリッドデータ挿入
        UA.MoveFirst
        i = 0
        Do While Not UA.EOF
            SafeSeek UB, Array("B_Year", "B_MonthDay", "B_JyoCD", "B_RaceNum", "B_Umaban", "B_KettoNum"), _
                        Array(UA("Year").value, UA("MonthDay").value, UA("JyoCD").value, UA("RaceNum").value, UA("Umaban").value, UA("KettoNum").value)
            SafeSeek UM, Array("KettoNum"), Array(UA("KettoNum").value)
            SafeSeek AV, Array("Year", "MonthDay", "JyoCD", "Kaiji", "Nichiji", "RaceNum", "Umaban"), _
                        Array(UA("Year").value, UA("MonthDay").value, UA("JyoCD").value, UA("Kaiji").value, UA("Nichiji").value, UA("RaceNum").value, UA("Umaban").value)
            DoEvents
            If mblnCancelFetching Then
                Exit Sub
            End If
            lngCP = 0
            With gd
                .SetItemMatrix i + rowOffset, lngCP, mSC.FitSpaceNum(UA("Wakuban"), 2), , ">^", , , gApp.GetWakubanColor(val(UA("Wakuban"))), Contrast(gApp.GetWakubanColor(val(UA("Wakuban"))))
                .SetItemMatrix i + rowOffset, lngCP, mSC.FitSpaceNum(UA("Umaban"), 2), , ">^"
                Call BameiAV(AV, UA, UB, UM, gd, i + rowOffset, lngCP)
                .SetItemMatrix i + rowOffset, lngCP, IIf(val(UB("DMTime")) = 0, "", mSC.MSSss(UB("DMTime"))), , ">^"
                .SetItemMatrix i + rowOffset, lngCP, IIf(val(UB("DMTime")) = 0, "", Format$(val(UB("DMGosaP")) / 100, "#0.00")), , ">^"
                .SetItemMatrix i + rowOffset, lngCP, IIf(val(UB("DMTime")) = 0, "", Format$(val(UB("DMGosaM")) / 100, "#0.00")), , ">^"
                .SetItemMatrix i + rowOffset, lngCP, IIf(val(UB("DMJyuni")) = 0, "", val(UB("DMJyuni"))), , ">^"
            End With
            UA.MoveNext
            
            If strDMKubun = "" Then
                strDMKubun = mSC.SE_DMKBN(UB("DMKubun"))
            End If
            
            i = i + 1
            If gd.Rows <= i + rowOffset Then
                gd.Rows = gd.Rows + 10
            End If
        Loop
    End If
    
    gd.Rows = i + rowOffset
    
    RaiseEvent FetchedMining(gd, strDMKubun)
    
    Exit Sub

Errorhandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   機能: 条件別グリッドを作る
'
'   備考: なし
'
Private Sub DirectMakeData_JokenBetu()
On Error GoTo Errorhandler
    
    Const rowOffset As Long = 2 '' データ行オフセット
        
    Dim lngCP     As Long '' カラムポインタ
    Dim lngNumSet As Long '' セット数
    Dim i         As Long '' ループカウンタ
    Dim j         As Long '' ループカウンタ
    Dim k         As Long '' ループカウンタ
    Dim r         As Single '' 赤色成分
    Dim b         As Single '' 青色成分
    Dim gd        As clsGridData
    
    Dim UA As ADODB.Recordset
    Dim UB As ADODB.Recordset
    Dim UM As ADODB.Recordset
    Dim AV As ADODB.Recordset

    Set UA = mRS_UMA_RACE_A
    Set UB = mRS_UMA_RACE_B
    Set UM = mRS_UMA
    Set AV = mRS_TORIKESI_JYOGAI
    
    Set gd = New clsGridData
    
    ' １～５着＆着外を１セットとして、何セット表示するかどうかは、馬場によって違う
    Select Case mBuf_RA.TrackCD
    Case "10" To "26", "28" ' 芝・ダート
        lngNumSet = 8
    Case "51" To "59"       ' 障害
        lngNumSet = 6
    Case Else               ' サンド・その他
        lngNumSet = 1
    End Select
    
    ' カラムヘッダ
    With gd
        
        .Rows = 28 + rowOffset ' ２段ヘッダ
        ' カラム数は馬場によって違う
        .Cols = 3 + (6 * lngNumSet)
        
        ' １行目 カラムヘッダ登録
        lngCP = 0
        .SetItemMatrix 0, lngCP, "枠", "枠番", ">-"
        .SetItemMatrix 0, lngCP, "番", "馬番", ">-"
        .SetItemMatrix 0, lngCP, "馬名", "馬名", "<-"
        For i = 0 To 5
            .SetItemMatrix 0, lngCP, "総合", "総合着回数(中央＋地方＋海外)", "^_"
        Next i
        ' 馬場あわせて選択
        Select Case mBuf_RA.TrackCD
        Case "10" To "22"        ' 芝
            For i = 0 To 5
                .SetItemMatrix 0, lngCP, "芝", "芝トラックでの総合着回数", "^_"
            Next i
            ' 左右直
            Select Case mBuf_RA.TrackCD
            Case "10"           ' 直線
                For i = 0 To 5
                    .SetItemMatrix 0, lngCP, "芝直", "芝・直線トラックでの着回数", "^_"
                Next i
            Case "11" To "16"   ' 左
                For i = 0 To 5
                    .SetItemMatrix 0, lngCP, "芝左", "芝・左回りトラックでの着回数", "^_"
                Next i
            Case Else           ' 右
                For i = 0 To 5
                    .SetItemMatrix 0, lngCP, "芝右", "芝・右回りトラックでの着回数", "^_"
                Next i
            End Select
            ' 距離
            Select Case mBuf_RA.KYORI
            Case Is <= "1600"
                For i = 0 To 5
                    .SetItemMatrix 0, lngCP, "芝1600m以下", "芝トラック1600m以下での着回数", "^_"
                Next i
            Case Is >= "2201"
                For i = 0 To 5
                    .SetItemMatrix 0, lngCP, "芝2201m以上", "芝トラック2201m以上での着回数", "^_"
                Next i
            Case Else
                For i = 0 To 5
                    .SetItemMatrix 0, lngCP, "芝1601m～2200m", "芝トラック1601m～2200mでの着回数", "^_"
                Next i
            End Select
            
            For i = 0 To 5
                .SetItemMatrix 0, lngCP, "芝良", "芝トラック、馬場状態「良」での着回数", "^_"
            Next i
            For i = 0 To 5
                .SetItemMatrix 0, lngCP, "芝稍重", "芝トラック、馬場状態「稍重」での着回数", "^_"
            Next i
            For i = 0 To 5
                .SetItemMatrix 0, lngCP, "芝重", "芝トラック、馬場状態「稍重」での着回数", "^_"
            Next i
            For i = 0 To 5
                .SetItemMatrix 0, lngCP, "芝不良", "芝トラック、馬場状態「不良」での着回数", "^_"
            Next i
        Case "23" To "26", "29"  ' ダート
            For i = 0 To 5
                .SetItemMatrix 0, lngCP, "ダート", "ダートトラックでの総合着回数", "^_"
            Next i
            ' 左右直
            Select Case mBuf_RA.TrackCD
            Case "29"           ' 直線
                For i = 0 To 5
                    .SetItemMatrix 0, lngCP, "ダート直", "ダート・直線トラックでの着回数", "^_"
                Next i
            Case "23", "25"     ' 左
                For i = 0 To 5
                    .SetItemMatrix 0, lngCP, "ダート左", "ダート・左回りトラックでの着回数", "^_"
                Next i
            Case "24", "26"     ' 右
                For i = 0 To 5
                    .SetItemMatrix 0, lngCP, "ダート右", "ダート・右回りトラックでの着回数", "^_"
                Next i
            End Select
            ' 距離
            Select Case mBuf_RA.KYORI
            Case Is <= "1600"
                For i = 0 To 5
                    .SetItemMatrix 0, lngCP, "ダート1600m以下", "ダートトラック1600m以下での着回数", "^_"
                Next i
            Case Is >= "2201"
                For i = 0 To 5
                    .SetItemMatrix 0, lngCP, "ダート2201m以上", "ダートトラック2201m以上での着回数", "^_"
                Next i
            Case Else
                For i = 0 To 5
                    .SetItemMatrix 0, lngCP, "ダート1601m～2200m", "ダートトラック1601m～2200mでの着回数", "^_"
                Next i
            End Select
            
            For i = 0 To 5
                .SetItemMatrix 0, lngCP, "ダート良", "ダートトラック、馬場状態「良」での着回数", "^_"
            Next i
            For i = 0 To 5
                .SetItemMatrix 0, lngCP, "ダート稍重", "ダートトラック、馬場状態「稍重」での着回数", "^_"
            Next i
            For i = 0 To 5
                .SetItemMatrix 0, lngCP, "ダート重", "ダートトラック、馬場状態「稍重」での着回数", "^_"
            Next i
            For i = 0 To 5
                .SetItemMatrix 0, lngCP, "ダート不良", "ダートトラック、馬場状態「不良」での着回数", "^_"
            Next i
        
        Case "51" To "59"        ' 障害
            For i = 0 To 5
                .SetItemMatrix 0, lngCP, "障害", "障害レースでの総合着回数", "^_"
            Next i
            ' 障害左右直は無い
            ' 障害距離別は無い
            
            For i = 0 To 5
                .SetItemMatrix 0, lngCP, "障害良", "障害レース、馬場状態「良」での着回数", "^_"
            Next i
            For i = 0 To 5
                .SetItemMatrix 0, lngCP, "障害稍重", "障害レース、馬場状態「稍重」での着回数", "^_"
            Next i
            For i = 0 To 5
                .SetItemMatrix 0, lngCP, "障害重", "障害レース、馬場状態「稍重」での着回数", "^_"
            Next i
            For i = 0 To 5
                .SetItemMatrix 0, lngCP, "障害不良", "障害レース、馬場状態「不良」での着回数", "^_"
            Next i
        End Select
                

        ' 2行目 カラムヘッダ登録
        lngCP = 0
        .SetItemMatrix 1, lngCP, "枠", "枠番", ">-"
        .SetItemMatrix 1, lngCP, "番", "馬番", ">-"
        .SetItemMatrix 1, lngCP, "馬名", , "<-"
        For i = 0 To lngNumSet - 1
            .SetItemMatrix 1, lngCP, "1着", .ItemMatrix(0, lngCP).ToolTip, "^_"
            .SetItemMatrix 1, lngCP, "2着", .ItemMatrix(0, lngCP).ToolTip, "^_"
            .SetItemMatrix 1, lngCP, "3着", .ItemMatrix(0, lngCP).ToolTip, "^_"
            .SetItemMatrix 1, lngCP, "4着", .ItemMatrix(0, lngCP).ToolTip, "^_"
            .SetItemMatrix 1, lngCP, "5着", .ItemMatrix(0, lngCP).ToolTip, "^_"
            .SetItemMatrix 1, lngCP, "着外", .ItemMatrix(0, lngCP).ToolTip, "^_"
        Next i
    End With
    
    ' グリッドデータ挿入
    UA.MoveFirst
    i = 0
    Do While Not UA.EOF
        DoEvents
        If mblnCancelFetching Then
            Exit Sub
        End If
        SafeSeek UB, Array("B_Year", "B_MonthDay", "B_JyoCD", "B_RaceNum", "B_Umaban", "B_KettoNum"), _
            Array(UA("Year").value, UA("MonthDay").value, UA("JyoCD").value, UA("RaceNum").value, UA("Umaban").value, UA("KettoNum").value)
        SafeSeek UM, Array("KettoNum"), Array(UA("KettoNum").value)
        SafeSeek AV, Array("Year", "MonthDay", "JyoCD", "Kaiji", "Nichiji", "RaceNum", "Umaban"), _
                    Array(UA("Year").value, UA("MonthDay").value, UA("JyoCD").value, UA("Kaiji").value, UA("Nichiji").value, UA("RaceNum").value, UA("Umaban").value)
        lngCP = 0
        With gd
            .SetItemMatrix i + rowOffset, lngCP, mSC.FitSpaceNum(UA("Wakuban"), 2), , ">^", , , gApp.GetWakubanColor(val(UA("Wakuban"))), Contrast(gApp.GetWakubanColor(val(UA("Wakuban"))))
            .SetItemMatrix i + rowOffset, lngCP, mSC.FitSpaceNum(UA("Umaban"), 2), , ">^"
            Call BameiAV(AV, UA, UB, UM, gd, i + rowOffset, lngCP)
            For j = 1 To 6
                .SetItemMatrix i + rowOffset, lngCP, mSC.ValStr(IfExist(UM, "SogoChakukaisu" & j)), , "^^"
            Next j
            
            ' 馬場あわせて選択
            Select Case mBuf_RA.TrackCD
            Case "10" To "22"        ' 芝
                For j = 1 To 6
                    If IfExist(UM, "Ba1Chakukaisu" & j) <> "" And IfExist(UM, "Ba2Chakukaisu" & j) <> "" And IfExist(UM, "Ba3Chakukaisu" & j) <> "" Then
                        .SetItemMatrix i + rowOffset, lngCP, val(UM("Ba1Chakukaisu" & j)) + val(UM("Ba2Chakukaisu" & j)) + val(UM("Ba3Chakukaisu" & j)), , "^_"
                    Else
                        .SetItemMatrix i + rowOffset, lngCP, "", , "^_"
                    End If
                Next j
                ' 左右直
                Select Case mBuf_RA.TrackCD
                Case "10"           ' 直線
                    For j = 1 To 6
                        .SetItemMatrix i + rowOffset, lngCP, mSC.ValStr(IfExist(UM, "Ba1Chakukaisu" & j)), , "^_" '芝直
                    Next j
                Case "11" To "16"   ' 左
                    For j = 1 To 6
                        .SetItemMatrix i + rowOffset, lngCP, mSC.ValStr(IfExist(UM, "Ba3Chakukaisu" & j)), , "^_" '芝左
                    Next j
                Case Else           ' 右
                    For j = 1 To 6
                        .SetItemMatrix i + rowOffset, lngCP, mSC.ValStr(IfExist(UM, "Ba2Chakukaisu" & j)), , "^_" '芝右
                    Next j
                End Select
                ' 距離
                Select Case mBuf_RA.KYORI
                Case Is <= "1600"
                    For j = 1 To 6
                        .SetItemMatrix i + rowOffset, lngCP, mSC.ValStr(IfExist(UM, "Kyori1Chakukaisu" & j)), , "^_"  '芝1600m以下
                    Next j
                Case Is >= "2201"
                    For j = 1 To 6
                        .SetItemMatrix i + rowOffset, lngCP, mSC.ValStr(IfExist(UM, "Kyori3Chakukaisu" & j)), , "^_"  '芝2201m以上
                    Next j
                Case Else
                    For j = 1 To 6
                        .SetItemMatrix i + rowOffset, lngCP, mSC.ValStr(IfExist(UM, "Kyori2Chakukaisu" & j)), , "^_"  '芝1601m～2200m
                    Next j
                End Select
                
                For j = 1 To 6
                    .SetItemMatrix i + rowOffset, lngCP, mSC.ValStr(IfExist(UM, "Jyotai1Chakukaisu" & j)), , "^_"  '芝良
                Next j
                For j = 1 To 6
                    .SetItemMatrix i + rowOffset, lngCP, mSC.ValStr(IfExist(UM, "Jyotai2Chakukaisu" & j)), , "^_"  '芝稍重
                Next j
                For j = 1 To 6
                    .SetItemMatrix i + rowOffset, lngCP, mSC.ValStr(IfExist(UM, "Jyotai3Chakukaisu" & j)), , "^_"  '芝重
                Next j
                For j = 1 To 6
                    .SetItemMatrix i + rowOffset, lngCP, mSC.ValStr(IfExist(UM, "Jyotai4Chakukaisu" & j)), , "^_"  '芝不良
                Next j
                Case "23" To "26", "29"  ' ダート
                For j = 1 To 6
                    If IfExist(UM, "Ba4Chakukaisu" & j) <> "" And IfExist(UM, "Ba5Chakukaisu" & j) <> "" And IfExist(UM, "Ba6Chakukaisu" & j) <> "" Then
                        .SetItemMatrix i + rowOffset, lngCP, val(UM("Ba4Chakukaisu" & j)) + val(UM("Ba5Chakukaisu" & j)) + val(UM("Ba6Chakukaisu" & j)), , "^_" 'ダート
                    Else
                        .SetItemMatrix i + rowOffset, lngCP, "", , "^_" 'ダート
                    End If
                Next j
                ' 左右直
                Select Case mBuf_RA.TrackCD
                Case "29"           ' 直線
                    For j = 1 To 6
                        .SetItemMatrix i + rowOffset, lngCP, mSC.ValStr(IfExist(UM, "Ba4Chakukaisu" & j)), , "^_" 'ダート直
                    Next j
                Case "23", "25"     ' 左
                    For j = 1 To 6
                        .SetItemMatrix i + rowOffset, lngCP, mSC.ValStr(IfExist(UM, "Ba6Chakukaisu" & j)), , "^_" 'ダート左
                    Next j
                Case "24", "26"     ' 右
                    For j = 1 To 6
                        .SetItemMatrix i + rowOffset, lngCP, mSC.ValStr(IfExist(UM, "Ba5Chakukaisu" & j)), , "^_" 'ダート右
                    Next j
                End Select
                ' 距離
                Select Case mBuf_RA.KYORI
                Case Is <= "1600"
                    For j = 1 To 6
                        .SetItemMatrix i + rowOffset, lngCP, mSC.ValStr(IfExist(UM, "Kyori4Chakukaisu" & j)), , "^_" 'ダート1600m以下
                    Next j
                Case Is >= "2201"
                    For j = 1 To 6
                        .SetItemMatrix i + rowOffset, lngCP, mSC.ValStr(IfExist(UM, "Kyori6Chakukaisu" & j)), , "^_" 'ダート2201m以上
                    Next j
                Case Else
                    For j = 1 To 6
                        .SetItemMatrix i + rowOffset, lngCP, mSC.ValStr(IfExist(UM, "Kyori5Chakukaisu" & j)), , "^_" 'ダート1601m～2200m
                    Next j
                End Select
                
                For j = 1 To 6
                    .SetItemMatrix i + rowOffset, lngCP, mSC.ValStr(IfExist(UM, "Jyotai5Chakukaisu" & j)), , "^_" 'ダート良
                Next j
                For j = 1 To 6
                    .SetItemMatrix i + rowOffset, lngCP, mSC.ValStr(IfExist(UM, "Jyotai6Chakukaisu" & j)), , "^_" 'ダート稍重
                Next j
                For j = 1 To 6
                    .SetItemMatrix i + rowOffset, lngCP, mSC.ValStr(IfExist(UM, "Jyotai7Chakukaisu" & j)), , "^_" 'ダート重
                Next j
                For j = 1 To 6
                    .SetItemMatrix i + rowOffset, lngCP, mSC.ValStr(IfExist(UM, "Jyotai8Chakukaisu" & j)), , "^_" 'ダート不良
                Next j
            
            Case "51" To "59"        ' 障害
                For j = 1 To 6
                    .SetItemMatrix i + rowOffset, lngCP, mSC.ValStr(IfExist(UM, "Ba7Chakukaisu" & j)), , "^_" '障害
                Next j
                ' 障害左右直は無い
                ' 障害距離別は無い
                
                For j = 1 To 6
                    .SetItemMatrix i + rowOffset, lngCP, mSC.ValStr(IfExist(UM, "Jyotai9Chakukaisu" & j)), , "^_" '障害良
                Next j
                For j = 1 To 6
                    .SetItemMatrix i + rowOffset, lngCP, mSC.ValStr(IfExist(UM, "Jyotai10Chakukaisu" & j)), , "^_" '障害稍重
                Next j
                For j = 1 To 6
                    .SetItemMatrix i + rowOffset, lngCP, mSC.ValStr(IfExist(UM, "Jyotai11Chakukaisu" & j)), , "^_" '障害重
                Next j
                For j = 1 To 6
                    .SetItemMatrix i + rowOffset, lngCP, mSC.ValStr(IfExist(UM, "Jyotai12Chakukaisu" & j)), , "^_" '障害不良
                Next j
            End Select

        End With
        UA.MoveNext
        i = i + 1
        If gd.Rows <= i + rowOffset Then
            gd.Rows = gd.Rows + 10
        End If
    Loop
    
    gd.Rows = i + rowOffset
    
    ' 背景色を設定する
    For i = 0 To ((gd.Cols - 3) / 6) - 1
        For j = 0 To 5
            For k = 2 To gd.Rows - 1
                gd.ItemMatrix(k, 3 + (i * 6) + j).BGColor = IIf(i Mod 2 = 0, RGB(230, 230, 255), RGB(245, 245, 255))
            Next k
        Next j
    Next i
    
    
    RaiseEvent FetchedJokenBetu(gd)
    Exit Sub
Errorhandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   機能: 成績グリッドを作る
'
'   備考: なし
'
Private Sub DirectMakeData_Seiseki()
On Error GoTo Errorhandler

    Const rowOffset As Long = 1 '' データ行オフセット
    
    Dim lngCP   As Long   '' カラムポインタ
    Dim i       As Long   '' ループカウンタ
    Dim j       As Long   '' ループカウンタ
    Dim lngFlag As Long
    Dim gd1     As clsGridData
    Dim gd2     As clsGridData
    Dim gd3     As clsGridData
    Dim UA      As ADODB.Recordset
    Dim UB      As ADODB.Recordset
    Dim UM      As ADODB.Recordset
    Dim AV      As ADODB.Recordset
    Dim setTmpForCellStr As String '' セル文字列作業用一時変数
    Dim KS      As ADODB.Recordset  '騎手マスタ

    Set gd1 = New clsGridData
    Set gd2 = New clsGridData
    Set gd3 = New clsGridData
    Set UA = mRS_UMA_RACE_A
    Set UB = mRS_UMA_RACE_B
    Set UM = mRS_UMA
    Set AV = mRS_TORIKESI_JYOGAI
    
    Set KS = mRS_KISHU
    
    With gd1
        .Rows = 28 + rowOffset
        .Cols = 22

        ' １行目 カラムヘッダ登録
        lngCP = 0
        .SetItemMatrix 0, lngCP, "隠", "ソート用", ">-"
        .SetItemMatrix 0, lngCP, "入", "入線順位", ">-"
        .SetItemMatrix 0, lngCP, "着", "確定着順", ">-"
        .SetItemMatrix 0, lngCP, "異常", "異常区分", ">-"
        .SetItemMatrix 0, lngCP, "枠", "枠番", ">-"
        .SetItemMatrix 0, lngCP, "番", "馬番", ">-"
        .SetItemMatrix 0, lngCP, "B", "ブリンカー"
        .SetItemMatrix 0, lngCP, "馬名", "馬名", "^-"
        .SetItemMatrix 0, lngCP, "性齢", "性齢"
        .SetItemMatrix 0, lngCP, "負担", "負担重量", ">-"
        .SetItemMatrix 0, lngCP, "馬体重", "馬体重", ">-"
        .SetItemMatrix 0, lngCP, "増減", "馬体重増減", ">-"
        .SetItemMatrix 0, lngCP, "騎手", "騎手名"
        .SetItemMatrix 0, lngCP, "タイム", "走破タイム", ">-"
        .SetItemMatrix 0, lngCP, "着差", "前馬との着差", ">-"
        .SetItemMatrix 0, lngCP, "1C", "1コーナーでの順位", ">-"
        .SetItemMatrix 0, lngCP, "2C", "2コーナーでの順位", ">-"
        .SetItemMatrix 0, lngCP, "3C", "3コーナーでの順位", ">-"
        .SetItemMatrix 0, lngCP, "4C", "4コーナーでの順位", ">-"
        If mBuf_RA.TrackCD >= "51" And mBuf_RA.TrackCD <= "59" Then
        '障害の時は上1Fと表示
            .SetItemMatrix 0, lngCP, "上1F", "上1ハロンタイム", ">-"
        Else
        '障害以外の場合は後3Fと表示
            .SetItemMatrix 0, lngCP, "後3F", "後3ハロンタイム", ">-"
        End If
        .SetItemMatrix 0, lngCP, "人", "単勝人気順", ">-"
        .SetItemMatrix 0, lngCP, "単オッズ", "単勝オッズ", ">-"
    End With
    
    ' グリッドデータ挿入
    UA.MoveFirst
    i = 0
    Do While Not UA.EOF
        SafeSeek UB, Array("B_Year", "B_MonthDay", "B_JyoCD", "B_RaceNum", "B_Umaban", "B_KettoNum"), _
                    Array(UA("Year").value, UA("MonthDay").value, UA("JyoCD").value, UA("RaceNum").value, UA("Umaban").value, UA("KettoNum").value)
        SafeSeek UM, Array("KettoNum"), Array(UA("KettoNum").value)
        SafeSeek AV, Array("Year", "MonthDay", "JyoCD", "Kaiji", "Nichiji", "RaceNum", "Umaban"), _
                    Array(UA("Year").value, UA("MonthDay").value, UA("JyoCD").value, UA("Kaiji").value, UA("Nichiji").value, UA("RaceNum").value, UA("Umaban").value)
        SafeSeek KS, Array("KisyuCode"), _
                    Array(UB("KisyuCode").value)
        
        DoEvents
        If mblnCancelFetching Then
            Exit Sub
        End If
    
        ' ３着確定、５着確定のデータの場合、確定外データは表示しない
        If Not ( _
            (mBuf_RA.head.DataKubun = "3" Or mBuf_RA.head.DataKubun = "4") And _
            UB("KakuteiJyuni") = "00" _
        ) Then
            
            lngCP = 0
            With gd1
                ' ソート用隠しカラムの設定
                setTmpForCellStr = ""  ' 一時的文字列変数を初期化
                If UB("KakuteiJyuni") <> "00" Then
                    setTmpForCellStr = UB("KakuteiJyuni")
                Else
                    setTmpForCellStr = "99"
                End If
                ' 異常区分同士はソート用順位を使用
                setTmpForCellStr = setTmpForCellStr & mSC.IJKB_Sort(UB("IjyoCD"))
                ' 同着同士の場合、着差が同着の方を最下位
                If UB("ChakusaCD") = "D  " Then
                    setTmpForCellStr = setTmpForCellStr & "1"
                Else
                    setTmpForCellStr = setTmpForCellStr & "0"
                End If
                setTmpForCellStr = setTmpForCellStr & UA("Umaban")
                .SetItemMatrix i + rowOffset, lngCP, setTmpForCellStr
                
                If UB("NyusenJyuni") <> "00" Then
                    .SetItemMatrix i + rowOffset, lngCP, mSC.ValStr(UB("NyusenJyuni")), , ">^"
                Else
                    .SetItemMatrix i + rowOffset, lngCP, "99", , ">^", , , &HFFFFFF, &HFFFFFF
                End If
                
                If UB("KakuteiJyuni") <> "00" Then
                    .SetItemMatrix i + rowOffset, lngCP, mSC.ValStr(UB("KakuteiJyuni")), , ">^", , , gApp.GetChakujyunColor(val(UB("KakuteiJyuni"))), SortString:=UB("KakuteiJyuni")
                Else
                    .SetItemMatrix i + rowOffset, lngCP, "99", , ">^", , , &HFFFFFF, &HFFFFFF, SortString:="99"
                End If
                .SetItemMatrix i + rowOffset, lngCP, mCC.IJKB2(UB("IjyoCD")), , ">^", SortString:=mSC.IJKB_Sort(UB("IjyoCD"))
                .SetItemMatrix i + rowOffset, lngCP, mSC.FitSpaceNum(UA("Wakuban"), 2), , ">^", , , gApp.GetWakubanColor(val(UA("Wakuban"))), Contrast(gApp.GetWakubanColor(val(UA("Wakuban"))))
                .SetItemMatrix i + rowOffset, lngCP, mSC.FitSpaceNum(UA("Umaban"), 2), , ">^", SortString:=UA("Umaban")
                .SetItemMatrix i + rowOffset, lngCP, IIf(UB("Blinker") = "0", "", "B")  'ブリンカー
                Call BameiAV(AV, UA, UB, UM, gd1, i + rowOffset, lngCP)
                .SetItemMatrix i + rowOffset, lngCP, mCC.SEIB4(UA("SexCD")) & mSC.ValStr(UA("Barei"))   '性齢
                If val(UB("FutanBefore")) = 0 Then
                    .SetItemMatrix i + rowOffset, lngCP, IIf(val(UB("Futan")) = 0, "", Format$(UB("Futan") / 10, "#0.0")), , ">-"
                Else
                    .SetItemMatrix i + rowOffset, lngCP, IIf(val(UB("Futan")) = 0, "", Format$(UB("Futan") / 10, "#0.0")), IIf(val(UB("FutanBefore")) = 0, "", Format$(UB("FutanBefore") / 10, "#0.0")) & "→" & IIf(val(UB("Futan")) = 0, "", Format$(UB("Futan") / 10, "#0.0")), ">-", , , , RGB(255, 0, 0)
                End If
                
                If val(UB("BaTaijyu")) < 2 Then
                    .SetItemMatrix i + rowOffset, lngCP, "     ", , ">-"
                ElseIf UB("BaTaijyu") = "999" Then
                    .SetItemMatrix i + rowOffset, lngCP, " --- ", , ">-"
                Else
                    .SetItemMatrix i + rowOffset, lngCP, val(UB("BaTaijyu")), , ">-"
                End If
                setTmpForCellStr = ""
                Select Case UB("ZogenSa")
                Case "999", "   "
                    setTmpForCellStr = ""
                Case "000"
                    setTmpForCellStr = "±0"
                Case Else
                    setTmpForCellStr = UB("ZogenFugo") & mSC.ValStr(UB("ZogenSa"))
                End Select
                .SetItemMatrix i + rowOffset, lngCP, setTmpForCellStr, , ">-"
                If Trim$(UB("KisyuRyakusyoBefore")) = "" Then
                    .SetItemMatrix i + rowOffset, lngCP, Trim$(UB("KisyuRyakusyo")), , , "KS", IfExist(KS, "KisyuCode")
                Else
                    .SetItemMatrix i + rowOffset, lngCP, Trim$(UB("KisyuRyakusyo")), Trim$(UB("KisyuRyakusyoBefore")) & "→" & Trim$(UB("KisyuRyakusyo")), , "KS", IfExist(KS, "KisyuCode"), , &HFF
                End If
                .SetItemMatrix i + rowOffset, lngCP, IIf(UB("Time") <> "0000", Format$(UB("Time"), "@:@@.@"), ""), , ">^"
                
                '着差コードのp ppの並び順
                setTmpForCellStr = "" ' 一時的文字列変数を初期化
                If UB("ChakusaCDP") <> "   " Then
                    If UB("ChakusaCDPP") <> "   " Then
                        setTmpForCellStr = mCC.CHKS1(UB("ChakusaCDPP")) & "+"
                    End If
                    setTmpForCellStr = setTmpForCellStr & mCC.CHKS1(UB("ChakusaCDP")) & "+"
                End If
                setTmpForCellStr = setTmpForCellStr & mCC.CHKS1(UB("ChakusaCD"))
                ' 「確定一着の同着以外の着差表示をブランクに」する処理
                ' 確定順位が１着で、かつ、同着ではない時に、着差欄は挿入しない。
                If UB("KakuteiJyuni") = "01" And UB("ChakusaCD") <> "D  " Then
                    .SetItemMatrix i + rowOffset, lngCP, ""
                Else
                    .SetItemMatrix i + rowOffset, lngCP, setTmpForCellStr, , ">^"
                End If
                .SetItemMatrix i + rowOffset, lngCP, JyuniNCProc(UB("Jyuni1c")), , ">^"
                .SetItemMatrix i + rowOffset, lngCP, JyuniNCProc(UB("Jyuni2c")), , ">^"
                .SetItemMatrix i + rowOffset, lngCP, JyuniNCProc(UB("Jyuni3c")), , ">^"
                .SetItemMatrix i + rowOffset, lngCP, JyuniNCProc(UB("Jyuni4c")), , ">^"
                If UB("HaronTimeL3") <> "999" And UB("HaronTimeL3") <> "   " Then
                    .SetItemMatrix i + rowOffset, lngCP, mSC.SSSs(UB("HaronTimeL3")), , ">^"
                Else
                    .SetItemMatrix i + rowOffset, lngCP, ""
                End If
                .SetItemMatrix i + rowOffset, lngCP, IIf(val(UB("Ninki")) = 0, "", val(UB("Ninki"))), , ">^"
                If (UB("Odds") = "0000") Then
                    .SetItemMatrix i + rowOffset, lngCP, ""
                ElseIf (IsNumeric(UB("Odds"))) Then '単勝オッズ
                    .SetItemMatrix i + rowOffset, lngCP, Format$(UB("Odds") / 10, "#0.0"), , ">-", , , , ColorHML(Format$(UB("odds") / 10, "#0.0"))
                Else
                    .SetItemMatrix i + rowOffset, lngCP, UB("Odds"), , ">-", , , , RGB(0, 0, 0)
                End If
            End With
            i = i + 1
            If gd1.Rows <= i + rowOffset Then
                gd1.Rows = gd1.Rows + 10
            End If
        End If
       UA.MoveNext
    Loop
    
    gd1.Rows = i + rowOffset
    
    Call DirectMakeData_Harai(gd2)
    Call DirectMakeData_Lap(gd3)
    
    RaiseEvent FetchedSeiseki(gd1, gd2, gd3, lngFlag)
    Exit Sub
Errorhandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   機能: 順位文字列の作成
'
'   備考: なし
'
Private Function JyuniNCProc(strJyuni As String) As String
    If strJyuni = "00" Then
        JyuniNCProc = ""
    Else
        JyuniNCProc = mSC.ValStr(strJyuni)
    End If
End Function

'
'   機能: 払戻グリッドを作る
'
'   備考: なし
'
Private Sub DirectMakeData_Harai(ByRef gd As clsGridData)
On Error GoTo Errorhandler

    Const rowOffset As Long = 1 '' データ行オフセット

    Dim strSQL As String
    Dim lngCP  As Long   '' カラムポインタ
    Dim lngRP  As Long   '' ロウポインタ
    Dim i      As Long
    Dim j      As Long
    Dim lngColor As Long
    
    Dim HR     As ADODB.Recordset
    
    Set HR = mRS_HARAI
    
    With mKey
        Call SafeSeek(HR, Array("Year", "MonthDay", "JyoCD", "Kaiji", "Nichiji", "RaceNum"), _
                            Array(.Year, .MonthDay, .JyoCD, .Kaiji, .Nichiji, .RaceNum))
    End With

    If HR.EOF Then
        Exit Sub
    End If

    With gd
        .Rows = 100
        .Cols = 4

        ' １行目 カラムヘッダ登録
        lngCP = 0
        .SetItemMatrix 0, lngCP, "式別", , ">-"
        .SetItemMatrix 0, lngCP, "馬番", , ">-"
        .SetItemMatrix 0, lngCP, "払戻金", , ">-"
        .SetItemMatrix 0, lngCP, "人気順", , ">-"
    End With

    ' グリッドデータ挿入
    lngRP = rowOffset
    
    Call HaraiSubFuseirituTokubarai(gd, lngRP, "単勝", RGB(223, 223, 255), HR("FuseirituFlag1") = "1", HR("TokubaraiFlag1") = "1")
    For i = 1 To 3
        lngCP = 0
        lngColor = RGB(223, 223, 255)
        If HR("PayTansyo" & "Pay" & i) <> Space(9) And HR("PayTansyoUmaban" & i) <> 0 Then
            gd.SetItemMatrix lngRP, lngCP, "単勝", , ">-", , , lngColor
            gd.SetItemMatrix lngRP, lngCP, mSC.ValStr(HR("PayTansyoUmaban" & i)), , ">-", , , lngColor
            gd.SetItemMatrix lngRP, lngCP, mSC.ValStr(HR("PayTansyoPay" & i)), , ">-", , , lngColor
            gd.SetItemMatrix lngRP, lngCP, val(HR("PayTansyoNinki" & i)), , ">-", , , lngColor
            lngRP = lngRP + 1
        End If
    Next i
    Call HaraiSubFuseirituTokubarai(gd, lngRP, "複勝", RGB(255, 223, 223), HR("FuseirituFlag2") = "1", HR("TokubaraiFlag2") = "1")
    For i = 1 To 5
        lngCP = 0
        lngColor = RGB(255, 223, 223)
        If HR("PayFukusyo" & "Pay" & i) <> Space(9) And HR("PayFukusyoUmaban" & i) <> 0 Then
            gd.SetItemMatrix lngRP, lngCP, "複勝", , ">-", , , RGB(255, 223, 223)
            gd.SetItemMatrix lngRP, lngCP, mSC.ValStr(HR("PayFukusyoUmaban" & i)), , ">-", , , lngColor
            gd.SetItemMatrix lngRP, lngCP, mSC.ValStr(HR("PayFukusyoPay" & i)), , ">-", , , lngColor
            gd.SetItemMatrix lngRP, lngCP, val(HR("PayFukusyoNinki" & i)), , ">-", , , lngColor
            lngRP = lngRP + 1
        End If
    Next i
    Call HaraiSubFuseirituTokubarai(gd, lngRP, "枠連", RGB(223, 255, 223), HR("FuseirituFlag3") = "1", HR("TokubaraiFlag3") = "1")
    For i = 1 To 3
        lngCP = 0
        lngColor = RGB(223, 255, 223)
        If HR("PayWakuren" & "Pay" & i) <> Space(9) And HR("PayWakurenKumiban" & i) <> 0 Then
            gd.SetItemMatrix lngRP, lngCP, "枠連", , ">-", , , lngColor
            gd.SetItemMatrix lngRP, lngCP, Format$(HR("PayWakurenKumiban" & i), "@-@"), , ">-", , , lngColor
            gd.SetItemMatrix lngRP, lngCP, mSC.ValStr(HR("PayWakurenPay" & i)), , ">-", , , lngColor
            gd.SetItemMatrix lngRP, lngCP, val(HR("PayWakurenNinki" & i)), , ">-", , , lngColor
            lngRP = lngRP + 1
        End If
    Next i
    Call HaraiSubFuseirituTokubarai(gd, lngRP, "馬連", RGB(255, 223, 255), HR("FuseirituFlag4") = "1", HR("TokubaraiFlag4") = "1")
    For i = 1 To 3
        lngCP = 0
        lngColor = RGB(255, 223, 255)
        If HR("PayUmaren" & "Pay" & i) <> Space(9) And HR("PayUmarenKumiban" & i) <> 0 Then
            gd.SetItemMatrix lngRP, lngCP, "馬連", , ">-", , , lngColor
            gd.SetItemMatrix lngRP, lngCP, val(Left$(HR("PayUmarenKumiban" & i), 2)) & "-" _
                                        & val(Right$(HR("PayUmarenKumiban" & i), 2)), , ">-", , , lngColor
            gd.SetItemMatrix lngRP, lngCP, mSC.ValStr(HR("PayUmarenPay" & i)), , ">-", , , lngColor
            gd.SetItemMatrix lngRP, lngCP, val(HR("PayUmarenNinki" & i)), , ">-", , , lngColor
            lngRP = lngRP + 1
        End If
    Next i
    Call HaraiSubFuseirituTokubarai(gd, lngRP, "ワイド", RGB(223, 255, 255), HR("FuseirituFlag5") = "1", HR("TokubaraiFlag5") = "1")
    For i = 1 To 7
        lngCP = 0
        lngColor = RGB(223, 255, 255)
        If HR("PayWide" & "Pay" & i) <> Space(9) And HR("PayWideKumiban" & i) <> 0 Then
            gd.SetItemMatrix lngRP, lngCP, "ワイド", , ">-", , , lngColor
            gd.SetItemMatrix lngRP, lngCP, val(Left$(HR("PayWideKumiban" & i), 2)) & "-" _
                                        & val(Right$(HR("PayWideKumiban" & i), 2)), , ">-", , , lngColor
            gd.SetItemMatrix lngRP, lngCP, mSC.ValStr(HR("PayWidePay" & i)), , ">-", , , lngColor
            gd.SetItemMatrix lngRP, lngCP, val(HR("PayWideNinki" & i)), , ">-", , , lngColor
            lngRP = lngRP + 1
        End If
    Next i
    Call HaraiSubFuseirituTokubarai(gd, lngRP, "馬単", RGB(255, 223, 191), HR("FuseirituFlag7") = "1", HR("TokubaraiFlag7") = "1")
    For i = 1 To 6
        lngCP = 0
        lngColor = RGB(255, 223, 191)
        If HR("PayUmatan" & "Pay" & i) <> Space(9) And HR("PayUmatanKumiban" & i) <> 0 Then
            gd.SetItemMatrix lngRP, lngCP, "馬単", , ">-", , , lngColor
            gd.SetItemMatrix lngRP, lngCP, val(Left$(HR("PayUmatanKumiban" & i), 2)) & "→" _
                                        & val(Right$(HR("PayUmatanKumiban" & i), 2)), , ">-", , , lngColor
            gd.SetItemMatrix lngRP, lngCP, mSC.ValStr(HR("PayUmatanPay" & i)), , ">-", , , lngColor
            gd.SetItemMatrix lngRP, lngCP, val(HR("PayUmatanNinki" & i)), , ">-", , , lngColor
            lngRP = lngRP + 1
        End If
    Next i
    Call HaraiSubFuseirituTokubarai(gd, lngRP, "３連複", RGB(255, 255, 223), HR("FuseirituFlag8") = "1", HR("TokubaraiFlag8") = "1")
    For i = 1 To 3
        lngCP = 0
        lngColor = RGB(255, 255, 223)
        If HR("PaySanrenpuku" & "Pay" & i) <> Space(9) And HR("PaySanrenpukuKumiban" & i) <> 0 Then
            gd.SetItemMatrix lngRP, lngCP, "３連複", , ">-", , , lngColor
            gd.SetItemMatrix lngRP, lngCP, val(Left$(HR("PaySanrenpukuKumiban" & i), 2)) & "-" _
                                        & val(Mid$(HR("PaySanrenpukuKumiban" & i), 3, 2)) & "-" _
                                        & val(Right$(HR("PaySanrenpukuKumiban" & i), 2)), , ">-", , , lngColor
            gd.SetItemMatrix lngRP, lngCP, mSC.ValStr(HR("PaySanrenpukuPay" & i)), , ">-", , , lngColor
            gd.SetItemMatrix lngRP, lngCP, val(HR("PaySanrenpukuNinki" & i)), , ">-", , , lngColor
            lngRP = lngRP + 1
        End If
    Next i

    If val(gApp.R_DBVersion) > 0 Then
        Call HaraiSubFuseirituTokubarai(gd, lngRP, "３連単", RGB(255, 255, 255), HR("FuseirituFlag9") = "1", HR("TokubaraiFlag9") = "1")
        For i = 1 To 6
            lngCP = 0
            lngColor = RGB(255, 255, 255)   'とりあえず
            If HR("PaySanrentanPay" & i) <> Space(9) And HR("PaySanrentanKumiban" & i) <> 0 Then 'Space(6)
                gd.SetItemMatrix lngRP, lngCP, "３連単", , ">-", , , lngColor
                gd.SetItemMatrix lngRP, lngCP, val(Mid$(HR("PaySanrentanKumiban" & i), 1, 2)) & "→" _
                                            & val(Mid$(HR("PaySanrentanKumiban" & i), 3, 2)) & "→" _
                                            & val(Mid$(HR("PaySanrentanKumiban" & i), 5, 2)), , ">-", , , lngColor
                gd.SetItemMatrix lngRP, lngCP, mSC.ValStr(HR("PaySanrentanPay" & i)), , ">-", , , lngColor
                gd.SetItemMatrix lngRP, lngCP, val(HR("PaySanrentanNinki" & i)), , ">-", , , lngColor
                lngRP = lngRP + 1
            End If
        Next i
    End If
    

    ' 返還情報

    ' グリッドマージに必要なため返還の行をプロパティに保存
    mlngHenkanRow = lngRP

    For i = 1 To 28
        If HR("HenkanUma" & i) = "1" Then
            lngCP = 0
            With gd
                .SetItemMatrix lngRP, lngCP, "返還", , ">^", , , RGB(192, 192, 192)
                .SetItemMatrix lngRP, lngCP, i & "番返還", , "<-", , , RGB(192, 192, 192)
                .SetItemMatrix lngRP, lngCP, i & "番返還", , "<-", , , RGB(192, 192, 192)
                .SetItemMatrix lngRP, lngCP, i & "番返還", , "<-", , , RGB(192, 192, 192)
            End With
            lngRP = lngRP + 1
        End If
    Next i
    For i = 1 To 8
        If HR("HenkanWaku" & i) = "1" Then
            lngCP = 0
            With gd
                .SetItemMatrix lngRP, lngCP, "返還", , ">^", , , RGB(192, 192, 192)
                .SetItemMatrix lngRP, lngCP, i & "枠返還", , "<-", , , RGB(192, 192, 192)
                .SetItemMatrix lngRP, lngCP, i & "枠返還", , "<-", , , RGB(192, 192, 192)
                .SetItemMatrix lngRP, lngCP, i & "枠返還", , "<-", , , RGB(192, 192, 192)
            End With
            lngRP = lngRP + 1
        End If
    Next i


    For i = 1 To 8
        If HR("HenkanDoWaku" & i) = "1" Then
            lngCP = 0
            With gd
                .SetItemMatrix lngRP, lngCP, "返還", , ">^", , , RGB(192, 192, 192)
                .SetItemMatrix lngRP, lngCP, i & "枠同枠返還", , "<-", , , RGB(192, 192, 192)
                .SetItemMatrix lngRP, lngCP, i & "枠同枠返還", , "<-", , , RGB(192, 192, 192)
                .SetItemMatrix lngRP, lngCP, i & "枠同枠返還", , "<-", , , RGB(192, 192, 192)
            End With
            lngRP = lngRP + 1
        End If
    Next i

    gd.Rows = lngRP

    Exit Sub
Errorhandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   機能: 不成立得払い情報をグリッドに書き込む
'
'   備考: 引き数 gd - 書き込む対象となるグリッドデータ
'                lngRP - 行
'                Shikibetu - 式別文字列
'                Color - 背景色
'                FureirituFlag - 不成立であるかどうかのブール値
'                TokubaraiFlag - 得払いであるかどうかのブール値
'         lngRP - 行 はインクリメントされる
'
Private Sub HaraiSubFuseirituTokubarai(ByRef gd As clsGridData, ByRef lngRP As Long, Shikibetu As String, color As Long, _
                                        FuseirituFlag As Boolean, TokubaraiFlag As Boolean)
On Error GoTo Errorhandler
    Dim lngCP As Long
    If FuseirituFlag Then
        lngCP = 0
        gd.SetItemMatrix lngRP, lngCP, Shikibetu, , ">-", , , color
        gd.SetItemMatrix lngRP, lngCP, "不成立", , "^-", , , color
        gd.SetItemMatrix lngRP, lngCP, "不成立", , "^-", , , color
        gd.SetItemMatrix lngRP, lngCP, "不成立", , "^-", , , color
        lngRP = lngRP + 1
    End If
    
    If TokubaraiFlag Then
        lngCP = 0
        gd.SetItemMatrix lngRP, lngCP, Shikibetu, , ">-", , , color
        gd.SetItemMatrix lngRP, lngCP, "特払", , ">-", , , color
        gd.SetItemMatrix lngRP, lngCP, "70", , ">-", , , color
        gd.SetItemMatrix lngRP, lngCP, "", , ">-", , , color
        lngRP = lngRP + 1
    End If
    Exit Sub
Errorhandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   機能: ﾗｯﾌﾟﾀｲﾑグリッドを作る
'
'   備考: なし
'
Private Sub DirectMakeData_Lap(ByRef gd As clsGridData)
On Error GoTo Errorhandler

    Const rowOffset As Long = 1 '' データ行オフセット
    
    Dim lngCP As Long   '' カラムポインタ
    Dim i     As Long   '' ループカウンタ
    Dim j     As Long   '' ループカウンタ
    Dim c     As Long   '' 表示したラップタイムのカウンタ
    
    Dim lngMeter As Long '' 表示するメートル
    
    ' トラック距離を200mで割った余りを最初に設定
    ' 200で割り切れる場合は200mから
    lngMeter = val(mBuf_RA.KYORI) Mod 200
    If lngMeter = 0 Then
        lngMeter = 200
    End If
    
    
    With gd
        .Rows = Bigger(5 + (-Int(-val(mBuf_RA.KYORI) / 200 / 5) * 2), 9)
        .Cols = 8
        
        c = 0
        Do While (lngMeter <= val(mBuf_RA.KYORI))
            ' 5列で2行おきに並べる
            ' メートル表示
            With .ItemMatrix(Int(c / 5) * 2, c Mod 5)
                .Text = lngMeter & "m"
                .BGColor = RGB(192, 192, 192)
            End With
            ' タイム表示
            With .ItemMatrix(Int(c / 5) * 2 + 1, c Mod 5)
                If mBuf_RA.LapTime(c) <> "000" Then
                    .Text = Format$(mBuf_RA.LapTime(c), "@@.@")
                End If
                .BGColor = RGB(223, 223, 223)
            End With
            ' カウントアップ
            c = c + 1
            ' 200m毎
            lngMeter = lngMeter + 200
        Loop
                
        mlngCornerRow = Bigger(Int((c - 1) / 5) * 2 + 3, 5)
        For j = 0 To 3
            If mBuf_RA.CornerInfo(j).Syukaisu <> 0 Then
                .ItemMatrix(mlngCornerRow + j, 0).Text = mBuf_RA.CornerInfo(j).Syukaisu & "周目" _
                        & mBuf_RA.CornerInfo(j).Corner & "コーナー " & Trim$(mBuf_RA.CornerInfo(j).Jyuni)
            End If
        Next j
        
        lngMeter = val(mBuf_RA.KYORI) Mod 200
         If mBuf_RA.HaronTimeS3 <> 0 Then
            .ItemMatrix(0, 6).Text = IIf(lngMeter = 0, "前３ハロン合計", Right("         " & lngMeter + 400, 9) & " 合計")
            .ItemMatrix(0, 7).Text = Format$(val(mBuf_RA.HaronTimeS3) / 10, "#0.0")
        End If
            
        If mBuf_RA.HaronTimeS4 <> 0 Then
            .ItemMatrix(1, 6).Text = IIf(lngMeter = 0, "前４ハロン合計", Right("         " & lngMeter + 600, 9) & " 合計")
            .ItemMatrix(1, 7).Text = Format$(val(mBuf_RA.HaronTimeS4) / 10, "#0.0")
        End If
        
        If mBuf_RA.HaronTimeL3 <> 0 Then
            .ItemMatrix(2, 6).Text = "後３ハロン合計"
            .ItemMatrix(2, 7).Text = Format$(val(mBuf_RA.HaronTimeL3) / 10, "#0.0")
        End If
            
        If mBuf_RA.HaronTimeL4 <> 0 Then
            .ItemMatrix(3, 6).Text = "後４ハロン合計"
            .ItemMatrix(3, 7).Text = Format$(val(mBuf_RA.HaronTimeL4) / 10, "#0.0")
        End If
    
        
    End With
    
    
    Exit Sub
Errorhandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   機能: 持ちタイムグリッドを作る
'
'   備考: なし
'
Private Sub DirectMakeData_MochiTime()
    Const rowOffset As Long = 2 '' データ行オフセット
    Dim gd     As clsGridData
    Dim strSQL As String
    Dim lngCP  As Long

    Dim cUA As ADODB.Connection
    Dim cUB As ADODB.Connection
    Dim cRA As ADODB.Connection
    
    Dim UA As ADODB.Recordset
    Dim UB As ADODB.Recordset
    Dim UM As ADODB.Recordset
    Dim AV As ADODB.Recordset
    
    Dim rUA As ADODB.Recordset
    Dim rUB As ADODB.Recordset
    Dim rRA As ADODB.Recordset
    
    Dim strMinTime  As String
    Dim strYear     As String
    Dim strMonthDay As String
    Dim strJyoCD    As String
    Dim strKaiji    As String
    Dim strNichiji  As String
    Dim strRaceNum  As String
    Dim strUmaban  As String
    
    Dim str2MinTime  As String
    Dim str2Year     As String
    Dim str2MonthDay As String
    Dim str2JyoCD    As String
    Dim str2Kaiji    As String
    Dim str2Nichiji  As String
    Dim str2RaceNum  As String
    Dim str2Umaban   As String
    Dim i As Long
    
    Set gd = New clsGridData
    
    Set UA = mRS_UMA_RACE_A
    Set UB = mRS_UMA_RACE_B
    Set UM = mRS_UMA
    Set AV = mRS_TORIKESI_JYOGAI

    Set cUA = New ADODB.Connection
    Set cUB = New ADODB.Connection
    Set cRA = New ADODB.Connection
    Set rUA = New ADODB.Recordset
    Set rUB = New ADODB.Recordset
    Set rRA = New ADODB.Recordset
    
    cUA.Open "Provider=Microsoft.Jet.OLEDB.4.0;" _
            & "Data Source=" & gApp.R_DBPath & "\subUMA_RACE_A.mdb"
    cUB.Open "Provider=Microsoft.Jet.OLEDB.4.0;" _
            & "Data Source=" & gApp.R_DBPath & "\subUMA_RACE_B.mdb"
    cRA.Open "Provider=Microsoft.Jet.OLEDB.4.0;" _
            & "Data Source=" & gApp.R_DBPath & "\subRACE.mdb"
    
    rUB.CursorLocation = adUseServer
    rUB.Open "UMA_RACE_B", cUB, adOpenKeyset, adLockReadOnly, adCmdTableDirect
    rUB.Index = "PrimaryKey"
    
    rRA.CursorLocation = adUseServer
    rRA.Open "RACE", cRA, adOpenKeyset, adLockReadOnly, adCmdTableDirect
    rRA.Index = "PrimaryKey"

    With gd
        .Rows = 28 + 2 ' ２段ヘッダ
        .Cols = 3 + (8 * 2)

        ' １行目 カラムヘッダ登録
        lngCP = 0
        .SetItemMatrix 0, lngCP, "枠", "枠番", ">-"
        .SetItemMatrix 0, lngCP, "番", "馬番", ">-"
        .SetItemMatrix 0, lngCP, "馬名", , "<-"
        For i = 0 To 7
            .SetItemMatrix 0, lngCP, mCC.KIBJ3(mBuf_RA.id.JyoCD) & " " & mCC.TRCK4(mBuf_RA.TrackCD) & " " & val(mBuf_RA.KYORI) & "m", , "<-" ' 中山 芝 1200m
        Next i
        For i = 0 To 7
            .SetItemMatrix 0, lngCP, mCC.TRCK4(mBuf_RA.TrackCD) & " " & val(mBuf_RA.KYORI) & "m", , "<-" ' 1200m
        Next i
        lngCP = 0
        .SetItemMatrix 1, lngCP, "枠", "枠番", ">-"
        .SetItemMatrix 1, lngCP, "番", "馬番", ">-"
        .SetItemMatrix 1, lngCP, "馬名", , "<-"
        For i = 0 To 1
            .SetItemMatrix 1, lngCP, "タイム", , ">-"
            .SetItemMatrix 1, lngCP, "後3F", "後3ハロン", ">-"
            .SetItemMatrix 1, lngCP, "年月日", , ">-"
            .SetItemMatrix 1, lngCP, "回場日", , ">-"
            .SetItemMatrix 1, lngCP, "着", "着順", ">-"
            .SetItemMatrix 1, lngCP, "異", "異常", "<-"
            .SetItemMatrix 1, lngCP, "馬場", "馬場状態", "<-"
            .SetItemMatrix 1, lngCP, "負担", "負担重量", ">-"
        Next i
    End With 'gd
    
    ' 出走馬ループ
    i = 0
    UA.MoveFirst
    Do While Not UA.EOF
        
        strMinTime = "99999"
        strYear = ""
        strMonthDay = ""
        strJyoCD = ""
        strKaiji = ""
        strNichiji = ""
        strRaceNum = ""
        strUmaban = ""
        
        str2MinTime = "99999"
        str2Year = ""
        str2MonthDay = ""
        str2JyoCD = ""
        str2Kaiji = ""
        str2Nichiji = ""
        str2RaceNum = ""
        str2Umaban = ""
        
        If UA("KettoNum") <> "0000000000" Then
            'KettoNumが初期値でないなら
            rUA.Open "SELECT * FROM UMA_RACE_A WHERE KettoNum='" & UA("KettoNum") & "'", cUA, adOpenKeyset, adLockReadOnly, adCmdText
            ' 過去レースループ
            Do While Not rUA.EOF
                ' バックグラウンド
                DoEvents
                If mblnCancelFetching Then
                    Exit Sub
                End If
                
                SafeSeek UM, Array("KettoNum"), Array(UA("KettoNum").value)
                
                ' 過去のみ
                If rUA("Year") & rUA("MonthDay") < (mKey.Year & mKey.MonthDay) Then
                    SafeSeek rRA, Array("Year", "MonthDay", "JyoCD", "Kaiji", "Nichiji", "RaceNum"), _
                                Array(rUA("Year").value, rUA("MonthDay").value, rUA("JyoCD").value, rUA("Kaiji").value, rUA("Nichiji").value, rUA("RaceNum").value)
                    ' レース情報が存在する場合のみ
                    If Not rRA.EOF Then
                        ' このレースと同じ距離 同じ（芝・ダート・障害）のみ
                        If rRA("Kyori") = mBuf_RA.KYORI And MochiTimeSubSameTrackCD(rRA("TrackCD").value, mBuf_RA.TrackCD) Then
                            SafeSeek rUB, Array("B_Year", "B_MonthDay", "B_JyoCD", "B_RaceNum", "B_Umaban", "B_KettoNum"), _
                                        Array(rUA("Year").value, rUA("MonthDay").value, rUA("JyoCD").value, rUA("RaceNum").value, rUA("Umaban").value, rUA("KettoNum").value)
                            ' ０秒でない場合のみ
                            If rUB("Time") <> 0 Then
                                ' より早いタイムならば
                                If rUB("Time") <= strMinTime Then
                                    strMinTime = rUB("Time")
                                    strYear = rUA("Year")
                                    strMonthDay = rUA("MonthDay")
                                    strJyoCD = rUA("JyoCD")
                                    strKaiji = rUA("Kaiji")
                                    strNichiji = rUA("Nichiji")
                                    strRaceNum = rUA("RaceNum")
                                    strUmaban = rUA("Umaban")
                                End If
                                ' 場所が同じ場合
                                If rUA("JyoCD") = mBuf_RA.id.JyoCD Then
                                    ' より早いタイムならば
                                    If rUB("Time") <= str2MinTime Then
                                        str2MinTime = rUB("Time")
                                        str2Year = rUA("Year")
                                        str2MonthDay = rUA("MonthDay")
                                        str2JyoCD = rUA("JyoCD")
                                        str2Kaiji = rUA("Kaiji")
                                        str2Nichiji = rUA("Nichiji")
                                        str2RaceNum = rUA("RaceNum")
                                        str2Umaban = rUA("Umaban")
                                    End If
                                End If
                            End If ' ０秒でない場合のみ
                        End If ' このレースと同じ距離のみ
                    End If ' レース情報が存在する場合のみ
                End If ' 過去のみ
                rUA.MoveNext
            Loop ' 過去レースループ
            rUA.Close
            
            SafeSeek UB, Array("B_Year", "B_MonthDay", "B_JyoCD", "B_RaceNum", "B_Umaban", "B_KettoNum"), _
                Array(UA("Year").value, UA("MonthDay").value, UA("JyoCD").value, UA("RaceNum").value, UA("Umaban").value, UA("KettoNum").value)
            SafeSeek UM, Array("KettoNum"), Array(UA("KettoNum").value)
            SafeSeek AV, Array("Year", "MonthDay", "JyoCD", "Kaiji", "Nichiji", "RaceNum", "Umaban"), _
                        Array(UA("Year").value, UA("MonthDay").value, UA("JyoCD").value, UA("Kaiji").value, UA("Nichiji").value, UA("RaceNum").value, UA("Umaban").value)
            
            lngCP = 0
            With gd
                .SetItemMatrix i + rowOffset, lngCP, mSC.FitSpaceNum(UA("Wakuban"), 2), , ">^", , , gApp.GetWakubanColor(val(UA("Wakuban"))), Contrast(gApp.GetWakubanColor(val(UA("Wakuban"))))
                .SetItemMatrix i + rowOffset, lngCP, mSC.FitSpaceNum(UA("Umaban"), 2), , ">^"
                Call BameiAV(AV, UA, UB, UM, gd, i + rowOffset, lngCP)
            End With
            ' 場所&芝＆距離　持ちタイム
            If str2Year <> "" Then
                strSQL = "SELECT * FROM UMA_RACE_A "
                strSQL = strSQL & " WHERE [Year]='" & str2Year & "'"
                strSQL = strSQL & " AND [MonthDay]='" & str2MonthDay & "'"
                strSQL = strSQL & " AND [JyoCD]='" & str2JyoCD & "'"
                strSQL = strSQL & " AND [Kaiji]='" & str2Kaiji & "'"
                strSQL = strSQL & " AND [Nichiji]='" & str2Nichiji & "'"
                strSQL = strSQL & " AND [RaceNum]='" & str2RaceNum & "'"
                strSQL = strSQL & " AND [Umaban]='" & str2Umaban & "'"
                strSQL = strSQL & " AND [KettoNum]='" & UA("KettoNum") & "'"
                rUA.Open strSQL, cUA, adOpenKeyset, adLockReadOnly, adCmdText
                rUB.Seek Array(rUA("Year"), rUA("MonthDay"), rUA("JyoCD"), rUA("RaceNum"), rUA("Umaban"), rUA("KettoNum"))
                rRA.Seek Array(rUA("Year"), rUA("MonthDay"), rUA("JyoCD"), rUA("Kaiji"), rUA("Nichiji"), rUA("RaceNum"))
                With gd
                    .SetItemMatrix i + rowOffset, lngCP, Format$(rUB("Time"), "@:@@.@"), , ">^"
                    .SetItemMatrix i + rowOffset, lngCP, IIf(rUB("HaronTimeL3") = "999", "", mSC.SSSs(rUB("HaronTimeL3"))), , ">^"
                    .SetItemMatrix i + rowOffset, lngCP, mSC.YMD3(rUA("Year") & rUA("MonthDay")), Trim$(rRA("Hondai")) & mCC.GRAD3(rRA("GradeCD")), ">^", "RA", rUA("Year") & rUA("MonthDay") & rUA("JyoCD") & rUA("Kaiji") & rUA("Nichiji") & rUA("RaceNum")
                    .SetItemMatrix i + rowOffset, lngCP, mSC.FitSpaceNum(rUA("Kaiji"), 2) & mCC.KIBJ3(rUA("JyoCD")) & mSC.FitSpaceNum(rUA("Nichiji"), 2), , ">^", "RA", rUA("Year") & rUA("MonthDay") & rUA("JyoCD") & rUA("Kaiji") & rUA("Nichiji") & rUA("RaceNum")
                    .SetItemMatrix i + rowOffset, lngCP, mSC.FitSpaceNum(rUB("KakuteiJyuni"), 2), , ">^", , , gApp.GetChakujyunColor(rUB("KakuteiJyuni"))
                    .SetItemMatrix i + rowOffset, lngCP, mCC.IJKB2(rUB("IJyoCD")), , "<^"
                    .SetItemMatrix i + rowOffset, lngCP, mSC.BabaGrid(rRA), , "<^"
                    .SetItemMatrix i + rowOffset, lngCP, Format$(rUB("Futan") / 10, "#0.0"), , ">^"
                End With ' gd
                rUA.Close
            Else
                lngCP = 11
            End If
            ' 距離&芝 持ちタイム
            If strYear <> "" Then
                strSQL = "SELECT * FROM UMA_RACE_A "
                strSQL = strSQL & " WHERE [Year]='" & strYear & "'"
                strSQL = strSQL & " AND [MonthDay]='" & strMonthDay & "'"
                strSQL = strSQL & " AND [JyoCD]='" & strJyoCD & "'"
                strSQL = strSQL & " AND [Kaiji]='" & strKaiji & "'"
                strSQL = strSQL & " AND [Nichiji]='" & strNichiji & "'"
                strSQL = strSQL & " AND [RaceNum]='" & strRaceNum & "'"
                strSQL = strSQL & " AND [Umaban]='" & strUmaban & "'"
                strSQL = strSQL & " AND [KettoNum]='" & UA("KettoNum") & "'"
                rUA.Open strSQL, cUA, adOpenKeyset, adLockReadOnly, adCmdText
                rUB.Seek Array(rUA("Year"), rUA("MonthDay"), rUA("JyoCD"), rUA("RaceNum"), rUA("Umaban"), rUA("KettoNum"))
                rRA.Seek Array(rUA("Year"), rUA("MonthDay"), rUA("JyoCD"), rUA("Kaiji"), rUA("Nichiji"), rUA("RaceNum"))
                With gd
                    .SetItemMatrix i + rowOffset, lngCP, Format$(rUB("Time"), "@:@@.@"), , ">^"
                    .SetItemMatrix i + rowOffset, lngCP, IIf(rUB("HaronTimeL3") = "999", "", mSC.SSSs(rUB("HaronTimeL3"))), , ">^"
                    .SetItemMatrix i + rowOffset, lngCP, mSC.YMD3(rUA("Year") & rUA("MonthDay")), Trim$(rRA("Hondai")) & mCC.GRAD3(rRA("GradeCD")), ">^", "RA", rUA("Year") & rUA("MonthDay") & rUA("JyoCD") & rUA("Kaiji") & rUA("Nichiji") & rUA("RaceNum")
                    .SetItemMatrix i + rowOffset, lngCP, mSC.FitSpaceNum(rUA("Kaiji"), 2) & mCC.KIBJ3(rUA("JyoCD")) & mSC.FitSpaceNum(rUA("Nichiji"), 2), , ">^", "RA", rUA("Year") & rUA("MonthDay") & rUA("JyoCD") & rUA("Kaiji") & rUA("Nichiji") & rUA("RaceNum")
                    .SetItemMatrix i + rowOffset, lngCP, mSC.FitSpaceNum(rUB("KakuteiJyuni"), 2), , ">^", , , gApp.GetChakujyunColor(rUB("KakuteiJyuni"))
                    .SetItemMatrix i + rowOffset, lngCP, mCC.IJKB2(rUB("IJyoCD")), , "<^"
                    .SetItemMatrix i + rowOffset, lngCP, mSC.BabaGrid(rRA), , "<^"
                    .SetItemMatrix i + rowOffset, lngCP, Format$(rUB("Futan") / 10, "#0.0"), , ">^"
                End With ' gd
                rUA.Close
            End If
        Else
            '血統番号が初期値の時は馬名まで表示
            SafeSeek UB, Array("B_Year", "B_MonthDay", "B_JyoCD", "B_RaceNum", "B_Umaban", "B_KettoNum"), _
                Array(UA("Year").value, UA("MonthDay").value, UA("JyoCD").value, UA("RaceNum").value, UA("Umaban").value, UA("KettoNum").value)
            SafeSeek UM, Array("KettoNum"), Array(UA("KettoNum").value)
            SafeSeek AV, Array("Year", "MonthDay", "JyoCD", "Kaiji", "Nichiji", "RaceNum", "Umaban"), _
                        Array(UA("Year").value, UA("MonthDay").value, UA("JyoCD").value, UA("Kaiji").value, UA("Nichiji").value, UA("RaceNum").value, UA("Umaban").value)
            lngCP = 0
            With gd
                .SetItemMatrix i + rowOffset, lngCP, mSC.FitSpaceNum(UA("Wakuban"), 2), , ">^", , , gApp.GetWakubanColor(val(UA("Wakuban"))), Contrast(gApp.GetWakubanColor(val(UA("Wakuban"))))
                .SetItemMatrix i + rowOffset, lngCP, mSC.FitSpaceNum(UA("Umaban"), 2), , ">^"
                Call BameiAV(AV, UA, UB, UM, gd, i + rowOffset, lngCP)
            End With
        End If
        
        UA.MoveNext
        i = i + 1
        If gd.Rows <= i + rowOffset Then
            gd.Rows = gd.Rows + 10
        End If
        
    Loop ' 出走馬ループ
    gd.Rows = i + rowOffset
    RaiseEvent FetchedMotiTIme(gd)
End Sub


'
'   機能: ２つのコードを範囲で比較する
'
'   備考: なし
'
Private Function MochiTimeSubSameTrackCD(a As String, b As String) As Boolean
    Select Case a
    Case "10" To "22"
        MochiTimeSubSameTrackCD = (b >= "10" And b <= "22")
    Case "23" To "26", "29"
        MochiTimeSubSameTrackCD = ((b >= "23" And b <= "26") Or b = "29")
    Case "27", "28", "51" To "59"
        MochiTimeSubSameTrackCD = (b = "27" Or b = "28" Or (b >= "51" And b <= "59"))
    End Select
End Function


'
'   機能: 過去走グリッドを作る
'
'   備考: なし
'
Private Sub DirectMakeData_Kako()
On Error GoTo Errorhandler
    Const rowOffset As Long = 1 '' データ行オフセット
    
    Dim gd     As clsGridData
    Dim strSQL As String
    Dim lngCP  As Long
    Dim lngKakoMax As Long
    
    Dim cUA As ADODB.Connection
    Dim cUB As ADODB.Connection
    Dim cRA As ADODB.Connection
    
    Dim UA As ADODB.Recordset
    Dim UB As ADODB.Recordset
    Dim UM As ADODB.Recordset
    Dim AV As ADODB.Recordset
    
    Dim rUA As ADODB.Recordset
    Dim rUB As ADODB.Recordset
    Dim rRA As ADODB.Recordset
    
    Dim i As Long
    
    mblnNowKakoFetching = True
    
    Set gd = New clsGridData
    Set UA = mRS_UMA_RACE_A
    Set UB = mRS_UMA_RACE_B
    Set UM = mRS_UMA
    Set AV = mRS_TORIKESI_JYOGAI

    ' イベントを拾わないように
    Set mAsyncCN_UMA_RACE_A = Nothing
    Set mCN_UMA_RACE_B = Nothing
    
    Set cUA = gApp.GetCN_UMA_RACE_A
    Set cUB = gApp.GetCN_UMA_RACE_B
    Set cRA = gApp.GetCN_RACE
    
    Set rUA = New ADODB.Recordset
    Set rUB = New ADODB.Recordset
    Set rRA = New ADODB.Recordset
    
    lngKakoMax = gApp.R_KakoNum
    If lngKakoMax < 1 Or lngKakoMax > 99 Then
        gApp.R_KakoNum = 5
        lngKakoMax = 5
    End If
    
    rUB.CursorLocation = adUseServer
    rUB.Open "UMA_RACE_B", cUB, adOpenKeyset, adLockReadOnly, adCmdTableDirect
    rUB.Index = "PrimaryKey"
    
    rRA.CursorLocation = adUseServer
    rRA.Open "RACE", cRA, adOpenKeyset, adLockReadOnly, adCmdTableDirect
    rRA.Index = "PrimaryKey"

    With gd
        .Cols = lngKakoMax + 3
        .Rows = 30 ' 大きめに
        ' カラムヘッダ挿入
        lngCP = 0
        .SetItemMatrix 0, lngCP, "枠", "枠番号", ">-"
        .SetItemMatrix 0, lngCP, "番", "馬番号", ">-"
        .SetItemMatrix 0, lngCP, "馬名", , "<-"
        For i = 0 To gApp.R_KakoNum - 1
            .SetItemMatrix 0, lngCP, IIf(i = 0, "前走", i + 1 & "走前")
        Next
    End With
    
    
    ' 出走馬ループ
    i = 0
    UA.MoveFirst
    Do While Not UA.EOF
    
        SafeSeek UB, Array("B_Year", "B_MonthDay", "B_JyoCD", "B_RaceNum", "B_Umaban", "B_KettoNum"), _
            Array(UA("Year").value, UA("MonthDay").value, UA("JyoCD").value, UA("RaceNum").value, UA("Umaban").value, UA("KettoNum").value)
        SafeSeek UM, Array("KettoNum"), Array(UA("KettoNum").value)
        SafeSeek AV, Array("Year", "MonthDay", "JyoCD", "Kaiji", "Nichiji", "RaceNum", "Umaban"), _
                    Array(UA("Year").value, UA("MonthDay").value, UA("JyoCD").value, UA("Kaiji").value, UA("Nichiji").value, UA("RaceNum").value, UA("Umaban").value)
        
                    
        strSQL = "SELECT * FROM UMA_RACE_A "
        strSQL = strSQL & "WHERE KettoNum='" & UA("KettoNum") & "'"
        strSQL = strSQL & " ORDER BY [Year] DESC, [MonthDay] DESC, [RaceNum] DESC, [JyoCD] DESC"
        rUA.Open strSQL, cUA, adOpenKeyset, adLockReadOnly, adCmdText
        lngCP = 0
        With gd
            .SetItemMatrix i + rowOffset, lngCP, mSC.FitSpaceNum(UA("Wakuban"), 2), "枠番", ">-", , , gApp.GetWakubanColor(CLng(UA("Wakuban"))), Contrast(gApp.GetWakubanColor(val(UA("Wakuban"))))
            .SetItemMatrix i + rowOffset, lngCP, mSC.FitSpaceNum(UA("Umaban"), 2), "馬番", ">-"
            Call BameiAV(AV, UA, UB, UM, gd, i + rowOffset, lngCP)
            .ItemMatrix(i + rowOffset, lngCP - 1).Text = vbCr & vbCr & .ItemMatrix(i + rowOffset, lngCP - 1).Text & vbCr & vbCr & vbCr
        End With 'gd
        
        If UA("KettoNum") <> "0000000000" Then
            ' 過去レースループ(KettoNumが初期値でないなら)
            Do While Not rUA.EOF
                ' バックグラウンド
                DoEvents
                If mblnCancelFetching Then
                    Exit Sub
                End If
                ' キャンセル感知
                If mblnCancelKakoFetching Then
                    mblnNowKakoFetching = False
                    mblnCancelKakoFetching = False
                    Exit Sub
                End If
                
                ' 中止レースは含まない
                If rUA("DataKubun") = "9" Then
                    rUA.MoveNext
                Else
                    ' 過去のみ
                    If rUA("Year") & rUA("MonthDay") & rUA("RaceNum") < (mKey.Year & mKey.MonthDay & mKey.RaceNum) Then
                        rRA.Seek Array(rUA("Year"), rUA("MonthDay"), rUA("JyoCD"), rUA("Kaiji"), rUA("Nichiji"), rUA("RaceNum"))
                        rUB.Seek Array(rUA("Year"), rUA("MonthDay"), rUA("JyoCD"), rUA("RaceNum"), rUA("Umaban"), rUA("KettoNum"))
                        gd.SetItemMatrix i + rowOffset, lngCP, getKakoStr(rUA, rUB, rRA), , "<^", "RA" _
                                    , rUA("Year") & rUA("MonthDay") & rUA("JyoCD") & rUA("Kaiji") & rUA("Nichiji") & rUA("RaceNum"), gApp.GetChakujyunColor(val(rUB("KakuteiJyuni"))), RGB(1, 1, 1)
                    End If ' 過去のみ
                    rUA.MoveNext
                    If lngCP - 3 > lngKakoMax Then
                        Exit Do
                    End If
                End If
            Loop ' 過去レースループ
        End If
        rUA.Close
        UA.MoveNext
        i = i + 1
        If gd.Rows <= i + rowOffset Then
            gd.Rows = gd.Rows + 10
        End If
    Loop ' 出走馬ループ
    gd.Rows = i + rowOffset
    RaiseEvent FetchedKako(gd)
    
    mblnNowKakoFetching = False
    
    Exit Sub
Errorhandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   機能: 過去送グリッド用文字列を生成する
'
'   備考: なし
'
Private Function getKakoStr(UA As ADODB.Recordset, UB As ADODB.Recordset, RA As ADODB.Recordset) As String
On Error GoTo Errorhandler
    Dim out As String
    Dim i   As Long
    
    '1行目
    out = Right$(mSC.YMD3(UA("Year") & UA("MonthDay")), 8) & " "
    out = out & mSC.FitSpaceNum(UA("Kaiji"), 2)
    out = out & mSC.FitSpaceStr(mCC.KIBJ3(UA("JyoCD")), 2)
    out = out & mSC.FitSpaceNum(UA("Nichiji"), 2) & " "
    out = out & IIf(RA("TenkoCD") <> 0, mSC.FitSpaceStr(mCC.TNKO1(RA("TenkoCD")), 2), Space(4)) '天候
    out = out & mSC.BabaGrid(RA)
    out = out & mSC.FitSpaceStr(mCC.IJKB5(UB("IJyoCD")), 1)
    out = out & mSC.FitSpaceNum(UB("KakuteiJyuni"), 2) & IIf(mSC.Emp(UB("KakuteiJyuni")), "　", "着")
    out = out & vbCrLf
    '2行目
    out = out & mSC.FitSpaceStr(IfExist(RA, "Ryakusyo3"), 3)
    If mCC.GRAD4(IfExist(RA, "GradeCD")) <> "" Then
        out = out & mCC.GRAD4(IfExist(RA, "GradeCD")) & " "
    Else
        out = out & mCC.KSJK5(IfExist(RA, "JyokenCD5")) & " "
    End If
    out = out & mSC.FitSpaceStr(mCC.TRCK5(IfExist(RA, "TrackCD")), 2)
    out = out & mSC.FitSpaceNum(IfExist(RA, "Kyori"), 4) & " "
    out = out & mSC.FitSpaceNum(UB("Ninki"), 2) & IIf(mSC.Emp(UB("Ninki")), "　", "人")
    out = out & mSC.FitSpaceNum(IfExist(RA, "TorokuTosu"), 2) & IIf(mSC.Emp(RA("TorokuTosu")), "　", "頭")
    out = out & vbCrLf
    '3行目
    out = out & IIf(val(UB("Time")) = 0, "      ", Format$(UB("Time"), "@:@@.@")) & " "
    out = out & IIf(UB("HaronTimeL3") = "999" Or UB("HaronTimeL3") = "000", "    ", Right$(Format$(UB("HaronTimeL3") / 10, "0.0"), 4))
    out = out & " " & mSC.FitSpaceStr(mCC.JRSB1(IfExist(RA, "JyuryoCD")), 3) & " "
    out = out & Right$(Space(2) & Format$(UB("Jyuni1c"), "##"), 2) _
                    & IIf(val(UB("Jyuni1c")) = 0 Or val(UB("Jyuni2c")) = 0, " ", "-") _
                    & Right$(Space(2) & Format$(UB("Jyuni2c"), "##"), 2) _
                    & IIf(val(UB("Jyuni2c")) = 0 Or val(UB("Jyuni3c")) = 0, " ", "-") _
                    & Right$(Space(2) & Format$(UB("Jyuni3c"), "##"), 2) _
                    & IIf(val(UB("Jyuni3c")) = 0 Or val(UB("Jyuni4c")) = 0, " ", "-") _
                    & Right$(Space(2) & Format$(UB("Jyuni4c"), "##"), 2)
    out = out & vbCrLf
    '4行目
    out = out & mSC.FitSpaceStr(UB("KisyuRyakusyo"), 4) & " "
    out = out & IIf(val(UB("Futan")) = 0, Space(6), Format$(UB("Futan") / 10, "#0.0") & "kg") & " "
    If val(UB("BaTaijyu")) < 2 Then
        out = out & Space(5)
    ElseIf UB("BaTaijyu") = "999" Then
        out = out & " --- "
    Else
        out = out & val(UB("BaTaijyu")) & "kg"
    End If
    If UB("ZogenFugo") = " " Then
        If UB("ZogenSa") = "   " Or UB("ZogenSa") = "999" Then
            out = out & Space(5)
        Else
            out = out & "±" & Left$(val(UB("ZogenSa")) & Space(3), 3)
        End If
    Else
        out = out & IIf(UB("ZogenSa") = "999", Space(5), UB("ZogenFugo") & Left$(val(UB("ZogenSa")) & Space(3), 3) & " ")
    End If
    out = out & IIf(UB("BLINKER") = "0", " ", "B") & " "
    out = out & mSC.SE_KRKH(UB("KyakusituKubun"))
    out = out & vbCrLf
    '5行目
    out = out & mSC.FitSpaceNum(UA("Wakuban"), 1) & "枠 "
    out = out & mSC.FitSpaceNum(UA("Umaban"), 2) & "番 "
    If UB("Odds") <> 0 Then
        out = out & "単勝" & Right$("   " & Format$(UB("Odds") / 10, "0.0"), 5)
    End If
    out = out & vbCrLf
    '6行目
    out = out & IIf(UB("TimeDiff") = "9999" Or UB("TimeDiff") = "0000", " ", Mid$(UB("TimeDiff"), 1, 1) _
                & Right$("  " & Format$(val(Mid$(UB("TimeDiff"), 2, 3)) / 10, "0.0"), 4))
    For i = 0 To 2
        If UB("KettoNum" & i + 1) <> String$(10, "0") Then
            out = out & mSC.FitSpaceStr(IIf(UB("KakuteiJyuni") = "01", "(", " ") _
                        & Trim$(UB("Bamei" & i + 1)) _
                        & IIf(UB("KakuteiJyuni") = "01", ")", " "), 11) & " "
        End If
    Next i
    
    getKakoStr = out
    Exit Function
Errorhandler:
    gApp.ErrLog
    Resume Next
End Function


'
'   機能: 取り消し除外を配慮した馬名フィールド書き込み
'
'   備考: lngCP - 列はインクリメントされる
'
Private Sub BameiAV(ByRef AV As ADODB.Recordset, _
                    ByRef UA As ADODB.Recordset, _
                    ByRef UB As ADODB.Recordset, _
                    ByRef UM As ADODB.Recordset, _
                    ByRef gd As clsGridData, _
                    ByVal row As Long, _
                    ByRef lngCP As Long)
On Error GoTo Errorhandler

    Dim StrikeFlag As Boolean


    ' 取り消し除外変更情報を配慮
    If UA("DataKubun") = "2" And Not AV.EOF Then
        StrikeFlag = (AV("DataKubun") = "1")
        gd.SetItemMatrix row, lngCP, Trim$(UA("Bamei")), mSC.AV_DKBN(AV("DataKubun")) & "(" & mSC.AV_JKBN(AV("JiyuKubun")) & ")", "<-", "UM", IfExist(UM, "KettoNum"), , &HFF, True
    Else
        StrikeFlag = (UB("IjyoCD") >= "1" And UB("IjyoCD") <= "3")
        gd.SetItemMatrix row, lngCP, Trim$(UA("Bamei")), mCC.IJKB1(UB("IJyoCD")), "<-", "UM", IfExist(UM, "KettoNum"), , , StrikeFlag
    End If
    Exit Sub
Errorhandler:
    gApp.ErrLog
    lngCP = lngCP + 1
End Sub


'
'   機能: DBVer1以上且つRA(DataKubun)<=2のとき、TC,CCに本レースの情報があればmBuf_RA構造体に上書きする
'
'   備考: 上書き対象はKyoriとTrackCD
'
Private Sub Set_CCTC()
On Error GoTo Errorhandler
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
Errorhandler:
    gApp.ErrLog
End Sub


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
