VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDataTK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "特別登録馬データ取得オブジェクト"
'
'   特別登録馬データ取得オブジェクト
'

Option Explicit

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   外部関数(イベント)
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

' イベント
Public Event FetchedKihon(GridData As clsGridData)
Public Event FetchedKetto(GridData As clsGridData)
Public Event FetchedKako(GridData As clsGridData)
Public Event FetchedJokenBetu(GridData As clsGridData)
Public Event FetchedMotiTIme(GridData As clsGridData)
Public Event NoUMARACE()


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部変数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

' 変換ユーティリティー
Private mCC As clsCodeConverter          '' JV-Data のコード変換メソッド群
Private mSC As clsStringConverter        '' 文字列変換メソッド群

' Viewerに提供する為のデータ格納変数
Private mstrLabels(0 To 8) As String      '' テキストデータ
Private mGridData(0 To 4)  As clsGridData '' グリッドデータ
Private mstrTitle          As String      '' タイトル
Private mstrRCKey           As String     '' レコード詳細画面のキー
Private mstrRCG1Key         As String     '' レコード詳細画面のキー
Private mblnIsG1            As Boolean    '' G1 レースかどうか

Private mCN_TOKU_RACE               As ADODB.Connection
Private WithEvents mAsyncCN_TOKU    As ADODB.Connection
Attribute mAsyncCN_TOKU.VB_VarHelpID = -1
Private mCN_UMA                     As ADODB.Connection
Private mCN_BANUSI                  As ADODB.Connection
Private mCN_CHOKYO                  As ADODB.Connection
Private mCN_SEISAN                  As ADODB.Connection
Private mCN_HANSYOKU                As ADODB.Connection

Private mRS_TOKU_RACE  As ADODB.Recordset
Private mRS_TOKU       As ADODB.Recordset
Private mRS_UMA        As ADODB.Recordset
Private mRS_BANUSI     As ADODB.Recordset
Private mRS_CHOKYO     As ADODB.Recordset
Private mRS_SEISAN     As ADODB.Recordset
Private mRS_HANSYOKU   As ADODB.Recordset
Private mRS_HARAI      As ADODB.Recordset
Private mRS_Time_K     As ADODB.Recordset
Private mRS_TimeJK     As ADODB.Recordset

Private mBuf           As JV_TK_TOKUUMA

Private mKey As clsKeyRA


Private mblnNowFetching         As Boolean  ' 取得中フラグ
Private mblnCancelFetching      As Boolean  ' キャンセルフラグ

Private mblnNowKakoFetching     As Boolean ' 過去走取得中フラグ
Private mblnCancelKakoFetching  As Boolean ' 過去走処理待ちフラグ



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
'   機能: グリッドデータを取得
'
'   備考: なし
'
Public Property Get GridDatas(Index As Integer) As clsGridData
    Set GridDatas = mGridData(Index)
End Property

'
'   機能: タイトルを取得
'
'   備考: なし
'
Public Property Get Title() As String
    Title = mstrTitle
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
'   機能: 過去データ取得状態を返す
'
'   備考: なし
'
Public Property Get NowKakoFetching() As Boolean
    NowKakoFetching = mblnNowKakoFetching
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
On Error GoTo ErrorHandler
    gApp.Log "TK CancelFetching"
    
        mblnCancelFetching = True
        
        freers mRS_TOKU_RACE
        freers mRS_TOKU
        freers mRS_UMA
        freers mRS_BANUSI
        freers mRS_CHOKYO
        freers mRS_SEISAN
        freers mRS_HANSYOKU
        freers mRS_Time_K
        freers mRS_TimeJK
        
freecn mCN_TOKU_RACE
freecn mAsyncCN_TOKU
freecn mCN_UMA
freecn mCN_BANUSI
freecn mCN_CHOKYO
freecn mCN_SEISAN
freecn mCN_HANSYOKU

        mblnNowFetching = False
'        ' データの取得終了を、アプリケーションに報告する
    Exit Sub
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   機能: キャンセル(過去データ)
'
'   備考: なし
'
Public Sub CancelKakoFetching()
    gApp.Log "RA CancelKakoFetching"

    mblnCancelKakoFetching = True
End Sub


'
'   機能: 取得を開始する(過去走)
'
'   備考: 初期表示用データの取得
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
Public Function Fetch(ByRef key As clsKeyRA) As Boolean
On Error GoTo ErrorHandler
    Dim rs           As ADODB.Recordset
    Dim lngNumRecord As Long            '' レース数
    Dim strSQL       As String          '' SQL文用
    Dim i            As Long            '' ループカウンタ
    Dim lngCP        As Long            '' セル入力用
    Dim lngRP        As Long            '' セル入力用
    Dim str          As String          '' 一次利用
    Dim intMDBIndex  As Integer         '' 分割MDBのインデクス番号

    Set mKey = key

    ' 取得中フラグを立てる
    mblnNowFetching = True

    '-------------------
    ' TOKU_RACE情報取得
    '-------------------
    

    Set mRS_TOKU_RACE = New ADODB.Recordset
    Set rs = mRS_TOKU_RACE
    strSQL = "SELECT * FROM TOKU_RACE " & key.SQLWHEREString
    rs.Open strSQL, mCN_TOKU_RACE, adOpenKeyset, adLockReadOnly, adCmdText
    
    
    ' データが無かったらFalseを返して終了
    If rs.EOF Or rs.BOF Then
        rs.Close
        Set rs = Nothing
        Fetch = False
        Exit Function
    End If

    ' RACEレコードをすべて構造体で保持する
    Call SetDataFromRS_TK(rs, mBuf)
    
    ' タイトル文字列を作る
    mstrTitle = key.ReadableString _
                & mCC.YOBI7(rs("YoubiCD")) & " " _
                & mCC.KIBJ3(rs("JyoCD")) _
                & mSC.KN1(rs("Kaiji") & rs("Nichiji"))
    
    ' レース情報を作る
    Call MakeData_Race
        
    rs.Close

    ' レコード詳細画面用のキー
    mstrRCKey = "1" & mKey.str & "0000" & mBuf.JyokenInfo.SyubetuCD & mBuf.KYORI & mBuf.TrackCD
    mstrRCG1Key = "2" & mKey.str & mBuf.RaceInfo.TokuNum & mBuf.JyokenInfo.SyubetuCD & mBuf.KYORI & mBuf.TrackCD

    '-----------------
    ' 非同期取得を行う
    '-----------------
    
    strSQL = "SELECT * FROM TOKU"
    strSQL = strSQL & " WHERE "
    strSQL = strSQL & "     [Year]   = '" & mKey.Year & "'"
    strSQL = strSQL & " AND MonthDay = '" & mKey.MonthDay & "'"
    strSQL = strSQL & " AND JyoCD    = '" & mKey.JyoCD & "'"
    strSQL = strSQL & " AND Kaiji    = '" & mKey.Kaiji & "'"
    strSQL = strSQL & " AND Nichiji  = '" & mKey.Nichiji & "'"
    strSQL = strSQL & " AND RaceNum  = '" & mKey.RaceNum & "'"
    strSQL = strSQL & " ORDER BY Num"
    
    gApp.Log strSQL
    mAsyncCN_TOKU.Execute strSQL, , adAsyncExecute

    gApp.Log "TK Fetch Complete"
    Fetch = True
    Exit Function
ErrorHandler:
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
    
    Set mCN_TOKU_RACE = gApp.GetCN_TOKU_RACE
    Set mAsyncCN_TOKU = gApp.GetCN_TOKU
    Set mCN_UMA = gApp.GetCN_UMA
    Set mCN_BANUSI = gApp.GetCN_BANUSI
    Set mCN_CHOKYO = gApp.GetCN_CHOKYO
    Set mCN_SEISAN = gApp.GetCN_SEISAN
    Set mCN_HANSYOKU = gApp.GetCN_HANSYOKU
 
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
'   機能: 特別登録馬レコードセットの取得完了イベントハンドラ
'
'   備考: なし
'
Private Sub mAsyncCN_TOKU_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
On Error GoTo ErrorHandler
    gApp.Log ">mAsyncCN_TOKU_ExecuteComplete"
    
    Dim strSQL As String
    Dim i As Long
    Dim blnFirst As Boolean
    
    If pError Is Nothing Then
    
        Set mRS_TOKU = pRecordset
        
        If mRS_TOKU.EOF Then
            RaiseEvent NoUMARACE
        Else
            
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
            
            Set mRS_SEISAN = New ADODB.Recordset
            If OpenTableDirect(mRS_SEISAN, mCN_SEISAN, "SEISAN") = False Then
                Set mRS_SEISAN = Nothing
            End If
        
            Set mRS_HANSYOKU = New ADODB.Recordset
            If OpenTableDirect(mRS_HANSYOKU, mCN_HANSYOKU, "HANSYOKU") = False Then
                Set mRS_HANSYOKU = Nothing
            End If
            
            Call DirectMakeData_Kihon
            If mblnCancelFetching Then
                mblnNowFetching = False
                Exit Sub
            End If

            Call DirectMakeData_Ketto
            If mblnCancelFetching Then
                mblnNowFetching = False
                Exit Sub
            End If

            Call DirectMakeData_JokenBetu
            If mblnCancelFetching Then
                mblnNowFetching = False
                Exit Sub
            End If
            
            Call DirectMakeData_MochiTime
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
    
    gApp.Log "<mAsyncCN_TOKU_ExecuteComplete"
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
    Dim str As String       '' 作業用文字列変数
    Dim i   As Long         '' ループカウンタ

    With mBuf

        ' 太字レース名等グレーバー　文字列生成
        str = ""
        str = str & mCC.KIBJ3(.id.JyoCD)               '' 競馬場名
        str = str & val(.id.RaceNum) & "R"             '' nR
        str = str & " "                                 '' 空白
        With .RaceInfo
            If .Nkai <> "000" Then
                str = str & "第" & val(.Nkai) & "回"    '' 第n回
            End If
        End With
        str = str & " "                                 '' 空白
        str = str & Trim$(.RaceInfo.Hondai)              '' 本題
        str = str & mCC.GRAD2(.GradeCD)                 '' グレードコード
        mstrLabels(0) = str

        ' データ区分
        mstrLabels(1) = mSC.TK_DKBN(.head.DataKubun)

        ' ヘッダ情報　文字列生成

        ' 年月日曜
        str = ""
        str = str & mSC.YMD2(.id.Year & .id.MonthDay) & mCC.YOBI7(.RaceInfo.YoubiCD) & "  "
        mstrLabels(2) = str

        ' RACEでは発走時間があったところ
        str = ""

        ' 競走条件
        With .JyokenInfo
            str = str & mCC.KSSB4(.SyubetuCD) & " " & mCC.KSJK1(.JyokenCD(4))
            str = str & "　　"
            str = str & mCC.KSKG1(.KigoCD, mBuf.id.Year)
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
        With mBuf.HandiDate
            str = str & "ハンディ発表日: " & mSC.YMD1(.Year & .Month & .Day)
        End With
        mstrLabels(5) = str

        str = ""

        ' 登録頭数
        If .TorokuTosu <> "000" Then
            str = str & "登録頭数 " & gSC.ValStr(.TorokuTosu) & "頭"
        Else
            str = str & "               "
        End If
        str = str & "  "

        ' 本賞金
        str = str & "  "

        mstrLabels(6) = str


        ' 入線頭数
        str = ""
        str = str & "             "
        str = str & "  "

        ' 付加賞金
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
'   機能: 基本情報リッドを作る
'
'   備考: なし
'
Private Sub DirectMakeData_Kihon()
On Error GoTo ErrorHandler
    
    Const rowOffset As Long = 1 '' データ行オフセット
    
    Dim lngCP As Long   '' カラムポインタ
    Dim i     As Long   '' ループカウンタ
    
    Dim gd As clsGridData
    Set gd = New clsGridData
    
    Dim TK As ADODB.Recordset
    Dim UM As ADODB.Recordset
    Dim CH As ADODB.Recordset
    Dim BN As ADODB.Recordset
    Dim BR As ADODB.Recordset

    Set TK = mRS_TOKU
    Set UM = mRS_UMA
    Set CH = mRS_CHOKYO
    Set BN = mRS_BANUSI
    Set BR = mRS_SEISAN
    
    With gd
        .Rows = 10 + rowOffset
        .Cols = 11
        
        ' カラムヘッダ挿入
        lngCP = 0
        .SetItemMatrix 0, lngCP, "馬記号"
        .SetItemMatrix 0, lngCP, "馬名"
        
        .SetItemMatrix 0, lngCP, "性齢"
        .SetItemMatrix 0, lngCP, "毛"
        .SetItemMatrix 0, lngCP, "負担", "単位kg", ">_"
        
        .SetItemMatrix 0, lngCP, "本賞金累計", , ">_"
        .SetItemMatrix 0, lngCP, "収得賞金累計", , ">_"
        .SetItemMatrix 0, lngCP, "調教師"
        
        .SetItemMatrix 0, lngCP, "馬主"
        .SetItemMatrix 0, lngCP, "生産者"
        
        .SetItemMatrix 0, lngCP, "交流区分"
    End With
    
    ' グリッド挿入
    TK.MoveFirst
    
    i = 0
    Do While Not TK.EOF
        DoEvents
        If mblnCancelFetching Then
            TK.Cancel
            UM.Cancel
            CH.Cancel
            BN.Cancel
            BR.Cancel
        
            Set TK = Nothing
            Set UM = Nothing
            Set CH = Nothing
            Set BN = Nothing
            Set BR = Nothing
            
            Set gd = Nothing
            
            gApp.Log "Kihon Cancel"
            Exit Sub
        End If
        
        SafeSeek UM, Array("KettoNum"), _
                    Array(TK("KettoNum").value)
        SafeSeek CH, Array("ChokyosiCode"), _
                    Array(TK("ChokyosiCode").value)
        If Not UM.EOF Then
            SafeSeek BN, Array("BanusiCode"), _
                        Array(UM("BanusiCode").value)
            SafeSeek BR, Array("BreederCode"), _
                        Array(UM("BreederCode").value)
        End If
        lngCP = 0
        With gd
            .SetItemMatrix i + rowOffset, lngCP, mCC.UMKG1(TK("UmaKigoCD")), mCC.UMKG3(TK("UmaKigoCD"))
            .SetItemMatrix i + rowOffset, lngCP, Trim$(TK("Bamei")), Trim$(TK("Bamei")), , "UM", IfExist(UM, "KettoNum")

            .SetItemMatrix i + rowOffset, lngCP, mCC.SEIB4(IfExist(UM, "SexCD")) '& Val(IfExist(UM,"Barei"))
            .SetItemMatrix i + rowOffset, lngCP, mCC.KEIR1(IfExist(UM, "KeiroCD"))
            .SetItemMatrix i + rowOffset, lngCP, Format$(val(TK("Futan")) / 10, "#0.0"), , ">-"

            .SetItemMatrix i + rowOffset, lngCP, mSC.Money3(val(IfExist(UM, "RuikeiHonsyoHeiti")) + val(IfExist(UM, "RuikeiHonsyoSyogai"))), , ">-"
            .SetItemMatrix i + rowOffset, lngCP, mSC.Money3(val(IfExist(UM, "RuikeiSyutokuHeichi")) + val(IfExist(UM, "RuikeiSyutokuSyogai"))), , ">-"
            .SetItemMatrix i + rowOffset, lngCP, mCC.TZSZ4(TK("TozaiCD")) & Trim$(TK("ChokyosiRyakusyo")), , , "CH", IfExist(CH, "ChokyosiCode")

            .SetItemMatrix i + rowOffset, lngCP, Trim$(IfExist(UM, "BanusiName")), , , "BN", IIf(UM.EOF, "", IfExist(BN, "BanusiCode"))
            .SetItemMatrix i + rowOffset, lngCP, Trim$(IfExist(UM, "BreederName")), , , "BR", IIf(UM.EOF, "", IfExist(BR, "BreederCode"))
            .SetItemMatrix i + rowOffset, lngCP, mSC.TK_KKBN(TK("Koryu"))
        End With
        
        TK.MoveNext

        i = i + 1
        If gd.Rows = i + rowOffset Then
            gd.Rows = gd.Rows + 10
        End If
    Loop
    gd.Rows = i + rowOffset
    
    RaiseEvent FetchedKihon(gd)
    Exit Sub
    
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: 血統グリッドを作る
'
'   備考: なし
'
Private Sub DirectMakeData_Ketto()
On Error GoTo ErrorHandler
    
    Const rowOffset As Long = 1 '' データ行オフセット
    
    Dim lngCP As Long   '' カラムポインタ
    Dim i     As Long   '' ループカウンタ
    
    
    Dim gd As clsGridData
    
    Dim TK As ADODB.Recordset
    Dim UM As ADODB.Recordset
    Dim HN As ADODB.Recordset

    Set gd = New clsGridData
    Set TK = mRS_TOKU
    Set UM = mRS_UMA
    Set HN = mRS_HANSYOKU
    
    With gd
        .Rows = 8 + rowOffset
        .Cols = 4
    
        ' カラムヘッダ登録
        lngCP = 0
        .SetItemMatrix 0, lngCP, "馬名"
        .SetItemMatrix 0, lngCP, "親(父/母)"
        .SetItemMatrix 0, lngCP, "二代"
        .SetItemMatrix 0, lngCP, "三代"
    
    End With
    
    ' グリッドデータ挿入
    TK.MoveFirst
    i = 0
    
    Do While Not TK.EOF
        DoEvents
        If mblnCancelFetching Then
            TK.Cancel
            UM.Cancel
            HN.Cancel
        
            Set TK = Nothing
            Set UM = Nothing
            Set HN = Nothing
            
            Set gd = Nothing
            
            gApp.Log "Ketto Cancel"
            Exit Sub
        End If
        SafeSeek UM, Array("KettoNum"), Array(TK("KettoNum").value)
        Call KettoSub(TK, UM, HN, gd, i, 0, "1", "3", "7")
        Call KettoSub(TK, UM, HN, gd, i, 1, "1", "3", "8")
        Call KettoSub(TK, UM, HN, gd, i, 2, "1", "4", "9")
        Call KettoSub(TK, UM, HN, gd, i, 3, "1", "4", "10")
        Call KettoSub(TK, UM, HN, gd, i, 4, "2", "5", "11")
        Call KettoSub(TK, UM, HN, gd, i, 5, "2", "5", "12")
        Call KettoSub(TK, UM, HN, gd, i, 6, "2", "6", "13")
        Call KettoSub(TK, UM, HN, gd, i, 7, "2", "6", "14")
        TK.MoveNext
        i = i + 1
        If gd.Rows <= i * 8 + rowOffset Then
            gd.Rows = gd.Rows + 8 * 4
        End If
    Loop
    gd.Rows = i * 8 + rowOffset
    
    RaiseEvent FetchedKetto(gd)
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: 血統情報作成用サブルーチン
'
'   備考: なし
'
Private Sub KettoSub(UA As ADODB.Recordset, UM As ADODB.Recordset, HN As ADODB.Recordset _
                    , gd As clsGridData, RowI As Long, RowP As Long _
                    , Field1 As String, Field2 As String, Field3 As String)
    Dim lngCP As Long
    
    Const rowOffset As Long = 1 '' データ行オフセット
    
    lngCP = 0
    With gd
        .SetItemMatrix (RowI * 8) + RowP + rowOffset, lngCP, Trim$(UA("Bamei")), Trim$(UA("Bamei")), , "UM", IfExist(UM, "KettoNum")
        
        If Not UM.EOF Then
            SafeSeek HN, Array("HansyokuNum"), Array(UM("Ketto3InfoHansyokuNum" & Field1).value)
            .SetItemMatrix (RowI * 8) + RowP + rowOffset, lngCP, KettoGetName(UM, HN, Field1), Trim$(IfExist(HN, "HansyokuNum")), , "HN", IfExist(HN, "HansyokuNum"), IIf(val(Field1) Mod 2 = 0, ColorMother, ColorFather)
            SafeSeek HN, Array("HansyokuNum"), Array(UM("Ketto3InfoHansyokuNum" & Field2).value)
            .SetItemMatrix (RowI * 8) + RowP + rowOffset, lngCP, KettoGetName(UM, HN, Field2), Trim$(IfExist(HN, "HansyokuNum")), , "HN", IfExist(HN, "HansyokuNum"), IIf(val(Field2) Mod 2 = 0, ColorMother, ColorFather)
            SafeSeek HN, Array("HansyokuNum"), Array(UM("Ketto3InfoHansyokuNum" & Field3).value)
            .SetItemMatrix (RowI * 8) + RowP + rowOffset, lngCP, KettoGetName(UM, HN, Field3), Trim$(IfExist(HN, "HansyokuNum")), , "HN", IfExist(HN, "HansyokuNum"), IIf(val(Field3) Mod 2 = 0, ColorMother, ColorFather)
        Else
            .SetItemMatrix (RowI * 8) + RowP + rowOffset, lngCP, Space(CLng(Field1)), Trim$(IfExist(HN, "HansyokuNum")), , , , IIf(val(Field1) Mod 2 = 0, ColorMother, ColorFather)
            .SetItemMatrix (RowI * 8) + RowP + rowOffset, lngCP, Space(CLng(Field2)), Trim$(IfExist(HN, "HansyokuNum")), , , , IIf(val(Field2) Mod 2 = 0, ColorMother, ColorFather)
            .SetItemMatrix (RowI * 8) + RowP + rowOffset, lngCP, Space(CLng(Field3)), Trim$(IfExist(HN, "HansyokuNum")), , , , IIf(val(Field3) Mod 2 = 0, ColorMother, ColorFather)
        End If
    End With
End Sub


'
'   機能: 血統情報から馬名を得る
'
'   備考: なし
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
            KettoGetName = Space(CLng(FieldNum))
        End If
    Else
        str = Trim$(UM("Ketto3InfoBamei" & FieldNum))
        If str <> "" Then
            KettoGetName = str
        Else
            KettoGetName = Space(CLng(FieldNum))
        End If
    End If
End Function

'
'   機能: 条件別グリッドを作る
'
'   備考: なし
'
Private Sub DirectMakeData_JokenBetu()
On Error GoTo ErrorHandler
    
    Const rowOffset As Long = 2 '' データ行オフセット
        
    Dim lngCP     As Long '' カラムポインタ
    Dim lngNumSet As Long '' セット数
    Dim i         As Long '' ループカウンタ
    Dim j         As Long '' ループカウンタ
    Dim k         As Long '' ループカウンタ
    Dim gd        As clsGridData
    
    Dim TK As ADODB.Recordset
    Dim UM As ADODB.Recordset

    Set TK = mRS_TOKU
    Set UM = mRS_UMA
    
    Set gd = New clsGridData
    
    ' １～５着＆着外を１セットとして、何セット表示するかどうかは、馬場によって違う
    Select Case mBuf.TrackCD
    Case "10" To "26", "28" ' 芝・ダート
        lngNumSet = 8
    Case "51" To "59"       ' 障害
        lngNumSet = 6
    Case Else               ' サンド・その他
        lngNumSet = 1
    End Select
    
    ' カラムヘッダ
    With gd
        
        .Rows = 26 + rowOffset ' ２段ヘッダ
        ' カラム数は馬場によって違う
        .Cols = 1 + (6 * lngNumSet)
        
        ' １行目 カラムヘッダ登録
        lngCP = 0
        .SetItemMatrix 0, lngCP, "馬名", , "<-"
        For i = 0 To 5
            .SetItemMatrix 0, lngCP, "総合", , "^_"
        Next i
        ' 馬場あわせて選択
        Select Case mBuf.TrackCD
        Case "10" To "22"        ' 芝
            For i = 0 To 5
                .SetItemMatrix 0, lngCP, "芝", , "^_"
            Next i
            ' 左右直
            Select Case mBuf.TrackCD
            Case "10"
                For i = 0 To 5
                    .SetItemMatrix 0, lngCP, "芝直", , "^_"
                Next i
            Case "11" To "16"
                For i = 0 To 5
                    .SetItemMatrix 0, lngCP, "芝左", , "^_"
                Next i
            Case Else
                For i = 0 To 5
                    .SetItemMatrix 0, lngCP, "芝右", , "^_"
                Next i
            End Select
            ' 距離
            Select Case mBuf.KYORI
            Case Is <= "1600"
                For i = 0 To 5
                    .SetItemMatrix 0, lngCP, "芝1600m以下", , "^_"
                Next i
            Case Is >= "2201"
                For i = 0 To 5
                    .SetItemMatrix 0, lngCP, "芝2201m以上", , "^_"
                Next i
            Case Else
                For i = 0 To 5
                    .SetItemMatrix 0, lngCP, "芝1601m～2200m", , "^_"
                Next i
            End Select
            
            For i = 0 To 5
                .SetItemMatrix 0, lngCP, "芝良", , "^_"
            Next i
            For i = 0 To 5
                .SetItemMatrix 0, lngCP, "芝稍重", , "^_"
            Next i
            For i = 0 To 5
                .SetItemMatrix 0, lngCP, "芝重", , "^_"
            Next i
            For i = 0 To 5
                .SetItemMatrix 0, lngCP, "芝不良", , "^_"
            Next i
        Case "23" To "26", "29"  ' ダート
            For i = 0 To 5
                .SetItemMatrix 0, lngCP, "ダート", , "^_"
            Next i
            ' 左右直
            Select Case mBuf.TrackCD
            Case "29"
                For i = 0 To 5
                    .SetItemMatrix 0, lngCP, "ダート直", , "^_"
                Next i
            Case "23", "25"
                For i = 0 To 5
                    .SetItemMatrix 0, lngCP, "ダート左", , "^_"
                Next i
            Case "24", "26"
                For i = 0 To 5
                    .SetItemMatrix 0, lngCP, "ダート右", , "^_"
                Next i
            End Select
            ' 距離
            Select Case mBuf.KYORI
            Case Is <= "1600"
                For i = 0 To 5
                    .SetItemMatrix 0, lngCP, "ダート1600m以下", , "^_"
                Next i
            Case Is >= "2201"
                For i = 0 To 5
                    .SetItemMatrix 0, lngCP, "ダート2201m以上", , "^_"
                Next i
            Case Else
                For i = 0 To 5
                    .SetItemMatrix 0, lngCP, "ダート1601m～2200m", , "^_"
                Next i
            End Select
            
            For i = 0 To 5
                .SetItemMatrix 0, lngCP, "ダート良", , "^_"
            Next i
            For i = 0 To 5
                .SetItemMatrix 0, lngCP, "ダート稍重", , "^_"
            Next i
            For i = 0 To 5
                .SetItemMatrix 0, lngCP, "ダート重", , "^_"
            Next i
            For i = 0 To 5
                .SetItemMatrix 0, lngCP, "ダート不良", , "^_"
            Next i
        
        Case "51" To "59"        ' 障害
            For i = 0 To 5
                .SetItemMatrix 0, lngCP, "障害", , "^_"
            Next i
            ' 障害左右直は無い
            ' 障害距離別は無い
            
            For i = 0 To 5
                .SetItemMatrix 0, lngCP, "障害良", , "^_"
            Next i
            For i = 0 To 5
                .SetItemMatrix 0, lngCP, "障害稍重", , "^_"
            Next i
            For i = 0 To 5
                .SetItemMatrix 0, lngCP, "障害重", , "^_"
            Next i
            For i = 0 To 5
                .SetItemMatrix 0, lngCP, "障害不良", , "^_"
            Next i
        End Select
                

        ' 2行目 カラムヘッダ登録
        lngCP = 0
        .SetItemMatrix 1, lngCP, "馬名", , "<-"
        For i = 0 To lngNumSet - 1
            .SetItemMatrix 1, lngCP, "1着", , "^_"
            .SetItemMatrix 1, lngCP, "2着", , "^_"
            .SetItemMatrix 1, lngCP, "3着", , "^_"
            .SetItemMatrix 1, lngCP, "4着", , "^_"
            .SetItemMatrix 1, lngCP, "5着", , "^_"
            .SetItemMatrix 1, lngCP, "着外", , "^_"
        Next i
    End With
    
    ' グリッドデータ挿入
    TK.MoveFirst
    i = 0
    Do While Not TK.EOF
        DoEvents
        If mblnCancelFetching Then
            TK.Cancel
            UM.Cancel
            
            Set TK = Nothing
            Set UM = Nothing
            
            Set gd = Nothing
            
            gApp.Log "JokenBetu Cancel"
            Exit Sub
        End If
        SafeSeek UM, Array("KettoNum"), Array(TK("KettoNum").value)
        lngCP = 0
        With gd
            .SetItemMatrix i + rowOffset, lngCP, Trim$(TK("Bamei")), Trim$(TK("Bamei")), "<^", "UM", IfExist(UM, "KettoNum")
            If Not UM.EOF Then
                For j = 1 To 6
                    .SetItemMatrix i + rowOffset, lngCP, mSC.ValStr(UM("SogoChakukaisu" & j)), , "^^"
                Next j
                
                ' 馬場あわせて選択
                Select Case mBuf.TrackCD
                Case "10" To "22"        ' 芝
                    For j = 1 To 6
                        If UM("Ba1Chakukaisu" & j) <> "" And UM("Ba2Chakukaisu" & j) <> "" And UM("Ba3Chakukaisu" & j) <> "" Then
                            .SetItemMatrix i + rowOffset, lngCP, val(UM("Ba1Chakukaisu" & j)) + val(UM("Ba2Chakukaisu" & j)) + val(UM("Ba3Chakukaisu" & j)), , "^_"
                        Else
                            .SetItemMatrix i + rowOffset, lngCP, "", , "^_"
                        End If
                    Next j
                    ' 左右直
                    Select Case mBuf.TrackCD
                    Case "10"
                        For j = 1 To 6
                            .SetItemMatrix i + rowOffset, lngCP, mSC.ValStr(UM("Ba1Chakukaisu" & j)), , "^_"  '芝直
                        Next j
                    Case "11" To "16"
                        For j = 1 To 6
                            .SetItemMatrix i + rowOffset, lngCP, mSC.ValStr(UM("Ba2Chakukaisu" & j)), , "^_"  '芝左
                        Next j
                    Case Else
                        For j = 1 To 6
                            .SetItemMatrix i + rowOffset, lngCP, mSC.ValStr(UM("Ba3Chakukaisu" & j)), , "^_"  '芝右
                        Next j
                    End Select
                    ' 距離
                    Select Case mBuf.KYORI
                    Case Is <= "1600"
                        For j = 1 To 6
                            .SetItemMatrix i + rowOffset, lngCP, mSC.ValStr(UM("Kyori1Chakukaisu" & j)), , "^_"  '芝1600m以下
                        Next j
                    Case Is >= "2201"
                        For j = 1 To 6
                            .SetItemMatrix i + rowOffset, lngCP, mSC.ValStr(UM("Kyori3Chakukaisu" & j)), , "^_"  '芝2201m以上
                        Next j
                    Case Else
                        For j = 1 To 6
                            .SetItemMatrix i + rowOffset, lngCP, mSC.ValStr(UM("Kyori2Chakukaisu" & j)), , "^_"  '芝1601m～2200m
                        Next j
                    End Select
                    
                    For j = 1 To 6
                        .SetItemMatrix i + rowOffset, lngCP, mSC.ValStr(UM("Jyotai1Chakukaisu" & j)), , "^_"  '芝良
                    Next j
                    For j = 1 To 6
                        .SetItemMatrix i + rowOffset, lngCP, mSC.ValStr(UM("Jyotai2Chakukaisu" & j)), , "^_"  '芝稍重
                    Next j
                    For j = 1 To 6
                        .SetItemMatrix i + rowOffset, lngCP, mSC.ValStr(UM("Jyotai3Chakukaisu" & j)), , "^_"  '芝重
                    Next j
                    For j = 1 To 6
                        .SetItemMatrix i + rowOffset, lngCP, mSC.ValStr(UM("Jyotai4Chakukaisu" & j)), , "^_"  '芝不良
                    Next j
                    Case "23" To "26", "29"  ' ダート
                    For j = 1 To 6
                        If UM("Ba4Chakukaisu" & j) <> "" And UM("Ba5Chakukaisu" & j) <> "" And UM("Ba6Chakukaisu" & j) <> "" Then
                            .SetItemMatrix i + rowOffset, lngCP, val(UM("Ba4Chakukaisu" & j)) + val(UM("Ba5Chakukaisu" & j)) + val(UM("Ba6Chakukaisu" & j)), , "^_" 'ダート
                        Else
                            .SetItemMatrix i + rowOffset, lngCP, "", , "^_" 'ダート
                        End If
                    Next j
                    ' 左右直
                    Select Case mBuf.TrackCD
                    Case "29"
                        For j = 1 To 6
                            .SetItemMatrix i + rowOffset, lngCP, mSC.ValStr(UM("Ba4Chakukaisu" & j)), , "^_"  'ダート直
                        Next j
                    Case "23", "25"
                        For j = 1 To 6
                            .SetItemMatrix i + rowOffset, lngCP, mSC.ValStr(UM("Ba5Chakukaisu" & j)), , "^_"  'ダート左
                        Next j
                    Case "24", "26"
                        For j = 1 To 6
                            .SetItemMatrix i + rowOffset, lngCP, mSC.ValStr(UM("Ba6Chakukaisu" & j)), , "^_"  'ダート右
                        Next j
                    End Select
                    ' 距離
                    Select Case mBuf.KYORI
                    Case Is <= "1600"
                        For j = 1 To 6
                            .SetItemMatrix i + rowOffset, lngCP, mSC.ValStr(UM("Kyori4Chakukaisu" & j)), , "^_"  'ダート1600m以下
                        Next j
                    Case Is >= "2201"
                        For j = 1 To 6
                            .SetItemMatrix i + rowOffset, lngCP, mSC.ValStr(UM("Kyori6Chakukaisu" & j)), , "^_"  'ダート2201m以上
                        Next j
                    Case Else
                        For j = 1 To 6
                            .SetItemMatrix i + rowOffset, lngCP, mSC.ValStr(UM("Kyori5Chakukaisu" & j)), , "^_"  'ダート1601m～2200m
                        Next j
                    End Select
                    
                    For j = 1 To 6
                        .SetItemMatrix i + rowOffset, lngCP, mSC.ValStr(UM("Jyotai5Chakukaisu" & j)), , "^_"  'ダート良
                    Next j
                    For j = 1 To 6
                        .SetItemMatrix i + rowOffset, lngCP, mSC.ValStr(UM("Jyotai6Chakukaisu" & j)), , "^_"  'ダート稍重
                    Next j
                    For j = 1 To 6
                        .SetItemMatrix i + rowOffset, lngCP, mSC.ValStr(UM("Jyotai7Chakukaisu" & j)), , "^_"  'ダート重
                    Next j
                    For j = 1 To 6
                        .SetItemMatrix i + rowOffset, lngCP, mSC.ValStr(UM("Jyotai8Chakukaisu" & j)), , "^_"  'ダート不良
                    Next j
                
                Case "51" To "59"        ' 障害
                    For j = 1 To 6
                        .SetItemMatrix i + rowOffset, lngCP, mSC.ValStr(UM("Ba7Chakukaisu" & j)), , "^_"  '障害
                    Next j
                    ' 障害左右直は無い
                    ' 障害距離別は無い
                    
                    For j = 1 To 6
                        .SetItemMatrix i + rowOffset, lngCP, mSC.ValStr(UM("Jyotai9Chakukaisu" & j)), , "^_"  '障害良
                    Next j
                    For j = 1 To 6
                        .SetItemMatrix i + rowOffset, lngCP, mSC.ValStr(UM("Jyotai10Chakukaisu" & j)), , "^_"  '障害稍重
                    Next j
                    For j = 1 To 6
                        .SetItemMatrix i + rowOffset, lngCP, mSC.ValStr(UM("Jyotai11Chakukaisu" & j)), , "^_"  '障害重
                    Next j
                    For j = 1 To 6
                        .SetItemMatrix i + rowOffset, lngCP, mSC.ValStr(UM("Jyotai12Chakukaisu" & j)), , "^_"  '障害不良
                    Next j
                End Select
            End If
        End With
        TK.MoveNext
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
                gd.ItemMatrix(k, 1 + (i * 6) + j).BGColor = IIf(i Mod 2 = 0, RGB(230, 230, 255), RGB(245, 245, 255))
            Next k
        Next j
    Next i
    
    
    RaiseEvent FetchedJokenBetu(gd)
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: 持ちタイムグリッドを作る
'
'   備考: なし
'
Private Sub DirectMakeData_MochiTime()
On Error GoTo ErrorHandler
    Const rowOffset As Long = 2 '' データ行オフセット
    Dim gd     As clsGridData
    Dim strSQL As String
    Dim lngCP  As Long

    Dim cUA As ADODB.Connection
    Dim cUB As ADODB.Connection
    Dim cRA As ADODB.Connection
    
    Dim TK As ADODB.Recordset
    Dim UM As ADODB.Recordset
    
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
    
    Set TK = mRS_TOKU
    Set UM = mRS_UMA

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
        .Rows = 300 + 2 ' ２段ヘッダ
        .Cols = 1 + (8 * 2)

        ' １行目 カラムヘッダ登録
        lngCP = 0
        .SetItemMatrix 0, lngCP, "馬名", , "<-"
        For i = 0 To 7
            .SetItemMatrix 0, lngCP, mCC.KIBJ3(mBuf.id.JyoCD) & " " & mCC.TRCK4(mBuf.TrackCD) & " " & val(mBuf.KYORI) & "m", , "<-" ' 中山 芝 1200m
        Next i
        For i = 0 To 7
            .SetItemMatrix 0, lngCP, mCC.TRCK4(mBuf.TrackCD) & " " & val(mBuf.KYORI) & "m", , "<-" ' 1200m
        Next i
        lngCP = 0
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
    TK.MoveFirst
    Do While Not TK.EOF
        
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
        
        If TK("KettoNum") <> "0000000000" Then
            'KettoNumが初期値でないなら
            rUA.Open "SELECT * FROM UMA_RACE_A WHERE KettoNum='" & TK("KettoNum") & "'", cUA, adOpenKeyset, adLockReadOnly, adCmdText
            ' 過去レースループ
            Do While Not rUA.EOF
                ' バックグラウンド
                DoEvents
                If mblnCancelFetching Then
                    rUA.Cancel
                    rUB.Cancel
                    rRA.Cancel
                    TK.Cancel
                    UM.Cancel
                    
                    rUA.Close
                    rUB.Close
                    rRA.Close
                    
                    Set rUA = Nothing
                    Set rUB = Nothing
                    Set rRA = Nothing
                    Set TK = Nothing
                    Set UM = Nothing
                    
                    Set gd = Nothing
                    
                    cUA.Close
                    cUB.Close
                    cRA.Close
                    
                    Set cUA = Nothing
                    Set cUB = Nothing
                    Set cRA = Nothing
                    
                    gApp.Log "Mochi Cancel"
                    Exit Sub
                End If
                
                ' 過去のみ
                If rUA("Year") & rUA("MonthDay") < (mKey.Year & mKey.MonthDay) Then
                    SafeSeek rRA, Array("Year", "MonthDay", "JyoCD", "Kaiji", "Nichiji", "RaceNum"), _
                                Array(rUA("Year").value, rUA("MonthDay").value, rUA("JyoCD").value, rUA("Kaiji").value, rUA("Nichiji").value, rUA("RaceNum").value)
                    ' レース情報が存在する場合のみ
                    If Not rRA.EOF Then
                        ' このレースと同じ距離のみ
                        If rRA("Kyori") = mBuf.KYORI And MochiTimeSubSameTrackCD(rRA("TrackCD").value, mBuf.TrackCD) Then
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
                                If rUA("JyoCD") = mBuf.id.JyoCD Then
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
            lngCP = 0
            SafeSeek UM, Array("KettoNum"), Array(TK("KettoNum").value)
            With gd
                .SetItemMatrix i + rowOffset, lngCP, Trim$(TK("Bamei")), Trim$(TK("Bamei")), "<^", "UM", IfExist(UM, "KettoNum")
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
                strSQL = strSQL & " AND [KettoNum]='" & TK("KettoNum") & "'"
                rUA.Open strSQL, cUA, adOpenKeyset, adLockReadOnly, adCmdText
                rUB.Seek Array(rUA("Year"), rUA("MonthDay"), rUA("JyoCD"), rUA("RaceNum"), rUA("Umaban"), rUA("KettoNum"))
                rRA.Seek Array(rUA("Year"), rUA("MonthDay"), rUA("JyoCD"), rUA("Kaiji"), rUA("Nichiji"), rUA("RaceNum"))
                With gd
                    .SetItemMatrix i + rowOffset, lngCP, Format$(rUB("Time"), "@:@@.@"), , ">^"
                    .SetItemMatrix i + rowOffset, lngCP, IIf(rUB("HaronTimeL3") = "999", "", mSC.SSSs(rUB("HaronTimeL3"))), , ">^"
                    .SetItemMatrix i + rowOffset, lngCP, mSC.YMD3(rUA("Year") & rUA("MonthDay")), Trim$(rRA("Hondai")) & mCC.GRAD3(rRA("GradeCD")), ">^", "RA", rUA("Year") & rUA("MonthDay") & rUA("JyoCD") & rUA("Kaiji") & rUA("Nichiji") & rUA("RaceNum")
                    .SetItemMatrix i + rowOffset, lngCP, mSC.ValStr(rUA("Kaiji")) & mCC.KIBJ3(rUA("JyoCD")) & mSC.ValStr(rUA("Nichiji")), , ">^", "RA", rUA("Year") & rUA("MonthDay") & rUA("JyoCD") & rUA("Kaiji") & rUA("Nichiji") & rUA("RaceNum")
                    .SetItemMatrix i + rowOffset, lngCP, mSC.ValStr(rUB("KakuteiJyuni")), , ">^", , , gApp.GetChakujyunColor(rUB("KakuteiJyuni"))
                    .SetItemMatrix i + rowOffset, lngCP, mCC.IJKB2(rUB("IJyoCD")), , "<^"
                    .SetItemMatrix i + rowOffset, lngCP, mSC.BabaGrid(rRA), , "<^"
                    .SetItemMatrix i + rowOffset, lngCP, Format$(rUB("Futan") / 10, "#0.0"), , ">^"
                End With ' gd
                rUA.Close
            Else
                lngCP = 9
            End If
            ' 距離 持ちタイム
            If strYear <> "" Then
                strSQL = "SELECT * FROM UMA_RACE_A "
                strSQL = strSQL & " WHERE [Year]='" & strYear & "'"
                strSQL = strSQL & " AND [MonthDay]='" & strMonthDay & "'"
                strSQL = strSQL & " AND [JyoCD]='" & strJyoCD & "'"
                strSQL = strSQL & " AND [Kaiji]='" & strKaiji & "'"
                strSQL = strSQL & " AND [Nichiji]='" & strNichiji & "'"
                strSQL = strSQL & " AND [RaceNum]='" & strRaceNum & "'"
                strSQL = strSQL & " AND [Umaban]='" & strUmaban & "'"
                strSQL = strSQL & " AND [KettoNum]='" & TK("KettoNum") & "'"
                rUA.Open strSQL, cUA, adOpenKeyset, adLockReadOnly, adCmdText
                rUB.Seek Array(rUA("Year"), rUA("MonthDay"), rUA("JyoCD"), rUA("RaceNum"), rUA("Umaban"), rUA("KettoNum"))
                rRA.Seek Array(rUA("Year"), rUA("MonthDay"), rUA("JyoCD"), rUA("Kaiji"), rUA("Nichiji"), rUA("RaceNum"))
                With gd
                    .SetItemMatrix i + rowOffset, lngCP, Format$(rUB("Time"), "@:@@.@"), , ">^"
                    .SetItemMatrix i + rowOffset, lngCP, IIf(rUB("HaronTimeL3") = "999", "", mSC.SSSs(rUB("HaronTimeL3"))), , ">^"
                    .SetItemMatrix i + rowOffset, lngCP, mSC.YMD3(rUA("Year") & rUA("MonthDay")), Trim$(rRA("Hondai")) & mCC.GRAD3(rRA("GradeCD")), ">^", "RA", rUA("Year") & rUA("MonthDay") & rUA("JyoCD") & rUA("Kaiji") & rUA("Nichiji") & rUA("RaceNum")
                    .SetItemMatrix i + rowOffset, lngCP, mSC.ValStr(rUA("Kaiji")) & mCC.KIBJ3(rUA("JyoCD")) & mSC.ValStr(rUA("Nichiji")), , ">^", "RA", rUA("Year") & rUA("MonthDay") & rUA("JyoCD") & rUA("Kaiji") & rUA("Nichiji") & rUA("RaceNum")
                    .SetItemMatrix i + rowOffset, lngCP, mSC.ValStr(rUB("KakuteiJyuni")), , ">^", , , gApp.GetChakujyunColor(rUB("KakuteiJyuni"))
                    .SetItemMatrix i + rowOffset, lngCP, mCC.IJKB2(rUB("IJyoCD")), , "<^"
                    .SetItemMatrix i + rowOffset, lngCP, mSC.BabaGrid(rRA), , "<^"
                    .SetItemMatrix i + rowOffset, lngCP, Format$(rUB("Futan") / 10, "#0.0"), , ">^"
                End With ' gd
                rUA.Close
            End If
        Else
            '血統番号が初期値の時は馬名まで表示
            lngCP = 0
            SafeSeek UM, Array("KettoNum"), Array(TK("KettoNum").value)
            With gd
                .SetItemMatrix i + rowOffset, lngCP, Trim$(TK("Bamei")), Trim$(TK("Bamei")), "<^", "UM", IfExist(UM, "KettoNum")
            End With
        End If
        
        gApp.Log "Mochitime" & i
        TK.MoveNext
        i = i + 1
        If gd.Rows >= i + rowOffset Then
            gd.Rows = gd.Rows + 10
        End If
    Loop ' 出走馬ループ
    gd.Rows = i + rowOffset
    RaiseEvent FetchedMotiTIme(gd)
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: トラックコードを判定
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
On Error GoTo ErrorHandler
    Const rowOffset As Long = 1 '' データ行オフセット
    
    Dim gd     As clsGridData
    Dim strSQL As String
    Dim lngCP  As Long
    Dim lngKakoMax As Long
    
    Dim cUA As ADODB.Connection
    Dim cUB As ADODB.Connection
    Dim cRA As ADODB.Connection
    
    Dim TK As ADODB.Recordset
    Dim UM As ADODB.Recordset
    
    Dim rUA As ADODB.Recordset
    Dim rUB As ADODB.Recordset
    Dim rRA As ADODB.Recordset
    
    Dim i As Long
    
    mblnNowKakoFetching = True
    
    Set gd = New clsGridData
    Set TK = mRS_TOKU
    Set UM = mRS_UMA

    ' イベントを拾わないように
    Set mAsyncCN_TOKU = Nothing
    
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
        .Cols = lngKakoMax + 1
        .Rows = 50 ' 大きめに
        ' カラムヘッダ挿入
        lngCP = 0
        .SetItemMatrix 0, lngCP, "馬名", , "<-"
        For i = 0 To gApp.R_KakoNum - 1
            .SetItemMatrix 0, lngCP, IIf(i = 0, "前走", i + 1 & "走前")
        Next
    End With
    
    
    ' 出走馬ループ
    i = 0
    TK.MoveFirst
    Do While Not TK.EOF
        strSQL = "SELECT * FROM UMA_RACE_A "
        strSQL = strSQL & "WHERE KettoNum='" & TK("KettoNum") & "'"
        strSQL = strSQL & " ORDER BY [Year] DESC, [MonthDay] DESC, [RaceNum] DESC, [JyoCD] DESC"
        rUA.Open strSQL, cUA, adOpenKeyset, adLockReadOnly, adCmdText
        lngCP = 0
        SafeSeek UM, Array("KettoNum"), Array(TK("KettoNum").value)

        With gd
            .SetItemMatrix i + rowOffset, lngCP, Trim$(TK("Bamei")), TK("KettoNum"), "<-", "UM", IfExist(UM, "KettoNum")
            .ItemMatrix(i + rowOffset, lngCP - 1).Text = vbCr & vbCr & .ItemMatrix(i + rowOffset, lngCP - 1).Text & vbCr & vbCr & vbCr
        End With 'gd
        If TK("KettoNum") <> "0000000000" Then
            ' 過去レースループ(KettoNumが初期値でないなら)
            Do While Not rUA.EOF
                ' バックグラウンド
                DoEvents
                If mblnCancelFetching Then
                    rUA.Cancel
                    rUB.Cancel
                    rRA.Cancel
                    TK.Cancel
                    UM.Cancel
                    
                    rUA.Close
                    rUB.Close
                    rRA.Close
                    
                    Set rUA = Nothing
                    Set rUB = Nothing
                    Set rRA = Nothing
                    Set TK = Nothing
                    Set UM = Nothing
                    
                    Set gd = Nothing
                    
                    cUA.Close
                    cUB.Close
                    cRA.Close
                    
                    Set cUA = Nothing
                    Set cUB = Nothing
                    Set cRA = Nothing
                    
                    gApp.Log "Kako Cancel"
                    Exit Sub
                End If
                ' キャンセル感知
                If mblnCancelKakoFetching Then
                    mblnNowKakoFetching = False
                    mblnCancelKakoFetching = False
                    Exit Sub
                End If
                
                ' 過去のみ
                If rUA("Year") & rUA("MonthDay") & rUA("RaceNum") < (mKey.Year & mKey.MonthDay & mKey.RaceNum) Then
                    rRA.Seek Array(rUA("Year"), rUA("MonthDay"), rUA("JyoCD"), rUA("Kaiji"), rUA("Nichiji"), rUA("RaceNum"))
                    rUB.Seek Array(rUA("Year"), rUA("MonthDay"), rUA("JyoCD"), rUA("RaceNum"), rUA("Umaban"), rUA("KettoNum"))
                    gd.SetItemMatrix i + rowOffset, lngCP, getKakoStr(rUA, rUB, rRA), , "<^", "RA" _
                                , rUA("Year") & rUA("MonthDay") & rUA("JyoCD") & rUA("Kaiji") & rUA("Nichiji") & rUA("RaceNum"), gApp.GetChakujyunColor(val(rUB("KakuteiJyuni"))), RGB(1, 1, 1)
                End If ' 過去のみ
                rUA.MoveNext
                If lngCP - 1 > lngKakoMax Then
                    Exit Do
                End If
            Loop ' 過去レースループ
        End If
        rUA.Close
        TK.MoveNext
        i = i + 1
        If gd.Rows >= i + rowOffset Then
            gd.Rows = gd.Rows + 10
        End If
    Loop ' 出走馬ループ
    gd.Rows = i + rowOffset
    RaiseEvent FetchedKako(gd)
    
    mblnNowKakoFetching = False
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: 過去送グリッド用文字列を生成する
'
'   備考: なし
'
Private Function getKakoStr(UA As ADODB.Recordset, UB As ADODB.Recordset, RA As ADODB.Recordset) As String
On Error GoTo ErrorHandler
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
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Function
