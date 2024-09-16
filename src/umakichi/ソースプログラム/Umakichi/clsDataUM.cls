VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDataUM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "競走馬 データクラス"
'
'   競走馬 データクラス
'

Option Explicit

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   外部関数(イベント)
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

' イベント
Public Event NoData()
Public Event NoTabData(Index As Long)

Public Event FetchedHeader()
Public Event FetchedLinkLabels()
Public Event FetchedKetto(gd As clsGridData)
Public Event FetchedKako(gd As clsGridData)
Public Event FetchedJokenBetu(gd2 As clsGridData, gd5 As clsGridData, gd6 As clsGridData)
Public Event FetchedTime(gd As clsGridData)
Public Event FetchedHanro(gd As clsGridData)

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部変数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

' 変換ユーティリティー
Private mCC As clsCodeConverter          '' JV-Data のコード変換メソッド群
Private mSC As clsStringConverter          '' JV-Data のコード変換メソッド群

' Viewerに提供する為のデータ格納変数
Private mstrLabels(0 To 8) As String      '' テキストデータ
Private mstrLinkLabels(0 To 2) As clsGridItem     '' リンクデータ

'モジュール内で使用する為のデータ格納変数
Private mbuf_UM          As JV_UM_UMA  ''競走馬構造体
Private mKey             As clsKeyUM   ''競走馬キー

'コネクション
Private mCN_UMA         As ADODB.Connection
Private mCN_CHOKYO      As ADODB.Connection
Attribute mCN_CHOKYO.VB_VarHelpID = -1
Private mCN_SEISAN      As ADODB.Connection
Private mCN_BANUSI      As ADODB.Connection
Attribute mCN_BANUSI.VB_VarHelpID = -1
Private mCN_HANSYOKU    As ADODB.Connection
Private mCN_HANRO       As ADODB.Connection
Attribute mCN_HANRO.VB_VarHelpID = -1
Private mCN_KISHU       As ADODB.Connection
Attribute mCN_KISHU.VB_VarHelpID = -1
Private mCN_RACE        As ADODB.Connection
Attribute mCN_RACE.VB_VarHelpID = -1
Private WithEvents mAsyncCN_UMA_RACE_A As ADODB.Connection
Attribute mAsyncCN_UMA_RACE_A.VB_VarHelpID = -1
Private mCN_UMA_RACE_B  As ADODB.Connection
Attribute mCN_UMA_RACE_B.VB_VarHelpID = -1

'レコードセット
Private mRS_UMA        As ADODB.Recordset
Private mRS_CHOKYO     As ADODB.Recordset
Private mRS_SEISAN     As ADODB.Recordset
Private mRS_BANUSI     As ADODB.Recordset
Private mRS_HANSYOKU   As ADODB.Recordset
Private mRS_HANRO      As ADODB.Recordset
Attribute mRS_HANRO.VB_VarHelpID = -1
Private mRS_KISHU      As ADODB.Recordset
Private mRS_RACE       As ADODB.Recordset
Attribute mRS_RACE.VB_VarHelpID = -1
Private mRS_UMA_RACE_A As ADODB.Recordset
Attribute mRS_UMA_RACE_A.VB_VarHelpID = -1
Private mRS_UMA_RACE_B As ADODB.Recordset
Attribute mRS_UMA_RACE_B.VB_VarHelpID = -1

Private mblnCancelFetching As Boolean


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
'   機能: リンクラベル情報を取得するプロパティ
'
'   備考: なし
'
Public Property Get LinkLabels(Index As Integer) As clsGridItem
    Set LinkLabels = mstrLinkLabels(Index)
End Property


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   外部関数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: 取得を開始する
'
'   備考: 初期表示用データの取得
'
Public Function Fetch(ByRef objKey As clsKeyUM) As Boolean
On Error GoTo ErrorHandler
    Dim rs              As ADODB.Recordset
    Dim strSQL          As String          '' SQL文用
    Dim i               As Long            '' ループカウンタ
    Dim lngCP           As Long            '' セル入力用
    Dim lngRP           As Long            '' セル入力用
    
    Set mKey = objKey
    
    Set rs = New ADODB.Recordset

    '-------------
    ' UMA 情報取得
    '-------------
    strSQL = "SELECT * FROM UMA " & objKey.SQLWHEREString
    
    rs.Open strSQL, mCN_UMA, adOpenKeyset, adLockReadOnly, adCmdText

    ' データが無かったらFalseを返して終了
    If rs.EOF Or rs.BOF Then
        rs.Close
        Set rs = Nothing
        Fetch = False
        RaiseEvent NoData
        Exit Function
    End If

    ' UMAレコードをすべて構造体で保持する
    Call SetDataFromRS_UM(rs, mbuf_UM)
    
    ' ヘッダ情報を設定する
    Call Set_Header(rs)
        
    ' 条件別成績　タブ(2) grid(2,5,6)
    Call Set_JokenBetu
        
    rs.Close
    
    Set rs = Nothing
    
    '-----------------
    ' 非同期取得を行う
    '-----------------
    'UMA_RACE_Aを開く
    strSQL = "SELECT * FROM UMA_RACE_A WHERE KettoNum='" & mbuf_UM.KettoNum & "'"
    strSQL = strSQL & " AND DataKubun<>'1' AND DataKubun<>'2'"
    mAsyncCN_UMA_RACE_A.Execute strSQL, , adAsyncExecute
        
    gApp.Log "UM Fetch Complete"
    
    Fetch = True
    Exit Function
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Function
    

'
'   機能: キャンセル
'
'   備考: なし
'
Public Sub CancelFetching()
On Error GoTo ErrorHandler
    mblnCancelFetching = True
    
    Call freers(mRS_UMA_RACE_A)

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
    
    Set mCN_UMA = gApp.GetCN_UMA
    Set mCN_CHOKYO = gApp.GetCN_CHOKYO
    Set mCN_BANUSI = gApp.GetCN_BANUSI
    Set mCN_SEISAN = gApp.GetCN_SEISAN
    Set mCN_HANSYOKU = gApp.GetCN_HANSYOKU
    Set mCN_HANRO = gApp.GetCN_HANRO
    Set mCN_KISHU = gApp.GetCN_KISHU
    Set mCN_RACE = gApp.GetCN_RACE
    Set mAsyncCN_UMA_RACE_A = gApp.GetCN_UMA_RACE_A
    Set mCN_UMA_RACE_B = gApp.GetCN_UMA_RACE_B
    
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
'   機能: 競走馬レースレコードセットの取得完了イベントハンドラ
'
'   備考: なし
'
Private Sub mAsyncCN_UMA_RACE_A_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
On Error GoTo ErrorHandler

    If pError Is Nothing Then
        Set mRS_UMA_RACE_A = pRecordset
    Else
        With pError
            gApp.Log .Description & .SQLState & .Source & .Number
        End With
    End If
    gApp.Log "<mAsyncCN_UMA_RACE_A_ExecuteComplete"

    DoEvents
    If mblnCancelFetching Then
        Exit Sub
    End If
    
    '各タブの設定へ
    ' ヘッダリンクを設定する
    Call Set_HeaderLink
    
    ' 血統　タブ(0)
    Call Set_Ketto

    ' 坂路調教　タブ(4)
    Call Set_Hanro(mbuf_UM.KettoNum)

    If Not mblnCancelFetching Then
        If mRS_UMA_RACE_A.EOF Then
            RaiseEvent NoTabData(1)
            RaiseEvent NoTabData(3)
        Else
            ' テーブルを開く(RACE,UMA_RACE_B)
            Call RS_open
        
            ' 過去成績　タブ(1)
            Call Set_Kako(mbuf_UM.KettoNum)
        
            ' 持ちタイム　タブ(3)
            Call Set_Time(mbuf_UM.KettoNum)
        
        End If
    End If
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: 競走馬ヘッダの設定
'
'   備考: なし
'
Private Sub Set_Header(rs As ADODB.Recordset)
    Dim i As Long

    With mbuf_UM
        
        mstrLabels(0) = ""
        mstrLabels(0) = mstrLabels(0) & mSC.YMD1(rs("MakeDate")) & "作成データ"   '作成年月日
        
        ' 太字レース名等グレーバー用文字列生成
        mstrLabels(1) = ""
        mstrLabels(1) = mstrLabels(1) & mCC.UMKG1(.UmaKigoCD) & Trim$(.BAMEI)  '馬名
    
        mstrLabels(2) = ""
        mstrLabels(2) = mstrLabels(2) & mCC.SEIB4(.SexCD)  '性別
        mstrLabels(2) = mstrLabels(2) & IIf(.DelKubun = 0, mSC.Age(.head.MakeDate.Year, .BirthDate.Year), "") & " "  '性別 年齢
        mstrLabels(2) = mstrLabels(2) & mCC.KEIR1(.KeiroCD) & " "  '毛色
        mstrLabels(2) = mstrLabels(2) & mCC.HINS2(.HinsyuCD) & " "  '品種
        mstrLabels(2) = mstrLabels(2) & IIf(val(rs("BirthDate")) = 0, "", mSC.YMD1(rs("BirthDate")) & "生 ") '生年月日
        mstrLabels(2) = mstrLabels(2) & mSC.UM_DelKubun(.DelKubun) & " "  '抹消区分
        mstrLabels(2) = mstrLabels(2) & mSC.YMD1(rs("DelDate"))  '抹消年月日
        mstrLabels(2) = mstrLabels(2) & vbCrLf  '改行
        mstrLabels(2) = mstrLabels(2) & IIf(val(rs("RegDate")) = 0, "", "競走馬登録：" & mSC.YMD1(rs("RegDate")))
        mstrLabels(2) = mstrLabels(2) & vbCrLf  '改行
        mstrLabels(2) = mstrLabels(2) & IIf(Trim$(rs("SanchiName")) = "", "", "産地　　　：" & Trim$(rs("SanchiName")))
        mstrLabels(2) = mstrLabels(2) & vbCrLf  '改行
        mstrLabels(2) = mstrLabels(2) & IIf(Trim$(rs("Syotai")) = "", "", "招待地域　：" & Trim$(rs("Syotai")))

        mstrLabels(3) = ""
        mstrLabels(3) = mstrLabels(3) & IIf(Trim$(.ChokyosiRyakusyo) = "", "", "調教師：") & vbCrLf
        mstrLabels(3) = mstrLabels(3) & IIf(Trim$(.BanusiName) = "", "", "馬主　：") & vbCrLf
        mstrLabels(3) = mstrLabels(3) & IIf(Trim$(.BreederName) = "", "", "生産者：")
        
        mstrLabels(4) = ""
        mstrLabels(4) = mstrLabels(4) & "平地本賞金　：" & mSC.Money2(.RuikeiHonsyoHeiti, 14)
        mstrLabels(4) = mstrLabels(4) & vbCrLf  '改行
        mstrLabels(4) = mstrLabels(4) & "平地付加賞金：" & mSC.Money2(.RuikeiFukaHeichi, 14)
        mstrLabels(4) = mstrLabels(4) & vbCrLf  '改行
        mstrLabels(4) = mstrLabels(4) & "平地収得賞金：" & mSC.Money2(.RuikeiSyutokuHeichi, 14)
        
        mstrLabels(5) = ""
        mstrLabels(5) = mstrLabels(5) & "障害本賞金　：" & mSC.Money2(.RuikeiHonsyoSyogai, 14)
        mstrLabels(5) = mstrLabels(5) & vbCrLf  '改行
        mstrLabels(5) = mstrLabels(5) & "障害付加賞金：" & mSC.Money2(.RuikeiFukaSyogai, 14)
        mstrLabels(5) = mstrLabels(5) & vbCrLf  '改行
        mstrLabels(5) = mstrLabels(5) & "障害収得賞金：" & mSC.Money2(.RuikeiSyutokuSyogai, 14)
        
        mstrLabels(6) = ""
        mstrLabels(6) = mstrLabels(6) & "逃げ回数：" & Format$(val(.Kyakusitu(0)), "@@@") & "回"
        mstrLabels(6) = mstrLabels(6) & vbCrLf    '改行
        mstrLabels(6) = mstrLabels(6) & "先行回数：" & Format$(val(.Kyakusitu(1)), "@@@") & "回"
        mstrLabels(6) = mstrLabels(6) & vbCrLf    '改行
        mstrLabels(6) = mstrLabels(6) & "差し回数：" & Format$(val(.Kyakusitu(2)), "@@@") & "回"
        mstrLabels(6) = mstrLabels(6) & vbCrLf    '改行
        mstrLabels(6) = mstrLabels(6) & "追込回数：" & Format$(val(.Kyakusitu(3)), "@@@") & "回"
        
        ' リンクデータオブジェクトを生成する
        For i = 0 To 2
            Set mstrLinkLabels(i) = New clsGridItem
        Next i
        
        ' リンク用文字列生成
        mstrLinkLabels(0).Text = Trim$(mCC.TZSZ4(.TozaiCD) & " " & .ChokyosiRyakusyo)
        mstrLinkLabels(0).Link = "CH"
        mstrLinkLabels(0).key = ""
        mstrLinkLabels(1).Text = Trim$(.BanusiName)
        mstrLinkLabels(1).Link = "BN"
        mstrLinkLabels(1).key = ""
        mstrLinkLabels(2).Text = Trim$(.BreederName)
        mstrLinkLabels(2).Link = "BR"
        mstrLinkLabels(2).key = ""

    End With
    
    RaiseEvent FetchedHeader

End Sub


'
'   機能: 競走馬ヘッダのリンク用文字列生成
'
'   備考: なし
'
Private Sub Set_HeaderLink()
    ' リンク用文字列生成
    With mbuf_UM
        mstrLinkLabels(0).key = DataExistCheck("CH", .ChokyosiCode)
        mstrLinkLabels(1).key = DataExistCheck("BN", .BanusiCode)
        mstrLinkLabels(2).key = DataExistCheck("BR", .BreederCode)
    End With
    
    RaiseEvent FetchedLinkLabels

End Sub


'
'   機能: データの有無をチェック
'
'   備考: なし
'
Private Function DataExistCheck(Link As String, key As String) As String
On Error GoTo ErrorHandler
    Dim rs As ADODB.Recordset
    Dim strSQL As String
    
    Set rs = New ADODB.Recordset
    
    ' レコードセットを取得する
    Select Case Link
    Case "CH"
        strSQL = "SELECT [ChokyosiCode] FROM CHOKYO WHERE [ChokyosiCode]='" & key & "'"
        rs.Open strSQL, mCN_CHOKYO, adOpenKeyset, adLockReadOnly, adCmdText
    Case "BN"
        strSQL = "SELECT [BanusiCode] FROM BANUSI WHERE [BanusiCode]='" & key & "'"
        rs.Open strSQL, mCN_BANUSI, adOpenKeyset, adLockReadOnly, adCmdText
    Case "BR"
        strSQL = "SELECT [BreederCode] FROM SEISAN WHERE [BreederCode]='" & key & "'"
        rs.Open strSQL, mCN_SEISAN, adOpenKeyset, adLockReadOnly, adCmdText
    End Select

    ' データが無かったら""を返して終了
    If rs.EOF Or rs.BOF Then
        rs.Close
        Set rs = Nothing
        DataExistCheck = ""
        Exit Function
    End If
    
    DataExistCheck = key
    
    rs.Close
    
    Set rs = Nothing
    
    Exit Function
ErrorHandler:
    gApp.ErrLog
End Function


'
'   機能: 条件別成績グリッドの作成
'
'   備考: なし
'
Private Sub Set_JokenBetu()
    Dim i As Long
    Dim j As Long
    Dim lngCP        As Long            '' セル入力用
    Dim lngRP        As Long            '' セル入力用
    Dim gd2 As clsGridData
    Dim gd5 As clsGridData
    Dim gd6 As clsGridData
    
    Set gd2 = New clsGridData
    Set gd5 = New clsGridData
    Set gd6 = New clsGridData
    
    ' 条件別成績　タブ
    '馬場別
    With gd2
        'ヘッダ
        .Rows = 7
        .Cols = 12
        lngCP = 0
        lngRP = 0
        .SetItemMatrix lngRP, lngCP, " "
        .SetItemMatrix lngRP, lngCP, " "
        .SetItemMatrix lngRP, lngCP, "1着"
        .SetItemMatrix lngRP, lngCP, "2着"
        .SetItemMatrix lngRP, lngCP, "3着"
        .SetItemMatrix lngRP, lngCP, "4着"
        .SetItemMatrix lngRP, lngCP, "5着"
        .SetItemMatrix lngRP, lngCP, "着外"
        .SetItemMatrix lngRP, lngCP, "合計"
        .SetItemMatrix lngRP, lngCP, "勝率"
        .SetItemMatrix lngRP, lngCP, "連対率"
        .SetItemMatrix lngRP, lngCP, "3着内率"
        For i = 0 To .Rows - 2
            .SetItemMatrix i + 1, 0, "馬" & vbCr & "場" & vbCr & "別"
        Next i
        .SetItemMatrix 1, 1, "　芝直"
        .SetItemMatrix 2, 1, "　芝右"
        .SetItemMatrix 3, 1, "　芝左"
        .SetItemMatrix 4, 1, "　ダ右"
        .SetItemMatrix 5, 1, "　ダ左"
        .SetItemMatrix 6, 1, "　障害"
        For i = 0 To .Rows - 2
            lngCP = 2
            For j = 0 To 9
                .SetItemMatrix i + 1, lngCP, mSC.Chakukaisu2(mbuf_UM.ChakuKaisuBa(IIf(i < 3, i, i + 1)).Chakukaisu, j)
            Next j
        Next i
    End With
    '馬場状態
    With gd5
        'ヘッダ
        .Rows = 13
        .Cols = 12
        lngCP = 0
        lngRP = 0
        .SetItemMatrix lngRP, lngCP, " "
        .SetItemMatrix lngRP, lngCP, " "
        .SetItemMatrix lngRP, lngCP, "1着"
        .SetItemMatrix lngRP, lngCP, "2着"
        .SetItemMatrix lngRP, lngCP, "3着"
        .SetItemMatrix lngRP, lngCP, "4着"
        .SetItemMatrix lngRP, lngCP, "5着"
        .SetItemMatrix lngRP, lngCP, "着外"
        .SetItemMatrix lngRP, lngCP, "合計"
        .SetItemMatrix lngRP, lngCP, "勝率"
        .SetItemMatrix lngRP, lngCP, "連対率"
        .SetItemMatrix lngRP, lngCP, "3着内率"
        For i = 0 To .Rows - 2
            .SetItemMatrix i + 1, 0, "馬" & vbCr & "場" & vbCr & "状" & vbCr & "態"
        Next i
        .SetItemMatrix 1, 1, "芝良"
        .SetItemMatrix 2, 1, "芝稍"
        .SetItemMatrix 3, 1, "芝重"
        .SetItemMatrix 4, 1, "芝不"
        .SetItemMatrix 5, 1, "ダ良"
        .SetItemMatrix 6, 1, "ダ稍"
        .SetItemMatrix 7, 1, "ダ重"
        .SetItemMatrix 8, 1, "ダ不"
        .SetItemMatrix 9, 1, "障良"
        .SetItemMatrix 10, 1, "障稍"
        .SetItemMatrix 11, 1, "障重"
        .SetItemMatrix 12, 1, "障不"
        For i = 0 To .Rows - 2
            lngCP = 2
            For j = 0 To 9
                .SetItemMatrix i + 1, lngCP, mSC.Chakukaisu2(mbuf_UM.ChakuKaisuJyotai(i).Chakukaisu, j)
            Next j
        Next i
    End With
    '距離別
    With gd6
        'ヘッダ
        .Rows = 7
        .Cols = 12
        lngCP = 0
        lngRP = 0
        .SetItemMatrix lngRP, lngCP, " "
        .SetItemMatrix lngRP, lngCP, " "
        .SetItemMatrix lngRP, lngCP, "1着"
        .SetItemMatrix lngRP, lngCP, "2着"
        .SetItemMatrix lngRP, lngCP, "3着"
        .SetItemMatrix lngRP, lngCP, "4着"
        .SetItemMatrix lngRP, lngCP, "5着"
        .SetItemMatrix lngRP, lngCP, "着外"
        .SetItemMatrix lngRP, lngCP, "合計"
        .SetItemMatrix lngRP, lngCP, "勝率"
        .SetItemMatrix lngRP, lngCP, "連対率"
        .SetItemMatrix lngRP, lngCP, "3着内率"
        For i = 0 To .Rows - 2
            .SetItemMatrix i + 1, 0, "距" & vbCr & "離" & vbCr & "別"
        Next i
        .SetItemMatrix 1, 1, "芝16下"
        .SetItemMatrix 2, 1, "芝22下"
        .SetItemMatrix 3, 1, "芝22超"
        .SetItemMatrix 4, 1, "ダ16下"
        .SetItemMatrix 5, 1, "ダ22下"
        .SetItemMatrix 6, 1, "ダ22超"
        For i = 0 To .Rows - 2
            lngCP = 2
            For j = 0 To 9
                .SetItemMatrix i + 1, lngCP, mSC.Chakukaisu2(mbuf_UM.ChakuKaisuKyori(i).Chakukaisu, j)
            Next j
        Next i
    End With

    RaiseEvent FetchedJokenBetu(gd2, gd5, gd6)

End Sub


'
'   機能: 血統グリッドの作成
'
'   備考: なし
'
Private Sub Set_Ketto()
On Error GoTo ErrorHandler
    Dim i As Long
    Dim j As Long
    Dim lngCP           As Long            '' セル入力用
    Dim lngRP           As Long            '' セル入力用
    Dim strHNname(14)   As String
    Dim gd              As clsGridData
    Dim lngIndex        As Long
    Dim strToolTip      As String
    Dim lngBGColor      As Long
    Dim HN              As ADODB.Recordset
    
    
    Set gd = New clsGridData
    
    Set mRS_HANSYOKU = New ADODB.Recordset
    Call OpenTableDirect(mRS_HANSYOKU, mCN_HANSYOKU, "HANSYOKU")
    
    Set HN = mRS_HANSYOKU
    
    With mbuf_UM
        For i = 0 To 13
            strHNname(i) = seek_HN(mRS_HANSYOKU, .Ketto3Info(i).HansyokuNum)
            DoEvents
            If mblnCancelFetching Then
                Exit Sub
            End If
        Next i
    End With
    
    ' 血統　タブ
    With gd
        .Rows = 9
        .Cols = 4
        ' カラムヘッダ登録
        lngCP = 1
        .SetItemMatrix 0, lngCP, "親(父/母)"
        .SetItemMatrix 0, lngCP, "二代"
        .SetItemMatrix 0, lngCP, "三代"
    End With
    With mbuf_UM
        ' カラム1
        For i = 0 To 3
            For j = 0 To 1

                lngIndex = i + j * 4 + 1
                
                SafeSeek HN, Array("HansyokuNum"), Array(.Ketto3Info(j).HansyokuNum)
                
                strToolTip = ""

                lngBGColor = IIf(j Mod 2 = 0, ColorFather, ColorMother)
                If Trim$(.Ketto3Info(j).BAMEI) = "" And HN.EOF Then
                ' 名前も繁殖馬名も無い場合
                    gd.SetItemMatrix lngIndex, 1, String(j, "."), strToolTip, , , , lngBGColor, lngBGColor
                Else
                ' 繁殖馬名がある場合
                    gd.SetItemMatrix lngIndex, 1, KettoGetName(HN, j + 0), strToolTip, , _
                                        "HN", IfExist(HN, "HansyokuNum"), lngBGColor
                End If
            Next j
        Next i
        ' カラム2
        For i = 0 To 1
            For j = 0 To 3
                lngIndex = i + j * 2 + 1
                
                SafeSeek HN, Array("HansyokuNum"), Array(.Ketto3Info(j + 2).HansyokuNum)
                
                strToolTip = ""
                
                lngBGColor = IIf(j Mod 2 = 0, ColorFather, ColorMother)
                If Trim$(.Ketto3Info(j + 2).BAMEI) = "" And HN.EOF Then
                ' 名前も繁殖馬名も無い場合
                    gd.SetItemMatrix lngIndex, 2, String(j, "."), strToolTip, , , , lngBGColor, lngBGColor
                Else
                ' 繁殖馬名がある場合
                    gd.SetItemMatrix lngIndex, 2, KettoGetName(HN, j + 2), strToolTip, , _
                                        "HN", IfExist(HN, "HansyokuNum"), lngBGColor
                End If
            Next j
        Next i
        ' カラム3
        For i = 0 To 7
            
            SafeSeek HN, Array("HansyokuNum"), Array(.Ketto3Info(i + 6).HansyokuNum)
            
            strToolTip = ""
            lngBGColor = IIf(i Mod 2 = 0, ColorFather, ColorMother)
            If Trim$(.Ketto3Info(i + 6).BAMEI) = "" And HN.EOF Then
            ' 名前も繁殖馬名も無い場合
                gd.SetItemMatrix i + 1, 3, String(i, "."), strToolTip, , , , lngBGColor, lngBGColor
            Else
            ' 繁殖馬名がある場合
                gd.SetItemMatrix i + 1, 3, KettoGetName(HN, i + 6), strToolTip, , _
                                    "HN", IfExist(HN, "HansyokuNum"), lngBGColor
            End If
        Next i
    End With
    
    mRS_HANSYOKU.Close

    RaiseEvent FetchedKetto(gd)
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub

'
'   機能: 繁殖馬に該当レコードが存在するか返す
'
'   備考: なし
'
Private Function seek_HN(rs As ADODB.Recordset, ByVal key As String) As String
On Error GoTo ErrorHandler
    Call SafeSeek(rs, Array("HansyokuNum"), Array(key))
    
    ' データが無かったら""を返して終了
    If rs.EOF Or rs.BOF Then
        seek_HN = ""
        Exit Function
    End If
    
    seek_HN = Trim$(rs("Bamei"))
        
    Exit Function
ErrorHandler:
    gApp.ErrLog
End Function


'
'   機能: 競走馬名を得る
'
'   備考: なし
'
Private Function KettoGetName(HN As ADODB.Recordset, FieldNum As String) As String
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
        str = Trim$(mbuf_UM.Ketto3Info(FieldNum).BAMEI)
        If str <> "" Then
            KettoGetName = str
        Else
            KettoGetName = Space(CLng(FieldNum))
        End If
    End If
End Function


'
'   機能: 坂路調教グリッドの作成
'
'   備考: なし
'
Private Sub Set_Hanro(key As String)
On Error GoTo ErrorHandler
    Const rowOffset As Long = 2 '' データ行オフセット
    
    Dim rs As ADODB.Recordset
    Dim strSQL As String
    Dim i As Long
    Dim j As Long
    Dim lngCP        As Long            '' セル入力用
    Dim lngRP        As Long            '' セル入力用
    Dim gd As clsGridData

    DoEvents
    If mblnCancelFetching Then
        Exit Sub
    End If
    
    ' 坂路調教　タブ

    ' 坂路調教取得
    ' レコードセットを取得する
    ' HANRO (坂路調教)
    strSQL = "SELECT * FROM HANRO WHERE [KettoNum]='" & key & "'"

    Set mRS_HANRO = New ADODB.Recordset
    mRS_HANRO.Open strSQL, mCN_HANRO, adOpenKeyset, adLockReadOnly, adCmdText
    
    Set rs = mRS_HANRO
    
    ' データが無かったらFalseを返して終了
    If rs.EOF Or rs.BOF Then
        rs.Close
        Set rs = Nothing
        RaiseEvent NoTabData(4)
        Exit Sub
    End If

    ' グリッドデータオブジェクトを生成する
    Set gd = New clsGridData

    With gd
        ' グリッドデータにロウ数、カラム数を設定
        .Rows = 10
        .Cols = 10
        
        lngCP = 0
        .SetItemMatrix 0, lngCP, "調教年月日曜日", , ">_"
        .SetItemMatrix 0, lngCP, "調教時刻", , ">_"
        .SetItemMatrix 0, lngCP, "ﾄﾚｾﾝ", "トレーニングセンター", "^_"
        .SetItemMatrix 0, lngCP, "４ハロン", "800m", "^_"
        .SetItemMatrix 0, lngCP, "４ハロン", "800m", "^_"
        .SetItemMatrix 0, lngCP, "３ハロン", "600m", "^_"
        .SetItemMatrix 0, lngCP, "３ハロン", "600m", "^_"
        .SetItemMatrix 0, lngCP, "２ハロン", "400m", "^_"
        .SetItemMatrix 0, lngCP, "２ハロン", "400m", "^_"
        .SetItemMatrix 0, lngCP, "１ハロン", "200m", "^_"
        lngCP = 0
        .SetItemMatrix 1, lngCP, "調教年月日曜日", , ">_"
        .SetItemMatrix 1, lngCP, "調教時刻", , ">_"
        .SetItemMatrix 1, lngCP, "ﾄﾚｾﾝ", "トレーニングセンター", "^_"
        .SetItemMatrix 1, lngCP, "合計", "0m～800m", ">_"
        .SetItemMatrix 1, lngCP, "ﾗｯﾌﾟﾀｲﾑ", "600m～800m", ">_"
        .SetItemMatrix 1, lngCP, "合計", "0m～600m", ">_"
        .SetItemMatrix 1, lngCP, "ﾗｯﾌﾟﾀｲﾑ", "600m～600m", ">_"
        .SetItemMatrix 1, lngCP, "合計", "0m～400m", ">_"
        .SetItemMatrix 1, lngCP, "ﾗｯﾌﾟﾀｲﾑ", "600m～400m", ">_"
        .SetItemMatrix 1, lngCP, "ﾗｯﾌﾟﾀｲﾑ", "0m～200m", ">_"
    End With

    rs.MoveFirst
    i = 0
    Do While Not rs.EOF
        DoEvents
        If mblnCancelFetching Then
            Exit Sub
        End If
        With gd
            lngRP = i + rowOffset
            lngCP = 0
            .SetItemMatrix lngRP, lngCP, mSC.YMD4(rs("ChokyoDate"))  '調教年月日曜日
            .SetItemMatrix lngRP, lngCP, mSC.HHNN1(rs("ChokyoTime"))  '調教時間
            .SetItemMatrix lngRP, lngCP, mSC.HC_TKBN(rs("TresenKubun"))  'ﾄﾚｾﾝ区分
            .SetItemMatrix lngRP, lngCP, mSC.SSSs(rs("HaronTime4"))  ' 4h 合計
            .SetItemMatrix lngRP, lngCP, mSC.SSSs(rs("LapTime4"))  ' 4h ﾗｯﾌﾟﾀｲﾑ
            .SetItemMatrix lngRP, lngCP, mSC.SSSs(rs("HaronTime3"))  ' 3h 合計
            .SetItemMatrix lngRP, lngCP, mSC.SSSs(rs("LapTime3"))  ' 3h ﾗｯﾌﾟﾀｲﾑ
            .SetItemMatrix lngRP, lngCP, mSC.SSSs(rs("HaronTime2"))  ' 2h 合計
            .SetItemMatrix lngRP, lngCP, mSC.SSSs(rs("LapTime2"))  ' 2h ﾗｯﾌﾟﾀｲﾑ
            .SetItemMatrix lngRP, lngCP, mSC.SSSs(rs("LapTime1"))  ' 1h ﾗｯﾌﾟﾀｲﾑ
        End With
        i = i + 1
        If gd.Rows = i + rowOffset Then
            gd.Rows = gd.Rows + 10
        End If
        rs.MoveNext
    Loop
    gd.Rows = i + rowOffset

    rs.Close

    Set rs = Nothing

    RaiseEvent FetchedHanro(gd)
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub

'
'   機能: レコードセットを開く
'
'   備考: なし
'
Private Function RS_open() As Boolean
On Error GoTo ErrorHandler
    
    'RACEを開く
    Set mRS_RACE = New ADODB.Recordset
    Call OpenTableDirect(mRS_RACE, mCN_RACE, "RACE")
    gApp.Log "UM mRS_RACE Complete Opened"
        
    'UMA_RACE_Bを開く
    Set mRS_UMA_RACE_B = New ADODB.Recordset
    Call OpenTableDirect(mRS_UMA_RACE_B, mCN_UMA_RACE_B, "UMA_RACE_B")
    gApp.Log "UM mRS_RACE_B Opened"
    
    'KISHUを開く
    Set mRS_KISHU = New ADODB.Recordset
    Call OpenTableDirect(mRS_KISHU, mCN_KISHU, "KISHU")
    gApp.Log "UM KISHU Opened"
    
    'UMAを開く
    Set mRS_UMA = New ADODB.Recordset
    Call OpenTableDirect(mRS_UMA, mCN_UMA, "UMA")
    gApp.Log "UM UMA Opened"
    
    Exit Function
ErrorHandler:
    gApp.ErrLog
End Function


'
'   機能: 過去走グリッドを作成
'
'   備考: なし
'
Private Sub Set_Kako(key As String)
On Error GoTo ErrorHandler
    Const rowOffset As Long = 1 '' データ行オフセット
    
    Dim buf_SE As JV_SE_RACE_UMA
    Dim i As Long
    Dim j As Long
    Dim lngCP        As Long            '' セル入力用
    Dim lngRP        As Long            '' セル入力用
    Dim strTmp As String
    Dim gd As clsGridData
    
    Dim RA As ADODB.Recordset
    Dim UA As ADODB.Recordset
    Dim UB As ADODB.Recordset
    Dim KS As ADODB.Recordset
    Dim UM As ADODB.Recordset
        
    Set RA = mRS_RACE
    Set UA = mRS_UMA_RACE_A
    Set UB = mRS_UMA_RACE_B
    Set KS = mRS_KISHU
    Set UM = mRS_UMA
    
    ' 過去成績　タブ
    ' グリッドデータオブジェクトを生成する
    Set gd = New clsGridData
    With gd
        
        ' グリッドデータにロウ数、カラム数を設定
        .Rows = 10 + 1
        .Cols = 27
        lngCP = 0
        .SetItemMatrix 0, lngCP, "年月日"
        .SetItemMatrix 0, lngCP, "回場日"
        .SetItemMatrix 0, lngCP, "レース名"
        .SetItemMatrix 0, lngCP, "条件"
        .SetItemMatrix 0, lngCP, "コース"
        .SetItemMatrix 0, lngCP, "馬場"
        .SetItemMatrix 0, lngCP, "習"
        .SetItemMatrix 0, lngCP, "騎手"
        .SetItemMatrix 0, lngCP, "負担"
        .SetItemMatrix 0, lngCP, "B"
        .SetItemMatrix 0, lngCP, "頭数"
        .SetItemMatrix 0, lngCP, "枠番"
        .SetItemMatrix 0, lngCP, "馬番"
        .SetItemMatrix 0, lngCP, "異常"
        .SetItemMatrix 0, lngCP, "着順"
        .SetItemMatrix 0, lngCP, "タイム"
        .SetItemMatrix 0, lngCP, "コーナー通過順"
        .SetItemMatrix 0, lngCP, "単オッズ"
        .SetItemMatrix 0, lngCP, "単人気"
        .SetItemMatrix 0, lngCP, "馬体重"
        .SetItemMatrix 0, lngCP, "増減差"
        .SetItemMatrix 0, lngCP, "獲得本賞金"
        .SetItemMatrix 0, lngCP, "獲得付加賞金"
        .SetItemMatrix 0, lngCP, "後3ハロン"
        .SetItemMatrix 0, lngCP, "1(2)着馬"
        .SetItemMatrix 0, lngCP, "タイム差"
        'col=27 "年月日"ソート用隠しカラム
    End With
    
    UA.MoveFirst
    i = 0
    Do While Not UA.EOF
        DoEvents
        If mblnCancelFetching Then
            Exit Sub
        End If
        Call SafeSeek(UB, Array("B_Year", "B_MonthDay", "B_JyoCD", "B_RaceNum", "B_Umaban", "B_KettoNum"), _
                    Array(UA("Year").value, UA("MonthDay").value, UA("JyoCD").value, UA("RaceNum").value, UA("Umaban").value, UA("KettoNum").value))
        Call SafeSeek(RA, Array("Year", "MonthDay", "JyoCD", "Kaiji", "Nichiji", "RaceNum"), _
                    Array(UA("Year").value, UA("MonthDay").value, UA("JyoCD").value, UA("Kaiji").value, UA("Nichiji").value, UA("RaceNum").value))
        ' 構造体に読み込む
        Call SetDataFromRS_SE_A(UA, buf_SE)
        Call SetDataFromRS_SE_B(UB, buf_SE)
        lngCP = 0
        With buf_SE
            lngCP = 0
            strTmp = UA("Year").value & UA("MonthDay").value & UA("JyoCD").value & UA("Kaiji").value & UA("Nichiji").value & UA("RaceNum").value
            With .id
                gd.SetItemMatrix i + rowOffset, lngCP, mSC.YMD3(.Year & .MonthDay), , , "RA", strTmp
                gd.SetItemMatrix i + rowOffset, lngCP, mSC.FitSpaceNum(.Kaiji, 2) & mCC.KIBJ3(.JyoCD) & mSC.FitSpaceNum(.Nichiji, 2), , , "RA", strTmp
            End With
            gd.SetItemMatrix i + rowOffset, lngCP, Trim$(RA("Ryakusyo6")) & mCC.GRAD2(RA("GradeCD")), Trim$(RA("Hondai")) & mCC.GRAD3(RA("GradeCD")), "<-", "RA"
            gd.SetItemMatrix i + rowOffset, lngCP, mCC.KSSB6(RA("SyubetuCD")) & mCC.KSJK4(RA("JyokenCD5")), "", "<-"
            gd.SetItemMatrix i + rowOffset, lngCP, mCC.TRCK5(RA("TrackCD")) & IIf(val(RA("Kyori")) = 0, "", val(RA("Kyori"))), "", "<-"
            gd.SetItemMatrix i + rowOffset, lngCP, mSC.BabaGrid(RA), , "<^"
            gd.SetItemMatrix i + rowOffset, lngCP, mCC.KSMN1(.MinaraiCD), mCC.KSMN3(.MinaraiCD)
            Call SafeSeek(KS, Array("KisyuCode"), Array(.KisyuCode))
            If val(.KisyuCodeBefore) = 0 Then
                gd.SetItemMatrix i + rowOffset, lngCP, Trim$(.KisyuRyakusyo), "", "<-", "KS", IIf(KS.EOF, "", .KisyuCode)
            Else
                gd.SetItemMatrix i + rowOffset, lngCP, Trim$(.KisyuRyakusyo), Trim$(.KisyuRyakusyoBefore) & "→" & Trim$(.KisyuRyakusyo) _
                    & "", "<-", "KS", IIf(KS.EOF, "", .KisyuCode), , &HFF
            End If
            gd.SetItemMatrix i + rowOffset, lngCP, IIf(val(.Futan) = 0, "", Format$(.Futan / 10, "0.0")), , ">-"
            gd.SetItemMatrix i + rowOffset, lngCP, IIf(.BLINKER = "0", "", "B")
            gd.SetItemMatrix i + rowOffset, lngCP, IIf(val(RA("SyussoTosu")) = 0, "", val(RA("SyussoTosu"))), , ">-"
            gd.SetItemMatrix i + rowOffset, lngCP, IIf(val(.Wakuban) = 0, "", .Wakuban), , ">-"
            gd.SetItemMatrix i + rowOffset, lngCP, IIf(val(.Umaban) = 0, "", val(.Umaban)), , ">-"
            gd.SetItemMatrix i + rowOffset, lngCP, mCC.IJKB2(.IJyoCD), mCC.IJKB1(.IJyoCD)
            
            gd.SetItemMatrix i + rowOffset, lngCP, IIf(val(.KakuteiJyuni) = 0, "", val(.KakuteiJyuni)), , ">-", , , gApp.GetChakujyunColor(CLng(.KakuteiJyuni))
            gd.SetItemMatrix i + rowOffset, lngCP, IIf(val(.TIME) = 0, "", Format$(.TIME, "@:@@.@")), , ">-"
            gd.SetItemMatrix i + rowOffset, lngCP, Right$(Space(2) & Format$(.Jyuni1c, "##"), 2) _
                                                    & IIf(val(.Jyuni1c) = 0 Or val(.Jyuni2c) = 0, " ", "-") _
                                                    & Right$(Space(2) & Format$(.Jyuni2c, "##"), 2) _
                                                    & IIf(val(.Jyuni2c) = 0 Or val(.Jyuni3c) = 0, " ", "-") _
                                                    & Right$(Space(2) & Format$(.Jyuni3c, "##"), 2) _
                                                    & IIf(val(.Jyuni3c) = 0 Or val(.Jyuni4c) = 0, " ", "-") _
                                                    & Right$(Space(2) & Format$(.Jyuni4c, "##"), 2) _
                                                    , "", ">-"
            If val(.Odds) = 0 Then
                strTmp = ""
            Else
                strTmp = Format$(.Odds / 10, "0.0")
            End If
            gd.SetItemMatrix i + rowOffset, lngCP, strTmp, , ">-"
            gd.SetItemMatrix i + rowOffset, lngCP, IIf(val(.Ninki) = 0, "", val(.Ninki)), , ">-"
            If val(UB("BaTaijyu")) < 2 Then
                gd.SetItemMatrix i + rowOffset, lngCP, "     ", , ">-"
            ElseIf UB("BaTaijyu") = "999" Then
                gd.SetItemMatrix i + rowOffset, lngCP, " --- ", , ">-"
            Else
                gd.SetItemMatrix i + rowOffset, lngCP, val(UB("BaTaijyu")) & "kg", , ">-"
            End If
            If .ZogenFugo = " " Then
                If .ZogenSa = "   " Or .ZogenSa = "999" Then
                    gd.SetItemMatrix i + rowOffset, lngCP, "", "", ">-"
                Else
                    gd.SetItemMatrix i + rowOffset, lngCP, "±" & val(.ZogenSa), "", ">-"
                End If
            Else
                gd.SetItemMatrix i + rowOffset, lngCP, IIf(.ZogenSa = "999", "", .ZogenFugo & val(.ZogenSa)), "", ">-"
            End If
            gd.SetItemMatrix i + rowOffset, lngCP, IIf(val(.Honsyokin) = 0, "", mSC.Money3(.Honsyokin)), , ">-"
            gd.SetItemMatrix i + rowOffset, lngCP, IIf(val(.Fukasyokin) = 0, "", mSC.Money3(.Fukasyokin)), , ">-"
            
            gd.SetItemMatrix i + rowOffset, lngCP, IIf(.HaronTimeL3 = "999", "", mSC.SSSs(.HaronTimeL3)), , ">-"
            With .ChakuUmaInfo(0) '
                Call SafeSeek(UM, Array("KettoNum"), Array(.KettoNum))
                If buf_SE.KakuteiJyuni = "01" Then
                    strTmp = "(" & Trim$(.BAMEI) & ")"
                Else
                    strTmp = Trim$(.BAMEI)
                End If
                If UM.EOF Then
                    gd.SetItemMatrix i + rowOffset, lngCP, strTmp, "", "<-"
                Else
                    gd.SetItemMatrix i + rowOffset, lngCP, strTmp, "", "<-", "UM", .KettoNum
                End If
            End With
            If .TimeDiff = "9999" Or .TimeDiff = "0000" Then
                gd.SetItemMatrix i + rowOffset, lngCP, "", , ">-"
            Else
                gd.SetItemMatrix i + rowOffset, lngCP, Mid$(.TimeDiff, 1, 1) _
                    & Format$(val(Mid$(.TimeDiff, 2, 3)) / 10, "#0.0"), , ">-"
            End If
            gd.SetItemMatrix i + rowOffset, lngCP, mSC.YMD2(.id.Year & .id.MonthDay)
        
            i = i + 1
            If gd.Rows = i + rowOffset Then
                gd.Rows = gd.Rows + 10
            End If
        End With
        UA.MoveNext
    Loop
    
    gd.Rows = i + rowOffset
    
    RaiseEvent FetchedKako(gd)
        
    Exit Sub
ErrorHandler:
    If Err.Number = 3265 Then
        Resume Next
    End If
    gApp.ErrLog
    Resume Next
End Sub
    

'
'   機能: 持ちタイムグリッドを作成
'
'   備考: なし
'
Private Sub Set_Time(key As String)
On Error GoTo ErrorHandler
    
    Const rowOffset As Long = 1 '' データ行オフセット
    
    Dim strTmp As String
    Dim lngRowNum As Long
    
    Dim i As Long
    Dim lngCP        As Long            '' セル入力用
    Dim lngRP        As Long            '' セル入力用
    Dim BGColor As Long
    Dim gd As clsGridData
    
    Dim RA As ADODB.Recordset
    Dim UA As ADODB.Recordset
    Dim UB As ADODB.Recordset
        
    Set RA = mRS_RACE
    Set UA = mRS_UMA_RACE_A
    Set UB = mRS_UMA_RACE_B
    
    Dim strHiddenCol As String
    
    ' 持ちタイム　タブ
                
    ' グリッドデータオブジェクトを生成する
    Set gd = New clsGridData
            
    With gd
        .Cols = 12
        .Rows = 10 + rowOffset
        
        ' カラムヘッダ登録
        lngCP = 0
        .SetItemMatrix 0, lngCP, "距離", , "<-"
        
        .SetItemMatrix 0, lngCP, "タイム", , ">-"
        .SetItemMatrix 0, lngCP, "後3ハロン", , ">-"
        .SetItemMatrix 0, lngCP, "年月日", , "<-"
        .SetItemMatrix 0, lngCP, "場回日", , "<-"
        .SetItemMatrix 0, lngCP, "着順", , ">-"
        .SetItemMatrix 0, lngCP, "異常", , "<-"
        .SetItemMatrix 0, lngCP, "馬場状態", , "<-"
        .SetItemMatrix 0, lngCP, "負担重量", , ">-"
        .SetItemMatrix 0, lngCP, "コース", , "<-"
    End With
    
    UA.MoveFirst
    i = 0
    Do While Not UA.EOF
        DoEvents
        If mblnCancelFetching Then
            Exit Sub
        End If
        Call SafeSeek(UB, Array("B_Year", "B_MonthDay", "B_JyoCD", "B_RaceNum", "B_Umaban", "B_KettoNum"), _
                    Array(UA("Year").value, UA("MonthDay").value, UA("JyoCD").value, UA("RaceNum").value, UA("Umaban").value, UA("KettoNum").value))
        Call SafeSeek(RA, Array("Year", "MonthDay", "JyoCD", "Kaiji", "Nichiji", "RaceNum"), _
                    Array(UA("Year").value, UA("MonthDay").value, UA("JyoCD").value, UA("Kaiji").value, UA("Nichiji").value, UA("RaceNum").value))
        If UB("Time") <> 0 Then
            If mCC.TRCK4(RA("TrackCD")) = "芝" Then
                BGColor = &HC0FFC0 '緑
                strHiddenCol = "1"
            ElseIf mCC.TRCK4(RA("TrackCD")) = "ダ" Then
                BGColor = RGB(251, 243, 206) '薄い黄土色1
                strHiddenCol = "2"
            ElseIf mCC.TRCK4(RA("TrackCD")) = "砂" Then
                BGColor = &HE0E0E0 '灰
                strHiddenCol = "3"
            ElseIf mCC.TRCK4(RA("TrackCD")) = "障" Then
                BGColor = &HFFFFFF '白
                strHiddenCol = "4"
            End If
            With gd
                lngCP = 0
                .SetItemMatrix i + rowOffset, lngCP, Trim$(RA("Kyori")), , , , , BGColor
                .SetItemMatrix i + rowOffset, lngCP, Format$(UB("Time"), "@:@@.@"), , ">^", , , BGColor
                .SetItemMatrix i + rowOffset, lngCP, IIf(UB("HaronTimeL3") = "999", "", mSC.SSSs(UB("HaronTimeL3"))), , ">^", , , BGColor
                strTmp = UA("Year").value & UA("MonthDay").value & UA("JyoCD").value & UA("Kaiji").value & UA("Nichiji").value & UA("RaceNum").value
                .SetItemMatrix i + rowOffset, lngCP, mSC.YMD4(UA("Year") & UA("MonthDay")), , ">^", "RA", strTmp, BGColor
                .SetItemMatrix i + rowOffset, lngCP, mCC.KIBJ3(UA("JyoCD")) & mSC.KN1(UA("Kaiji") & UA("Nichiji")), , ">^", "RA", strTmp, BGColor
                .SetItemMatrix i + rowOffset, lngCP, mSC.ValStr(UB("KakuteiJyuni")), , ">^", , , IIf(UB("KakuteiJyuni") <= 3, gApp.GetChakujyunColor(UB("KakuteiJyuni")), BGColor)
                .SetItemMatrix i + rowOffset, lngCP, mCC.IJKB2(UB("IJyoCD")), mCC.IJKB1(UB("IJyoCD")), "<^", , , BGColor
                .SetItemMatrix i + rowOffset, lngCP, mSC.BabaGrid(RA), , "<^", , , BGColor
                .SetItemMatrix i + rowOffset, lngCP, Format$(UB("Futan") / 10, "#0.0"), , ">^", , , BGColor
                .SetItemMatrix i + rowOffset, lngCP, mCC.TRCK5(RA("TrackCD")), mCC.TRCK1(RA("TrackCD")), ">^", , , BGColor
                .SetItemMatrix i + rowOffset, lngCP, mCC.TRCK4(RA("TrackCD")), , ">^", , , BGColor
                
                ' ソート用に隠しカラムを作成
                Dim tmpVal As String
                tmpVal = Trim$(RA("Kyori"))
                If Len(tmpVal) < 4 Then
                    strHiddenCol = strHiddenCol & "0" & tmpVal
                Else
                    strHiddenCol = strHiddenCol & tmpVal
                End If
                .SetItemMatrix i + rowOffset, lngCP, strHiddenCol & UB("Time") & UB("HaronTimeL3"), , ">^", , , BGColor
                
                i = i + 1
                If gd.Rows = i + rowOffset Then
                    gd.Rows = gd.Rows + 10
                End If
            End With
        End If
        UA.MoveNext
    Loop
    
    If i = 0 Then
        RaiseEvent NoTabData(3)
        Exit Sub
    End If
    gd.Rows = i + rowOffset
    
    RaiseEvent FetchedTime(gd)
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub

