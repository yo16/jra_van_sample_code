VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDataCH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "調教師 データクラス"
'
'   調教師 データクラス
'

Option Explicit

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   外部関数(イベント)
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

' イベント
Public Event NoData()
Public Event NoDataUMA()

Public Event FetchCompleteSYUYOU(gd As clsGridData)
Public Event FetchCompleteUMA(gd As clsGridData)

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部変数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

' 変換ユーティリティー
Private mCC As clsCodeConverter          '' JV-Data のコード変換メソッド群
Private mSC As clsStringConverter          '' JV-Data のコード変換メソッド群

' Viewerに提供する為のデータ格納変数
Private mstrLabels(0 To 5) As String     '' テキストデータ
Private mGridData(0 To 4) As clsGridData '' グリッドデータ
'モジュール内で使用する為のデータ格納変数
Private mBuf_CH As JV_CH_CHOKYOSI
Private mKey As clsKeyCH                '' キー

'コネクション
Private WithEvents mAsyncCN_UMA As ADODB.Connection
Attribute mAsyncCN_UMA.VB_VarHelpID = -1
Private mCN_RACE As ADODB.Connection
Attribute mCN_RACE.VB_VarHelpID = -1
'レコードセット
Private mRS_UMA As ADODB.Recordset
Private mRS_RACE As ADODB.Recordset

Private mblnCancelFetching As Boolean


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   プロパティ
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: ラベル用文字列を返す
'
'   備考: なし
'
Public Property Get Labels(Index As Integer) As String
    Labels = mstrLabels(Index)
End Property

'
'   機能: グリッドデータを保存する
'
'   備考: なし
'
Public Property Get GridDatas(Index As Integer) As clsGridData
    Set GridDatas = mGridData(Index)
End Property


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   外部関数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: データを取得する
'
'   備考: なし
'
Public Function Fetch(ByRef key As clsKeyCH) As Boolean
On Error GoTo ErrorHandler
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim strSQL As String
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim m As Long
    Dim lngCP        As Long            '' セル入力用
    Dim lngRP        As Long            '' セル入力用
    
    Set mKey = key
    
    ' グリッドデータオブジェクトを生成する
    For i = 0 To 3
        Set mGridData(i) = New clsGridData
    Next i

    'CHOKYO (調教師マスタ)
    Set cn = gApp.GetCN_CHOKYO
    strSQL = "SELECT * FROM CHOKYO " & key.SQLWHEREString
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenKeyset, adLockReadOnly, adCmdText
    
    ' データが無かったらFalseを返して終了
    If rs.EOF Or rs.BOF Then
        rs.Close
        Set rs = Nothing
        Fetch = False
        RaiseEvent NoData
        Exit Function
    End If

    Call SetDataFromRS_CH(rs, mBuf_CH)
    
    With mBuf_CH
        
        mstrLabels(0) = ""
        mstrLabels(0) = mstrLabels(0) & mSC.YMD1(rs("MakeDate")) & "作成データ"   '作成年月日
        
        ' 太字レース名等グレーバー用文字列生成
        mstrLabels(1) = ""
        mstrLabels(1) = mstrLabels(1) & Trim$(.ChokyosiName)
        mstrLabels(2) = ""
        If .TozaiCD <> 4 Then
            mstrLabels(2) = mstrLabels(2) & mSC.CTrim(.ChokyosiNameKana)
        End If
    
        ' ラベル１用文字列生成
        mstrLabels(3) = ""
        mstrLabels(3) = mstrLabels(3) & mSC.CH_DelKubun(.DelKubun)  '抹消区分
        mstrLabels(3) = mstrLabels(3) & mSC.YMD1(rs("DelDate"))  '抹消年月日
        mstrLabels(3) = mstrLabels(3) & vbCrLf  '改行
        mstrLabels(3) = mstrLabels(3) & IIf(val(.TozaiCD) = 0, "", "所属　　：" & mCC.TZSZ5(.TozaiCD)) & vbCrLf  '東西所属
        mstrLabels(3) = mstrLabels(3) & IIf(val(rs("IssueDate")) = 0, "", "初免許　：" & mSC.YMD1(rs("IssueDate"))) & vbCrLf  '免許交付
        mstrLabels(3) = mstrLabels(3) & IIf(val(rs("BirthDate")) = 0, "", "生年月日：" & mSC.YMD1(rs("BirthDate")))   '生年月日
        
        ' ラベル２用文字列生成
        mstrLabels(4) = ""
        mstrLabels(4) = mstrLabels(4) & IIf(val(.SexCD) = 0, "", "性別　　：" & mSC.CH_SexCD(.SexCD) & " ") & vbCrLf  '性別
        mstrLabels(4) = mstrLabels(4) & IIf(Trim$(.Syotai) = "", "", "招待地域：" & Trim$(.Syotai))

    End With
    
    rs.Close
    
    ' 以下Grid処理
    'CHOKYO_SEISEKI (調教師マスタ_成績)
    Set cn = gApp.GetCN_CHOKYO_SEISEKI
    strSQL = "SELECT * FROM CHOKYO_SEISEKI " & key.SQLWHEREString
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenKeyset, adLockReadOnly, adCmdText
    
    Call SetDataFromRS_CH_SEISEKI(rs, mBuf_CH)

    ' 総合成績　タブ
    With mGridData(0)
        .Rows = 10
        .Cols = 13
        lngCP = 0
        lngRP = 0
        .SetItemMatrix lngRP, lngCP, " "
        .SetItemMatrix lngRP, lngCP, "本賞金"
        .SetItemMatrix lngRP, lngCP, "付加賞金"
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
        
        .SetItemMatrix 1, 0, "本年(" & mBuf_CH.HonZenRuikei(0).SetYear & "年)"
        .SetItemMatrix 2, 0, "平地"
        .SetItemMatrix 3, 0, "障害"
        .SetItemMatrix 4, 0, "前年(" & mBuf_CH.HonZenRuikei(1).SetYear & "年)"
        .SetItemMatrix 5, 0, "平地"
        .SetItemMatrix 6, 0, "障害"
        .SetItemMatrix 7, 0, "累計　  　　"
        .SetItemMatrix 8, 0, "平地"
        .SetItemMatrix 9, 0, "障害"
    End With
    'データ
    For i = 0 To 2 '0:本年 1:前年 2:累計
        ' 平地
        With mBuf_CH.HonZenRuikei(i)
            mGridData(0).ItemMatrix(3 * i + 2, 1).Text = mSC.Money(.HonSyokinHeichi)
            mGridData(0).ItemMatrix(3 * i + 2, 2).Text = mSC.Money(.FukaSyokinHeichi)
            For j = 0 To 9
                mGridData(0).ItemMatrix(3 * i + 2, j + 3).Text = mSC.Chakukaisu2(.ChakuKaisuHeichi.Chakukaisu, j)
            Next j
        End With
        ' 障害
        With mBuf_CH.HonZenRuikei(i)
            mGridData(0).ItemMatrix(3 * i + 3, 1).Text = mSC.Money(.HonSyokinSyogai)
            mGridData(0).ItemMatrix(3 * i + 3, 2).Text = mSC.Money(.FukaSyokinSyogai)
            For j = 0 To 9
                mGridData(0).ItemMatrix(3 * i + 3, j + 3).Text = mSC.Chakukaisu2(.ChakuKaisuSyogai.Chakukaisu, j)
            Next j
        End With
    Next i
    
    
    ' 距離別成績　タブ
    With mGridData(1)
        .Rows = 10
        .Cols = 2 * 10 + 1
        
        'カラムヘッダ
        lngCP = 0
        lngRP = 0
        .SetItemMatrix lngRP, lngCP, " "
        For i = 0 To 9
            .SetItemMatrix lngRP, lngCP, "　芝　"
        Next i
        For i = 0 To 9
            .SetItemMatrix lngRP, lngCP, "ダート"
        Next i
        lngCP = 0
        lngRP = 1
        .SetItemMatrix lngRP, lngCP, " "
        For i = 0 To 1
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
        Next i
        ' ロウヘッダ
        lngRP = 1
        .SetItemMatrix lngRP + 1, 0, "本年　　"
        .SetItemMatrix lngRP + 2, 0, "～1600M"
        .SetItemMatrix lngRP + 3, 0, "～2200M"
        .SetItemMatrix lngRP + 4, 0, "2200M～"
        .SetItemMatrix lngRP + 5, 0, "前年　　"
        .SetItemMatrix lngRP + 6, 0, "～1600M"
        .SetItemMatrix lngRP + 7, 0, "～2200M"
        .SetItemMatrix lngRP + 8, 0, "2200M～"
    End With
    'データ
    For i = 0 To 1 '0:本年 1:前年
        For j = 0 To 2 '0:～1600M 1:～2200M 2:2200M～
            With mBuf_CH.HonZenRuikei(i).ChakuKaisuKyori(j) '芝
                For k = 0 To 9
                    mGridData(1).ItemMatrix(4 * i + j + 3, k + 1).Text = mSC.Chakukaisu2(.Chakukaisu, k)
                Next k
            End With
        Next j
        For j = 0 To 2
            With mBuf_CH.HonZenRuikei(i).ChakuKaisuKyori(j + 3) 'ダート
                For k = 0 To 9
                    mGridData(1).ItemMatrix(4 * i + j + 3, k + 11).Text = mSC.Chakukaisu2(.Chakukaisu, k)
                Next k
            End With
        Next j
    Next i
    
    ' 競馬場別成績　タブ
    With mGridData(2)
        .Rows = 13
        .Cols = 4 * 10 + 1
        'カラムヘッダ
        lngCP = 0
        lngRP = 0
        .SetItemMatrix lngRP, lngCP, " "
        For i = 0 To 19
            .SetItemMatrix lngRP, lngCP, "本年度"
        Next i
        For i = 0 To 19
            .SetItemMatrix lngRP, lngCP, "前年度"
        Next i
        lngCP = 0
        lngRP = 1
        .SetItemMatrix lngRP, lngCP, " "
        For i = 0 To 1
            For j = 0 To 9
                .SetItemMatrix lngRP, lngCP, "平地"
            Next j
            For j = 0 To 9
                .SetItemMatrix lngRP, lngCP, "障害"
            Next j
        Next i
        lngCP = 0
        lngRP = 2
        .SetItemMatrix lngRP, lngCP, " "
        For i = 0 To 3
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
        Next i
        ' ロウヘッダ
        lngRP = 3
        For i = 0 To 9
            .SetItemMatrix lngRP + i, 0, mCC.KIBJ3(Right$("0" & CStr(i + 1), 2)), mCC.KIBJ5(Right$("0" & CStr(i + 1), 2))
        Next i
    End With
    'データ
    For i = 0 To 1 '0:本年 1:前年
        For j = 0 To 1 '0:平地 1:障害
            For k = 0 To 9 '0:札幌～9:小倉
                With mBuf_CH.HonZenRuikei(i).ChakuKaisuJyo(j + 2 * k)
                    For m = 0 To 9
                        mGridData(2).ItemMatrix(k + 3, 10 * (2 * i + j) + m + 1).Text = mSC.Chakukaisu2(.Chakukaisu, m)
                    Next m
                End With
            Next k
        Next j
    Next i
    
    rs.Close
    
    Set rs = Nothing
    
    Set mAsyncCN_UMA = gApp.GetCN_UMA
    strSQL = "SELECT * FROM UMA WHERE [ChokyosiCode]='" & mKey.str & "'"
    strSQL = strSQL & " ORDER BY [BirthDate]"
    mAsyncCN_UMA.Execute strSQL, , adAsyncExecute
    
    gApp.Log "CH Fetch Complete"
    
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
    
    Call freers(mRS_UMA)

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
End Sub


'
'   機能: 調教師マスタレコードセットの取得完了イベントハンドラ
'
'   備考: Fetchから抜けて裏で行われる処理
'
Private Sub mAsyncCN_UMA_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
On Error GoTo ErrorHandler

    If pError Is Nothing Then
        Set mRS_UMA = pRecordset
        
        If mRS_UMA.EOF Then
            RaiseEvent NoDataUMA
        Else
            Call Set_Syuyou
            Call Set_Uma
        End If
    Else
        With pError
            gApp.Log .Description & .SQLState & .Source & .Number
        End With
        RaiseEvent NoDataUMA
    End If
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Sub



'
'   機能: 調教師画面用データを作る(主要勝鞍）
'
'   備考: なし
'
Private Sub Set_Syuyou()
On Error GoTo ErrorHandler
    Dim gd      As clsGridData
    Dim UM      As ADODB.Recordset
    Dim RA      As ADODB.Recordset
    Dim lngCP        As Long            '' セル入力用
    Dim lngRP        As Long            '' セル入力用
    Dim i As Long
    Dim j As Long
    
    'RACEを開く
    Set mRS_RACE = New ADODB.Recordset
    Set mCN_RACE = gApp.GetCN_RACE
    Call OpenTableDirect(mRS_RACE, mCN_RACE, "RACE")
    gApp.Log "CH mRS_RACE Complete Opened"
    
    Set UM = mRS_UMA
    Set RA = mRS_RACE
    
    ' 主要勝鞍　タブ
    
    ' グリッドデータオブジェクトを生成する
    Set gd = New clsGridData
    ' 主要勝鞍　タブ
    With gd
        .Rows = 4
        .Cols = 6
        'カラムヘッダ
        lngCP = 0
        lngRP = 0
        .SetItemMatrix lngRP, lngCP, "年月日"
        .SetItemMatrix lngRP, lngCP, "場回日"
        .SetItemMatrix lngRP, lngCP, "R"
        .SetItemMatrix lngRP, lngCP, "競走"
        .SetItemMatrix lngRP, lngCP, "出走"
        .SetItemMatrix lngRP, lngCP, "馬名"
    End With
    'データ
    For i = 0 To 2
        lngCP = 0
        lngRP = i + 1
        With mBuf_CH.SaikinJyusyo(i)
            With .SaikinJyusyoid
                Call SafeSeek(RA, Array("Year", "MonthDay", "JyoCD", "Kaiji", "Nichiji", "RaceNum"), _
                        Array(.Year, .MonthDay, .JyoCD, .Kaiji, .Nichiji, .RaceNum))
            End With
            If RA.EOF Then
                With .SaikinJyusyoid
                    gd.SetItemMatrix lngRP, lngCP, mSC.YMD4(.Year & .MonthDay)
                    gd.SetItemMatrix lngRP, lngCP, mCC.KIBJ3(.JyoCD) & mSC.KN1(.Kaiji & .Nichiji)
                    gd.SetItemMatrix lngRP, lngCP, IIf(val(.RaceNum) = 0, "", val(.RaceNum))
                End With
                gd.SetItemMatrix lngRP, lngCP, Trim$(.Ryakusyo6) & mCC.GRAD2(.GradeCD), Trim$(.Hondai)
            Else
                With .SaikinJyusyoid
                    gd.SetItemMatrix lngRP, lngCP, mSC.YMD4(.Year & .MonthDay), , , "RA", .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum
                    gd.SetItemMatrix lngRP, lngCP, mCC.KIBJ3(.JyoCD) & mSC.KN1(.Kaiji & .Nichiji), , , "RA", .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum
                    gd.SetItemMatrix lngRP, lngCP, IIf(val(.RaceNum) = 0, "", val(.RaceNum)), , , "RA", .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum
                End With
                gd.SetItemMatrix lngRP, lngCP, Trim$(.Ryakusyo6) & mCC.GRAD2(.GradeCD), Trim$(RA("Hondai")) & mCC.GRAD3(RA("GradeCD")), , "RA", _
                    .SaikinJyusyoid.Year & .SaikinJyusyoid.MonthDay & .SaikinJyusyoid.JyoCD & .SaikinJyusyoid.Kaiji & .SaikinJyusyoid.Nichiji & .SaikinJyusyoid.RaceNum
            End If
            gd.SetItemMatrix lngRP, lngCP, IIf(val(.SyussoTosu) = 0, "", Right$(" " & val(.SyussoTosu), 2) & "頭")
            gd.SetItemMatrix lngRP, lngCP, Trim$(.BAMEI), , , "UM", Fetch_KettoNum(.KettoNum)
        End With
    Next i
    RA.Close
    Set RA = Nothing
    
    RaiseEvent FetchCompleteSYUYOU(gd)
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: 競走馬グリッドを作る
'
'   備考: なし
'
Private Sub Set_Uma()
On Error GoTo ErrorHandler
    Dim gd      As clsGridData
    Dim rs      As ADODB.Recordset
    Dim BN      As ADODB.Recordset
    Dim BR      As ADODB.Recordset
    Dim lngCP        As Long            '' セル入力用
    Dim lngRP        As Long            '' セル入力用
    Set rs = mRS_UMA
    
    Dim buf_UM As JV_UM_UMA
    Dim i As Long
    Dim j As Long
    
    ' 競走馬　タブ
    
    ' グリッドデータオブジェクトを生成する
    Set gd = New clsGridData
    
    ' 馬主リンク確認用
    Set BN = New ADODB.Recordset
    Call OpenTableDirect(BN, gApp.GetCN_BANUSI, "BANUSI")
    
    ' 生産者SEISANリンク確認用
    Set BR = New ADODB.Recordset
    Call OpenTableDirect(BR, gApp.GetCN_SEISAN, "SEISAN")

    With gd
        ' グリッドデータにロウ数、カラム数を設定
        .Rows = 10
        .Cols = 8
        
        'カラムヘッダ
        lngCP = 0
        lngRP = 0
        .SetItemMatrix lngRP, lngCP, "馬記号"
        .SetItemMatrix lngRP, lngCP, "馬名"
        .SetItemMatrix lngRP, lngCP, "性齢"
        .SetItemMatrix lngRP, lngCP, "生年月日"
        .SetItemMatrix lngRP, lngCP, "馬主"
        .SetItemMatrix lngRP, lngCP, "生産者"
        .SetItemMatrix lngRP, lngCP, "産地"
        .SetItemMatrix lngRP, lngCP, "抹消"
    End With
    rs.MoveFirst
    
    lngRP = 1
    Do While Not rs.EOF
        ' バックグラウンド
        DoEvents
        If mblnCancelFetching Then
            Exit Sub
        End If
        
        Call SafeSeek(BN, Array("BanusiCode"), Array(rs("BanusiCode").value))
        Call SafeSeek(BR, Array("BreederCode"), Array(rs("BreederCode").value))

        lngCP = 0
        With gd
            .SetItemMatrix lngRP, lngCP, mCC.UMKG1(rs("UmaKigoCD")), mCC.UMKG3(rs("UmaKigoCD"))
            .SetItemMatrix lngRP, lngCP, Trim$(rs("Bamei")), , , "UM", rs("KettoNum")
            .SetItemMatrix lngRP, lngCP, mCC.SEIB4(rs("SexCD")) & mSC.Age(mBuf_CH.head.MakeDate.Year, rs("BirthDate"))
            .SetItemMatrix lngRP, lngCP, mSC.YMD1(rs("BirthDate"))
            .SetItemMatrix lngRP, lngCP, Trim$(rs("BanusiName")), , , "BN", IfExist(BN, "BanusiCode")
            .SetItemMatrix lngRP, lngCP, Trim$(rs("BreederName")), , , "BR", IfExist(BR, "BreederCode")
            .SetItemMatrix lngRP, lngCP, Trim$(rs("SanchiName"))
            .SetItemMatrix lngRP, lngCP, Trim$(IIf(val(rs("DelKubun")) = 1, "抹消", ""))
        End With
        rs.MoveNext
        lngRP = lngRP + 1
        If lngRP >= gd.Rows Then
            gd.Rows = gd.Rows + 10
        End If
    Loop
    rs.Close
  
    Set rs = Nothing
    
    gd.Rows = lngRP
    
    RaiseEvent FetchCompleteUMA(gd)
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: 血統番号を取得する
'
'   備考: なし
'
Private Function Fetch_KettoNum(key As String) As String
On Error GoTo ErrorHandler
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim strSQL As String
    
    Set rs = New ADODB.Recordset
    
    ' レコードセットを取得する
    strSQL = "SELECT [Bamei] FROM UMA WHERE [KettoNum]='" & key & "'"
    Set cn = gApp.GetCN_UMA
    rs.Open strSQL, cn, adOpenKeyset, adLockReadOnly, adCmdText

    ' データが無かったらstrKeyを返して終了
    If rs.EOF Or rs.BOF Then
        rs.Close
        Set rs = Nothing
        Fetch_KettoNum = ""
        Exit Function
    End If
    
    Fetch_KettoNum = key
    
    rs.Close
    
    Set rs = Nothing
    
    Exit Function
ErrorHandler:
    gApp.ErrLog
End Function

