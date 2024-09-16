VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDataKS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "騎手 データクラス"
'
'   騎手 データクラス
'

Option Explicit

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   外部関数(イベント)
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

' イベント
Public Event NoData()

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部変数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

' 変換ユーティリティー
Private mCC As clsCodeConverter          '' JV-Data のコード変換メソッド群
Private mSC As clsStringConverter          '' JV-Data のコード変換メソッド群

' Viewerに提供する為のデータ格納変数
Private mstrLabels(0 To 4) As String     '' テキストデータ
Private mstrLinkLabels(0 To 0) As clsGridItem     '' リンクデータ
Private mGridData(0 To 4) As clsGridData '' グリッドデータ

'コネクション
Private mCN_RACE As ADODB.Connection
Private mCN_UMA As ADODB.Connection
Attribute mCN_UMA.VB_VarHelpID = -1

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
'   機能: リンクラベル用文字列を返す
'
'   備考: なし
'
Public Property Get LinkLabels(Index As Integer) As clsGridItem
    Set LinkLabels = mstrLinkLabels(Index)
End Property

'
'   機能: グリッドデータを返す
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
Public Function Fetch(ByRef objKey As clsKeyKS) As Boolean
On Error GoTo ErrorHandler
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim CH As ADODB.Recordset
    Dim RA As ADODB.Recordset
    Dim UM As ADODB.Recordset
    Dim buf_KS As JV_KS_KISYU
    Dim strSQL As String
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim m As Long
    Dim lngCP        As Long            '' セル入力用
    Dim lngRP        As Long            '' セル入力用
    
    ' グリッドデータオブジェクトを生成する
    For i = 0 To 4
        Set mGridData(i) = New clsGridData
    Next i

    'KISHU (騎手マスタ)
    Set cn = gApp.GetCN_KISHU
    strSQL = "SELECT * FROM KISHU " & objKey.SQLWHEREString
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenKeyset, adLockReadOnly, adCmdText

    ' データが無かったらFalseを返して終了
    If rs.EOF Or rs.BOF Then
        rs.Close
        Fetch = False
        RaiseEvent NoData
        Exit Function
    End If

    Call SetDataFromRS_KS(rs, buf_KS)
    
    With buf_KS
        'CHOKYO (調教師マスタ)
        strSQL = "SELECT * FROM CHOKYO WHERE [ChokyosiCode]='" & .ChokyosiCode & "'"
        Set CH = New ADODB.Recordset
        CH.Open strSQL, gApp.GetCN_CHOKYO, adOpenKeyset, adLockReadOnly, adCmdText
        
        mstrLabels(0) = ""
        mstrLabels(0) = mstrLabels(0) & mSC.YMD1(rs("MakeDate")) & "作成データ"   '作成年月日
        
        ' 太字レース名等グレーバー用文字列生成
        mstrLabels(1) = ""
        mstrLabels(1) = mstrLabels(1) & Trim$(.KisyuName) & mCC.KSMN1(.MinaraiCD)
        mstrLabels(2) = ""
        If .TozaiCD <> 4 Then
            mstrLabels(2) = mstrLabels(2) & mSC.CTrim(.KisyuNameKana)
        End If
            
        ' ラベル１用文字列生成
        mstrLabels(3) = ""
        mstrLabels(3) = mstrLabels(3) & mSC.UM_DelKubun(.DelKubun)  '抹消区分
        mstrLabels(3) = mstrLabels(3) & mSC.YMD1(rs("DelDate"))  '抹消年月日
        mstrLabels(3) = mstrLabels(3) & vbCrLf  '改行
        mstrLabels(3) = mstrLabels(3) & IIf(val(.TozaiCD) = 0, "", "所属　　：" & mCC.TZSZ5(.TozaiCD)) & vbCrLf  '東西所属
        mstrLabels(3) = mstrLabels(3) & IIf(val(.SikakuCD) = 0, "", "資格　　：" & mCC.KSSK1(.SikakuCD)) & vbCrLf  '騎乗資格
        mstrLabels(3) = mstrLabels(3) & IIf(val(rs("IssueDate")) = 0, "", "初免許　：" & mSC.YMD1(rs("IssueDate"))) & vbCrLf  '免許交付
        mstrLabels(3) = mstrLabels(3) & IIf(val(rs("BirthDate")) = 0, "", "生年月日：" & mSC.YMD1(rs("BirthDate")))   '生年月日

        ' ラベル２用文字列生成
        mstrLabels(4) = ""
        mstrLabels(4) = mstrLabels(4) & IIf(val(.SexCD) = 0, "", "性別　　　：" & mSC.KS_SexCD(.SexCD) & " ") & vbCrLf  '性別
        If Not CH.EOF Then
            mstrLabels(4) = mstrLabels(4) & "所属調教師："
            Set mstrLinkLabels(0) = New clsGridItem
            mstrLinkLabels(0).Text = Trim$(.ChokyosiRyakusyo)
            mstrLinkLabels(0).Link = "CH"
            mstrLinkLabels(0).key = .ChokyosiCode
            mstrLinkLabels(0).ToolTip = ""
        End If
        mstrLabels(4) = mstrLabels(4) & vbCrLf
        mstrLabels(4) = mstrLabels(4) & IIf(Trim$(.Syotai) = "", "", "招待地域　：" & Trim$(.Syotai))
    
    End With
    
    rs.Close
    CH.Close

    ' 以下Grid処理
    'KISHU_SEISEKI (騎手マスタ_成績)
    Set cn = gApp.GetCN_KISHU_SEISEKI
    strSQL = "SELECT * FROM KISHU_SEISEKI " & objKey.SQLWHEREString
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenKeyset, adLockReadOnly, adCmdText
    
    Call SetDataFromRS_KS_SEISEKI(rs, buf_KS)

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
        
        .SetItemMatrix 1, 0, "本年(" & buf_KS.HonZenRuikei(0).SetYear & "年)"
        .SetItemMatrix 2, 0, "平地"
        .SetItemMatrix 3, 0, "障害"
        .SetItemMatrix 4, 0, "前年(" & buf_KS.HonZenRuikei(1).SetYear & "年)"
        .SetItemMatrix 5, 0, "平地"
        .SetItemMatrix 6, 0, "障害"
        .SetItemMatrix 7, 0, "累計　  　　"
        .SetItemMatrix 8, 0, "平地"
        .SetItemMatrix 9, 0, "障害"
    End With
    'データ
    For i = 0 To 2 '0:本年 1:前年 2:累計
        ' 平地
        With buf_KS.HonZenRuikei(i)
            mGridData(0).ItemMatrix(3 * i + 2, 1).Text = mSC.Money(.HonSyokinHeichi)
            mGridData(0).ItemMatrix(3 * i + 2, 2).Text = mSC.Money(.FukaSyokinHeichi)
            For j = 0 To 9
                mGridData(0).ItemMatrix(3 * i + 2, j + 3).Text = mSC.Chakukaisu2(.ChakuKaisuHeichi.Chakukaisu, j)
            Next j
        End With
        ' 障害
        With buf_KS.HonZenRuikei(i)
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
            With buf_KS.HonZenRuikei(i).ChakuKaisuKyori(j) '芝
                For k = 0 To 9
                    mGridData(1).ItemMatrix(4 * i + j + 3, k + 1).Text = mSC.Chakukaisu2(.Chakukaisu, k)
                Next k
            End With
        Next j
        For j = 0 To 2
            With buf_KS.HonZenRuikei(i).ChakuKaisuKyori(j + 3) 'ダート
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
                With buf_KS.HonZenRuikei(i).ChakuKaisuJyo(j + 2 * k)
                    For m = 0 To 9
                        mGridData(2).ItemMatrix(k + 3, 10 * (2 * i + j) + m + 1).Text = mSC.Chakukaisu2(.Chakukaisu, m)
                    Next m
                End With
            Next k
        Next j
    Next i
    
    'RACEを開く
    Set RA = New ADODB.Recordset
    Set mCN_RACE = gApp.GetCN_RACE
    Call OpenTableDirect(RA, mCN_RACE, "RACE")
    gApp.Log "KS mRS_RACE Complete Opened"
    'UMAを開く
    Set UM = New ADODB.Recordset
    Set mCN_UMA = gApp.GetCN_UMA
    Call OpenTableDirect(UM, mCN_UMA, "UMA")
    gApp.Log "KS mRS_UMA Complete Opened"
    
    ' 初騎乗　タブ
    With mGridData(3)
        .Rows = 5
        .Cols = 9
        'カラムヘッダ
        lngCP = 0
        lngRP = 0
        .SetItemMatrix lngRP, lngCP, " "
        .SetItemMatrix lngRP, lngCP, " "
        .SetItemMatrix lngRP, lngCP, "年月日"
        .SetItemMatrix lngRP, lngCP, "場回日"
        .SetItemMatrix lngRP, lngCP, "R"
        .SetItemMatrix lngRP, lngCP, "出走"
        .SetItemMatrix lngRP, lngCP, "馬名"
        .SetItemMatrix lngRP, lngCP, "確定"
        .SetItemMatrix lngRP, lngCP, "異常"
        ' ロウヘッダ
        lngRP = 1
        .SetItemMatrix lngRP + 0, 0, "初騎乗情報"
        .SetItemMatrix lngRP + 1, 0, "初騎乗情報"
        .SetItemMatrix lngRP + 2, 0, "初勝利情報"
        .SetItemMatrix lngRP + 3, 0, "初勝利情報"
        lngRP = 1
        .SetItemMatrix lngRP + 0, 1, "平地"
        .SetItemMatrix lngRP + 1, 1, "障害"
        .SetItemMatrix lngRP + 2, 1, "平地"
        .SetItemMatrix lngRP + 3, 1, "障害"
    End With
    'データ
    For i = 0 To 1
        lngCP = 2
        lngRP = i + 1
        With buf_KS.HatuKiJyo(i)
            With .Hatukijyoid
                Call SafeSeek(RA, Array("Year", "MonthDay", "JyoCD", "Kaiji", "Nichiji", "RaceNum"), _
                        Array(.Year, .MonthDay, .JyoCD, .Kaiji, .Nichiji, .RaceNum))
            End With
            If RA.EOF Then
                With .Hatukijyoid
                    mGridData(3).SetItemMatrix lngRP, lngCP, mSC.YMD4(.Year & .MonthDay)
                    mGridData(3).SetItemMatrix lngRP, lngCP, mCC.KIBJ3(.JyoCD) & mSC.KN1(.Kaiji & .Nichiji)
                    mGridData(3).SetItemMatrix lngRP, lngCP, IIf(val(.RaceNum) = 0, "", val(.RaceNum))
                End With
            Else
                With .Hatukijyoid
                    mGridData(3).SetItemMatrix lngRP, lngCP, mSC.YMD4(.Year & .MonthDay), , , "RA", .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum
                    mGridData(3).SetItemMatrix lngRP, lngCP, mCC.KIBJ3(.JyoCD) & mSC.KN1(.Kaiji & .Nichiji), , , "RA", .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum
                    mGridData(3).SetItemMatrix lngRP, lngCP, IIf(val(.RaceNum) = 0, "", val(.RaceNum)), , , "RA", .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum
                End With
            End If
            mGridData(3).SetItemMatrix lngRP, lngCP, IIf(val(.SyussoTosu) = 0, "", Right$(" " & val(.SyussoTosu), 2) & "頭")
            Call SafeSeek(UM, Array("KettoNum"), Array(.KettoNum))
            If UM.EOF Then
                mGridData(3).SetItemMatrix lngRP, lngCP, Trim$(.BAMEI)
            Else
                mGridData(3).SetItemMatrix lngRP, lngCP, Trim$(.BAMEI), , , "UM", .KettoNum
            End If
            mGridData(3).SetItemMatrix lngRP, lngCP, IIf(val(.KakuteiJyuni) = 0, "", val(.KakuteiJyuni) & "着")
            mGridData(3).SetItemMatrix lngRP, lngCP, mCC.IJKB2(.IJyoCD), mCC.IJKB1(.IJyoCD)
        End With
        lngCP = 2
        lngRP = i + 3
        With buf_KS.HatuSyori(i)
            With .Hatusyoriid
                Call SafeSeek(RA, Array("Year", "MonthDay", "JyoCD", "Kaiji", "Nichiji", "RaceNum"), _
                        Array(.Year, .MonthDay, .JyoCD, .Kaiji, .Nichiji, .RaceNum))
            End With
            If RA.EOF Then
                With .Hatusyoriid
                    mGridData(3).SetItemMatrix lngRP, lngCP, mSC.YMD4(.Year & .MonthDay)
                    mGridData(3).SetItemMatrix lngRP, lngCP, mCC.KIBJ3(.JyoCD) & mSC.KN1(.Kaiji & .Nichiji)
                    mGridData(3).SetItemMatrix lngRP, lngCP, IIf(val(.RaceNum) = 0, "", val(.RaceNum))
                End With
            Else
                With .Hatusyoriid
                    mGridData(3).SetItemMatrix lngRP, lngCP, mSC.YMD4(.Year & .MonthDay), , , "RA", .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum
                    mGridData(3).SetItemMatrix lngRP, lngCP, mCC.KIBJ3(.JyoCD) & mSC.KN1(.Kaiji & .Nichiji), , , "RA", .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum
                    mGridData(3).SetItemMatrix lngRP, lngCP, IIf(val(.RaceNum) = 0, "", val(.RaceNum)), , , "RA", .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum
                End With
            End If
            mGridData(3).SetItemMatrix lngRP, lngCP, IIf(val(.SyussoTosu) = 0, "", Right$(" " & val(.SyussoTosu), 2) & "頭")
            Call SafeSeek(UM, Array("KettoNum"), Array(.KettoNum))
            If UM.EOF Then
                mGridData(3).SetItemMatrix lngRP, lngCP, Trim$(.BAMEI)
            Else
                mGridData(3).SetItemMatrix lngRP, lngCP, Trim$(.BAMEI), , , "UM", .KettoNum
            End If
        End With
    Next i

    
    ' 主要勝鞍　タブ
    With mGridData(4)
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
        With buf_KS.SaikinJyusyo(i)
            With .SaikinJyusyoid
                Call SafeSeek(RA, Array("Year", "MonthDay", "JyoCD", "Kaiji", "Nichiji", "RaceNum"), _
                        Array(.Year, .MonthDay, .JyoCD, .Kaiji, .Nichiji, .RaceNum))
            End With
            If RA.EOF Then
                With .SaikinJyusyoid
                    mGridData(4).SetItemMatrix lngRP, lngCP, mSC.YMD4(.Year & .MonthDay)
                    mGridData(4).SetItemMatrix lngRP, lngCP, mCC.KIBJ3(.JyoCD) & mSC.KN1(.Kaiji & .Nichiji)
                    mGridData(4).SetItemMatrix lngRP, lngCP, IIf(val(.RaceNum) = 0, "", val(.RaceNum))
                End With
                mGridData(4).SetItemMatrix lngRP, lngCP, Trim$(.Ryakusyo6) & mCC.GRAD2(.GradeCD), Trim$(.Hondai)
            Else
                With .SaikinJyusyoid
                    mGridData(4).SetItemMatrix lngRP, lngCP, mSC.YMD4(.Year & .MonthDay), , , "RA", .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum
                    mGridData(4).SetItemMatrix lngRP, lngCP, mCC.KIBJ3(.JyoCD) & mSC.KN1(.Kaiji & .Nichiji), , , "RA", .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum
                    mGridData(4).SetItemMatrix lngRP, lngCP, IIf(val(.RaceNum) = 0, "", val(.RaceNum)), , , "RA", .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum
                End With
                mGridData(4).SetItemMatrix lngRP, lngCP, Trim$(.Ryakusyo6) & mCC.GRAD2(.GradeCD), Trim$(RA("Hondai")) & mCC.GRAD3(RA("GradeCD")), , "RA", _
                    .SaikinJyusyoid.Year & .SaikinJyusyoid.MonthDay & .SaikinJyusyoid.JyoCD & .SaikinJyusyoid.Kaiji & .SaikinJyusyoid.Nichiji & .SaikinJyusyoid.RaceNum
            End If
            mGridData(4).SetItemMatrix lngRP, lngCP, IIf(val(.SyussoTosu) = 0, "", Right$(" " & val(.SyussoTosu), 2) & "頭")
            Call SafeSeek(UM, Array("KettoNum"), Array(.KettoNum))
            If UM.EOF Then
                mGridData(4).SetItemMatrix lngRP, lngCP, Trim$(.BAMEI)
            Else
                mGridData(4).SetItemMatrix lngRP, lngCP, Trim$(.BAMEI), , , "UM", .KettoNum
            End If
        End With
    Next i

    rs.Close
    Set rs = Nothing
    RA.Close
    Set RA = Nothing
    UM.Close
    Set UM = Nothing
    
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

