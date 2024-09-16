VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStringConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "各種文字列の変換クラス"
'
'   文字列変換クラス
'

Option Explicit

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   外部関数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: ヘッダ馬場情報
'
'   備考: なし
'
Public Function BabaHeader(RA As ADODB.Recordset) As String
    
    Select Case RA("TrackCD")
        Case "00", "27", "28"           '初期値,砂
            BabaHeader = ""
        Case "52"                       '芝→ダート
            BabaHeader = IIf(RA("SibaBabaCD") = "0", "", "芝:" & gCC.BBJT1(RA("SibaBabaCD"))) _
                & " " & IIf(RA("DirtBabaCD") = "0", "", "ダート:" & gCC.BBJT1(RA("DirtBabaCD")))
        Case "10" To "22", "51" To "59" '芝
            BabaHeader = IIf(RA("SibaBabaCD") = "0", "", "芝:" & gCC.BBJT1(RA("SibaBabaCD")))
        Case "23" To "29"               'ダート
            BabaHeader = IIf(RA("DirtBabaCD") = "0", "", "ダート:" & gCC.BBJT1(RA("DirtBabaCD")))
    End Select

End Function


'
'   機能: グリッド馬場情報
'
'   備考: なし
'
Public Function BabaGrid(RA As ADODB.Recordset) As String
        
    Select Case RA("TrackCD")
        Case "00", "27", "28"           '初期値,砂
            BabaGrid = ""
        Case "52"                       '芝→ダート
            BabaGrid = IIf(RA("SibaBabaCD") = "0", "", "芝" & gCC.BBJT4(RA("SibaBabaCD"))) _
                & IIf(RA("SibaBabaCD") = "0" Or RA("DirtBabaCD") = "0", "", ":") _
                & IIf(RA("DirtBabaCD") = "0", "", "ダ" & gCC.BBJT4(RA("DirtBabaCD")))
        Case "10" To "22", "51" To "59" '芝
            BabaGrid = IIf(RA("SibaBabaCD") = "0", "", gCC.BBJT4(RA("SibaBabaCD")))
        Case "23" To "29"               'ダート
            BabaGrid = IIf(RA("DirtBabaCD") = "0", "", gCC.BBJT4(RA("DirtBabaCD")))
    End Select

End Function


'
'   機能: 前後スペース削除、(文字間)連続スペースを半角スペース3つに変換する
'
'   備考: （全角半角混在の場合）今のところ中のスペースが"_","＿"の順番にあると想定
'
Public Function CTrim(str As String) As String
    Dim strTmp As String
    Dim strOut As String
    Dim lngTmp As Long
    
    strTmp = Trim$(str)
    
    If InStr(strTmp, " ") Then
        lngTmp = InStr(strTmp, " ")
        strTmp = Replace(strTmp, " ", "")
        strTmp = Replace(strTmp, "　", "")
        strOut = Mid$(strTmp, 1, lngTmp - 1) & "   " & Mid$(strTmp, lngTmp)
    ElseIf InStr(strTmp, "　") Then
        lngTmp = InStr(strTmp, "　")
        strTmp = Replace(strTmp, " ", "")
        strTmp = Replace(strTmp, "　", "")
        strOut = Mid$(strTmp, 1, lngTmp - 1) & "   " & Mid$(strTmp, lngTmp)
    Else
        strOut = strTmp
    End If
    CTrim = strOut
End Function


'
'   機能: 空データの判断
'
'   備考: " ","0"だけのデータはTrue:データなし
'
Public Function Emp(str As String) As Boolean
    If str = "" Then
        Emp = True
    ElseIf Space(Len(str)) = str Then
        Emp = True
    ElseIf String$(Len(str), "0") = str Then
        Emp = True
    Else
        Emp = False
    End If
End Function


'
'   機能: 文字列を数値の文字列にする
'
'   備考: "" -> "", "000" -> "0", "0020" -> "20"
'         str(val("XXX")) と似てるが、 全部"0"の時以外のエラー文字では、0にならない
'
Public Function ValStr(strVal As String) As String
    Dim out As String
    out = val(strVal)
    If out = "0" Then
        If strVal <> String$(Len(strVal), "0") Or strVal = "" Then
            out = ""
        End If
    End If
    ValStr = out
End Function


'
'   機能: 年月日のフォーマットに合わせる
'
'   備考: 返り値 yyyy年mm月dd日
'
Public Function YMD1(strYMD As String) As String
    If val(strYMD) = 0 Then
        YMD1 = ""
    Else
        YMD1 = Mid$(strYMD, 1, 4) & "年" _
                & Right$(" " & val(Mid$(strYMD, 5, 2)), 2) & "月" _
                & Right$(" " & val(Mid$(strYMD, 7, 2)), 2) & "日"
    End If
End Function


'
'   機能: 年月日のフォーマットに合わせる
'
'   備考: 返り値 yyyy/mm/dd
'
Public Function YMD2(strYMD As String) As String
    If val(strYMD) = 0 Then
        YMD2 = ""
    Else
        YMD2 = Mid$(strYMD, 1, 4) & "/" _
                & Right$(" " & val(Mid$(strYMD, 5, 2)), 2) & "/" _
                & Right$(" " & val(Mid$(strYMD, 7, 2)), 2)
    End If
End Function


'
'   機能: 年月日のフォーマットに合わせる
'
'   備考: 返り値 yy/mm/dd
'
Public Function YMD3(strYMD As String) As String
    If val(strYMD) = 0 Then
        YMD3 = ""
    Else
        YMD3 = Format$(Right$(strYMD, 6), "@@/@@/@@")
    End If
End Function


'
'   機能: 年月日(曜)のフォーマットに合わせる
'
'   備考: 返り値 yyyy年mm月dd日(w)
'
Public Function YMD4(strYMD As String) As String
    If val(strYMD) = 0 Then
        YMD4 = ""
    Else
        YMD4 = Mid$(strYMD, 1, 4) & "年" _
                & Right$(" " & val(Mid$(strYMD, 5, 2)), 2) & "月" _
                & Right$(" " & val(Mid$(strYMD, 7, 2)), 2) & "日" _
                & "(" & Format$(Format$(strYMD, "####/##/##"), "AAA") & ")"
    End If
End Function


'
'   機能: 月日時分のフォーマットに合わせる
'
'   備考: 返り値 mm月dd日HH時MM分
'
Public Function MDHM(str As String) As String
    If val(str) = 0 Then
        MDHM = ""
    Else
        MDHM = Right$(" " & val(Mid$(str, 1, 2)), 2) & "月" _
                & Right$(" " & val(Mid$(str, 3, 2)), 2) & "日" _
                & Right$(" " & val(Mid$(str, 5, 2)), 2) & "時" _
                & Right$(" " & val(Mid$(str, 7, 2)), 2) & "分"
    End If
End Function


'
'   機能: 999.9秒のフォーマットに合わせる
'
'   備考: 0のデータが来たら""を返す（ハロン用）
'
Public Function SSSs(strSSSs As String) As String
    If val(strSSSs) = 0 Then
        SSSs = ""
    Else
        SSSs = Format$(val(strSSSs) / 10, "##0.0")
    End If
End Function


'
'   機能: 9:99.99秒のフォーマットに合わせる
'
'   備考: 返り値 9:99.99
'
Public Function MSSss(strMSSss As String) As String
    MSSss = val(Mid$(strMSSss, 1, 1)) & ":" _
            & Right$("0" & val(Mid$(strMSSss, 2, 2)), 2) & "." _
            & Right$("0" & val(Mid$(strMSSss, 4, 2)), 2)
End Function


'
'   機能: 99.99秒のフォーマットに合わせる
'
'   備考: 返り値 99.99
'
Public Function SSpp(strSSpp As String) As String
    SSpp = Format$(strSSpp, "@@.@@")
End Function


'
'   機能: 時分のフォーマットに合わせる
'
'   備考: 返り値 hh時nn分
'
Public Function HHNN1(strHHNN As String) As String
    HHNN1 = Right$("  " & val(Mid$(strHHNN, 1, 2)), 2) & "時" _
          & Right$("  " & val(Mid$(strHHNN, 3, 2)), 2) & "分"
End Function


'
'   機能: 時分のフォーマットに合わせる
'
'   備考: 返り値 hh:nn
'
Public Function HHNN2(strHHNN As String) As String
    HHNN2 = Right$("  " & val(Mid$(strHHNN, 1, 2)), 2) & ":" _
          & Mid$(strHHNN, 3, 2)
End Function


'
'   機能: 回日のフォーマットに合わせる
'
'   備考: 返り値 kk回nn日
'
Public Function KN1(strKN As String) As String
    If val(strKN) = 0 Then
        KN1 = ""
    Else
        KN1 = Right$(" " & val(Mid$(strKN, 1, 2)), 2) & "回" _
                & Right$(" " & val(Mid$(strKN, 3, 2)), 2) & "日"
    End If
End Function


'
'   機能: 数値にスペースを挿入してデータ幅を合わせる
'
'   備考: 返り値 右寄せ数字文字列（データが0の場合はスペースのみで表示）
'
Public Function FitSpaceNum(str As String, num As Integer) As String
    If val(str) = 0 Then
        FitSpaceNum = Space(num)
    Else
        FitSpaceNum = Right$(Space(num) & val(str), num)
    End If
End Function


'
'   機能: 文字列に全角スペースを挿入してデータ幅を合わせる
'
'   備考: 返り値 左寄せ文字列
'
Public Function FitSpaceStr(str As String, num As Integer) As String
    FitSpaceStr = Left$(Trim$(str) & String$(num, "　"), num)
End Function


'
'   機能: 馬年齢を返す
'
'   備考: なし
'
Public Function Age(ARG1 As String, ARG2 As String) As String
    Dim lngTmp As Long
    
    If val(ARG1) = 0 Or val(ARG2) = 0 Then
        Age = ""
        gApp.Log "Age :引き数に0があります（馬齢を表示できません）"
    Else
        lngTmp = Left$(ARG1, 4) - Left$(ARG2, 4)
        If lngTmp >= 0 Then
            Age = Right$(" " & lngTmp, 2)
        Else
            Age = ""
            gApp.Log "Age :年齢がマイナスです"
        End If
    End If
End Function


'
'   機能: 賞金の統一フォーマットを返す
'
'   備考: なし
'
Public Function Money(ARG1 As String) As String
    Money = Format$(val(ARG1) * 100, "#,###,###,##0円")
End Function


'
'   機能: 賞金の統一フォーマットを返す
'
'   備考: なし
'
Public Function Money2(ARG1 As String, num As Long) As String
    Money2 = Right$(Space(num - 2) & Format$(val(ARG1) * 100, "#,###,###,##0円"), num)
End Function


'
'   機能: 賞金の統一フォーマットを返す
'
'   備考: なし
'
Public Function Money3(ARG1 As String) As String
    If ARG1 <> "" Then
        Money3 = Format$(val(ARG1) * 100, "#,###,###,##0")
    End If
End Function


'
'   機能: 着回数から各種表示文字列を返す
'
'   備考: 返り値 option - 1:着回数 2:合計 3:勝率 4:連対率
'
Public Function Chakukaisu(vStr As String, vOption As Integer) As String
    Dim i As Integer
    Dim intDataLen As Integer
    Dim sngGet(6) As Single
    Dim sngTmp As Single
    Dim strOut As String
    
    If Len(vStr) = 3 * 6 Then
        intDataLen = 3
    ElseIf Len(vStr) = 6 * 6 Then
        intDataLen = 6
    ElseIf Len(vStr) = 0 Then
        Chakukaisu = ""
    Else
#If DBG Then
  Debug.Print "着回数データが不適です"
#End If
    End If
    
    If Len(vStr) Then
        For i = 0 To 5
            sngGet(i) = Mid$(vStr, intDataLen * i + 1, intDataLen)
        Next i
        
        Select Case vOption
        Case 1
            strOut = ""
            For i = 0 To 5
                strOut = strOut & IIf(i = 0, "", "-") & Right$(Space(intDataLen) & sngGet(i), intDataLen)
            Next i
            Chakukaisu = strOut
        Case 2
            For i = 0 To 5
                sngTmp = sngTmp + sngGet(i)
            Next i
            Chakukaisu = Right$(Space(intDataLen) & sngTmp, intDataLen)
        Case 3
            Chakukaisu = "未"
        Case 4
            Chakukaisu = "未"
        Case Else
#If DBG Then
  Debug.Print "Chakukaisu Optionが不適です"
#End If
        End Select
    End If
End Function


'
'   機能: 着回数から各種表示文字列を返す
'
'   備考: 返り値 option - 0-5:着回数 6:合計 7:勝率 8:連対率 9:3着内率
'
Public Function Chakukaisu2(ARG() As String, vOption As Long) As String
    Dim i As Integer
    Dim intDataLen As Integer
    Dim sngGet(0 To 5) As Single
    Dim sngSum As Single
    Dim strOut As String
    
    If Len(ARG(0)) = 3 Then
        intDataLen = 3
    ElseIf Len(ARG(0)) = 6 Then
        intDataLen = 6
    ElseIf Len(ARG(0)) = 0 Then
        Chakukaisu2 = ""
    End If
    
    If Len(ARG(0)) Then
        For i = 0 To 5
            sngGet(i) = CSng(ARG(i))
        Next i
        sngSum = sngGet(0) + sngGet(1) + sngGet(2) + sngGet(3) + sngGet(4) + sngGet(5)
        sngSum = IIf(sngSum = 0, -1, sngSum)
        Select Case vOption
        Case 6
            Chakukaisu2 = CStr(IIf(sngSum = -1, 0, sngSum))
        Case 7
            Chakukaisu2 = Format$(sngGet(0) / sngSum, "0.000")
        Case 8
            Chakukaisu2 = Format$((sngGet(0) + sngGet(1)) / sngSum, "0.000")
        Case 9
            Chakukaisu2 = Format$((sngGet(0) + sngGet(1) + sngGet(2)) / sngSum, "0.000")
        Case Else
            Chakukaisu2 = CStr(sngGet(vOption))
        End Select
    End If
End Function


'
'   機能: 坂路調教(HC) - トレセン区分
'
'   備考: なし
'
Public Function HC_TKBN(strCD As String) As String
    Select Case strCD
    Case "0"
        HC_TKBN = "美浦"
    Case "1"
        HC_TKBN = "栗東"
    End Select
End Function


'
'   機能: レース詳細(RA) - レコード更新区分
'
'   備考: なし
'
Public Function RA_RKBN(strCD As String) As String
    Select Case strCD
    Case "0"
        RA_RKBN = ""
    Case "1"
        RA_RKBN = "基準タイム"
    Case "2"
        RA_RKBN = "レコード"
    End Select
End Function


'
'   機能: レース詳細(RA) - データ区分
'
'   備考: なし
'
Public Function RA_DKBN(strCD As String) As String
    Select Case strCD
    Case "1"
        RA_DKBN = "出走馬名表"
    Case "2"
        RA_DKBN = "出馬表"
    Case "3"
        RA_DKBN = "速報成績(3着まで確定)"
    Case "4"
        RA_DKBN = "速報成績(5着まで確定)"
    Case "5"
        RA_DKBN = "速報成績(全馬着順確定)"
    Case "6"
        RA_DKBN = "速報成績(全馬着順+ラップタイム)"
    Case "7"
        RA_DKBN = "成績"
    Case "A"
        RA_DKBN = "地方競馬"
    Case "B"
        RA_DKBN = "海外国際レース"
    Case "9"
        RA_DKBN = "レース中止"
    Case "0"
        RA_DKBN = "該当レコード削除(提供ミスなどの理由による)"
    End Select
End Function


'
'   機能: 馬毎レース情報(SE) - マイニング区分
'
'   備考: なし
'
Public Function SE_DMKBN(strCD As String) As String
    Select Case strCD
    Case "1"
        SE_DMKBN = "前日"
    Case "2"
        SE_DMKBN = "当日"
    Case "3"
        SE_DMKBN = "直前"
    End Select
End Function


'
'   機能: 特別登録馬(TK) - データ区分
'
'   備考: なし
'
Public Function TK_DKBN(strCD As String) As String
    Select Case strCD
    Case "1"
        TK_DKBN = "ハンデ発表前"
    Case "2"
        TK_DKBN = "ハンデ発表後"
    Case "0"
        TK_DKBN = "該当レコード削除(提供ミスなどの理由による)"
    End Select
End Function


'
'   機能: 特別登録馬(TK) - 　　交流区分
'
'   備考: なし
'
Public Function TK_KKBN(strCD As String) As String
    Select Case strCD
    Case "0"
        TK_KKBN = ""
    Case "1"
        TK_KKBN = "地方馬"
    Case "2"
        TK_KKBN = "外国馬"
    End Select
End Function


'
'   機能: 天候馬場状態(WE) - 変更識別
'
'   備考: なし
'
Public Function WE_HKBN(strCD As String) As String
    Select Case strCD
    Case "1"
        WE_HKBN = "天候馬場初期状態"
    Case "2"
        WE_HKBN = "天候変更"
    Case "3"
        WE_HKBN = "馬場状態変更"
    End Select
End Function


'
'   機能: 出走取消・競走除外(AV) - データ区分
'
'   備考: なし
'
Public Function AV_DKBN(strCD As String) As String
    Select Case strCD
    Case "1"
        AV_DKBN = "出走取消"
    Case "2"
        AV_DKBN = "競走除外"
    Case "0"
        AV_DKBN = "該当レコード削除(提供ミスなどの理由による)"
    End Select
End Function


'
'   機能: 出走取消・競走除外(AV) - 事由区分
'
'   備考: なし
'
Public Function AV_JKBN(strCD As String) As String
    Select Case strCD
    Case "000"
        AV_JKBN = ""
    Case "001"
        AV_JKBN = "疾病"
    Case "002"
        AV_JKBN = "事故"
    Case "003"
        AV_JKBN = "その他"
    End Select
End Function


'
'   機能: 馬毎レース情報(SE) - 今回レース脚質判定
'
'   備考: なし
'
Public Function SE_KRKH(strCD As String) As String
    Select Case strCD
    Case "1"
        SE_KRKH = "逃"
    Case "2"
        SE_KRKH = "先"
    Case "3"
        SE_KRKH = "差"
    Case "4"
        SE_KRKH = "追"
    Case "0"
        SE_KRKH = "　"
    End Select
End Function


'
'   機能: 競走馬マスタ(UM) - 競走馬抹消区分
'
'   備考: なし
'
Public Function UM_DelKubun(strCD As String) As String
    Select Case strCD
    Case "0"
        UM_DelKubun = "" '現役非表示の為
    Case "1"
        UM_DelKubun = "抹消"
    End Select
End Function


'
'   機能: 騎手マスタ(KS) - 性別区分
'
'   備考: なし
'
Public Function KS_SexCD(strCD As String) As String
    Select Case strCD
    Case "1"
        KS_SexCD = "男性"
    Case "2"
        KS_SexCD = "女性"
    End Select
End Function


'
'   機能: 調教師マスタ(CH) - 調教師抹消区分
'
'   備考: なし
'
Public Function CH_DelKubun(strCD As String) As String
    Select Case strCD
    Case "0"
        CH_DelKubun = "" '現役非表示の為
    Case "1"
        CH_DelKubun = "抹消"
    End Select
End Function


'
'   機能: 調教師マスタ(CH) - 性別区分
'
'   備考: なし
'
Public Function CH_SexCD(strCD As String) As String
    Select Case strCD
    Case "1"
        CH_SexCD = "男性"
    Case "2"
        CH_SexCD = "女性"
    End Select
End Function


'
'   機能: 繁殖馬マスタ(HN) - 繁殖馬抹消区分
'
'   備考: なし
'
Public Function HN_DelKubun(strCD As String) As String
    Select Case strCD
    Case "0"
        HN_DelKubun = "" '現役非表示の為
    Case "1"
        HN_DelKubun = "抹消"
    End Select
End Function


'
'   機能: 繁殖馬マスタ(HN) - 繁殖馬持込区分
'
'   備考: なし
'
Public Function HN_HansyokuMochiKubun(strCD As String) As String
    Select Case strCD
    Case "0"
        HN_HansyokuMochiKubun = "内国産"
    Case "1"
        HN_HansyokuMochiKubun = "持込"
    Case "2"
        HN_HansyokuMochiKubun = "輸入内国産扱い"
    Case "3"
        HN_HansyokuMochiKubun = "輸入"
    Case "9"
        HN_HansyokuMochiKubun = "その他"
    End Select
End Function


'
'   機能: 異常区分コード の 略名(2文字) をソート順位用文字列にする
'
'   備考: ・異常区分の優先順位
'               1 競争中止
'               2 失格
'               3 発走除外
'               4 競争除外
'               5 出走取消
'               ※：降着、落馬再騎乗については、着が設定されますので、基本的に考慮不要
Public Function IJKB_Sort(strCD As String) As String
    Select Case strCD
    Case "0"
        IJKB_Sort = "9"
    Case "1" ' "出走取消"
        IJKB_Sort = "5"
    Case "2" ' "発走除外"
        IJKB_Sort = "3"
    Case "3" ' "競走除外"
        IJKB_Sort = "4"
    Case "4" ' "競走中止"
        IJKB_Sort = "1"
    Case "5" ' "失格"
        IJKB_Sort = "2"
    Case "6" ' "落馬再騎乗"
        IJKB_Sort = "9"
    Case "7" ' "降着"
        IJKB_Sort = "9"
    End Select
End Function

