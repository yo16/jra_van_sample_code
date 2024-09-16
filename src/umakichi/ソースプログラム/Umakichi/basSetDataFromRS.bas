Attribute VB_Name = "basSetDataFromRS"
'
'   構造体にレコードデータを取得するモジュール
'

Option Explicit

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   外部関数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: 構造体にデータをセットする－RAレコード
'
'   備考: なし
'
Public Sub SetDataFromRS_RA(ByRef rs As ADODB.Recordset, ByRef buf As JV_RA_RACE)
    Dim i As Integer                                '' ループカウンタ
    Dim j As Integer                                '' ループカウンタ
    Dim k As Integer                                '' ループカウンタ
    
    With buf
        With .head
            .RecordSpec = rs("RecordSpec")                                                      '' レコード種別
            .DataKubun = rs("DataKubun")                                                        '' データ区分
            With .MakeDate
                .Year = Mid$(rs("MakeDate"), 1, 4)                                               '' 年
                .Month = Mid$(rs("MakeDate"), 5, 2)                                              '' 月
                .Day = Mid$(rs("MakeDate"), 7, 2)                                                '' 日
            End With ' MakeDate
        End With ' head
        With .id
            .Year = rs("Year")
            .MonthDay = rs("MonthDay")
            .JyoCD = rs("JyoCD")
            .Kaiji = rs("Kaiji")
            .Nichiji = rs("Nichiji")
            .RaceNum = rs("RaceNum")
        End With
        With .RaceInfo
            .YoubiCD = rs("YoubiCD")                                                            '' 曜日コード
            .TokuNum = rs("TokuNum")                                                            '' 特別競走番号
            .Hondai = rs("Hondai")                                                              '' 競走名本題
            .Fukudai = rs("Fukudai")                                                            '' 競走名副題
            .Kakko = rs("Kakko")                                                                '' 競走名カッコ内
            .HondaiEng = rs("HondaiEng")                                                        '' 競走名本題欧字
            .FukudaiEng = rs("FukudaiEng")                                                      '' 競走名副題欧字
            .KakkoEng = rs("KakkoEng")                                                          '' 競走名カッコ内欧字
            .Ryakusyo10 = rs("Ryakusyo10")                                                      '' 競走名略称１０字
            .Ryakusyo6 = rs("Ryakusyo6")                                                        '' 競走名略称６字
            .Ryakusyo3 = rs("Ryakusyo3")                                                        '' 競走名略称３字
            .Kubun = rs("Kubun")                                                                '' 競走名区分
            .Nkai = rs("Nkai")                                                                  '' 重賞回次[第N回]
        End With ' RaceInfo
        .GradeCD = rs("GradeCD")                                                                '' グレードコード
        .GradeCDBefore = rs("GradeCDBefore")                                                    '' 変更前グレードコード
        With .JyokenInfo
            .SyubetuCD = rs("SyubetuCD")                                                        '' 競走種別コード
            .KigoCD = rs("KigoCD")                                                              '' 競走記号コード
            .JyuryoCD = rs("JyuryoCD")                                                          '' 重量種別コード
            For j = 0 To 4
                .JyokenCD(j) = rs("JyokenCD" & j + 1)                                           '' 競走条件コード
            Next j
        End With ' JyokenInfo
        .JyokenName = rs("JyokenName")                                                          '' 競走条件名称
        .KYORI = rs("Kyori")                                                                    '' 距離
        .KyoriBefore = rs("KyoriBefore")                                                        '' 変更前距離
        .TrackCD = rs("TrackCD")                                                                '' トラックコード
        .TrackCDBefore = rs("TrackCDBefore")                                                    '' 変更前トラックコード
        .CourseKubunCD = rs("CourseKubunCD")                                                    '' コース区分
        .CourseKubunCDBefore = rs("CourseKubunCDBefore")                                        '' 変更前コース区分
        For i = 0 To 6
            .Honsyokin(i) = rs("Honsyokin" & i + 1)                                             '' 本賞金
        Next i
        For i = 0 To 4
            .HonsyokinBefore(i) = rs("HonsyokinBefore" & i + 1)                                 '' 変更前本賞金
        Next i
        For i = 0 To 4
            .Fukasyokin(i) = rs("Fukasyokin" & i + 1)                                           '' 付加賞金
        Next i
        For i = 0 To 2
            .FukasyokinBefore(i) = rs("FukasyokinBefore" & i + 1)                               '' 変更前付加賞金
        Next i
        .HassoTime = rs("HassoTime")                                                            '' 発走時刻
        .HassoTimeBefore = rs("HassoTimeBefore")                                                '' 変更前発走時刻
        .TorokuTosu = rs("TorokuTosu")                                                          '' 登録頭数
        .SyussoTosu = rs("SyussoTosu")                                                          '' 出走頭数
        .NyusenTosu = rs("NyusenTosu")                                                          '' 入線頭数
        With .TenkoBaba
            .TenkoCD = rs("TenkoCD")                                                            '' 天候コード
            .SibaBabaCD = rs("SibaBabaCD")                                                      '' 芝馬場状態コード
            .DirtBabaCD = rs("DirtBabaCD")                                                      '' ダート馬場状態コード
        End With ' TenkoBaba
        For i = 0 To 24
            .LapTime(i) = rs("LapTime" & i + 1)                                                 '' ラップタイム
        Next i
        .SyogaiMileTime = rs("SyogaiMileTime")                                                  '' 障害マイルタイム
        .HaronTimeS3 = rs("HaronTimeS3")                                                        '' 前３ハロンタイム
        .HaronTimeS4 = rs("HaronTimeS4")                                                        '' 前４ハロンタイム
        .HaronTimeL3 = rs("HaronTimeL3")                                                        '' 後３ハロンタイム
        .HaronTimeL4 = rs("HaronTimeL4")                                                        '' 後４ハロンタイム
        For i = 0 To 3
            With .CornerInfo(i)
                .Corner = rs("Corner" & i + 1)                                                  '' コーナー
                .Syukaisu = rs("Syukaisu" & i + 1)                                              '' 周回数
                .Jyuni = rs("Jyuni" & i + 1)                                                    '' 各通過順位
            End With ' CornerInfo
        Next i
        .RecordUpKubun = rs("RecordUpKubun")                                                    '' レコード更新区分
        .CRLF = vbCrLf 'CRLF
    End With
End Sub


'
'   機能: 構造体にデータをセットする－AVレコード
'
'   備考: なし
'
Public Sub SetDataFromRS_AV(ByRef rs As ADODB.Recordset, ByRef buf As JV_AV_INFO)
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer

    With buf
        With .head
            .RecordSpec = rs("RecordSpec")                                                      '' レコード種別
            .DataKubun = rs("DataKubun")                                                       '' データ区分
            With .MakeDate
                .Year = Mid$(rs("MakeDate"), 1, 4)                                               '' 年
                .Month = Mid$(rs("MakeDate"), 5, 2)                                              '' 月
                .Day = Mid$(rs("MakeDate"), 7, 2)                                                '' 日
            End With ' MakeDate
        End With ' head
        With .id
            .Year = rs("Year")                                                                 '' 開催年
            .MonthDay = rs("MonthDay")                                                         '' 開催月日
            .JyoCD = rs("JyoCD")                                                               '' 競馬場コード
            .Kaiji = rs("Kaiji")                                                               '' 開催回[第N回]
            .Nichiji = rs("Nichiji")                                                           '' 開催日目[N日目]
            .RaceNum = rs("RaceNum")                                                           '' レース番号
        End With ' id
        With .HappyoTime
            .Month = Mid$(rs("HappyoTime"), 1, 2)                                               '' 月
            .Day = Mid$(rs("HappyoTime"), 3, 2)                                                 '' 日
            .Hour = Mid$(rs("HappyoTime"), 5, 2)                                                '' 時
            .Minute = Mid$(rs("HappyoTime"), 7, 2)                                              '' 分
        End With ' HappyoTime
        .Umaban = rs("Umaban")                                                                 '' 馬番
        .BAMEI = rs("Bamei")                                                                   '' 馬名
        .JiyuKubun = rs("JiyuKubun")                                                           '' 事由区分
        .CRLF = vbCrLf 'CRLF
    End With ' buf
End Sub


'
'   機能: 構造体にデータをセットする－BNレコード
'
'   備考: なし
'
Public Sub SetDataFromRS_BN(ByRef rs As ADODB.Recordset, ByRef buf As JV_BN_BANUSI)
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer

    With buf
        With .head
            .RecordSpec = rs("RecordSpec")                                                     '' レコード種別
            .DataKubun = rs("DataKubun")                                                       '' データ区分
            With .MakeDate
                .Year = Mid$(rs("MakeDate"), 1, 4)                                               '' 年
                .Month = Mid$(rs("MakeDate"), 5, 2)                                              '' 月
                .Day = Mid$(rs("MakeDate"), 7, 2)                                                '' 日
            End With ' MakeDate
        End With ' head
        .BanusiCode = rs("BanusiCode")
        .BanusiName_Co = rs("BanusiName_Co")                    '' 馬主コード
        .BanusiName = rs("BanusiName")                                                         '' 馬主名
        .BanusiNameKana = rs("BanusiNameKana")                                                 '' 馬主名半角カナ
        .BanusiNameEng = rs("BanusiNameEng")                                                   '' 馬主名欧字
        .Fukusyoku = rs("Fukusyoku")                                                           '' 服色標示
        With .HonRuikei(0)
            .SetYear = rs("H_SetYear")                                                         '' 設定年
            .HonSyokinTotal = rs("H_HonSyokinTotal")                                           '' 本賞金合計
            .Fukasyokin = rs("H_Fukasyokin")                                                   '' 付加賞金合計
            For j = 0 To 5
                .Chakukaisu(j) = rs("H_Chakukaisu" & j + 1)                               '' 着回数
            Next j
        End With ' HonRuikei(0)
        With .HonRuikei(1)
            .SetYear = rs("R_SetYear")                                                         '' 設定年
            .HonSyokinTotal = rs("R_HonSyokinTotal")                                           '' 本賞金合計
            .Fukasyokin = rs("R_Fukasyokin")                                                   '' 付加賞金合計
            For j = 0 To 5
                .Chakukaisu(j) = rs("R_Chakukaisu" & j + 1)                               '' 着回数
            Next j
        End With ' HonRuikei(1)
        .CRLF = vbCrLf 'CRLF
    End With ' buf
End Sub


'
'   機能: 構造体にデータをセットする－BRレコード
'
'   備考: なし
'
Public Sub SetDataFromRS_BR(ByRef rs As ADODB.Recordset, ByRef buf As JV_BR_BREEDER)
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer

    With buf
        With .head
            .RecordSpec = rs("RecordSpec")                                                     '' レコード種別
            .DataKubun = rs("DataKubun")                                                       '' データ区分
            With .MakeDate
                .Year = Mid$(rs("MakeDate"), 1, 4)                                               '' 年
                .Month = Mid$(rs("MakeDate"), 5, 2)                                              '' 月
                .Day = Mid$(rs("MakeDate"), 7, 2)                                                '' 日
            End With ' MakeDate
        End With ' head
        .BreederCode = rs("BreederCode")                                                       '' 生産者コード
        .BreederName_Co = rs("BreederName_Co")
        .BreederName = rs("BreederName")                                                       '' 生産者名
        .BreederNameKana = rs("BreederNameKana")                                               '' 生産者名半角カナ
        .BreederNameEng = rs("BreederNameEng")                                                 '' 生産者名欧字
        .Address = rs("Address")                                                               '' 生産者住所自治省名
        With .HonRuikei(0)
            .SetYear = rs("H_SetYear")                                                         '' 設定年
            .HonSyokinTotal = rs("H_HonSyokinTotal")                                           '' 本賞金合計
            .Fukasyokin = rs("H_Fukasyokin")                                                   '' 付加賞金合計
            For j = 0 To 5
                .Chakukaisu(j) = rs("H_Chakukaisu" & j + 1)                               '' 着回数
            Next j
        End With ' HonRuikei(0)
        With .HonRuikei(1)
            .SetYear = rs("R_SetYear")                                                         '' 設定年
            .HonSyokinTotal = rs("R_HonSyokinTotal")                                           '' 本賞金合計
            .Fukasyokin = rs("R_Fukasyokin")                                                   '' 付加賞金合計
            For j = 0 To 5
                .Chakukaisu(j) = rs("R_Chakukaisu" & j + 1)                               '' 着回数
            Next j
        End With ' HonRuikei(1)
        .CRLF = vbCrLf 'CRLF
    End With ' buf
End Sub


'
'   機能: 構造体にデータをセットする－CHレコード
'
'   備考: なし
'
Public Sub SetDataFromRS_CH(ByRef rs As ADODB.Recordset, ByRef buf As JV_CH_CHOKYOSI)
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer

    With buf
        With .head
            .RecordSpec = rs("RecordSpec")                                                     '' レコード種別
            .DataKubun = rs("DataKubun")                                                       '' データ区分
            With .MakeDate
                .Year = Mid$(rs("MakeDate"), 1, 4)                                               '' 年
                .Month = Mid$(rs("MakeDate"), 5, 2)                                              '' 月
                .Day = Mid$(rs("MakeDate"), 7, 2)                                                '' 日
            End With ' MakeDate
        End With ' head
        .ChokyosiCode = rs("ChokyosiCode")                                                     '' 調教師コード
        .DelKubun = rs("DelKubun")                                                             '' 調教師抹消区分
        With .IssueDate
                .Year = Mid$(rs("IssueDate"), 1, 4)                                               '' 年
                .Month = Mid$(rs("IssueDate"), 5, 2)                                              '' 月
                .Day = Mid$(rs("IssueDate"), 7, 2)                                                '' 日
        End With ' IssueDate
        With .DelDate
                .Year = Mid$(rs("DelDate"), 1, 4)                                               '' 年
                .Month = Mid$(rs("DelDate"), 5, 2)                                              '' 月
                .Day = Mid$(rs("DelDate"), 7, 2)                                                '' 日
        End With ' DelDate
        With .BirthDate
                .Year = Mid$(rs("BirthDate"), 1, 4)                                               '' 年
                .Month = Mid$(rs("BirthDate"), 5, 2)                                              '' 月
                .Day = Mid$(rs("BirthDate"), 7, 2)                                                '' 日
        End With ' BirthDate
        .ChokyosiName = rs("ChokyosiName")                                                     '' 調教師名漢字
        .ChokyosiNameKana = rs("ChokyosiNameKana")                                             '' 調教師名半角カナ
        .ChokyosiRyakusyo = rs("ChokyosiRyakusyo")                                             '' 調教師名略称
        .ChokyosiNameEng = rs("ChokyosiNameEng")                                               '' 調教師名欧字
        .SexCD = rs("SexCD")                                                                   '' 性別区分
        .TozaiCD = rs("TozaiCD")                                                               '' 調教師東西所属コード
        .Syotai = rs("Syotai")                                                                 '' 招待地域名
        For i = 0 To 2
            With .SaikinJyusyo(i)
                With .SaikinJyusyoid
                    .Year = Mid$(rs("SaikinJyusyo" & i + 1 & "SaikinJyusyoid"), 1, 4)
                    .MonthDay = Mid$(rs("SaikinJyusyo" & i + 1 & "SaikinJyusyoid"), 5, 4)
                    .JyoCD = Mid$(rs("SaikinJyusyo" & i + 1 & "SaikinJyusyoid"), 9, 2)
                    .Kaiji = Mid$(rs("SaikinJyusyo" & i + 1 & "SaikinJyusyoid"), 11, 2)
                    .Nichiji = Mid$(rs("SaikinJyusyo" & i + 1 & "SaikinJyusyoid"), 13, 2)
                    .RaceNum = Mid$(rs("SaikinJyusyo" & i + 1 & "SaikinJyusyoid"), 15, 2)
                End With ' SaikinJyusyoid
                .Hondai = rs("SaikinJyusyo" & i + 1 & "Hondai")                                '' 競走名本題
                .Ryakusyo10 = rs("SaikinJyusyo" & i + 1 & "Ryakusyo10")                        '' 競走名略称10字
                .Ryakusyo6 = rs("SaikinJyusyo" & i + 1 & "Ryakusyo6")                          '' 競走名略称6字
                .Ryakusyo3 = rs("SaikinJyusyo" & i + 1 & "Ryakusyo3")                          '' 競走名略称3字
                .GradeCD = rs("SaikinJyusyo" & i + 1 & "GradeCD")                              '' グレードコード
                .SyussoTosu = rs("SaikinJyusyo" & i + 1 & "SyussoTosu")                        '' 出走頭数
                .KettoNum = rs("SaikinJyusyo" & i + 1 & "KettoNum")                            '' 血統登録番号
                .BAMEI = rs("SaikinJyusyo" & i + 1 & "Bamei")                                  '' 馬名
            End With ' SaikinJyusyo
        Next i

    End With ' buf
End Sub
        

'
'   機能: 構造体にデータをセットする－CH_SEISEKIレコード
'
'   備考: なし
'
Public Sub SetDataFromRS_CH_SEISEKI(ByRef rs As ADODB.Recordset, ByRef buf As JV_CH_CHOKYOSI)
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer

    With buf
        .ChokyosiCode = rs("ChokyosiCode")                                                 '' 調教師コード
        For i = 0 To 2
            rs.Filter = "Num='" & i & "'"  ' ADO Recordset Function
            With .HonZenRuikei(i)
                .SetYear = rs("SetYear")                                                       '' 設定年
                .HonSyokinHeichi = rs("HonSyokinHeichi")                                       '' 平地本賞金合計
                .HonSyokinSyogai = rs("HonSyokinSyogai")                                       '' 障害本賞金合計
                .FukaSyokinHeichi = rs("FukaSyokinHeichi")                                     '' 平地付加賞金合計
                .FukaSyokinSyogai = rs("FukaSyokinSyogai")                                     '' 障害付加賞金合計
                With .ChakuKaisuHeichi
                    For k = 0 To 5
                        .Chakukaisu(k) = rs("HeichiChakukaisu" & k + 1)
                    Next k
                End With ' ChakuKaisuHeichi
                With .ChakuKaisuSyogai
                    For k = 0 To 5
                        .Chakukaisu(k) = rs("SyogaiChakukaisu" & k + 1)
                    Next k
                End With ' ChakuKaisuSyogai
                For j = 0 To 5
                    With .ChakuKaisuKyori(j)
                        For k = 0 To 5
                            .Chakukaisu(k) = rs("Kyori" & j + 1 & "Chakukaisu" & k + 1)
                        Next k
                    End With ' ChakuKaisuKyori
                Next j
            End With ' HonZenRuikei
            With .HonZenRuikei(i)
                For j = 0 To 19
                    With .ChakuKaisuJyo(j)
                        For k = 0 To 5
                            .Chakukaisu(k) = rs("Jyo" & j + 1 & "Chakukaisu" & k + 1)
                        Next k
                    End With ' ChakuKaisuJyo
                Next j
            End With ' HonZenRuikei
        Next i
        .CRLF = vbCrLf 'CRLF
    End With ' buf
End Sub


'
'   機能: 構造体にデータをセットする－DMレコード
'
'   備考: なし
'
Public Sub SetDataFromRS_DM(ByRef rs As ADODB.Recordset, ByRef buf As JV_DM_INFO)
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer

    With buf
        With .head
            .RecordSpec = rs("RecordSpec")                                                     '' レコード種別
            .DataKubun = rs("DataKubun")                                                       '' データ区分
            With .MakeDate
                .Year = Mid$(rs("MakeDate"), 1, 4)                                               '' 年
                .Month = Mid$(rs("MakeDate"), 5, 2)                                              '' 月
                .Day = Mid$(rs("MakeDate"), 7, 2)                                                '' 日
            End With ' MakeDate
        End With ' head
        With .id
            .Year = rs("Year")                                                                 '' 開催年
            .MonthDay = rs("MonthDay")                                                         '' 開催月日
            .JyoCD = rs("JyoCD")                                                               '' 競馬場コード
            .Kaiji = rs("Kaiji")                                                               '' 開催回[第N回]
            .Nichiji = rs("Nichiji")                                                           '' 開催日目[N日目]
            .RaceNum = rs("RaceNum")                                                           '' レース番号
        End With ' id
        With .MakeHM
            .Hour = Mid$(rs("MakeHM"), 1, 2)                                                               '' 時
            .Minute = Mid$(rs("MakeHM"), 3, 2)                                                           '' 分
        End With ' MakeHM
        For i = 0 To 17
            With .DMInfo(i)
                .Umaban = rs("Umaban" & i + 1)                                                       '' 馬番
                .DMTime = rs("DMTime" & i + 1)                                                        '' 予想走破タイム
                .DMGosaP = rs("DMGosaP" & i + 1)                                                      '' 予想誤差(信頼度)＋
                .DMGosaM = rs("DMGosaM" & i + 1)                                                      '' 予想誤差(信頼度)－
            End With ' DMInfo
        Next i
        .CRLF = vbCrLf 'CRLF
    End With ' buf
End Sub


'
'   機能: 構造体にデータをセットする－H1レコード
'
'   備考: なし
'
Public Sub SetDataFromRS_H1_(ByRef rs As ADODB.Recordset, ByRef buf As JV_H1_HYOSU_ZENKAKE)
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer

    With buf
        With .head
            .RecordSpec = rs("RecordSpec")                                                     '' レコード種別
            .DataKubun = rs("DataKubun")                                                       '' データ区分
            With .MakeDate
                .Year = Mid$(rs("MakeDate"), 1, 4)                                               '' 年
                .Month = Mid$(rs("MakeDate"), 5, 2)                                              '' 月
                .Day = Mid$(rs("MakeDate"), 7, 2)                                                '' 日
            End With ' MakeDate
        End With ' head
        With .id
            .Year = rs("Year")                                                                 '' 開催年
            .MonthDay = rs("MonthDay")                                                         '' 開催月日
            .JyoCD = rs("JyoCD")                                                               '' 競馬場コード
            .Kaiji = rs("Kaiji")                                                               '' 開催回[第N回]
            .Nichiji = rs("Nichiji")                                                           '' 開催日目[N日目]
            .RaceNum = rs("RaceNum")                                                           '' レース番号
        End With ' id
        .TorokuTosu = rs("TorokuTosu")                                                         '' 登録頭数
        .SyussoTosu = rs("SyussoTosu")                                                         '' 出走頭数
        For i = 0 To 6
            .HatubaiFlag(i) = rs("HatubaiFlag" & i + 1)                                   '' 発売フラグ
        Next i
        .FukuChakuBaraiKey = rs("FukuChakuBaraiKey")                                           '' 複勝着払キー
        For i = 0 To 27
            .HenkanUma(i) = rs("HenkanUma" & i + 1)                                       '' 返還馬番情報(馬番01～28)
        Next i
        For i = 0 To 7
            .HenkanWaku(i) = rs("HenkanWaku" & i + 1)                                     '' 返還枠番情報(枠番1～8)
        Next i
        For i = 0 To 7
            .HenkanDoWaku(i) = rs("HenkanDoWaku" & i + 1)                                 '' 返還同枠情報(枠番1～8)
        Next i
        For i = 0 To 13
            .HyoTotal(i) = rs("HyoTotal" & i + 1)                                         '' 票数合計
        Next i
        
        For i = 0 To 27
            With .HyoTansyo(i)
                .Umaban = rs("Umaban")                                                         '' 馬番
                .Hyo = rs("Hyo")                                                               '' 票数
                .Ninki = rs("Ninki")                                                           '' 人気
            End With ' HyoTansyo
            With .HyoFukusyo(i)
                .Hyo = rs("Hyo")                                                               '' 票数
                .Ninki = rs("Ninki")                                                           '' 人気
            End With ' HyoFukusyo
        Next i
    
        For i = 0 To 35
            With .HyoWakuren(i)
                .Umaban = rs("Kumi")                                                         '' 馬番
                .Hyo = rs("Hyo")                                                               '' 票数
                .Ninki = rs("Ninki")                                                           '' 人気
            End With ' HyoWakuren
        Next i
        
        For i = 0 To 152
            With .HyoUmaren(i)
                .Kumi = rs("Kumi")                                                         '' 馬番
                .Hyo = rs("Hyo")                                                               '' 票数
                .Ninki = rs("Ninki")                                                           '' 人気
            End With ' HyoUmaren
            With .HyoWide(i)
                .Hyo = rs("Hyo")                                                               '' 票数
                .Ninki = rs("Ninki")                                                           '' 人気
            End With ' HyoWide
        Next i
    
        For i = 0 To 305
            With .HyoUmatan(i)
                .Kumi = rs("Kumi")                                                        '' 馬番
                .Hyo = rs("Hyo")                                                               '' 票数
                .Ninki = rs("Ninki")                                                           '' 人気
            End With ' HyoUmatan
        Next i
    
        For i = 0 To 815
            With .HyoSanrenpuku(i)
                .Kumi = rs("Kumi")                                                         '' 馬番
                .Hyo = rs("Hyo")                                                               '' 票数
                .Ninki = rs("Ninki")                                                           '' 人気
            End With ' HyoSanrenpuku
        Next i

    End With ' buf
End Sub


'
'   機能: 構造体にデータをセットする－H1レコード(HYOSU_TANPUKU)
'
'   備考: なし
'
Public Sub SetDataFromRS_H1_TANPUKU(ByRef rs As ADODB.Recordset, ByRef buf As JV_H1_HYOSU_ZENKAKE)
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer

    With buf
        With .head
            .RecordSpec = rs("RecordSpec")                                                     '' レコード種別
            .DataKubun = rs("DataKubun")                                                       '' データ区分
            With .MakeDate
                .Year = Mid$(rs("MakeDate"), 1, 4)                                               '' 年
                .Month = Mid$(rs("MakeDate"), 5, 2)                                              '' 月
                .Day = Mid$(rs("MakeDate"), 7, 2)                                                '' 日
            End With ' MakeDate
        End With ' head
        With .id
            .Year = rs("Year")                                                                 '' 開催年
            .MonthDay = rs("MonthDay")                                                         '' 開催月日
            .JyoCD = rs("JyoCD")                                                               '' 競馬場コード
            .Kaiji = rs("Kaiji")                                                               '' 開催回[第N回]
            .Nichiji = rs("Nichiji")                                                           '' 開催日目[N日目]
            .RaceNum = rs("RaceNum")                                                           '' レース番号
        End With ' id
        .TorokuTosu = rs("TorokuTosu")                                                         '' 登録頭数
        .SyussoTosu = rs("SyussoTosu")                                                         '' 出走頭数
        For i = 0 To 6
            .HatubaiFlag(i) = rs("HatubaiFlag" & i + 1)                                   '' 発売フラグ
        Next i
        .FukuChakuBaraiKey = rs("FukuChakuBaraiKey")                                           '' 複勝着払キー
        For i = 0 To 27
            .HenkanUma(i) = rs("HenkanUma" & i + 1)                                       '' 返還馬番情報(馬番01～28)
        Next i
        For i = 0 To 7
            .HenkanWaku(i) = rs("HenkanWaku" & i + 1)                                     '' 返還枠番情報(枠番1～8)
        Next i
        For i = 0 To 7
            .HenkanDoWaku(i) = rs("HenkanDoWaku" & i + 1)                                 '' 返還同枠情報(枠番1～8)
        Next i
        For i = 0 To 13
            .HyoTotal(i) = rs("HyoTotal" & i + 1)                                         '' 票数合計
        Next i
        
        For i = 0 To 27
            With .HyoTansyo(i)
                .Umaban = rs("Umaban")                                                         '' 馬番
                .Hyo = rs("Hyo")                                                               '' 票数
                .Ninki = rs("Ninki")                                                           '' 人気
            End With ' HyoTansyo
            With .HyoFukusyo(i)
                .Hyo = rs("Hyo")                                                               '' 票数
                .Ninki = rs("Ninki")                                                           '' 人気
            End With ' HyoFukusyo
        Next i
    
        For i = 0 To 35
            With .HyoWakuren(i)
                .Umaban = rs("Kumi")                                                         '' 馬番
                .Hyo = rs("Hyo")                                                               '' 票数
                .Ninki = rs("Ninki")                                                           '' 人気
            End With ' HyoWakuren
        Next i
        
        For i = 0 To 152
            With .HyoUmaren(i)
                .Kumi = rs("Kumi")                                                         '' 馬番
                .Hyo = rs("Hyo")                                                               '' 票数
                .Ninki = rs("Ninki")                                                           '' 人気
            End With ' HyoUmaren
            With .HyoWide(i)
                .Hyo = rs("Hyo")                                                               '' 票数
                .Ninki = rs("Ninki")                                                           '' 人気
            End With ' HyoWide
        Next i
    
        For i = 0 To 305
            With .HyoUmatan(i)
                .Kumi = rs("Kumi")                                                        '' 馬番
                .Hyo = rs("Hyo")                                                               '' 票数
                .Ninki = rs("Ninki")                                                           '' 人気
            End With ' HyoUmatan
        Next i
    
        For i = 0 To 815
            With .HyoSanrenpuku(i)
                .Kumi = rs("Kumi")                                                         '' 馬番
                .Hyo = rs("Hyo")                                                               '' 票数
                .Ninki = rs("Ninki")                                                           '' 人気
            End With ' HyoSanrenpuku
        Next i

    End With ' buf
End Sub


'
'   機能: 構造体にデータをセットする－H1レコード(WAKU)
'
'   備考: なし
'
Public Sub SetDataFromRS_H1_WAKU(ByRef rs As ADODB.Recordset, ByRef buf As JV_H1_HYOSU_ZENKAKE)
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer

    With buf
        With .head
            .RecordSpec = rs("RecordSpec")                                                     '' レコード種別
            .DataKubun = rs("DataKubun")                                                       '' データ区分
            With .MakeDate
                .Year = Mid$(rs("MakeDate"), 1, 4)                                               '' 年
                .Month = Mid$(rs("MakeDate"), 5, 2)                                              '' 月
                .Day = Mid$(rs("MakeDate"), 7, 2)                                                '' 日
            End With ' MakeDate
        End With ' head
        With .id
            .Year = rs("Year")                                                                 '' 開催年
            .MonthDay = rs("MonthDay")                                                         '' 開催月日
            .JyoCD = rs("JyoCD")                                                               '' 競馬場コード
            .Kaiji = rs("Kaiji")                                                               '' 開催回[第N回]
            .Nichiji = rs("Nichiji")                                                           '' 開催日目[N日目]
            .RaceNum = rs("RaceNum")                                                           '' レース番号
        End With ' id
        .TorokuTosu = rs("TorokuTosu")                                                         '' 登録頭数
        .SyussoTosu = rs("SyussoTosu")                                                         '' 出走頭数
        For i = 0 To 6
            .HatubaiFlag(i) = rs("HatubaiFlag" & i + 1)                                   '' 発売フラグ
        Next i
        .FukuChakuBaraiKey = rs("FukuChakuBaraiKey")                                           '' 複勝着払キー
        For i = 0 To 27
            .HenkanUma(i) = rs("HenkanUma" & i + 1)                                       '' 返還馬番情報(馬番01～28)
        Next i
        For i = 0 To 7
            .HenkanWaku(i) = rs("HenkanWaku" & i + 1)                                     '' 返還枠番情報(枠番1～8)
        Next i
        For i = 0 To 7
            .HenkanDoWaku(i) = rs("HenkanDoWaku" & i + 1)                                 '' 返還同枠情報(枠番1～8)
        Next i
        For i = 0 To 13
            .HyoTotal(i) = rs("HyoTotal" & i + 1)                                         '' 票数合計
        Next i
        
        For i = 0 To 27
            With .HyoTansyo(i)
                .Umaban = rs("Umaban")                                                         '' 馬番
                .Hyo = rs("Hyo")                                                               '' 票数
                .Ninki = rs("Ninki")                                                           '' 人気
            End With ' HyoTansyo
            With .HyoFukusyo(i)
                .Hyo = rs("Hyo")                                                               '' 票数
                .Ninki = rs("Ninki")                                                           '' 人気
            End With ' HyoFukusyo
        Next i
    
        For i = 0 To 35
            With .HyoWakuren(i)
                .Umaban = rs("Kumi")                                                         '' 馬番
                .Hyo = rs("Hyo")                                                               '' 票数
                .Ninki = rs("Ninki")                                                           '' 人気
            End With ' HyoWakuren
        Next i
        
        For i = 0 To 152
            With .HyoUmaren(i)
                .Kumi = rs("Kumi")                                                         '' 馬番
                .Hyo = rs("Hyo")                                                               '' 票数
                .Ninki = rs("Ninki")                                                           '' 人気
            End With ' HyoUmaren
            With .HyoWide(i)
                .Hyo = rs("Hyo")                                                               '' 票数
                .Ninki = rs("Ninki")                                                           '' 人気
            End With ' HyoWide
        Next i
    
        For i = 0 To 305
            With .HyoUmatan(i)
                .Kumi = rs("Kumi")                                                        '' 馬番
                .Hyo = rs("Hyo")                                                               '' 票数
                .Ninki = rs("Ninki")                                                           '' 人気
            End With ' HyoUmatan
        Next i
    
        For i = 0 To 815
            With .HyoSanrenpuku(i)
                .Kumi = rs("Kumi")                                                         '' 馬番
                .Hyo = rs("Hyo")                                                               '' 票数
                .Ninki = rs("Ninki")                                                           '' 人気
            End With ' HyoSanrenpuku
        Next i

    End With ' buf
End Sub


'
'   機能: 構造体にデータをセットする－H1レコード(UMAREN)
'
'   備考: なし
'
Public Sub SetDataFromRS_H1_UMAREN(ByRef rs As ADODB.Recordset, ByRef buf As JV_H1_HYOSU_ZENKAKE)
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer

    With buf
        With .head
            .RecordSpec = rs("RecordSpec")                                                     '' レコード種別
            .DataKubun = rs("DataKubun")                                                       '' データ区分
            With .MakeDate
                .Year = Mid$(rs("MakeDate"), 1, 4)                                               '' 年
                .Month = Mid$(rs("MakeDate"), 5, 2)                                              '' 月
                .Day = Mid$(rs("MakeDate"), 7, 2)                                                '' 日
            End With ' MakeDate
        End With ' head
        With .id
            .Year = rs("Year")                                                                 '' 開催年
            .MonthDay = rs("MonthDay")                                                         '' 開催月日
            .JyoCD = rs("JyoCD")                                                               '' 競馬場コード
            .Kaiji = rs("Kaiji")                                                               '' 開催回[第N回]
            .Nichiji = rs("Nichiji")                                                           '' 開催日目[N日目]
            .RaceNum = rs("RaceNum")                                                           '' レース番号
        End With ' id
        .TorokuTosu = rs("TorokuTosu")                                                         '' 登録頭数
        .SyussoTosu = rs("SyussoTosu")                                                         '' 出走頭数
        For i = 0 To 6
            .HatubaiFlag(i) = rs("HatubaiFlag" & i + 1)                                   '' 発売フラグ
        Next i
        .FukuChakuBaraiKey = rs("FukuChakuBaraiKey")                                           '' 複勝着払キー
        For i = 0 To 27
            .HenkanUma(i) = rs("HenkanUma" & i + 1)                                       '' 返還馬番情報(馬番01～28)
        Next i
        For i = 0 To 7
            .HenkanWaku(i) = rs("HenkanWaku" & i + 1)                                     '' 返還枠番情報(枠番1～8)
        Next i
        For i = 0 To 7
            .HenkanDoWaku(i) = rs("HenkanDoWaku" & i + 1)                                 '' 返還同枠情報(枠番1～8)
        Next i
        For i = 0 To 13
            .HyoTotal(i) = rs("HyoTotal" & i + 1)                                         '' 票数合計
        Next i
        
        For i = 0 To 27
            With .HyoTansyo(i)
                .Umaban = rs("Umaban")                                                         '' 馬番
                .Hyo = rs("Hyo")                                                               '' 票数
                .Ninki = rs("Ninki")                                                           '' 人気
            End With ' HyoTansyo
            With .HyoFukusyo(i)
                .Hyo = rs("Hyo")                                                               '' 票数
                .Ninki = rs("Ninki")                                                           '' 人気
            End With ' HyoFukusyo
        Next i
    
        For i = 0 To 35
            With .HyoWakuren(i)
                .Umaban = rs("Kumi")                                                         '' 馬番
                .Hyo = rs("Hyo")                                                               '' 票数
                .Ninki = rs("Ninki")                                                           '' 人気
            End With ' HyoWakuren
        Next i
        
        For i = 0 To 152
            With .HyoUmaren(i)
                .Kumi = rs("Kumi")                                                         '' 馬番
                .Hyo = rs("Hyo")                                                               '' 票数
                .Ninki = rs("Ninki")                                                           '' 人気
            End With ' HyoUmaren
            With .HyoWide(i)
                .Hyo = rs("Hyo")                                                               '' 票数
                .Ninki = rs("Ninki")                                                           '' 人気
            End With ' HyoWide
        Next i
    
        For i = 0 To 305
            With .HyoUmatan(i)
                .Kumi = rs("Kumi")                                                        '' 馬番
                .Hyo = rs("Hyo")                                                               '' 票数
                .Ninki = rs("Ninki")                                                           '' 人気
            End With ' HyoUmatan
        Next i
    
        For i = 0 To 815
            With .HyoSanrenpuku(i)
                .Kumi = rs("Kumi")                                                         '' 馬番
                .Hyo = rs("Hyo")                                                               '' 票数
                .Ninki = rs("Ninki")                                                           '' 人気
            End With ' HyoSanrenpuku
        Next i

    End With ' buf
End Sub


'
'   機能: 構造体にデータをセットする－H1レコード(WIDE)
'
'   備考: なし
'
Public Sub SetDataFromRS_H1_WIDE(ByRef rs As ADODB.Recordset, ByRef buf As JV_H1_HYOSU_ZENKAKE)
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer

    With buf
        With .head
            .RecordSpec = rs("RecordSpec")                                                     '' レコード種別
            .DataKubun = rs("DataKubun")                                                       '' データ区分
            With .MakeDate
                .Year = Mid$(rs("MakeDate"), 1, 4)                                               '' 年
                .Month = Mid$(rs("MakeDate"), 5, 2)                                              '' 月
                .Day = Mid$(rs("MakeDate"), 7, 2)                                                '' 日
            End With ' MakeDate
        End With ' head
        With .id
            .Year = rs("Year")                                                                 '' 開催年
            .MonthDay = rs("MonthDay")                                                         '' 開催月日
            .JyoCD = rs("JyoCD")                                                               '' 競馬場コード
            .Kaiji = rs("Kaiji")                                                               '' 開催回[第N回]
            .Nichiji = rs("Nichiji")                                                           '' 開催日目[N日目]
            .RaceNum = rs("RaceNum")                                                           '' レース番号
        End With ' id
        .TorokuTosu = rs("TorokuTosu")                                                         '' 登録頭数
        .SyussoTosu = rs("SyussoTosu")                                                         '' 出走頭数
        For i = 0 To 6
            .HatubaiFlag(i) = rs("HatubaiFlag" & i + 1)                                   '' 発売フラグ
        Next i
        .FukuChakuBaraiKey = rs("FukuChakuBaraiKey")                                           '' 複勝着払キー
        For i = 0 To 27
            .HenkanUma(i) = rs("HenkanUma" & i + 1)                                       '' 返還馬番情報(馬番01～28)
        Next i
        For i = 0 To 7
            .HenkanWaku(i) = rs("HenkanWaku" & i + 1)                                     '' 返還枠番情報(枠番1～8)
        Next i
        For i = 0 To 7
            .HenkanDoWaku(i) = rs("HenkanDoWaku" & i + 1)                                 '' 返還同枠情報(枠番1～8)
        Next i
        For i = 0 To 13
            .HyoTotal(i) = rs("HyoTotal" & i + 1)                                         '' 票数合計
        Next i
        
        For i = 0 To 27
            With .HyoTansyo(i)
                .Umaban = rs("Umaban")                                                         '' 馬番
                .Hyo = rs("Hyo")                                                               '' 票数
                .Ninki = rs("Ninki")                                                           '' 人気
            End With ' HyoTansyo
            With .HyoFukusyo(i)
                .Hyo = rs("Hyo")                                                               '' 票数
                .Ninki = rs("Ninki")                                                           '' 人気
            End With ' HyoFukusyo
        Next i
    
        For i = 0 To 35
            With .HyoWakuren(i)
                .Umaban = rs("Kumi")                                                         '' 馬番
                .Hyo = rs("Hyo")                                                               '' 票数
                .Ninki = rs("Ninki")                                                           '' 人気
            End With ' HyoWakuren
        Next i
        
        For i = 0 To 152
            With .HyoUmaren(i)
                .Kumi = rs("Kumi")                                                         '' 馬番
                .Hyo = rs("Hyo")                                                               '' 票数
                .Ninki = rs("Ninki")                                                           '' 人気
            End With ' HyoUmaren
            With .HyoWide(i)
                .Hyo = rs("Hyo")                                                               '' 票数
                .Ninki = rs("Ninki")                                                           '' 人気
            End With ' HyoWide
        Next i
    
        For i = 0 To 305
            With .HyoUmatan(i)
                .Kumi = rs("Kumi")                                                        '' 馬番
                .Hyo = rs("Hyo")                                                               '' 票数
                .Ninki = rs("Ninki")                                                           '' 人気
            End With ' HyoUmatan
        Next i
    
        For i = 0 To 815
            With .HyoSanrenpuku(i)
                .Kumi = rs("Kumi")                                                         '' 馬番
                .Hyo = rs("Hyo")                                                               '' 票数
                .Ninki = rs("Ninki")                                                           '' 人気
            End With ' HyoSanrenpuku
        Next i

    End With ' buf
End Sub


'
'   機能: 構造体にデータをセットする－H1レコード(UMATAN)
'
'   備考: なし
'
Public Sub SetDataFromRS_H1_UMATAN(ByRef rs As ADODB.Recordset, ByRef buf As JV_H1_HYOSU_ZENKAKE)
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer

    With buf
        With .head
            .RecordSpec = rs("RecordSpec")                                                     '' レコード種別
            .DataKubun = rs("DataKubun")                                                       '' データ区分
            With .MakeDate
                .Year = Mid$(rs("MakeDate"), 1, 4)                                               '' 年
                .Month = Mid$(rs("MakeDate"), 5, 2)                                              '' 月
                .Day = Mid$(rs("MakeDate"), 7, 2)                                                '' 日
            End With ' MakeDate
        End With ' head
        With .id
            .Year = rs("Year")                                                                 '' 開催年
            .MonthDay = rs("MonthDay")                                                         '' 開催月日
            .JyoCD = rs("JyoCD")                                                               '' 競馬場コード
            .Kaiji = rs("Kaiji")                                                               '' 開催回[第N回]
            .Nichiji = rs("Nichiji")                                                           '' 開催日目[N日目]
            .RaceNum = rs("RaceNum")                                                           '' レース番号
        End With ' id
        .TorokuTosu = rs("TorokuTosu")                                                         '' 登録頭数
        .SyussoTosu = rs("SyussoTosu")                                                         '' 出走頭数
        For i = 0 To 6
            .HatubaiFlag(i) = rs("HatubaiFlag" & i + 1)                                   '' 発売フラグ
        Next i
        .FukuChakuBaraiKey = rs("FukuChakuBaraiKey")                                           '' 複勝着払キー
        For i = 0 To 27
            .HenkanUma(i) = rs("HenkanUma" & i + 1)                                       '' 返還馬番情報(馬番01～28)
        Next i
        For i = 0 To 7
            .HenkanWaku(i) = rs("HenkanWaku" & i + 1)                                     '' 返還枠番情報(枠番1～8)
        Next i
        For i = 0 To 7
            .HenkanDoWaku(i) = rs("HenkanDoWaku" & i + 1)                                 '' 返還同枠情報(枠番1～8)
        Next i
        For i = 0 To 13
            .HyoTotal(i) = rs("HyoTotal" & i + 1)                                         '' 票数合計
        Next i
        
        For i = 0 To 27
            With .HyoTansyo(i)
                .Umaban = rs("Umaban")                                                         '' 馬番
                .Hyo = rs("Hyo")                                                               '' 票数
                .Ninki = rs("Ninki")                                                           '' 人気
            End With ' HyoTansyo
            With .HyoFukusyo(i)
                .Hyo = rs("Hyo")                                                               '' 票数
                .Ninki = rs("Ninki")                                                           '' 人気
            End With ' HyoFukusyo
        Next i
    
        For i = 0 To 35
            With .HyoWakuren(i)
                .Umaban = rs("Kumi")                                                         '' 馬番
                .Hyo = rs("Hyo")                                                               '' 票数
                .Ninki = rs("Ninki")                                                           '' 人気
            End With ' HyoWakuren
        Next i
        
        For i = 0 To 152
            With .HyoUmaren(i)
                .Kumi = rs("Kumi")                                                         '' 馬番
                .Hyo = rs("Hyo")                                                               '' 票数
                .Ninki = rs("Ninki")                                                           '' 人気
            End With ' HyoUmaren
            With .HyoWide(i)
                .Hyo = rs("Hyo")                                                               '' 票数
                .Ninki = rs("Ninki")                                                           '' 人気
            End With ' HyoWide
        Next i
    
        For i = 0 To 305
            With .HyoUmatan(i)
                .Kumi = rs("Kumi")                                                        '' 馬番
                .Hyo = rs("Hyo")                                                               '' 票数
                .Ninki = rs("Ninki")                                                           '' 人気
            End With ' HyoUmatan
        Next i
    
        For i = 0 To 815
            With .HyoSanrenpuku(i)
                .Kumi = rs("Kumi")                                                         '' 馬番
                .Hyo = rs("Hyo")                                                               '' 票数
                .Ninki = rs("Ninki")                                                           '' 人気
            End With ' HyoSanrenpuku
        Next i

    End With ' buf
End Sub


'
'   機能: 構造体にデータをセットする－H1レコード(SANREN)
'
'   備考: なし
'
Public Sub SetDataFromRS_H1_SANREN(ByRef rs As ADODB.Recordset, ByRef buf As JV_H1_HYOSU_ZENKAKE)
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer

    With buf
        With .head
            .RecordSpec = rs("RecordSpec")                                                     '' レコード種別
            .DataKubun = rs("DataKubun")                                                       '' データ区分
            With .MakeDate
                .Year = Mid$(rs("MakeDate"), 1, 4)                                               '' 年
                .Month = Mid$(rs("MakeDate"), 5, 2)                                              '' 月
                .Day = Mid$(rs("MakeDate"), 7, 2)                                                '' 日
            End With ' MakeDate
        End With ' head
        With .id
            .Year = rs("Year")                                                                 '' 開催年
            .MonthDay = rs("MonthDay")                                                         '' 開催月日
            .JyoCD = rs("JyoCD")                                                               '' 競馬場コード
            .Kaiji = rs("Kaiji")                                                               '' 開催回[第N回]
            .Nichiji = rs("Nichiji")                                                           '' 開催日目[N日目]
            .RaceNum = rs("RaceNum")                                                           '' レース番号
        End With ' id
        .TorokuTosu = rs("TorokuTosu")                                                         '' 登録頭数
        .SyussoTosu = rs("SyussoTosu")                                                         '' 出走頭数
        For i = 0 To 6
            .HatubaiFlag(i) = rs("HatubaiFlag" & i + 1)                                   '' 発売フラグ
        Next i
        .FukuChakuBaraiKey = rs("FukuChakuBaraiKey")                                           '' 複勝着払キー
        For i = 0 To 27
            .HenkanUma(i) = rs("HenkanUma" & i + 1)                                       '' 返還馬番情報(馬番01～28)
        Next i
        For i = 0 To 7
            .HenkanWaku(i) = rs("HenkanWaku" & i + 1)                                     '' 返還枠番情報(枠番1～8)
        Next i
        For i = 0 To 7
            .HenkanDoWaku(i) = rs("HenkanDoWaku" & i + 1)                                 '' 返還同枠情報(枠番1～8)
        Next i
        For i = 0 To 13
            .HyoTotal(i) = rs("HyoTotal" & i + 1)                                         '' 票数合計
        Next i
        
        For i = 0 To 27
            With .HyoTansyo(i)
                .Umaban = rs("Umaban")                                                         '' 馬番
                .Hyo = rs("Hyo")                                                               '' 票数
                .Ninki = rs("Ninki")                                                           '' 人気
            End With ' HyoTansyo
            With .HyoFukusyo(i)
                .Hyo = rs("Hyo")                                                               '' 票数
                .Ninki = rs("Ninki")                                                           '' 人気
            End With ' HyoFukusyo
        Next i
    
        For i = 0 To 35
            With .HyoWakuren(i)
                .Umaban = rs("Kumi")                                                         '' 馬番
                .Hyo = rs("Hyo")                                                               '' 票数
                .Ninki = rs("Ninki")                                                           '' 人気
            End With ' HyoWakuren
        Next i
        
        For i = 0 To 152
            With .HyoUmaren(i)
                .Kumi = rs("Kumi")                                                         '' 馬番
                .Hyo = rs("Hyo")                                                               '' 票数
                .Ninki = rs("Ninki")                                                           '' 人気
            End With ' HyoUmaren
            With .HyoWide(i)
                .Hyo = rs("Hyo")                                                               '' 票数
                .Ninki = rs("Ninki")                                                           '' 人気
            End With ' HyoWide
        Next i
    
        For i = 0 To 305
            With .HyoUmatan(i)
                .Kumi = rs("Kumi")                                                        '' 馬番
                .Hyo = rs("Hyo")                                                               '' 票数
                .Ninki = rs("Ninki")                                                           '' 人気
            End With ' HyoUmatan
        Next i
    
        For i = 0 To 815
            With .HyoSanrenpuku(i)
                .Kumi = rs("Kumi")                                                         '' 馬番
                .Hyo = rs("Hyo")                                                               '' 票数
                .Ninki = rs("Ninki")                                                           '' 人気
            End With ' HyoSanrenpuku
        Next i

    End With ' buf
End Sub


'
'   機能: 構造体にデータをセットする－HCレコード
'
'   備考: なし
'
Public Sub SetDataFromRS_HC(ByRef rs As ADODB.Recordset, ByRef buf As JV_HC_HANRO)
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer

    With buf
        With .head
            .RecordSpec = rs("RecordSpec")                              '' レコード種別
            .DataKubun = rs("DataKubun")                                '' データ区分
            With .MakeDate
                .Year = Mid$(rs("MakeDate"), 1, 4)                       '' 年
                .Month = Mid$(rs("MakeDate"), 5, 2)                      '' 月
                .Day = Mid$(rs("MakeDate"), 7, 2)                        '' 日
            End With ' MakeDate
        End With ' head
        .TresenKubun = rs("TresenKubun")                                '' トレセン区分
        With .ChokyoDate
            .Year = Mid$(rs("ChokyoDate"), 1, 4)                         '' 年
            .Month = Mid$(rs("ChokyoDate"), 5, 2)                        '' 月
            .Day = Mid$(rs("ChokyoDate"), 7, 2)                          '' 日
        End With ' ChokyoDate
        .ChokyoTime = rs("ChokyoTime")                                  '' 調教時刻
        .KettoNum = rs("KettoNum")                                      '' 血統登録番号
        .HaronTime4 = rs("HaronTime4")                                  '' 4ハロンタイム合計(800M-0M)
        .LapTime4 = rs("LapTime4")                                      '' ラップタイム(800M-600M)
        .HaronTime3 = rs("HaronTime3")                                  '' 3ハロンタイム合計(600M-0M)
        .LapTime3 = rs("LapTime3")                                      '' ラップタイム(600M-400M)
        .HaronTime2 = rs("HaronTime2")                                  '' 2ハロンタイム合計(400M-0M)
        .LapTime2 = rs("LapTime2")                                      '' ラップタイム(400M-200M)
        .LapTime1 = rs("LapTime1")                                      '' ラップタイム(200M-0M)
        .CRLF = vbCrLf 'CRLF
    End With ' buf
End Sub


'
'   機能: 構造体にデータをセットする－HNレコード
'
'   備考: なし
'
Public Sub SetDataFromRS_HN(ByRef rs As ADODB.Recordset, ByRef buf As JV_HN_HANSYOKU)
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer

    With buf
        With .head
            .RecordSpec = rs("RecordSpec")                              '' レコード種別
            .DataKubun = rs("DataKubun")                                '' データ区分
            With .MakeDate
                .Year = Mid$(rs("MakeDate"), 1, 4)                       '' 年
                .Month = Mid$(rs("MakeDate"), 5, 2)                      '' 月
                .Day = Mid$(rs("MakeDate"), 7, 2)                        '' 日
            End With ' MakeDate
        End With ' head
        .HansyokuNum = rs("HansyokuNum")                                '' 繁殖登録番号
        .Reserved = rs("reserved")                                      '' 予備
        .KettoNum = rs("KettoNum")                                      '' 血統登録番号
        .DelKubun = rs("DelKubun")                                      '' 繁殖馬抹消区分
        .BAMEI = rs("Bamei")                                            '' 馬名
        .BameiKana = rs("BameiKana")                                    '' 馬名半角カナ
        .BameiEng = rs("BameiEng")                                      '' 馬名欧字
        .BirthYear = rs("BirthYear")                                    '' 生年
        .SexCD = rs("SexCD")                                            '' 性別コード
        .HinsyuCD = rs("HinsyuCD")                                      '' 品種コード
        .KeiroCD = rs("KeiroCD")                                        '' 毛色コード
        .HansyokuMochiKubun = rs("HansyokuMochiKubun")                  '' 繁殖馬持込区分
        .ImportYear = rs("ImportYear")                                  '' 輸入年
        .SanchiName = rs("SanchiName")                                  '' 産地名
        .HansyokuFNum = rs("HansyokuFNum")                              '' 父馬繁殖登録番号
        .HansyokuMNum = rs("HansyokuMNum")                              '' 母馬繁殖登録番号
        .CRLF = vbCrLf 'CRLF
    End With ' buf
End Sub


'
'   機能: 構造体にデータをセットする－HRレコード
'
'   備考: なし
'
Public Sub SetDataFromRS_HR(ByRef rs As ADODB.Recordset, ByRef buf As JV_HR_PAY)
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer

    With buf
        With .head
            .RecordSpec = rs("RecordSpec")                              '' レコード種別
            .DataKubun = rs("DataKubun")                                '' データ区分
            With .MakeDate
                .Year = Mid$(rs("MakeDate"), 1, 4)                       '' 年
                .Month = Mid$(rs("MakeDate"), 5, 2)                      '' 月
                .Day = Mid$(rs("MakeDate"), 7, 2)                        '' 日
            End With ' MakeDate
        End With ' head
        With .id
            .Year = rs("Year")                                          '' 開催年
            .MonthDay = rs("MonthDay")                                  '' 開催月日
            .JyoCD = rs("JyoCD")                                        '' 競馬場コード
            .Kaiji = rs("Kaiji")
            .Nichiji = rs("Nichiji")                                    '' 開催日目[N日目]
            .RaceNum = rs("RaceNum")                                    '' レース番号
        End With ' id
        .TorokuTosu = rs("TorokuTosu")                                  '' 登録頭数
        .SyussoTosu = rs("SyussoTosu")                                  '' 出走頭数
        For i = 0 To 8
            .FuseirituFlag(i) = rs("FuseirituFlag" & i + 1)             '' 不成立フラグ
        Next i
        For i = 0 To 8
            .TokubaraiFlag(i) = rs("TokubaraiFlag" & i + 1)             '' 特払フラグ
        Next i
        For i = 0 To 8
            .HenkanFlag(i) = rs("HenkanFlag" & i + 1)                   '' 返還フラグ
        Next i
        For i = 0 To 27
            .HenkanUma(i) = rs("HenkanUma" & i + 1)                     '' 返還馬番情報(馬番01～28)
        Next i
        For i = 0 To 7
            .HenkanWaku(i) = rs("HenkanWaku" & i + 1)                   '' 返還枠番情報(枠番1～8)
        Next i
        For i = 0 To 7
            .HenkanDoWaku(i) = rs("HenkanDoWaku" & i + 1)               '' 返還同枠情報(枠番1～8)
        Next i
        For i = 0 To 2
            With .PayTansyo(i)
                .Umaban = rs("PayTansyoUmaban" & i + 1)                 '' 馬番
                .Pay = rs("PayTansyoPay" & i + 1)                       '' 払戻金
                .Ninki = rs("PayTansyoNinki" & i + 1)                   '' 人気順
            End With ' PayTansyo
        Next i
        For i = 0 To 4
            With .PayFukusyo(i)
                .Umaban = rs("PayFukusyoUmaban" & i + 1)                '' 馬番
                .Pay = rs("PayFukusyoPay" & i + 1)                      '' 払戻金
                .Ninki = rs("PayFukusyoNinki" & i + 1)                  '' 人気順
            End With ' PayFukusyo
        Next i
        For i = 0 To 2
            With .PayWakuren(i)
                .Umaban = rs("PayWakurenKumiban" & i + 1)               '' 馬番
                .Pay = rs("PayWakurenPay" & i + 1)                      '' 払戻金
                .Ninki = rs("PayWakurenNinki" & i + 1)                  '' 人気順
            End With ' PayWakuren
        Next i
        For i = 0 To 2
            With .PayUmaren(i)
                .Kumi = rs("PayUmarenKumiban" & i + 1)                  '' 組番
                .Pay = rs("PayUmarenPay" & i + 1)                       '' 払戻金
                .Ninki = rs("PayUmarenNinki" & i + 1)                   '' 人気順
            End With ' PayUmaren
        Next i
        For i = 0 To 6
            With .PayWide(i)
                .Kumi = rs("PayWideKumiban" & i + 1)                    '' 組番
                .Pay = rs("PayWidePay" & i + 1)                         '' 払戻金
                .Ninki = rs("PayWideNinki" & i + 1)                     '' 人気順
            End With ' PayWide
        Next i
        For i = 0 To 2
            With .PayReserved1(i)
                .Kumi = rs("PayReserved1Kumiban" & i + 1)               '' 組番
                .Pay = rs("PayReserved1Pay" & i + 1)                    '' 払戻金
                .Ninki = rs("PayReserved1Ninki" & i + 1)                '' 人気順
            End With ' PayReserved1
        Next i
        For i = 0 To 5
            With .PayUmatan(i)
                .Kumi = rs("PayUmatanKumiban" & i + 1)                  '' 組番
                .Pay = rs("PayUmatanPay" & i + 1)                       '' 払戻金
                .Ninki = rs("PayUmatanNinki" & i + 1)                   '' 人気順
            End With ' PayUmatan
        Next i
        For i = 0 To 2
            With .PaySanrenpuku(i)
                .Kumi = rs("PaySanrenpukuKumiban" & i + 1)              '' 組番
                .Pay = rs("PaySanrenpukuPay" & i + 1)                   '' 払戻金
                .Ninki = rs("PaySanrenpukuNinki" & i + 1)               '' 人気順
            End With ' PaySanrenpuku
        Next i
        For i = 0 To 5
            With .PaySanrentan(i)
                .Kumi = rs("PaySanrentanKumiban" & i + 1)               '' 組番
                .Pay = rs("PaySanrentanPay" & i + 1)                    '' 払戻金
                .Ninki = rs("PaySanrentanNinki" & i + 1)                '' 人気順
            End With ' PaySanrentan
            
        Next i
        .CRLF = vbCrLf 'CRLF
    End With ' buf
End Sub


'
'   機能: 構造体にデータをセットする－JCレコード
'
'   備考: なし
'
Public Sub SetDataFromRS_JC(ByRef rs As ADODB.Recordset, ByRef buf As JV_JC_INFO)
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer

    With buf
        With .head
            .RecordSpec = rs("RecordSpec")                              '' レコード種別
            .DataKubun = rs("DataKubun")                                '' データ区分
            With .MakeDate
                .Year = Mid$(rs("MakeDate"), 1, 4)                       '' 年
                .Month = Mid$(rs("MakeDate"), 5, 2)                      '' 月
                .Day = Mid$(rs("MakeDate"), 7, 2)                        '' 日
            End With ' MakeDate
        End With ' head
        With .id
            .Year = rs("Year")                                          '' 開催年
            .MonthDay = rs("MonthDay")                                  '' 開催月日
            .JyoCD = rs("JyoCD")                                        '' 競馬場コード
            .Kaiji = rs("Kaiji")                                        '' 開催回[第N回]
            .Nichiji = rs("Nichiji")                                    '' 開催日目[N日目]
            .RaceNum = rs("RaceNum")                                    '' レース番号
        End With ' id
        With .HappyoTime
            .Month = Mid$(rs("HappyoTime"), 1, 2)                        '' 月
            .Day = Mid$(rs("HappyoTime"), 3, 2)                          '' 日
            .Hour = Mid$(rs("HappyoTime"), 5, 2)                         '' 時
            .Minute = Mid$(rs("HappyoTime"), 7, 2)                       '' 分
        End With ' HappyoTime
        .Umaban = rs("Umaban")                                          '' 馬番
        .BAMEI = rs("Bamei")                                            '' 馬名
        With .JCInfoAfter
            .Futan = rs("AtoFutan")                                        '' 負担重量
            .KisyuCode = rs("AtoKisyuCode")                                '' 騎手コード
            .KisyuName = rs("AtoKisyuName")                                '' 騎手名
            .MinaraiCD = rs("AtoMinaraiCD")                                '' 騎手見習コード
        End With ' JCInfoAfter
        With .JCInfoBefore
            .Futan = rs("MaeFutan")                                        '' 負担重量
            .KisyuCode = rs("MaeKisyuCode")                                '' 騎手コード
            .KisyuName = rs("MaeKisyuName")                                '' 騎手名
            .MinaraiCD = rs("MaeMinaraiCD")                                '' 騎手見習コード
        End With ' JCInfoBefore
        .CRLF = vbCrLf 'CRLF
    End With ' buf
End Sub


'
'   機能: 構造体にデータをセットする－KSレコード
'
'   備考: なし
'
Public Sub SetDataFromRS_KS(ByRef rs As ADODB.Recordset, ByRef buf As JV_KS_KISYU)
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer

    With buf
        With .head
            .RecordSpec = rs("RecordSpec")                              '' レコード種別
            .DataKubun = rs("DataKubun")                                '' データ区分
            With .MakeDate
                .Year = Mid$(rs("MakeDate"), 1, 4)                       '' 年
                .Month = Mid$(rs("MakeDate"), 5, 2)                      '' 月
                .Day = Mid$(rs("MakeDate"), 7, 2)                        '' 日
            End With ' MakeDate
        End With ' head
        .KisyuCode = rs("KisyuCode")                                    '' 騎手コード
        .DelKubun = rs("DelKubun")                                      '' 騎手抹消区分
        With .IssueDate
            .Year = Mid$(rs("IssueDate"), 1, 4)                          '' 年
            .Month = Mid$(rs("IssueDate"), 5, 2)                         '' 月
            .Day = Mid$(rs("IssueDate"), 7, 2)                           '' 日
        End With ' IssueDate
        With .DelDate
            .Year = Mid$(rs("DelDate"), 1, 4)                            '' 年
            .Month = Mid$(rs("DelDate"), 5, 2)                           '' 月
            .Day = Mid$(rs("DelDate"), 7, 2)                             '' 日
        End With ' DelDate
        With .BirthDate
            .Year = Mid$(rs("BirthDate"), 1, 4)                          '' 年
            .Month = Mid$(rs("BirthDate"), 5, 2)                         '' 月
            .Day = Mid$(rs("BirthDate"), 7, 2)                           '' 日
        End With ' BirthDate
        .KisyuName = rs("KisyuName")                                    '' 騎手名漢字
        .Reserved = rs("reserved")                                      '' 予備
        .KisyuNameKana = rs("KisyuNameKana")                            '' 騎手名半角カナ
        .KisyuRyakusyo = rs("KisyuRyakusyo")                            '' 騎手名略称
        .KisyuNameEng = rs("KisyuNameEng")                              '' 騎手名欧字
        .SexCD = rs("SexCD")                                            '' 性別区分
        .SikakuCD = rs("SikakuCD")                                      '' 騎乗資格コード
        .MinaraiCD = rs("MinaraiCD")                                    '' 騎手見習コード
        .TozaiCD = rs("TozaiCD")                                        '' 騎手東西所属コード
        .Syotai = rs("Syotai")                                          '' 招待地域名
        .ChokyosiCode = rs("ChokyosiCode")                              '' 所属調教師コード
        .ChokyosiRyakusyo = rs("ChokyosiRyakusyo")                      '' 所属調教師名略称
        For i = 0 To 1
            With .HatuKiJyo(i)
                With .Hatukijyoid
                    .Year = Mid$(rs("HatuKiJyo" & i + 1 & "Hatukijyoid"), 1, 4)
                    .MonthDay = Mid$(rs("HatuKiJyo" & i + 1 & "Hatukijyoid"), 5, 4)
                    .JyoCD = Mid$(rs("HatuKiJyo" & i + 1 & "Hatukijyoid"), 9, 2)
                    .Kaiji = Mid$(rs("HatuKiJyo" & i + 1 & "Hatukijyoid"), 11, 2)
                    .Nichiji = Mid$(rs("HatuKiJyo" & i + 1 & "Hatukijyoid"), 13, 2)
                    .RaceNum = Mid$(rs("HatuKiJyo" & i + 1 & "Hatukijyoid"), 15, 2)
                End With ' Hatukijyoid
                .SyussoTosu = rs("HatuKiJyo" & i + 1 & "SyussoTosu")            '' 出走頭数
                .KettoNum = rs("HatuKiJyo" & i + 1 & "KettoNum")                '' 血統登録番号
                .BAMEI = rs("HatuKiJyo" & i + 1 & "Bamei")                      '' 馬名
                .KakuteiJyuni = rs("HatuKiJyo" & i + 1 & "KakuteiJyuni")        '' 確定着順
                .IJyoCD = rs("HatuKiJyo" & i + 1 & "IJyoCD")                    '' 異常区分コード
            End With ' HatuKiJyo
        Next i
        For i = 0 To 1
            With .HatuSyori(i)
                With .Hatusyoriid
                    .Year = Mid$(rs("HatuSyori" & i + 1 & "Hatusyoriid"), 1, 4)
                    .MonthDay = Mid$(rs("HatuSyori" & i + 1 & "Hatusyoriid"), 5, 4)
                    .JyoCD = Mid$(rs("HatuSyori" & i + 1 & "Hatusyoriid"), 9, 2)
                    .Kaiji = Mid$(rs("HatuSyori" & i + 1 & "Hatusyoriid"), 11, 2)
                    .Nichiji = Mid$(rs("HatuSyori" & i + 1 & "Hatusyoriid"), 13, 2)
                    .RaceNum = Mid$(rs("HatuSyori" & i + 1 & "Hatusyoriid"), 15, 2)
                End With ' Hatusyoriid
                .SyussoTosu = rs("HatuSyori" & i + 1 & "SyussoTosu")            '' 出走頭数
                .KettoNum = rs("HatuSyori" & i + 1 & "KettoNum")                '' 血統登録番号
                .BAMEI = rs("HatuSyori" & i + 1 & "Bamei")                      '' 馬名
            End With ' HatuSyori
        Next i
        For i = 0 To 2
            With .SaikinJyusyo(i)
                With .SaikinJyusyoid
                    .Year = Mid$(rs("SaikinJyusyo" & i + 1 & "SaikinJyusyoid"), 1, 4)
                    .MonthDay = Mid$(rs("SaikinJyusyo" & i + 1 & "SaikinJyusyoid"), 5, 4)
                    .JyoCD = Mid$(rs("SaikinJyusyo" & i + 1 & "SaikinJyusyoid"), 9, 2)
                    .Kaiji = Mid$(rs("SaikinJyusyo" & i + 1 & "SaikinJyusyoid"), 11, 2)
                    .Nichiji = Mid$(rs("SaikinJyusyo" & i + 1 & "SaikinJyusyoid"), 13, 2)
                    .RaceNum = Mid$(rs("SaikinJyusyo" & i + 1 & "SaikinJyusyoid"), 15, 2)
                End With ' SaikinJyusyoid
                .Hondai = rs("SaikinJyusyo" & i + 1 & "Hondai")                 '' 競走名本題
                .Ryakusyo10 = rs("SaikinJyusyo" & i + 1 & "Ryakusyo10")         '' 競走名略称10字
                .Ryakusyo6 = rs("SaikinJyusyo" & i + 1 & "Ryakusyo6")           '' 競走名略称6字
                .Ryakusyo3 = rs("SaikinJyusyo" & i + 1 & "Ryakusyo3")           '' 競走名略称3字
                .GradeCD = rs("SaikinJyusyo" & i + 1 & "GradeCD")               '' グレードコード
                .SyussoTosu = rs("SaikinJyusyo" & i + 1 & "SyussoTosu")         '' 出走頭数
                .KettoNum = rs("SaikinJyusyo" & i + 1 & "KettoNum")             '' 血統登録番号
                .BAMEI = rs("SaikinJyusyo" & i + 1 & "Bamei")                   '' 馬名
            End With ' SaikinJyusyo
        Next i
    End With ' buf
End Sub
        

'
'   機能: 構造体にデータをセットする－KS_SEISEKIレコード
'
'   備考: なし
'
Public Sub SetDataFromRS_KS_SEISEKI(ByRef rs As ADODB.Recordset, ByRef buf As JV_KS_KISYU)
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer

    With buf
        .KisyuCode = rs("KisyuCode")                                        '' 調教師コード
        For i = 0 To 2
            rs.Filter = "Num='" & i & "'"  ' ADO Recordset Function
            With .HonZenRuikei(i)
                .SetYear = rs("SetYear")                                    '' 設定年
                .HonSyokinHeichi = rs("HonSyokinHeichi")                    '' 平地本賞金合計
                .HonSyokinSyogai = rs("HonSyokinSyogai")                    '' 障害本賞金合計
                .FukaSyokinHeichi = rs("FukaSyokinHeichi")                  '' 平地付加賞金合計
                .FukaSyokinSyogai = rs("FukaSyokinSyogai")                  '' 障害付加賞金合計
                With .ChakuKaisuHeichi
                    For k = 0 To 5
                        .Chakukaisu(k) = rs("HeichiChakukaisu" & k + 1)
                    Next k
                End With ' ChakuKaisuHeichi
                With .ChakuKaisuSyogai
                    For k = 0 To 5
                        .Chakukaisu(k) = rs("SyogaiChakukaisu" & k + 1)
                    Next k
                End With ' ChakuKaisuSyogai
                For j = 0 To 5
                    With .ChakuKaisuKyori(j)
                        For k = 0 To 5
                            .Chakukaisu(k) = rs("Kyori" & j + 1 & "Chakukaisu" & k + 1)
                        Next k
                    End With ' ChakuKaisuKyori
                Next j
            End With ' HonZenRuikei
            With .HonZenRuikei(i)
                For j = 0 To 19
                    With .ChakuKaisuJyo(j)
                        For k = 0 To 5
                            .Chakukaisu(k) = rs("Jyo" & j + 1 & "Chakukaisu" & k + 1)
                        Next k
                    End With ' ChakuKaisuJyo
                Next j
            End With ' HonZenRuikei
        Next i
        .CRLF = vbCrLf 'CRLF
    End With ' buf
End Sub


'
'   機能: 構造体にデータをセットする－O1レコード
'
'   備考: なし
'
Public Sub SetDataFromRS_O1(ByRef rs As ADODB.Recordset, ByRef buf As JV_O1_ODDS_TANFUKUWAKU)
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim dic_O1_Tan As Scripting.Dictionary
    Dim dic_O1_Fuku As Scripting.Dictionary
    Dim dic_O1_Waku As Scripting.Dictionary
    Dim strKey As String
    
    Set dic_O1_Tan = New Dictionary
    Set dic_O1_Fuku = New Dictionary
    Set dic_O1_Waku = New Dictionary
    
    With buf
        With .head
            .RecordSpec = rs("RecordSpec")                          '' レコード種別
            .DataKubun = rs("DataKubun")                            '' データ区分
            With .MakeDate
                .Year = Mid$(rs("MakeDate"), 1, 4)                   '' 年
                .Month = Mid$(rs("MakeDate"), 5, 2)                  '' 月
                .Day = Mid$(rs("MakeDate"), 7, 2)                    '' 日
            End With ' MakeDate
        End With ' head
        With .id
            .Year = rs("Year")                                      '' 開催年
            .MonthDay = rs("MonthDay")                              '' 開催月日
            .JyoCD = rs("JyoCD")                                    '' 競馬場コード
            .Kaiji = rs("Kaiji")                                    '' 開催回[第N回]
            .Nichiji = rs("Nichiji")                                '' 開催日目[N日目]
            .RaceNum = rs("RaceNum")                                '' レース番号
        End With ' id
        With .HappyoTime
            .Month = Mid$(rs("HappyoTime"), 1, 2)                    '' 月
            .Day = Mid$(rs("HappyoTime"), 3, 2)                      '' 日
            .Hour = Mid$(rs("HappyoTime"), 5, 2)                     '' 時
            .Minute = Mid$(rs("HappyoTime"), 7, 2)                   '' 分
        End With ' HappyoTime
        .TorokuTosu = rs("TorokuTosu")                              '' 登録頭数
        .SyussoTosu = rs("SyussoTosu")                              '' 出走頭数
        .TansyoFlag = rs("TansyoFlag")                              '' 発売フラグ
        .FukusyoFlag = rs("FukusyoFlag")                            '' 発売フラグ
        .WakurenFlag = rs("WakurenFlag")                            '' 発売フラグ　枠連
        .FukuChakuBaraiKey = rs("FukuChakuBaraiKey")                '' 複勝着払キー
        .TotalHyosuTansyo = rs("TotalHyosuTansyo")                  '' 単勝票数合計
        .TotalHyosuFukusyo = rs("TotalHyosuFukusyo")                '' 複勝票数合計
        .TotalHyosuWakuren = rs("TotalHyosuWakuren")                '' 枠連票数合計
        
        Call MakeDic(dic_O1_Tan, rs, "TanData", 28, 2, 8)
        Call MakeDic(dic_O1_Fuku, rs, "FukuData", 28, 2, 12)
        For i = 0 To 27
            strKey = Format$(i + 1, "00")
            With .OddsTansyoInfo(i)
                If dic_O1_Tan.Exists(strKey) Then
                    .Umaban = strKey                                '' 馬番
                    .Odds = Mid$(dic_O1_Tan.item(strKey), 1, 4)      '' オッズ
                    .Ninki = Mid$(dic_O1_Tan.item(strKey), 5, 2)     '' 人気順
                Else
                    .Umaban = Space(2)
                    .Odds = Space(4)
                    .Ninki = Space(2)
                End If
            End With ' OddsTansyoInfo
            With .OddsFukusyoInfo(i)
                If dic_O1_Fuku.Exists(strKey) Then
                    .Umaban = strKey
                    .OddsLow = Mid$(dic_O1_Fuku.item(strKey), 1, 4)  '' 最低オッズ
                    .OddsHigh = Mid$(dic_O1_Fuku.item(strKey), 5, 4) '' 最高オッズ
                    .Ninki = Mid$(dic_O1_Fuku.item(strKey), 9, 2)    '' 人気順
                Else
                    .Umaban = Space(2)
                    .OddsLow = Space(4)
                    .OddsHigh = Space(4)
                    .Ninki = Space(2)
                End If
            End With ' OddsFukusyoInfo
        Next i
        
        Call MakeDic(dic_O1_Waku, rs, "WakuData", 36, 2, 9)
        k = 0
        For i = 1 To 8
            For j = i To 8
                strKey = Format$(i * 10 + j, "00")
                With .OddsWakurenInfo(k)
                    If dic_O1_Waku.Exists(strKey) Then
                        .Kumi = strKey                              '' 組
                        .Odds = Mid$(dic_O1_Waku.item(strKey), 1, 5) '' オッズ
                        .Ninki = Mid$(dic_O1_Waku.item(strKey), 6, 2) '' 人気順
                    Else
                        .Kumi = Space(2)
                        .Odds = Space(5)
                        .Ninki = Space(2)
                    End If
                End With ' OddsWakurenInfo
            k = k + 1
            Next j
        Next i
        .CRLF = vbCrLf 'CRLF
    End With ' buf
    
    Set dic_O1_Tan = Nothing
    Set dic_O1_Fuku = Nothing
    Set dic_O1_Waku = Nothing
    
End Sub


'
'   機能: 構造体にデータをセットする－O2レコード
'
'   備考: なし
'
Public Sub SetDataFromRS_O2(ByRef rs As ADODB.Recordset, ByRef buf As JV_O2_ODDS_UMAREN)
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim dic_O2 As Scripting.Dictionary
    Dim strKey As String

    Set dic_O2 = New Dictionary
    
    With buf
        With .head
            .RecordSpec = rs("RecordSpec")                          '' レコード種別
            .DataKubun = rs("DataKubun")                            '' データ区分
            With .MakeDate
                .Year = Mid$(rs("MakeDate"), 1, 4)                   '' 年
                .Month = Mid$(rs("MakeDate"), 5, 2)                  '' 月
                .Day = Mid$(rs("MakeDate"), 7, 2)                    '' 日
            End With ' MakeDate
        End With ' head
        With .id
            .Year = rs("Year")                                      '' 開催年
            .MonthDay = rs("MonthDay")                              '' 開催月日
            .JyoCD = rs("JyoCD")                                    '' 競馬場コード
            .Kaiji = rs("Kaiji")                                    '' 開催回[第N回]
            .Nichiji = rs("Nichiji")                                '' 開催日目[N日目]
            .RaceNum = rs("RaceNum")                                '' レース番号
        End With ' id
        With .HappyoTime
            .Month = Mid$(rs("HappyoTime"), 1, 2)                    '' 月
            .Day = Mid$(rs("HappyoTime"), 3, 2)                      '' 日
            .Hour = Mid$(rs("HappyoTime"), 5, 2)                     '' 時
            .Minute = Mid$(rs("HappyoTime"), 7, 2)                   '' 分
        End With ' HappyoTime
        .TorokuTosu = rs("TorokuTosu")                              '' 登録頭数
        .SyussoTosu = rs("SyussoTosu")                              '' 出走頭数
        .UmarenFlag = rs("UmarenFlag")                              '' 発売フラグ　馬連
        .TotalHyosuUmaren = rs("TotalHyosuUmaren")                  '' 馬連票数合計
        
        Call MakeDic(dic_O2, rs, "Data", 153, 4, 13)
        k = 0
        For i = 1 To 17
            For j = i + 1 To 18
                strKey = Format$(i * 100 + j, "0000")
                With .OddsUmarenInfo(k)
                    If dic_O2.Exists(strKey) Then
                        .Kumi = strKey                              '' 組
                        .Odds = Mid$(dic_O2.item(strKey), 1, 6)      '' オッズ
                        .Ninki = Mid$(dic_O2.item(strKey), 7, 3)     '' 人気順
                    Else
                        .Kumi = Space(4)
                        .Odds = Space(6)
                        .Ninki = Space(3)
                    End If
                End With ' OddsUmarenInfo
                k = k + 1
            Next j
        Next i
        .CRLF = vbCrLf 'CRLF
    End With ' buf
    
    Set dic_O2 = Nothing
End Sub


'
'   機能: 構造体にデータをセットする－O3レコード
'
'   備考: なし
'
Public Sub SetDataFromRS_O3(ByRef rs As ADODB.Recordset, ByRef buf As JV_O3_ODDS_WIDE)
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim dic_O3 As Scripting.Dictionary
    Dim strKey As String
    
    Set dic_O3 = New Dictionary
    
    With buf
        With .head
            .RecordSpec = rs("RecordSpec")                          '' レコード種別
            .DataKubun = rs("DataKubun")                            '' データ区分
            With .MakeDate
                .Year = Mid$(rs("MakeDate"), 1, 4)                   '' 年
                .Month = Mid$(rs("MakeDate"), 5, 2)                  '' 月
                .Day = Mid$(rs("MakeDate"), 7, 2)                    '' 日
            End With ' MakeDate
        End With ' head
        With .id
            .Year = rs("Year")                                      '' 開催年
            .MonthDay = rs("MonthDay")                              '' 開催月日
            .JyoCD = rs("JyoCD")                                    '' 競馬場コード
            .Kaiji = rs("Kaiji")                                    '' 開催回[第N回]
            .Nichiji = rs("Nichiji")                                '' 開催日目[N日目]
            .RaceNum = rs("RaceNum")                                '' レース番号
        End With ' id
        With .HappyoTime
            .Month = Mid$(rs("HappyoTime"), 1, 2)                    '' 月
            .Day = Mid$(rs("HappyoTime"), 3, 2)                      '' 日
            .Hour = Mid$(rs("HappyoTime"), 5, 2)                     '' 時
            .Minute = Mid$(rs("HappyoTime"), 7, 2)                   '' 分
        End With ' HappyoTime
        .TorokuTosu = rs("TorokuTosu")                              '' 登録頭数
        .SyussoTosu = rs("SyussoTosu")                              '' 出走頭数
        .WideFlag = rs("WideFlag")                                  '' 発売フラグ　ワイド
        .TotalHyosuWide = rs("TotalHyosuWide")                      '' ワイド票数合計
                
        Call MakeDic(dic_O3, rs, "Data", 153, 4, 17)
        k = 0
        For i = 1 To 17
            For j = i + 1 To 18
                strKey = Format$(i * 100 + j, "0000")
                With .OddsWideInfo(k)
                    If dic_O3.Exists(strKey) Then
                        .Kumi = strKey                              '' 組番
                        .OddsLow = Mid$(dic_O3.item(strKey), 1, 5)   '' 最低オッズ
                        .OddsHigh = Mid$(dic_O3.item(strKey), 6, 5)  '' 最高オッズ
                        .Ninki = Mid$(dic_O3.item(strKey), 11, 3)    '' 人気順
                    Else
                        .Kumi = Space(4)
                        .OddsLow = Space(5)
                        .OddsHigh = Space(5)
                        .Ninki = Space(3)
                    End If
                End With ' OddsWideInfo
                k = k + 1
            Next j
        Next i
        .CRLF = vbCrLf 'CRLF
    End With ' buf
    
    Set dic_O3 = Nothing
End Sub


'
'   機能: 構造体にデータをセットする－O4レコード
'
'   備考: なし
'
Public Sub SetDataFromRS_O4(ByRef rs As ADODB.Recordset, ByRef buf As JV_O4_ODDS_UMATAN)
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim dic_O4 As Scripting.Dictionary
    Dim strKey As String
    
    Set dic_O4 = New Dictionary

    With buf
        With .head
            .RecordSpec = rs("RecordSpec")                              '' レコード種別
            .DataKubun = rs("DataKubun")                                '' データ区分
            With .MakeDate
                .Year = Mid$(rs("MakeDate"), 1, 4)                       '' 年
                .Month = Mid$(rs("MakeDate"), 5, 2)                      '' 月
                .Day = Mid$(rs("MakeDate"), 7, 2)                        '' 日
            End With ' MakeDate
        End With ' head
        With .id
            .Year = rs("Year")                                          '' 開催年
            .MonthDay = rs("MonthDay")                                  '' 開催月日
            .JyoCD = rs("JyoCD")                                        '' 競馬場コード
            .Kaiji = rs("Kaiji")                                        '' 開催回[第N回]
            .Nichiji = rs("Nichiji")                                    '' 開催日目[N日目]
            .RaceNum = rs("RaceNum")                                    '' レース番号
        End With ' id
        With .HappyoTime
            .Month = Mid$(rs("HappyoTime"), 1, 2)                        '' 月
            .Day = Mid$(rs("HappyoTime"), 3, 2)                          '' 日
            .Hour = Mid$(rs("HappyoTime"), 5, 2)                         '' 時
            .Minute = Mid$(rs("HappyoTime"), 7, 2)                       '' 分
        End With ' HappyoTime
        .TorokuTosu = rs("TorokuTosu")                                  '' 登録頭数
        .SyussoTosu = rs("SyussoTosu")                                  '' 出走頭数
        .UmatanFlag = rs("UmatanFlag")                                  '' 発売フラグ　馬単
        .TotalHyosuUmatan = rs("TotalHyosuUmatan")                      '' 馬単票数合計
        
        Call MakeDic(dic_O4, rs, "Data", 306, 4, 13)
        k = 0
        For i = 1 To 18
            For j = 1 To 18
                If (j <> i) Then
                    strKey = Format$(i * 100 + j, "0000")
                    With .OddsUmatanInfo(k)
                        If dic_O4.Exists(strKey) Then
                            .Kumi = strKey                              '' 組番
                            .Odds = Mid$(dic_O4.item(strKey), 1, 6)      '' オッズ
                            .Ninki = Mid$(dic_O4.item(strKey), 7, 3)     '' 人気順
                        Else
                            .Kumi = Space(4)
                            .Odds = Space(6)
                            .Ninki = Space(3)
                        End If
                    End With ' OddsUmatanInfo
                    k = k + 1
                End If
                
            Next j
        Next i
        .CRLF = vbCrLf 'CRLF
    End With ' buf
    
    Set dic_O4 = Nothing
End Sub


'
'   機能: 構造体にデータをセットする－O5レコード
'
'   備考: なし
'
Public Sub SetDataFromRS_O5(ByRef rs As ADODB.Recordset, ByRef buf As JV_O5_ODDS_SANREN)
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim p As Long
    Dim dic_O5 As Scripting.Dictionary
    Dim strKey As String
    
    Set dic_O5 = New Dictionary
    
    With buf
        With .head
            .RecordSpec = rs("RecordSpec")                              '' レコード種別
            .DataKubun = rs("DataKubun")                                '' データ区分
            With .MakeDate
                .Year = Mid$(rs("MakeDate"), 1, 4)                       '' 年
                .Month = Mid$(rs("MakeDate"), 5, 2)                      '' 月
                .Day = Mid$(rs("MakeDate"), 7, 2)                        '' 日
            End With ' MakeDate
        End With ' head
        With .id
            .Year = rs("Year")                                          '' 開催年
            .MonthDay = rs("MonthDay")                                  '' 開催月日
            .JyoCD = rs("JyoCD")                                        '' 競馬場コード
            .Kaiji = rs("Kaiji")                                        '' 開催回[第N回]
            .Nichiji = rs("Nichiji")                                    '' 開催日目[N日目]
            .RaceNum = rs("RaceNum")                                    '' レース番号
        End With ' id
        With .HappyoTime
            .Month = Mid$(rs("HappyoTime"), 1, 2)                        '' 月
            .Day = Mid$(rs("HappyoTime"), 3, 2)                          '' 日
            .Hour = Mid$(rs("HappyoTime"), 5, 2)                         '' 時
            .Minute = Mid$(rs("HappyoTime"), 7, 2)                       '' 分
        End With ' HappyoTime
        .TorokuTosu = rs("TorokuTosu")                                  '' 登録頭数
        .SyussoTosu = rs("SyussoTosu")                                  '' 出走頭数
        .SanrenpukuFlag = rs("SanrenFlag")                              '' 発売フラグ　3連複
        .TotalHyosuSanrenpuku = rs("TotalHyosuSanren")                  '' 3連複票数合計
        
        Call MakeDic(dic_O5, rs, "Data", 816, 6, 15)
        p = 0
        For i = 1 To 16
            For j = i + 1 To 17
                For k = j + 1 To 18
                    strKey = Format$(i * 10000 + j * 100 + k, "000000")
                    With .OddsSanrenInfo(p)
                        If dic_O5.Exists(strKey) Then
                            .Kumi = strKey                              '' 組番
                            .Odds = Mid$(dic_O5.item(strKey), 1, 6)      '' オッズ
                            .Ninki = Mid$(dic_O5.item(strKey), 7, 3)     '' 人気順
                        Else
                            .Kumi = Space(6)
                            .Odds = Space(6)
                            .Ninki = Space(3)
                        End If
                    End With ' OddsSanrenInfo
                    p = p + 1
                Next k
            Next j
        Next i
        .CRLF = vbCrLf 'CRLF
    End With ' buf
    
    Set dic_O5 = Nothing
End Sub


'
'   機能: 構造体にデータをセットする－RCレコード
'
'   備考: なし
'
Public Sub SetDataFromRS_RC(ByRef rs As ADODB.Recordset, ByRef buf As JV_RC_RECORD)
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer

    With buf
        With .head
            .RecordSpec = rs("RecordSpec")                                                     '' レコード種別
            .DataKubun = rs("DataKubun")                                                       '' データ区分
            With .MakeDate
                .Year = Mid$(rs("MakeDate"), 1, 4)                                               '' 年
                .Month = Mid$(rs("MakeDate"), 5, 2)                                              '' 月
                .Day = Mid$(rs("MakeDate"), 7, 2)                                                '' 日
            End With ' MakeDate
        End With ' head
        .RecInfoKubun = rs("RecInfoKubun")                                                     '' レコード識別区分
        With .id
            .Year = rs("Year")                                                                 '' 開催年
            .MonthDay = rs("MonthDay")                                                         '' 開催月日
            .JyoCD = rs("JyoCD")                                                               '' 競馬場コード
            .Kaiji = rs("Kaiji")                                                               '' 開催回[第N回]
            .Nichiji = rs("Nichiji")                                                           '' 開催日目[N日目]
            .RaceNum = rs("RaceNum")                                                           '' レース番号
        End With ' id
        .TokuNum = Mid$(rs("TokuNum_SyubetuCD"), 1, 4)                                                             '' 特別競走番号
        .SyubetuCD = Mid$(rs("TokuNum_SyubetuCD"), 5, 2)                                        '' 競走種別コード
        .Hondai = rs("Hondai")                                                                 '' 競走名本題
        .GradeCD = rs("GradeCD")                                                               '' グレードコード
        .KYORI = rs("Kyori")                                                                   '' 距離
        .TrackCD = rs("TrackCD")                                                               '' 競走種別コード
        .RecKubun = rs("RecKubun")                                                             '' レコード区分
        .RecTime = rs("RecTime")                                                               '' レコードタイム
        With .TenkoBaba
            .TenkoCD = rs("TenkoCD")                                                           '' 天候コード
            .SibaBabaCD = rs("SibaBabaCD")                                                     '' 芝馬場状態コード
            .DirtBabaCD = rs("DirtBabaCD")                                                     '' ダート馬場状態コード
        End With ' TenkoBaba
        With .RecUmaInfo(0)
            .KettoNum = rs("RecUmaKettoNum1")                                                  '' 血統登録番号
            .BAMEI = rs("RecUmaBamei1")                                                        '' 馬名
            .UmaKigoCD = rs("RecUmaUmaKigoCD1")                                                '' 馬記号コード
            .SexCD = rs("RecUmaSexCD1")                                                        '' 性別コード
            .ChokyosiCode = rs("RecUmaChokyosiCode1")                                          '' 調教師コード
            .ChokyosiName = rs("RecUmaChokyosiName1")                                          '' 調教師名
            .Futan = rs("RecUmaFutan1")                                                        '' 負担重量
            .KisyuCode = rs("RecUmaKisyuCode1")                                                '' 騎手コード
            .KisyuName = rs("RecUmaKisyuName1")                                                '' 騎手名
        End With ' RecUmaInfo
        With .RecUmaInfo(1)
            .KettoNum = rs("RecUmaKettoNum2")                                                  '' 血統登録番号
            .BAMEI = rs("RecUmaBamei2")                                                        '' 馬名
            .UmaKigoCD = rs("RecUmaUmaKigoCD2")                                                '' 馬記号コード
            .SexCD = rs("RecUmaSexCD2")                                                        '' 性別コード
            .ChokyosiCode = rs("RecUmaChokyosiCode2")                                          '' 調教師コード
            .ChokyosiName = rs("RecUmaChokyosiName2")                                          '' 調教師名
            .Futan = rs("RecUmaFutan2")                                                        '' 負担重量
            .KisyuCode = rs("RecUmaKisyuCode2")                                                '' 騎手コード
            .KisyuName = rs("RecUmaKisyuName2")                                                '' 騎手名
        End With ' RecUmaInfo
        With .RecUmaInfo(2)
            .KettoNum = rs("RecUmaKettoNum3")                                                  '' 血統登録番号
            .BAMEI = rs("RecUmaBamei3")                                                        '' 馬名
            .UmaKigoCD = rs("RecUmaUmaKigoCD3")                                                '' 馬記号コード
            .SexCD = rs("RecUmaSexCD3")                                                        '' 性別コード
            .ChokyosiCode = rs("RecUmaChokyosiCode3")                                          '' 調教師コード
            .ChokyosiName = rs("RecUmaChokyosiName3")                                          '' 調教師名
            .Futan = rs("RecUmaFutan3")                                                        '' 負担重量
            .KisyuCode = rs("RecUmaKisyuCode3")                                                '' 騎手コード
            .KisyuName = rs("RecUmaKisyuName3")                                                '' 騎手名
        End With ' RecUmaInfo
        .CRLF = vbCrLf 'CRLF
    End With ' buf
End Sub


'
'   機能: 構造体にデータをセットする－SEレコード
'
'   備考: なし
'
Public Sub SetDataFromRS_SE(ByRef rs As ADODB.Recordset, ByRef buf As JV_SE_RACE_UMA)
On Error GoTo ErrorHandler
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer

    With buf
        With .head
            .RecordSpec = rs("RecordSpec")                                                     '' レコード種別
            .DataKubun = rs("DataKubun")                                                       '' データ区分
            With .MakeDate
                .Year = Mid$(rs("MakeDate"), 1, 4)                                               '' 年
                .Month = Mid$(rs("MakeDate"), 5, 2)                                              '' 月
                .Day = Mid$(rs("MakeDate"), 7, 2)                                                '' 日
            End With ' MakeDate
        End With ' head
        With .id
            .Year = rs("Year")                                                                 '' 開催年
            .MonthDay = rs("MonthDay")                                                         '' 開催月日
            .JyoCD = rs("JyoCD")                                                               '' 競馬場コード
            .Kaiji = rs("Kaiji")                                                               '' 開催回[第N回]
            .Nichiji = rs("Nichiji")                                                           '' 開催日目[N日目]
            .RaceNum = rs("RaceNum")                                                           '' レース番号
        End With ' id
        .Wakuban = rs("Wakuban")                                                               '' 枠番
        .Umaban = rs("Umaban")                                                                 '' 馬番
        .KettoNum = rs("KettoNum")                                                             '' 血統登録番号
        .BAMEI = rs("Bamei")                                                                   '' 馬名
        .UmaKigoCD = rs("UmaKigoCD")                                                           '' 馬記号コード
        .SexCD = rs("SexCD")                                                                   '' 性別コード
        .HinsyuCD = rs("HinsyuCD")                                                             '' 品種コード
        .KeiroCD = rs("KeiroCD")                                                               '' 毛色コード
        .Barei = rs("Barei")                                                                   '' 馬齢
        .TozaiCD = rs("TozaiCD")                                                               '' 東西所属コード
        .ChokyosiCode = rs("ChokyosiCode")                                                     '' 調教師コード
        .ChokyosiRyakusyo = rs("ChokyosiRyakusyo")                                             '' 調教師名略称
        .BanusiCode = rs("BanusiCode")                                                         '' 馬主コード
        .BanusiName = rs("BanusiName")                                                         '' 馬主名
        .Fukusyoku = rs("Fukusyoku")                                                           '' 服色標示
        .reserved1 = rs("reserved1")                                                           '' 予備
        .Futan = rs("Futan")                                                                   '' 負担重量
        .FutanBefore = rs("FutanBefore")                                                       '' 変更前負担重量
        .BLINKER = rs("Blinker")                                                               '' ブリンカー使用区分
        .reserved2 = rs("reserved2")                                                           '' 予備
        .KisyuCode = rs("KisyuCode")                                                           '' 騎手コード
        .KisyuCodeBefore = rs("KisyuCodeBefore")                                               '' 変更前騎手コード
        .KisyuRyakusyo = rs("KisyuRyakusyo")                                                   '' 騎手名略称
        .KisyuRyakusyoBefore = rs("KisyuRyakusyoBefore")                                       '' 変更前騎手名略称
        .MinaraiCD = rs("MinaraiCD")                                                           '' 騎手見習コード
        .MinaraiCDBefore = rs("MinaraiCDBefore")                                               '' 変更前騎手見習コード
        .BaTaijyu = rs("BaTaijyu")                                                             '' 馬体重
        .ZogenFugo = rs("ZogenFugo")                                                           '' 増減符号
        .ZogenSa = rs("ZogenSa")                                                               '' 増減差
        .IJyoCD = rs("IJyoCD")                                                                 '' 異常区分コード
        .NyusenJyuni = rs("NyusenJyuni")                                                       '' 入線順位
        .KakuteiJyuni = rs("KakuteiJyuni")                                                     '' 確定着順
        .DochakuKubun = rs("DochakuKubun")                                                     '' 同着区分
        .DochakuTosu = rs("DochakuTosu")                                                       '' 同着頭数
        .TIME = rs("Time")                                                                     '' 走破タイム
        .ChakusaCD = rs("ChakusaCD")                                                           '' 着差コード
        .ChakusaCDP = rs("ChakusaCDP")                                                         '' +着差コード
        .ChakusaCDPP = rs("ChakusaCDPP")                                                       '' ++着差コード
        .Jyuni1c = rs("Jyuni1c")                                                               '' 1コーナーでの順位
        .Jyuni2c = rs("Jyuni2c")                                                               '' 2コーナーでの順位
        .Jyuni3c = rs("Jyuni3c")                                                               '' 3コーナーでの順位
        .Jyuni4c = rs("Jyuni4c")                                                               '' 4コーナーでの順位
        .Odds = rs("Odds")                                                                     '' 単勝オッズ
        .Ninki = rs("Ninki")                                                                   '' 単勝人気順
        .Honsyokin = rs("Honsyokin")                                                           '' 獲得本賞金
        .Fukasyokin = rs("Fukasyokin")                                                         '' 獲得付加賞金
        .reserved3 = rs("reserved3")                                                           '' 予備
        .reserved4 = rs("reserved4")                                                           '' 予備
        .HaronTimeL4 = rs("HaronTimeL4")                                                       '' 後４ハロンタイム
        .HaronTimeL3 = rs("HaronTimeL3")                                                       '' 後３ハロンタイム
        For i = 0 To 2
            With .ChakuUmaInfo(i)
                .KettoNum = rs("KettoNum" & i + 1)              '' 血統登録番号                                                    '' 血統登録番号（相手馬1）
                .BAMEI = rs("Bamei" & i + 1)                    '' 馬名                                            '' 馬名
            End With ' ChakuUmaInfo
        Next i
        .TimeDiff = rs("TimeDiff")                                                             '' タイム差
        .RecordUpKubun = rs("RecordUpKubun")                                                   '' レコード更新区分
        .DMKubun = rs("DMKubun")                                                               '' マイニング区分
        .DMTime = rs("DMTime")                                                                 '' マイニング予想走破タイム
        .DMGosaP = rs("DMGosaP")                                                               '' 予測誤差(信頼度)＋
        .DMGosaM = rs("DMGosaM")                                                               '' 予測誤差(信頼度)－
        .DMJyuni = rs("DMJyuni")                                                               '' マイニング予想順位
        .KyakusituKubun = rs("KyakusituKubun")                                                 '' 今回レース脚質判定
        .CRLF = vbCrLf 'CRLF
    End With ' buf
    
    Exit Sub
ErrorHandler:
    If Err.Number = 3265 Then
        Resume Next
    End If
    gApp.ErrLog
End Sub


'
'   機能: 構造体にデータをセットする－SEレコード(A)
'
'   備考: なし
'
Public Sub SetDataFromRS_SE_A(ByRef rs As ADODB.Recordset, ByRef buf As JV_SE_RACE_UMA)
On Error GoTo ErrorHandler
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer

    With buf
        With .head
            .RecordSpec = rs("RecordSpec")                                                     '' レコード種別
            .DataKubun = rs("DataKubun")                                                       '' データ区分
            With .MakeDate
                .Year = Mid$(rs("MakeDate"), 1, 4)                                               '' 年
                .Month = Mid$(rs("MakeDate"), 5, 2)                                              '' 月
                .Day = Mid$(rs("MakeDate"), 7, 2)                                                '' 日
            End With ' MakeDate
        End With ' head
        With .id
            .Year = rs("Year")                                                                 '' 開催年
            .MonthDay = rs("MonthDay")                                                         '' 開催月日
            .JyoCD = rs("JyoCD")                                                               '' 競馬場コード
            .Kaiji = rs("Kaiji")                                                               '' 開催回[第N回]
            .Nichiji = rs("Nichiji")                                                           '' 開催日目[N日目]
            .RaceNum = rs("RaceNum")                                                           '' レース番号
        End With ' id
        .Wakuban = rs("Wakuban")                                                               '' 枠番
        .Umaban = rs("Umaban")                                                                 '' 馬番
        .KettoNum = rs("KettoNum")                                                             '' 血統登録番号
        .BAMEI = rs("Bamei")                                                                   '' 馬名
        .UmaKigoCD = rs("UmaKigoCD")                                                           '' 馬記号コード
        .SexCD = rs("SexCD")                                                                   '' 性別コード
        .HinsyuCD = rs("HinsyuCD")                                                             '' 品種コード
        .KeiroCD = rs("KeiroCD")                                                               '' 毛色コード
        .Barei = rs("Barei")                                                                   '' 馬齢
        .TozaiCD = rs("TozaiCD")                                                               '' 東西所属コード
        .ChokyosiCode = rs("ChokyosiCode")                                                     '' 調教師コード
        .ChokyosiRyakusyo = rs("ChokyosiRyakusyo")                                             '' 調教師名略称
        .BanusiCode = rs("BanusiCode")                                                         '' 馬主コード
        .BanusiName = rs("BanusiName")                                                         '' 馬主名
        .Fukusyoku = rs("Fukusyoku")                                                           '' 服色標示
        .reserved1 = rs("reserved1")                                                           '' 予備

    End With ' buf
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: 構造体にデータをセットする－SEレコード(B)
'
'   備考: なし
'
Public Sub SetDataFromRS_SE_B(ByRef rs As ADODB.Recordset, ByRef buf As JV_SE_RACE_UMA)
On Error GoTo ErrorHandler
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer

    With buf
        .Futan = rs("Futan")                                                                   '' 負担重量
        .FutanBefore = rs("FutanBefore")                                                       '' 変更前負担重量
        .BLINKER = rs("Blinker")                                                               '' ブリンカー使用区分
        .reserved2 = rs("reserved2")                                                           '' 予備
        .KisyuCode = rs("KisyuCode")                                                           '' 騎手コード
        .KisyuCodeBefore = rs("KisyuCodeBefore")                                               '' 変更前騎手コード
        .KisyuRyakusyo = rs("KisyuRyakusyo")                                                   '' 騎手名略称
        .KisyuRyakusyoBefore = rs("KisyuRyakusyoBefore")                                       '' 変更前騎手名略称
        .MinaraiCD = rs("MinaraiCD")                                                           '' 騎手見習コード
        .MinaraiCDBefore = rs("MinaraiCDBefore")                                               '' 変更前騎手見習コード
        .BaTaijyu = rs("BaTaijyu")                                                             '' 馬体重
        .ZogenFugo = rs("ZogenFugo")                                                           '' 増減符号
        .ZogenSa = rs("ZogenSa")                                                               '' 増減差
        .IJyoCD = rs("IJyoCD")                                                                 '' 異常区分コード
        .NyusenJyuni = rs("NyusenJyuni")                                                       '' 入線順位
        .KakuteiJyuni = rs("KakuteiJyuni")                                                     '' 確定着順
        .DochakuKubun = rs("DochakuKubun")                                                     '' 同着区分
        .DochakuTosu = rs("DochakuTosu")                                                       '' 同着頭数
        .TIME = rs("Time")                                                                     '' 走破タイム
        .ChakusaCD = rs("ChakusaCD")                                                           '' 着差コード
        .ChakusaCDP = rs("ChakusaCDP")                                                         '' +着差コード
        .ChakusaCDPP = rs("ChakusaCDPP")                                                       '' ++着差コード
        .Jyuni1c = rs("Jyuni1c")                                                               '' 1コーナーでの順位
        .Jyuni2c = rs("Jyuni2c")                                                               '' 2コーナーでの順位
        .Jyuni3c = rs("Jyuni3c")                                                               '' 3コーナーでの順位
        .Jyuni4c = rs("Jyuni4c")                                                               '' 4コーナーでの順位
        .Odds = rs("Odds")                                                                     '' 単勝オッズ
        .Ninki = rs("Ninki")                                                                   '' 単勝人気順
        .Honsyokin = rs("Honsyokin")                                                           '' 獲得本賞金
        .Fukasyokin = rs("Fukasyokin")                                                         '' 獲得付加賞金
        .reserved3 = rs("reserved3")                                                           '' 予備
        .reserved4 = rs("reserved4")                                                           '' 予備
        .HaronTimeL4 = rs("HaronTimeL4")                                                       '' 後４ハロンタイム
        .HaronTimeL3 = rs("HaronTimeL3")                                                       '' 後３ハロンタイム
        For i = 0 To 2
            With .ChakuUmaInfo(i)
                .KettoNum = rs("KettoNum" & i + 1)              '' 血統登録番号                                                    '' 血統登録番号（相手馬1）
                .BAMEI = rs("Bamei" & i + 1)                    '' 馬名                                            '' 馬名
            End With ' ChakuUmaInfo
        Next i
        .TimeDiff = rs("TimeDiff")                                                             '' タイム差
        .RecordUpKubun = rs("RecordUpKubun")                                                   '' レコード更新区分
        .DMKubun = rs("DMKubun")                                                               '' マイニング区分
        .DMTime = rs("DMTime")                                                                 '' マイニング予想走破タイム
        .DMGosaP = rs("DMGosaP")                                                               '' 予測誤差(信頼度)＋
        .DMGosaM = rs("DMGosaM")                                                               '' 予測誤差(信頼度)－
        .DMJyuni = rs("DMJyuni")                                                               '' マイニング予想順位
        .KyakusituKubun = rs("KyakusituKubun")                                                 '' 今回レース脚質判定
        .CRLF = vbCrLf 'CRLF
    End With ' buf
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: 構造体にデータをセットする－SKレコード
'
'   備考: なし
'
Public Sub SetDataFromRS_SK(ByRef rs As ADODB.Recordset, ByRef buf As JV_SK_SANKU)
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer

    With buf
        With .head
            .RecordSpec = rs("RecordSpec")                                                     '' レコード種別
            .DataKubun = rs("DataKubun")                                                       '' データ区分
            With .MakeDate
                .Year = Mid$(rs("MakeDate"), 1, 4)                                               '' 年
                .Month = Mid$(rs("MakeDate"), 5, 2)                                              '' 月
                .Day = Mid$(rs("MakeDate"), 7, 2)                                                '' 日
            End With ' MakeDate
        End With ' head
        .KettoNum = rs("KettoNum")                                                             '' 血統登録番号
        With .BirthDate
            .Year = Mid$(rs("BirthDate"), 1, 4)                                               '' 年
            .Month = Mid$(rs("BirthDate"), 5, 2)                                              '' 月
            .Day = Mid$(rs("BirthDate"), 7, 2)                                                '' 日
        End With ' BirthDate
        .SexCD = rs("SexCD")                                                                   '' 性別コード
        .HinsyuCD = rs("HinsyuCD")                                                             '' 品種コード
        .KeiroCD = rs("KeiroCD")                                                               '' 毛色コード
        .SankuMochiKubun = rs("SankuMochiKubun")                                               '' 産駒持込区分
        .ImportYear = rs("ImportYear")                                                         '' 輸入年
        .BreederCode = rs("BreederCode")                                                       '' 生産者コード
        .SanchiName = rs("SanchiName")                                                         '' 産地名
        .HansyokuNum(0) = rs("FNum")                                                           '' 父繁殖登録番号
        .HansyokuNum(1) = rs("MNum")                                                           '' 母繁殖登録番号
        .HansyokuNum(2) = rs("FFNum")                                                          '' 父父繁殖登録番号
        .HansyokuNum(3) = rs("FMNum")                                                          '' 父母繁殖登録番号
        .HansyokuNum(4) = rs("MFNum")                                                          '' 母父繁殖登録番号
        .HansyokuNum(5) = rs("MMNum")                                                          '' 母母繁殖登録番号
        .HansyokuNum(6) = rs("FFFNum")                                                         '' 父父父繁殖登録番号
        .HansyokuNum(7) = rs("FFMNum")                                                         '' 父父母繁殖登録番号
        .HansyokuNum(8) = rs("FMFNum")                                                         '' 父母父繁殖登録番号
        .HansyokuNum(9) = rs("FMMNum")                                                         '' 父母母繁殖登録番号
        .HansyokuNum(10) = rs("MFFNum")                                                        '' 母父父繁殖登録番号
        .HansyokuNum(11) = rs("MFMNum")                                                        '' 母父母繁殖登録番号
        .HansyokuNum(12) = rs("MMFNum")                                                        '' 母母父繁殖登録番号
        .HansyokuNum(13) = rs("MMMNum")                                                        '' 母母母繁殖登録番号
        .CRLF = vbCrLf 'CRLF
    End With ' buf
End Sub


'
'   機能: 構造体にデータをセットする－TKレコード
'
'   備考: なし
'
Public Sub SetDataFromRS_TK(ByRef rs As ADODB.Recordset, ByRef buf As JV_TK_TOKUUMA)
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer

    With buf
        With .head
            .RecordSpec = rs("RecordSpec")                                                      '' レコード種別
            .DataKubun = rs("DataKubun")                                                        '' データ区分
            With .MakeDate
                .Year = Mid$(rs("MakeDate"), 1, 4)                                              '' 年
                .Month = Mid$(rs("MakeDate"), 5, 2)                                             '' 月
                .Day = Mid$(rs("MakeDate"), 7, 2)                                               '' 日
            End With ' MakeDate
        End With ' head
        With .id
            .Year = rs("Year")                                                                 '' 開催年
            .MonthDay = rs("MonthDay")                                                         '' 開催月日
            .JyoCD = rs("JyoCD")                                                               '' 競馬場コード
            .Kaiji = rs("Kaiji")                                                               '' 開催回[第N回]
            .Nichiji = rs("Nichiji")                                                           '' 開催日目[N日目]
            .RaceNum = rs("RaceNum")                                                           '' レース番号
        End With ' id
        With .RaceInfo
            .YoubiCD = rs("YoubiCD")                                                            '' 曜日コード
            .TokuNum = rs("TokuNum")                                                            '' 特別競走番号
            .Hondai = rs("Hondai")                                                              '' 競走名本題
            .Fukudai = rs("Fukudai")                                                            '' 競走名副題
            .Kakko = rs("Kakko")                                                                '' 競走名カッコ内
            .HondaiEng = rs("HondaiEng")                                                        '' 競走名本題欧字
            .FukudaiEng = rs("FukudaiEng")                                                      '' 競走名副題欧字
            .KakkoEng = rs("KakkoEng")                                                          '' 競走名カッコ内欧字
            .Ryakusyo10 = rs("Ryakusyo10")                                                      '' 競走名略称１０字
            .Ryakusyo6 = rs("Ryakusyo6")                                                        '' 競走名略称６字
            .Ryakusyo3 = rs("Ryakusyo3")                                                        '' 競走名略称３字
            .Kubun = rs("Kubun")                                                                '' 競走名区分
            .Nkai = rs("Nkai")                                                                  '' 重賞回次[第N回]
        End With ' RaceInfo
        .GradeCD = rs("GradeCD")                                                                '' グレードコード
        With .JyokenInfo
            .SyubetuCD = rs("SyubetuCD")                                                        '' 競走種別コード
            .KigoCD = rs("KigoCD")                                                              '' 競走記号コード
            .JyuryoCD = rs("JyuryoCD")                                                          '' 重量種別コード
            For j = 0 To 4
                .JyokenCD(j) = rs("JyokenCD" & j + 1)                                          '' 競走条件コード
            Next j
        End With ' JyokenInfo
        .KYORI = rs("Kyori")                                                                    '' 距離
        .TrackCD = rs("TrackCD")                                                                '' トラックコード
        .CourseKubunCD = rs("CourseKubunCD")                                                    '' コース区分
        With .HandiDate
            .Year = Mid$(rs("HandiDate"), 1, 4)                                              '' 年
            .Month = Mid$(rs("HandiDate"), 5, 2)                                             '' 月
            .Day = Mid$(rs("HandiDate"), 7, 2)                                               '' 日
        End With ' HandiDate
        .TorokuTosu = rs("TorokuTosu")                                                          '' 登録頭数
        .CRLF = vbCrLf 'CRLF
    End With ' buf

End Sub


'
'   機能: 構造体にデータをセットする－TK_UMAINFOレコード
'
'   備考: なし
'
Public Sub SetDataFromRS_TK_UMAINFO(ByRef rs As ADODB.Recordset, ByRef buf As JV_TK_TOKUUMA)
    Dim i As Integer
    
    ' 全部再登録する
    rs.MoveFirst
    i = 0
    Do While Not rs.EOF
        With buf
            With .TokuUmaInfo(i)
                .num = rs("Num")                                                '' 連番
                .KettoNum = rs("KettoNum")                                      '' 血統登録番号
                .BAMEI = rs("Bamei")                                            '' 馬名
                .UmaKigoCD = rs("UmaKigoCD")                                    '' 馬記号コード
                .SexCD = rs("SexCD")                                            '' 性別コード
                .TozaiCD = rs("TozaiCD")                                        '' 調教師東西所属コード
                .ChokyosiCode = rs("ChokyosiCode")                              '' 調教師コード
                .ChokyosiRyakusyo = rs("ChokyosiRyakusyo")                      '' 調教師名略称
                .Futan = rs("Futan")                                            '' 負担重量
                .Koryu = rs("Koryu")                                            '' 交流区分
            End With ' TokuUmaInfo
            rs.MoveNext
        End With
        i = i + 1
    Loop
    
End Sub


'
'   機能: 構造体にデータをセットする－UMレコード
'
'   備考: なし
'
Public Sub SetDataFromRS_UM(ByRef rs As ADODB.Recordset, ByRef buf As JV_UM_UMA)
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer

    With buf
        With .head
            .RecordSpec = rs("RecordSpec")                                                     '' レコード種別
            .DataKubun = rs("DataKubun")                                                       '' データ区分
            With .MakeDate
                .Year = Mid$(rs("MakeDate"), 1, 4)                                              '' 年
                .Month = Mid$(rs("MakeDate"), 5, 2)                                             '' 月
                .Day = Mid$(rs("MakeDate"), 7, 2)                                               '' 日
            End With ' MakeDate
        End With ' head
        .KettoNum = rs("KettoNum")                                                             '' 血統登録番号
        .DelKubun = rs("DelKubun")                                                             '' 競走馬抹消区分
        With .RegDate
            .Year = Mid$(rs("RegDate"), 1, 4)                                                    '' 年月日
            .Month = Mid$(rs("RegDate"), 5, 2)                                                           '' 年月日
            .Day = Mid$(rs("RegDate"), 7, 2)                                                             '' 年月日
        End With ' RegDate
        With .DelDate
            .Year = Mid$(rs("DelDate"), 1, 4)                                                    '' 年月日
            .Month = Mid$(rs("DelDate"), 5, 2)                                                           '' 年月日
            .Day = Mid$(rs("DelDate"), 7, 2)                                                             '' 年月日
        End With ' DelDate
        With .BirthDate
            .Year = Mid$(rs("BirthDate"), 1, 4)                                                    '' 年月日
            .Month = Mid$(rs("BirthDate"), 5, 2)                                                           '' 年月日
            .Day = Mid$(rs("BirthDate"), 7, 2)                                                             '' 年月日
        End With ' BirthDate
        .BAMEI = rs("Bamei")                                                                   '' 馬名
        .BameiKana = rs("BameiKana")                                                           '' 馬名半角カナ
        .BameiEng = rs("BameiEng")                                                             '' 馬名欧字
        .UmaKigoCD = rs("UmaKigoCD")                                                           '' 馬記号コード
        .SexCD = rs("SexCD")                                                                   '' 性別コード
        .HinsyuCD = rs("HinsyuCD")                                                             '' 品種コード
        .KeiroCD = rs("KeiroCD")                                                               '' 毛色コード
        For i = 0 To 13
            With .Ketto3Info(i)
                .HansyokuNum = rs("Ketto3InfoHansyokuNum" & i + 1)                                             '' 繁殖登録番号
                .BAMEI = rs("Ketto3InfoBamei" & i + 1)                                                           '' 馬名
            End With ' Ketto3Info
        Next i
        .TozaiCD = rs("TozaiCD")                                                               '' 東西所属コード
        .ChokyosiCode = rs("ChokyosiCode")                                                     '' 調教師コード
        .ChokyosiRyakusyo = rs("ChokyosiRyakusyo")                                             '' 調教師名略称
        .Syotai = rs("Syotai")                                                                 '' 招待地域名
        .BreederCode = rs("BreederCode")                                                       '' 生産者コード
        .BreederName = rs("BreederName")                                                       '' 生産者名
        .SanchiName = rs("SanchiName")                                                         '' 産地名
        .BanusiCode = rs("BanusiCode")                                                         '' 馬主コード
        .BanusiName = rs("BanusiName")                                                         '' 馬主名
        .RuikeiHonsyoHeiti = rs("RuikeiHonsyoHeiti")                                           '' 平地本賞金累計
        .RuikeiHonsyoSyogai = rs("RuikeiHonsyoSyogai")                                         '' 障害本賞金累計
        .RuikeiFukaHeichi = rs("RuikeiFukaHeichi")                                             '' 平地付加賞金累計
        .RuikeiFukaSyogai = rs("RuikeiFukaSyogai")                                             '' 障害付加賞金累計
        .RuikeiSyutokuHeichi = rs("RuikeiSyutokuHeichi")                                       '' 平地収得賞金累計
        .RuikeiSyutokuSyogai = rs("RuikeiSyutokuSyogai")                                       '' 障害収得賞金累計
        With .ChakuSogo
            For i = 0 To 5
                .Chakukaisu(i) = rs("SogoChakukaisu" & i + 1)
            Next i
        End With ' ChakuSogo
        With .ChakuChuo
            For i = 0 To 5
                .Chakukaisu(i) = rs("ChuoChakukaisu" & i + 1)
            Next i
        End With ' ChakuChuo
        For i = 0 To 6
            With .ChakuKaisuBa(i)
                For j = 0 To 5
                    .Chakukaisu(j) = rs("Ba" & i + 1 & "Chakukaisu" & j + 1)
                Next j
            End With ' ChakuKaisuBa
        Next i
        For i = 0 To 11
            With .ChakuKaisuJyotai(i)
                For j = 0 To 5
                    .Chakukaisu(j) = rs("Jyotai" & i + 1 & "Chakukaisu" & j + 1)
                Next j
            End With ' ChakuKaisuJyotai
        Next i
        For i = 0 To 5
            With .ChakuKaisuKyori(i)
                For j = 0 To 5
                    .Chakukaisu(j) = rs("Kyori" & i + 1 & "Chakukaisu" & j + 1)
                Next j
            End With ' ChakuKaisuKyori
        Next i
        For i = 0 To 3
            .Kyakusitu(i) = rs("Kyakusitu" & i + 1)                                                    '' 脚質傾向
        Next i
        .RaceCount = rs("RaceCount")                                                           '' 登録レース数
        .CRLF = vbCrLf 'CRLF
    End With ' buf
End Sub


'
'   機能: 構造体にデータをセットする－WEレコード
'
'   備考: なし
'
Public Sub SetDataFromRS_WE(ByRef rs As ADODB.Recordset, ByRef buf As JV_WE_WEATHER)
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer

    With buf
        With .head
            .RecordSpec = rs("RecordSpec")                                                     '' レコード種別
            .DataKubun = rs("DataKubun")                                                       '' データ区分
            With .MakeDate
                .Year = Mid$(rs("MakeDate"), 1, 4)                                               '' 年
                .Month = Mid$(rs("MakeDate"), 5, 2)                                              '' 月
                .Day = Mid$(rs("MakeDate"), 7, 2)                                                '' 日
            End With ' MakeDate
        End With ' head
        With .id
            .Year = rs("Year")                                                                 '' 開催年
            .MonthDay = rs("MonthDay")                                                         '' 開催月日
            .JyoCD = rs("JyoCD")                                                               '' 競馬場コード
            .Kaiji = rs("Kaiji")                                                               '' 開催回[第N回]
            .Nichiji = rs("Nichiji")                                                           '' 開催日目[N日目]
        End With ' id
        With .HappyoTime
            .Month = Mid$(rs("HappyoTime"), 1, 2)                                               '' 月
            .Day = Mid$(rs("HappyoTime"), 3, 2)                                                 '' 日
            .Hour = Mid$(rs("HappyoTime"), 5, 2)                                                '' 時
            .Minute = Mid$(rs("HappyoTime"), 7, 2)                                              '' 分
        End With ' HappyoTime
        .HenkoID = rs("HenkoID")                                                               '' 変更識別
        With .TenkoBaba
            .TenkoCD = rs("AtoTenkoCD")                                                           '' 天候コード
            .SibaBabaCD = rs("AtoSibaBabaCD")                                                     '' 芝馬場状態コード
            .DirtBabaCD = rs("AtoDirtBabaCD")                                                     '' ダート馬場状態コード
        End With ' TenkoBaba
        With .TenkoBabaBefore
            .TenkoCD = rs("MaeTenkoCD")                                                           '' 天候コード
            .SibaBabaCD = rs("MaeSibaBabaCD")                                                     '' 芝馬場状態コード
            .DirtBabaCD = rs("MaeDirtBabaCD")                                                     '' ダート馬場状態コード
        End With ' TenkoBabaBefore
        .CRLF = vbCrLf 'CRLF
    End With ' buf
End Sub


'
'   機能: 構造体にデータをセットする－WHレコード
'
'   備考: なし
'
Public Sub SetDataFromRS_WH(ByRef rs As ADODB.Recordset, ByRef buf As JV_WH_BATAIJYU)
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer

    With buf
        With .head
            .RecordSpec = rs("RecordSpec")                                                     '' レコード種別
            .DataKubun = rs("DataKubun")                                                       '' データ区分
            With .MakeDate
                .Year = Mid$(rs("MakeDate"), 1, 4)                                               '' 年
                .Month = Mid$(rs("MakeDate"), 5, 2)                                              '' 月
                .Day = Mid$(rs("MakeDate"), 7, 2)                                                '' 日
            End With ' MakeDate
        End With ' head
        With .id
            .Year = rs("Year")                                                                 '' 開催年
            .MonthDay = rs("MonthDay")                                                         '' 開催月日
            .JyoCD = rs("JyoCD")                                                               '' 競馬場コード
            .Kaiji = rs("Kaiji")                                                               '' 開催回[第N回]
            .Nichiji = rs("Nichiji")                                                           '' 開催日目[N日目]
            .RaceNum = rs("RaceNum")                                                           '' レース番号
        End With ' id
        With .HappyoTime
            .Month = Mid$(rs("HappyoTime"), 1, 2)                                               '' 月
            .Day = Mid$(rs("HappyoTime"), 3, 2)                                                 '' 日
            .Hour = Mid$(rs("HappyoTime"), 5, 2)                                                '' 時
            .Minute = Mid$(rs("HappyoTime"), 7, 2)                                              '' 分
        End With ' HappyoTime
        For i = 0 To 17
            With .BataijyuInfo(i)
                .Umaban = rs("Umaban" & i + 1)                                            '' 馬番
                .BAMEI = rs("Bamei" & i + 1)                                              '' 馬名
                .BaTaijyu = rs("BaTaijyu" & i + 1)                                        '' 馬体重
                .ZogenFugo = rs("ZogenFugo" & i + 1)                                      '' 増減符号
                .ZogenSa = rs("ZogenSa" & i + 1)                                          '' 増減差
            End With ' BataijyuInfo
        Next i
        .CRLF = vbCrLf 'CRLF
    End With ' buf
End Sub


'
'   機能: 構造体にデータをセットする－YSレコード
'
'   備考: なし
'
Public Sub SetDataFromRS_YS(ByRef rs As ADODB.Recordset, ByRef buf As JV_YS_SCHEDULE)
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer

    With buf
        With .head
            .RecordSpec = rs("RecordSpec")                                                     '' レコード種別
            .DataKubun = rs("DataKubun")                                                       '' データ区分
            With .MakeDate
                .Year = Mid$(rs("MakeDate"), 1, 4)                                               '' 年
                .Month = Mid$(rs("MakeDate"), 5, 2)                                              '' 月
                .Day = Mid$(rs("MakeDate"), 7, 2)                                                '' 日
            End With ' MakeDate
        End With ' head
        With .id
            .Year = rs("Year")                                                                 '' 開催年
            .MonthDay = rs("MonthDay")                                                         '' 開催月日
            .JyoCD = rs("JyoCD")                                                               '' 競馬場コード
            .Kaiji = rs("Kaiji")                                                               '' 開催回[第N回]
            .Nichiji = rs("Nichiji")                                                           '' 開催日目[N日目]
        End With ' id
        .YoubiCD = rs("YoubiCD")                                                               '' 曜日コード
        With .JyusyoInfo(0)
            .TokuNum = rs("Jyusyo1TokuNum")                                                    '' 特別競走番号
            .Hondai = rs("Jyusyo1Hondai")                                                      '' 競走名本題
            .Ryakusyo10 = rs("Jyusyo1Ryakusyo10")                                              '' 競走名略称10字
            .Ryakusyo6 = rs("Jyusyo1Ryakusyo6")                                                '' 競走名略称6字
            .Ryakusyo3 = rs("Jyusyo1Ryakusyo3")                                                '' 競走名略称3字
            .Nkai = rs("Jyusyo1Nkai")                                                          '' 重賞回次[第N回]
            .GradeCD = rs("Jyusyo1GradeCD")                                                    '' グレードコード
            .SyubetuCD = rs("Jyusyo1SyubetuCD")                                                '' 競走種別コード
            .KigoCD = rs("Jyusyo1KigoCD")                                                      '' 競走記号コード
            .JyuryoCD = rs("Jyusyo1JyuryoCD")                                                  '' 重量種別コード
            .KYORI = rs("Jyusyo1Kyori")                                                        '' 距離
            .TrackCD = rs("Jyusyo1TrackCD")                                                    '' トラックコード
        End With ' JyusyoInfo(0)
        With .JyusyoInfo(1)
            .TokuNum = rs("Jyusyo2TokuNum")                                                    '' 特別競走番号
            .Hondai = rs("Jyusyo2Hondai")                                                      '' 競走名本題
            .Ryakusyo10 = rs("Jyusyo2Ryakusyo10")                                              '' 競走名略称10字
            .Ryakusyo6 = rs("Jyusyo2Ryakusyo6")                                                '' 競走名略称6字
            .Ryakusyo3 = rs("Jyusyo2Ryakusyo3")                                                '' 競走名略称3字
            .Nkai = rs("Jyusyo2Nkai")                                                          '' 重賞回次[第N回]
            .GradeCD = rs("Jyusyo2GradeCD")                                                    '' グレードコード
            .SyubetuCD = rs("Jyusyo2SyubetuCD")                                                '' 競走種別コード
            .KigoCD = rs("Jyusyo2KigoCD")                                                      '' 競走記号コード
            .JyuryoCD = rs("Jyusyo2JyuryoCD")                                                  '' 重量種別コード
            .KYORI = rs("Jyusyo2Kyori")                                                        '' 距離
            .TrackCD = rs("Jyusyo2TrackCD")                                                    '' トラックコード
        End With ' JyusyoInfo(1)
        With .JyusyoInfo(2)
            .TokuNum = rs("Jyusyo3TokuNum")                                                    '' 特別競走番号
            .Hondai = rs("Jyusyo3Hondai")                                                      '' 競走名本題
            .Ryakusyo10 = rs("Jyusyo3Ryakusyo10")                                              '' 競走名略称10字
            .Ryakusyo6 = rs("Jyusyo3Ryakusyo6")                                                '' 競走名略称6字
            .Ryakusyo3 = rs("Jyusyo3Ryakusyo3")                                                '' 競走名略称3字
            .Nkai = rs("Jyusyo3Nkai")                                                          '' 重賞回次[第N回]
            .GradeCD = rs("Jyusyo3GradeCD")                                                    '' グレードコード
            .SyubetuCD = rs("Jyusyo3SyubetuCD")                                                '' 競走種別コード
            .KigoCD = rs("Jyusyo3KigoCD")                                                      '' 競走記号コード
            .JyuryoCD = rs("Jyusyo3JyuryoCD")                                                  '' 重量種別コード
            .KYORI = rs("Jyusyo3Kyori")                                                        '' 距離
            .TrackCD = rs("Jyusyo3TrackCD")                                                    '' トラックコード
        End With ' JyusyoInfo(2)
        .CRLF = vbCrLf 'CRLF
    End With ' buf
End Sub


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部関数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: レコードセットから辞書を作成する
'
'   備考: なし
'
Private Sub MakeDic(ByRef dic As Dictionary, _
    rs As ADODB.Recordset, _
    field As String, _
    numBlocks As Long, _
    keyLen As Long, _
    blkLen As Long)

    Dim i As Long
    Dim p As Long
    Dim buf As String

    If IsNull(rs(field)) Then
        Exit Sub
    End If
    
    buf = rs(field)
    
    p = 1
    For i = 0 To numBlocks - 1
        If p > Len(buf) Then Exit For
        Call dic.Add(Mid$(buf, p, keyLen), Mid$(buf, p + keyLen, blkLen - keyLen))
        p = p + blkLen
    Next i

End Sub
