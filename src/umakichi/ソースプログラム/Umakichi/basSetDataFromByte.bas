Attribute VB_Name = "basSetDataFromByte"
'
'   データセット関数
'

Option Explicit

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   外部関数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: 構造体にデータをセットする－特別登録馬
'
'   備考: なし
'
    Public Sub SetDataFromByte_TK(ByRef bytBuf() As Byte, ByRef mBuf As JV_TK_TOKUUMA)
    Dim i As Integer                                '' ループカウンタ
    Dim j As Integer                                '' ループカウンタ
    Dim k As Integer                                '' ループカウンタ
    Dim p As Long                                   '' 切り分け開始位置
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMidByte(bytBuf, p, 2)              '' レコード種別
            .DataKubun = IncMidByte(bytBuf, p, 1)               '' データ区分
            With .MakeDate
                .Year = IncMidByte(bytBuf, p, 4)                '' 年
                .Month = IncMidByte(bytBuf, p, 2)               '' 月
                .Day = IncMidByte(bytBuf, p, 2)                 '' 日
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMidByte(bytBuf, p, 4)                    '' 開催年
            .MonthDay = IncMidByte(bytBuf, p, 4)                '' 開催月日
            .JyoCD = IncMidByte(bytBuf, p, 2)                   '' 競馬場コード
            .Kaiji = IncMidByte(bytBuf, p, 2)                   '' 開催回[第N回]
            .Nichiji = IncMidByte(bytBuf, p, 2)                 '' 開催日目[N日目]
            .RaceNum = IncMidByte(bytBuf, p, 2)                 '' レース番号
        End With ' id
        With .RaceInfo
            .YoubiCD = IncMidByte(bytBuf, p, 1)                 '' 曜日コード
            .TokuNum = IncMidByte(bytBuf, p, 4)                 '' 特別競走番号
            .Hondai = IncMidByte(bytBuf, p, 60)                 '' 競走名本題
            .Fukudai = IncMidByte(bytBuf, p, 60)                '' 競走名副題
            .Kakko = IncMidByte(bytBuf, p, 60)                  '' 競走名カッコ内
            .HondaiEng = IncMidByte(bytBuf, p, 120)             '' 競走名本題欧字
            .FukudaiEng = IncMidByte(bytBuf, p, 120)            '' 競走名副題欧字
            .KakkoEng = IncMidByte(bytBuf, p, 120)              '' 競走名カッコ内欧字
            .Ryakusyo10 = IncMidByte(bytBuf, p, 20)             '' 競走名略称１０字
            .Ryakusyo6 = IncMidByte(bytBuf, p, 12)              '' 競走名略称６字
            .Ryakusyo3 = IncMidByte(bytBuf, p, 6)               '' 競走名略称３字
            .Kubun = IncMidByte(bytBuf, p, 1)                   '' 競走名区分
            .Nkai = IncMidByte(bytBuf, p, 3)                    '' 重賞回次[第N回]
        End With ' RaceInfo
        .GradeCD = IncMidByte(bytBuf, p, 1)                     '' グレードコード
        With .JyokenInfo
            .SyubetuCD = IncMidByte(bytBuf, p, 2)               '' 競走種別コード
            .KigoCD = IncMidByte(bytBuf, p, 3)                  '' 競走記号コード
            .JyuryoCD = IncMidByte(bytBuf, p, 1)                '' 重量種別コード
            For j = 0 To 4
                .JyokenCD(j) = IncMidByte(bytBuf, p, 3)         '' 競走条件コード
            Next j
        End With ' JyokenInfo
        .KYORI = IncMidByte(bytBuf, p, 4)                       '' 距離
        .TrackCD = IncMidByte(bytBuf, p, 2)                     '' トラックコード
        .CourseKubunCD = IncMidByte(bytBuf, p, 2)               '' コース区分
        With .HandiDate
            .Year = IncMidByte(bytBuf, p, 4)                    '' 年
            .Month = IncMidByte(bytBuf, p, 2)                   '' 月
            .Day = IncMidByte(bytBuf, p, 2)                     '' 日
        End With ' HandiDate
        .TorokuTosu = IncMidByte(bytBuf, p, 3)                  '' 登録頭数
        For i = 0 To 299
            With .TokuUmaInfo(i)
                .num = IncMidByte(bytBuf, p, 3)                 '' 連番
                .KettoNum = IncMidByte(bytBuf, p, 10)           '' 血統登録番号
                .BAMEI = IncMidByte(bytBuf, p, 36)              '' 馬名
                .UmaKigoCD = IncMidByte(bytBuf, p, 2)           '' 馬記号コード
                .SexCD = IncMidByte(bytBuf, p, 1)               '' 性別コード
                .TozaiCD = IncMidByte(bytBuf, p, 1)             '' 調教師東西所属コード
                .ChokyosiCode = IncMidByte(bytBuf, p, 5)        '' 調教師コード
                .ChokyosiRyakusyo = IncMidByte(bytBuf, p, 8)    '' 調教師名略称
                .Futan = IncMidByte(bytBuf, p, 3)               '' 負担重量
                .Koryu = IncMidByte(bytBuf, p, 1)               '' 交流区分
            End With ' TokuUmaInfo
        Next i
        .CRLF = IncMidByte(bytBuf, p, 2)                        '' レコード区切
    End With

    End Sub


'
'   機能: 構造体にデータをセットする－レース詳細
'
'   備考: なし
'
    Public Sub SetDataFromByte_RA(ByRef bytBuf() As Byte, ByRef mBuf As JV_RA_RACE)
    Dim i As Integer                                '' ループカウンタ
    Dim j As Integer                                '' ループカウンタ
    Dim k As Integer                                '' ループカウンタ
    Dim p As Long                                   '' 切り分け開始位置
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMidByte(bytBuf, p, 2)              '' レコード種別
            .DataKubun = IncMidByte(bytBuf, p, 1)               '' データ区分
            With .MakeDate
                .Year = IncMidByte(bytBuf, p, 4)                '' 年
                .Month = IncMidByte(bytBuf, p, 2)               '' 月
                .Day = IncMidByte(bytBuf, p, 2)                 '' 日
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMidByte(bytBuf, p, 4)                    '' 開催年
            .MonthDay = IncMidByte(bytBuf, p, 4)                '' 開催月日
            .JyoCD = IncMidByte(bytBuf, p, 2)                   '' 競馬場コード
            .Kaiji = IncMidByte(bytBuf, p, 2)                   '' 開催回[第N回]
            .Nichiji = IncMidByte(bytBuf, p, 2)                 '' 開催日目[N日目]
            .RaceNum = IncMidByte(bytBuf, p, 2)                 '' レース番号
        End With ' id
        With .RaceInfo
            .YoubiCD = IncMidByte(bytBuf, p, 1)                 '' 曜日コード
            .TokuNum = IncMidByte(bytBuf, p, 4)                 '' 特別競走番号
            .Hondai = IncMidByte(bytBuf, p, 60)                 '' 競走名本題
            .Fukudai = IncMidByte(bytBuf, p, 60)                '' 競走名副題
            .Kakko = IncMidByte(bytBuf, p, 60)                  '' 競走名カッコ内
            .HondaiEng = IncMidByte(bytBuf, p, 120)             '' 競走名本題欧字
            .FukudaiEng = IncMidByte(bytBuf, p, 120)            '' 競走名副題欧字
            .KakkoEng = IncMidByte(bytBuf, p, 120)              '' 競走名カッコ内欧字
            .Ryakusyo10 = IncMidByte(bytBuf, p, 20)             '' 競走名略称１０字
            .Ryakusyo6 = IncMidByte(bytBuf, p, 12)              '' 競走名略称６字
            .Ryakusyo3 = IncMidByte(bytBuf, p, 6)               '' 競走名略称３字
            .Kubun = IncMidByte(bytBuf, p, 1)                   '' 競走名区分
            .Nkai = IncMidByte(bytBuf, p, 3)                    '' 重賞回次[第N回]
        End With ' RaceInfo
        .GradeCD = IncMidByte(bytBuf, p, 1)                     '' グレードコード
        .GradeCDBefore = IncMidByte(bytBuf, p, 1)               '' 変更前グレードコード
        With .JyokenInfo
            .SyubetuCD = IncMidByte(bytBuf, p, 2)               '' 競走種別コード
            .KigoCD = IncMidByte(bytBuf, p, 3)                  '' 競走記号コード
            .JyuryoCD = IncMidByte(bytBuf, p, 1)                '' 重量種別コード
            For j = 0 To 4
                .JyokenCD(j) = IncMidByte(bytBuf, p, 3)         '' 競走条件コード
            Next j
        End With ' JyokenInfo
        .JyokenName = IncMidByte(bytBuf, p, 60)                 '' 競走条件名称
        .KYORI = IncMidByte(bytBuf, p, 4)                       '' 距離
        .KyoriBefore = IncMidByte(bytBuf, p, 4)                 '' 変更前距離
        .TrackCD = IncMidByte(bytBuf, p, 2)                     '' トラックコード
        .TrackCDBefore = IncMidByte(bytBuf, p, 2)               '' 変更前トラックコード
        .CourseKubunCD = IncMidByte(bytBuf, p, 2)               '' コース区分
        .CourseKubunCDBefore = IncMidByte(bytBuf, p, 2)         '' 変更前コース区分
        For i = 0 To 6
            .Honsyokin(i) = IncMidByte(bytBuf, p, 8)            '' 本賞金
        Next i
        For i = 0 To 4
            .HonsyokinBefore(i) = IncMidByte(bytBuf, p, 8)      '' 変更前本賞金
        Next i
        For i = 0 To 4
            .Fukasyokin(i) = IncMidByte(bytBuf, p, 8)           '' 付加賞金
        Next i
        For i = 0 To 2
            .FukasyokinBefore(i) = IncMidByte(bytBuf, p, 8)     '' 変更前付加賞金
        Next i
        .HassoTime = IncMidByte(bytBuf, p, 4)                   '' 発走時刻
        .HassoTimeBefore = IncMidByte(bytBuf, p, 4)             '' 変更前発走時刻
        .TorokuTosu = IncMidByte(bytBuf, p, 2)                  '' 登録頭数
        .SyussoTosu = IncMidByte(bytBuf, p, 2)                  '' 出走頭数
        .NyusenTosu = IncMidByte(bytBuf, p, 2)                  '' 入線頭数
        With .TenkoBaba
            .TenkoCD = IncMidByte(bytBuf, p, 1)                 '' 天候コード
            .SibaBabaCD = IncMidByte(bytBuf, p, 1)              '' 芝馬場状態コード
            .DirtBabaCD = IncMidByte(bytBuf, p, 1)              '' ダート馬場状態コード
        End With ' TenkoBaba
        For i = 0 To 24
            .LapTime(i) = IncMidByte(bytBuf, p, 3)              '' ラップタイム
        Next i
        .SyogaiMileTime = IncMidByte(bytBuf, p, 4)              '' 障害マイルタイム
        .HaronTimeS3 = IncMidByte(bytBuf, p, 3)                 '' 前３ハロンタイム
        .HaronTimeS4 = IncMidByte(bytBuf, p, 3)                 '' 前４ハロンタイム
        .HaronTimeL3 = IncMidByte(bytBuf, p, 3)                 '' 後３ハロンタイム
        .HaronTimeL4 = IncMidByte(bytBuf, p, 3)                 '' 後４ハロンタイム
        For i = 0 To 3
            With .CornerInfo(i)
                .Corner = IncMidByte(bytBuf, p, 1)              '' コーナー
                .Syukaisu = IncMidByte(bytBuf, p, 1)            '' 周回数
                .Jyuni = IncMidByte(bytBuf, p, 70)              '' 各通過順位
            End With ' CornerInfo
        Next i
        .RecordUpKubun = IncMidByte(bytBuf, p, 1)               '' レコード更新区分
        .CRLF = IncMidByte(bytBuf, p, 2)        '' レコード区切り
    End With

    End Sub


'
'   機能: 構造体にデータをセットする－馬毎レース情報
'
'   備考: なし
'
    Public Sub SetDataFromByte_SE(ByRef bytBuf() As Byte, ByRef mBuf As JV_SE_RACE_UMA)
    Dim i As Integer                                '' ループカウンタ
    Dim p As Long                                   '' 切り分け開始位置
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMidByte(bytBuf, p, 2)      '' レコード種別
            .DataKubun = IncMidByte(bytBuf, p, 1)       '' データ区分
            With .MakeDate
                .Year = IncMidByte(bytBuf, p, 4)        '' 年
                .Month = IncMidByte(bytBuf, p, 2)       '' 月
                .Day = IncMidByte(bytBuf, p, 2)         '' 日
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMidByte(bytBuf, p, 4)            '' 開催年
            .MonthDay = IncMidByte(bytBuf, p, 4)        '' 開催月日
            .JyoCD = IncMidByte(bytBuf, p, 2)           '' 競馬場コード
            .Kaiji = IncMidByte(bytBuf, p, 2)           '' 開催回[第N回]
            .Nichiji = IncMidByte(bytBuf, p, 2)         '' 開催日目[N日目]
            .RaceNum = IncMidByte(bytBuf, p, 2)         '' レース番号
        End With ' id
        .Wakuban = IncMidByte(bytBuf, p, 1)             '' 枠番
        .Umaban = IncMidByte(bytBuf, p, 2)              '' 馬番
        .KettoNum = IncMidByte(bytBuf, p, 10)           '' 血統登録番号
        .BAMEI = IncMidByte(bytBuf, p, 36)              '' 馬名
        .UmaKigoCD = IncMidByte(bytBuf, p, 2)           '' 馬記号コード
        .SexCD = IncMidByte(bytBuf, p, 1)               '' 性別コード
        .HinsyuCD = IncMidByte(bytBuf, p, 1)            '' 品種コード
        .KeiroCD = IncMidByte(bytBuf, p, 2)             '' 毛色コード
        .Barei = IncMidByte(bytBuf, p, 2)               '' 馬齢
        .TozaiCD = IncMidByte(bytBuf, p, 1)             '' 東西所属コード
        .ChokyosiCode = IncMidByte(bytBuf, p, 5)        '' 調教師コード
        .ChokyosiRyakusyo = IncMidByte(bytBuf, p, 8)    '' 調教師名略称
        .BanusiCode = IncMidByte(bytBuf, p, 6)          '' 馬主コード
        .BanusiName = IncMidByte(bytBuf, p, 64)         '' 馬主名
        .Fukusyoku = IncMidByte(bytBuf, p, 60)          '' 服色標示
        .reserved1 = IncMidByte(bytBuf, p, 60)          '' 予備
        .Futan = IncMidByte(bytBuf, p, 3)               '' 負担重量
        .FutanBefore = IncMidByte(bytBuf, p, 3)         '' 変更前負担重量
        .BLINKER = IncMidByte(bytBuf, p, 1)             '' ブリンカー使用区分
        .reserved2 = IncMidByte(bytBuf, p, 1)           '' 予備
        .KisyuCode = IncMidByte(bytBuf, p, 5)           '' 騎手コード
        .KisyuCodeBefore = IncMidByte(bytBuf, p, 5)     '' 変更前騎手コード
        .KisyuRyakusyo = IncMidByte(bytBuf, p, 8)       '' 騎手名略称
        .KisyuRyakusyoBefore = IncMidByte(bytBuf, p, 8) '' 変更前騎手名略称
        .MinaraiCD = IncMidByte(bytBuf, p, 1)           '' 騎手見習コード
        .MinaraiCDBefore = IncMidByte(bytBuf, p, 1)     '' 変更前騎手見習コード
        .BaTaijyu = IncMidByte(bytBuf, p, 3)            '' 馬体重
        .ZogenFugo = IncMidByte(bytBuf, p, 1)           '' 増減符号
        .ZogenSa = IncMidByte(bytBuf, p, 3)             '' 増減差
        .IJyoCD = IncMidByte(bytBuf, p, 1)              '' 異常区分コード
        .NyusenJyuni = IncMidByte(bytBuf, p, 2)         '' 入線順位
        .KakuteiJyuni = IncMidByte(bytBuf, p, 2)        '' 確定着順
        .DochakuKubun = IncMidByte(bytBuf, p, 1)        '' 同着区分
        .DochakuTosu = IncMidByte(bytBuf, p, 1)         '' 同着頭数
        .TIME = IncMidByte(bytBuf, p, 4)                '' 走破タイム
        .ChakusaCD = IncMidByte(bytBuf, p, 3)           '' 着差コード
        .ChakusaCDP = IncMidByte(bytBuf, p, 3)          '' +着差コード
        .ChakusaCDPP = IncMidByte(bytBuf, p, 3)         '' ++着差コード
        .Jyuni1c = IncMidByte(bytBuf, p, 2)             '' 1コーナーでの順位
        .Jyuni2c = IncMidByte(bytBuf, p, 2)             '' 2コーナーでの順位
        .Jyuni3c = IncMidByte(bytBuf, p, 2)             '' 3コーナーでの順位
        .Jyuni4c = IncMidByte(bytBuf, p, 2)             '' 4コーナーでの順位
        .Odds = IncMidByte(bytBuf, p, 4)                '' 単勝オッズ
        .Ninki = IncMidByte(bytBuf, p, 2)               '' 単勝人気順
        .Honsyokin = IncMidByte(bytBuf, p, 8)           '' 獲得本賞金
        .Fukasyokin = IncMidByte(bytBuf, p, 8)          '' 獲得付加賞金
        .reserved3 = IncMidByte(bytBuf, p, 3)           '' 予備
        .reserved4 = IncMidByte(bytBuf, p, 3)           '' 予備
        .HaronTimeL4 = IncMidByte(bytBuf, p, 3)         '' 後４ハロンタイム
        .HaronTimeL3 = IncMidByte(bytBuf, p, 3)         '' 後３ハロンタイム
        For i = 0 To 2
            With .ChakuUmaInfo(i)
                .KettoNum = IncMidByte(bytBuf, p, 10)   '' 血統登録番号
                .BAMEI = IncMidByte(bytBuf, p, 36)      '' 馬名
            End With ' ChakuUmaInfo
        Next i
        .TimeDiff = IncMidByte(bytBuf, p, 4)            '' タイム差
        .RecordUpKubun = IncMidByte(bytBuf, p, 1)       '' レコード更新区分
        .DMKubun = IncMidByte(bytBuf, p, 1)             '' マイニング区分
        .DMTime = IncMidByte(bytBuf, p, 5)              '' マイニング予想走破タイム
        .DMGosaP = IncMidByte(bytBuf, p, 4)             '' 予測誤差(信頼度)＋
        .DMGosaM = IncMidByte(bytBuf, p, 4)             '' 予測誤差(信頼度)－
        .DMJyuni = IncMidByte(bytBuf, p, 2)             '' マイニング予想順位
        .KyakusituKubun = IncMidByte(bytBuf, p, 1)      '' 今回レース脚質判定
        .CRLF = IncMidByte(bytBuf, p, 2)                '' レコード区切り
    End With

    End Sub


'
'   機能: 構造体にデータをセットする－払戻
'
'   備考: なし
'
    Public Sub SetDataFromByte_HR(bytBuf() As Byte, ByRef mBuf As JV_HR_PAY)
    Dim i As Integer                                '' ループカウンタ
    Dim j As Integer                                '' ループカウンタ
    Dim k As Integer                                '' ループカウンタ
    Dim p As Long                                   '' 切り分け開始位置
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMidByte(bytBuf, p, 2)          '' レコード種別
            .DataKubun = IncMidByte(bytBuf, p, 1)           '' データ区分
            With .MakeDate
                .Year = IncMidByte(bytBuf, p, 4)            '' 年
                .Month = IncMidByte(bytBuf, p, 2)           '' 月
                .Day = IncMidByte(bytBuf, p, 2)             '' 日
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMidByte(bytBuf, p, 4)                '' 開催年
            .MonthDay = IncMidByte(bytBuf, p, 4)            '' 開催月日
            .JyoCD = IncMidByte(bytBuf, p, 2)               '' 競馬場コード
            .Kaiji = IncMidByte(bytBuf, p, 2)               '' 開催回[第N回]
            .Nichiji = IncMidByte(bytBuf, p, 2)             '' 開催日目[N日目]
            .RaceNum = IncMidByte(bytBuf, p, 2)             '' レース番号
        End With ' id
        .TorokuTosu = IncMidByte(bytBuf, p, 2)              '' 登録頭数
        .SyussoTosu = IncMidByte(bytBuf, p, 2)              '' 出走頭数
        For i = 0 To 8
            .FuseirituFlag(i) = IncMidByte(bytBuf, p, 1)    '' 不成立フラグ
        Next i
        For i = 0 To 8
            .TokubaraiFlag(i) = IncMidByte(bytBuf, p, 1)    '' 特払フラグ
        Next i
        For i = 0 To 8
            .HenkanFlag(i) = IncMidByte(bytBuf, p, 1)       '' 返還フラグ
        Next i
        For i = 0 To 27
            .HenkanUma(i) = IncMidByte(bytBuf, p, 1)        '' 返還馬番情報(馬番01～28)
        Next i
        For i = 0 To 7
            .HenkanWaku(i) = IncMidByte(bytBuf, p, 1)       '' 返還枠番情報(枠番1～8)
        Next i
        For i = 0 To 7
            .HenkanDoWaku(i) = IncMidByte(bytBuf, p, 1)     '' 返還同枠情報(枠番1～8)
        Next i
        For i = 0 To 2
            With .PayTansyo(i)
                .Umaban = IncMidByte(bytBuf, p, 2)          '' 馬番
                .Pay = IncMidByte(bytBuf, p, 9)             '' 払戻金
                .Ninki = IncMidByte(bytBuf, p, 2)           '' 人気順
            End With ' PayTansyo
        Next i
        For i = 0 To 4
            With .PayFukusyo(i)
                .Umaban = IncMidByte(bytBuf, p, 2)          '' 馬番
                .Pay = IncMidByte(bytBuf, p, 9)             '' 払戻金
                .Ninki = IncMidByte(bytBuf, p, 2)           '' 人気順
            End With ' PayFukusyo
        Next i
        For i = 0 To 2
            With .PayWakuren(i)
                .Umaban = IncMidByte(bytBuf, p, 2)          '' 馬番
                .Pay = IncMidByte(bytBuf, p, 9)             '' 払戻金
                .Ninki = IncMidByte(bytBuf, p, 2)           '' 人気順
            End With ' PayWakuren
        Next i
        For i = 0 To 2
            With .PayUmaren(i)
                .Kumi = IncMidByte(bytBuf, p, 4)            '' 組番
                .Pay = IncMidByte(bytBuf, p, 9)             '' 払戻金
                .Ninki = IncMidByte(bytBuf, p, 3)           '' 人気順
            End With ' PayUmaren
        Next i
        For i = 0 To 6
            With .PayWide(i)
                .Kumi = IncMidByte(bytBuf, p, 4)            '' 組番
                .Pay = IncMidByte(bytBuf, p, 9)             '' 払戻金
                .Ninki = IncMidByte(bytBuf, p, 3)           '' 人気順
            End With ' PayWide
        Next i
        For i = 0 To 2
            With .PayReserved1(i)
                .Kumi = IncMidByte(bytBuf, p, 4)            '' 組番
                .Pay = IncMidByte(bytBuf, p, 9)             '' 払戻金
                .Ninki = IncMidByte(bytBuf, p, 3)           '' 人気順
            End With ' PayReserved1
        Next i
        For i = 0 To 5
            With .PayUmatan(i)
                .Kumi = IncMidByte(bytBuf, p, 4)            '' 組番
                .Pay = IncMidByte(bytBuf, p, 9)             '' 払戻金
                .Ninki = IncMidByte(bytBuf, p, 3)           '' 人気順
            End With ' PayUmatan
        Next i
        For i = 0 To 2
            With .PaySanrenpuku(i)
                .Kumi = IncMidByte(bytBuf, p, 6)            '' 組番
                .Pay = IncMidByte(bytBuf, p, 9)             '' 払戻金
                .Ninki = IncMidByte(bytBuf, p, 3)           '' 人気順
            End With ' PaySanrenpuku
        Next i
        For i = 0 To 5
            With .PaySanrentan(i)
                .Kumi = IncMidByte(bytBuf, p, 6)            '' 組番
                .Pay = IncMidByte(bytBuf, p, 9)             '' 払戻金
                .Ninki = IncMidByte(bytBuf, p, 4)           '' 人気順
            End With ' PaySanrentan
        Next i
        .CRLF = IncMidByte(bytBuf, p, 2)        '' レコード区切り
    End With
   
    End Sub


'
'   機能: 構造体にデータをセットする－票数（全掛式）
'
'   備考: なし
'
    Public Sub SetDataFromByte_H1(bytBuf() As Byte, ByRef mBuf As JV_H1_HYOSU_ZENKAKE)
    Dim i As Integer                                '' ループカウンタ
    Dim j As Integer                                '' ループカウンタ
    Dim k As Integer                                '' ループカウンタ
    Dim p As Long                                   '' 切り分け開始位置
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMidByte(bytBuf, p, 2)      '' レコード種別
            .DataKubun = IncMidByte(bytBuf, p, 1)       '' データ区分
            With .MakeDate
                .Year = IncMidByte(bytBuf, p, 4)        '' 年
                .Month = IncMidByte(bytBuf, p, 2)       '' 月
                .Day = IncMidByte(bytBuf, p, 2)         '' 日
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMidByte(bytBuf, p, 4)            '' 開催年
            .MonthDay = IncMidByte(bytBuf, p, 4)        '' 開催月日
            .JyoCD = IncMidByte(bytBuf, p, 2)           '' 競馬場コード
            .Kaiji = IncMidByte(bytBuf, p, 2)           '' 開催回[第N回]
            .Nichiji = IncMidByte(bytBuf, p, 2)         '' 開催日目[N日目]
            .RaceNum = IncMidByte(bytBuf, p, 2)         '' レース番号
        End With ' id
        .TorokuTosu = IncMidByte(bytBuf, p, 2)          '' 登録頭数
        .SyussoTosu = IncMidByte(bytBuf, p, 2)          '' 出走頭数
        For i = 0 To 6
            .HatubaiFlag(i) = IncMidByte(bytBuf, p, 1)  '' 発売フラグ
        Next i
        .FukuChakuBaraiKey = IncMidByte(bytBuf, p, 1)   '' 複勝着払キー
        For i = 0 To 27
            .HenkanUma(i) = IncMidByte(bytBuf, p, 1)    '' 返還馬番情報(馬番01～28)
        Next i
        For i = 0 To 7
            .HenkanWaku(i) = IncMidByte(bytBuf, p, 1)   '' 返還枠番情報(枠番1～8)
        Next i
        For i = 0 To 7
            .HenkanDoWaku(i) = IncMidByte(bytBuf, p, 1) '' 返還同枠情報(枠番1～8)
        Next i
        For i = 0 To 27
            With .HyoTansyo(i)
                .Umaban = IncMidByte(bytBuf, p, 2)      '' 馬番
                .Hyo = IncMidByte(bytBuf, p, 11)        '' 票数
                .Ninki = IncMidByte(bytBuf, p, 2)       '' 人気
            End With ' HyoTansyo
        Next i
        For i = 0 To 27
            With .HyoFukusyo(i)
                .Umaban = IncMidByte(bytBuf, p, 2)      '' 馬番
                .Hyo = IncMidByte(bytBuf, p, 11)        '' 票数
                .Ninki = IncMidByte(bytBuf, p, 2)       '' 人気
            End With ' HyoFukusyo
        Next i
        For i = 0 To 35
            With .HyoWakuren(i)
                .Umaban = IncMidByte(bytBuf, p, 2)      '' 馬番
                .Hyo = IncMidByte(bytBuf, p, 11)        '' 票数
                .Ninki = IncMidByte(bytBuf, p, 2)       '' 人気
            End With ' HyoWakuren
        Next i
        For i = 0 To 152
            With .HyoUmaren(i)
                .Kumi = IncMidByte(bytBuf, p, 4)        '' 組番
                .Hyo = IncMidByte(bytBuf, p, 11)        '' 票数
                .Ninki = IncMidByte(bytBuf, p, 3)       '' 人気
            End With ' HyoUmaren
        Next i
        For i = 0 To 152
            With .HyoWide(i)
                .Kumi = IncMidByte(bytBuf, p, 4)        '' 組番
                .Hyo = IncMidByte(bytBuf, p, 11)        '' 票数
                .Ninki = IncMidByte(bytBuf, p, 3)       '' 人気
            End With ' HyoWide
        Next i
        For i = 0 To 305
            With .HyoUmatan(i)
                .Kumi = IncMidByte(bytBuf, p, 4)        '' 組番
                .Hyo = IncMidByte(bytBuf, p, 11)        '' 票数
                .Ninki = IncMidByte(bytBuf, p, 3)       '' 人気
            End With ' HyoUmatan
        Next i
        For i = 0 To 815
            With .HyoSanrenpuku(i)
                .Kumi = IncMidByte(bytBuf, p, 6)        '' 組番
                .Hyo = IncMidByte(bytBuf, p, 11)        '' 票数
                .Ninki = IncMidByte(bytBuf, p, 3)       '' 人気
            End With ' HyoSanrenpuku
        Next i
        For i = 0 To 13
            .HyoTotal(i) = IncMidByte(bytBuf, p, 11)    '' 票数合計
        Next i
        .CRLF = IncMidByte(bytBuf, p, 2)                '' レコード区切り
    End With
    
    End Sub


'
'   機能: 構造体にデータをセットする－票数６（３連単）
'
'   備考: なし
'
    Public Sub SetDataFromByte_H6(bytBuf() As Byte, ByRef mBuf As JV_H6_HYOSU_SANRENTAN)
    Dim i As Integer                                '' ループカウンタ
    Dim p As Long                                   '' 切り分け開始位置
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMidByte(bytBuf, p, 2)      '' レコード種別
            .DataKubun = IncMidByte(bytBuf, p, 1)       '' データ区分
            With .MakeDate
                .Year = IncMidByte(bytBuf, p, 4)        '' 年
                .Month = IncMidByte(bytBuf, p, 2)       '' 月
                .Day = IncMidByte(bytBuf, p, 2)         '' 日
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMidByte(bytBuf, p, 4)            '' 開催年
            .MonthDay = IncMidByte(bytBuf, p, 4)        '' 開催月日
            .JyoCD = IncMidByte(bytBuf, p, 2)           '' 競馬場コード
            .Kaiji = IncMidByte(bytBuf, p, 2)           '' 開催回[第N回]
            .Nichiji = IncMidByte(bytBuf, p, 2)         '' 開催日目[N日目]
            .RaceNum = IncMidByte(bytBuf, p, 2)         '' レース番号
        End With ' id
        .TorokuTosu = IncMidByte(bytBuf, p, 2)          '' 登録頭数
        .SyussoTosu = IncMidByte(bytBuf, p, 2)          '' 出走頭数
        .HatubaiFlag = IncMidByte(bytBuf, p, 1)         '' 発売フラグ 3連単
        For i = 0 To 17
            .HenkanUma(i) = IncMidByte(bytBuf, p, 1)    '' 返還馬番情報(馬番01～18)
        Next i
        For i = 0 To 4895
            With .HyoSanrentan(i)
                .Kumi = IncMidByte(bytBuf, p, 6)        '' 組番
                .Hyo = IncMidByte(bytBuf, p, 11)        '' 票数
                .Ninki = IncMidByte(bytBuf, p, 4)       '' 人気
            End With ' HyoSanrentan
        Next i
        .TotalHyoSanrentan = IncMidByte(bytBuf, p, 11)    '' 3連単票数合計
        .TotalHyoSanrentanHenkan = IncMidByte(bytBuf, p, 11) '' 3連単返還票数合計
        .CRLF = IncMidByte(bytBuf, p, 2)                  '' レコード区切り
    End With
    End Sub
    
    
'
'   機能: 構造体にデータをセットする－オッズ（単複枠）
'
'   備考: なし
'
    Public Sub SetDataFromByte_O1(bytBuf() As Byte, ByRef mBuf As JV_O1_ODDS_TANFUKUWAKU)
    Dim i As Integer                                '' ループカウンタ
    Dim j As Integer                                '' ループカウンタ
    Dim k As Integer                                '' ループカウンタ
    Dim p As Long                                   '' 切り分け開始位置
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMidByte(bytBuf, p, 2)      '' レコード種別
            .DataKubun = IncMidByte(bytBuf, p, 1)       '' データ区分
            With .MakeDate
                .Year = IncMidByte(bytBuf, p, 4)        '' 年
                .Month = IncMidByte(bytBuf, p, 2)       '' 月
                .Day = IncMidByte(bytBuf, p, 2) '' 日
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMidByte(bytBuf, p, 4)            '' 開催年
            .MonthDay = IncMidByte(bytBuf, p, 4)        '' 開催月日
            .JyoCD = IncMidByte(bytBuf, p, 2)           '' 競馬場コード
            .Kaiji = IncMidByte(bytBuf, p, 2)           '' 開催回[第N回]
            .Nichiji = IncMidByte(bytBuf, p, 2)         '' 開催日目[N日目]
            .RaceNum = IncMidByte(bytBuf, p, 2)         '' レース番号
        End With ' id
        With .HappyoTime
            .Month = IncMidByte(bytBuf, p, 2)           '' 月
            .Day = IncMidByte(bytBuf, p, 2)             '' 日
            .Hour = IncMidByte(bytBuf, p, 2)            '' 時
            .Minute = IncMidByte(bytBuf, p, 2)          '' 分
        End With ' HappyoTime
        .TorokuTosu = IncMidByte(bytBuf, p, 2)          '' 登録頭数
        .SyussoTosu = IncMidByte(bytBuf, p, 2)          '' 出走頭数
        .TansyoFlag = IncMidByte(bytBuf, p, 1)          '' 発売フラグ
        .FukusyoFlag = IncMidByte(bytBuf, p, 1)         '' 発売フラグ
        .WakurenFlag = IncMidByte(bytBuf, p, 1)         '' 発売フラグ　枠連
        .FukuChakuBaraiKey = IncMidByte(bytBuf, p, 1)   '' 複勝着払キー
        For i = 0 To 27
            With .OddsTansyoInfo(i)
                .Umaban = IncMidByte(bytBuf, p, 2)      '' 馬番
                .Odds = IncMidByte(bytBuf, p, 4)        '' オッズ
                .Ninki = IncMidByte(bytBuf, p, 2)       '' 人気順
            End With ' OddsTansyoInfo
        Next i
        For i = 0 To 27
            With .OddsFukusyoInfo(i)
                .Umaban = IncMidByte(bytBuf, p, 2)      '' 馬番
                .OddsLow = IncMidByte(bytBuf, p, 4)     '' 最低オッズ
                .OddsHigh = IncMidByte(bytBuf, p, 4)    '' 最高オッズ
                .Ninki = IncMidByte(bytBuf, p, 2)       '' 人気順
            End With ' OddsFukusyoInfo
        Next i
        For i = 0 To 35
            With .OddsWakurenInfo(i)
                .Kumi = IncMidByte(bytBuf, p, 2)        '' 組
                .Odds = IncMidByte(bytBuf, p, 5)        '' オッズ
                .Ninki = IncMidByte(bytBuf, p, 2)       '' 人気順
            End With ' OddsWakurenInfo
        Next i
        .TotalHyosuTansyo = IncMidByte(bytBuf, p, 11)   '' 単勝票数合計
        .TotalHyosuFukusyo = IncMidByte(bytBuf, p, 11)  '' 複勝票数合計
        .TotalHyosuWakuren = IncMidByte(bytBuf, p, 11)  '' 枠連票数合計
        .CRLF = IncMidByte(bytBuf, p, 2)                '' レコード区切り
    End With

    End Sub


'
'   機能: 構造体にデータをセットする－オッズ（馬連）
'
'   備考: なし
'
    Public Sub SetDataFromByte_O2(bytBuf() As Byte, ByRef mBuf As JV_O2_ODDS_UMAREN)
    Dim i As Integer                                '' ループカウンタ
    Dim j As Integer                                '' ループカウンタ
    Dim k As Integer                                '' ループカウンタ
    Dim p As Long                                   '' 切り分け開始位置
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMidByte(bytBuf, p, 2)      '' レコード種別
            .DataKubun = IncMidByte(bytBuf, p, 1)       '' データ区分
            With .MakeDate
                .Year = IncMidByte(bytBuf, p, 4)        '' 年
                .Month = IncMidByte(bytBuf, p, 2)       '' 月
                .Day = IncMidByte(bytBuf, p, 2) '' 日
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMidByte(bytBuf, p, 4)    '' 開催年
            .MonthDay = IncMidByte(bytBuf, p, 4)        '' 開催月日
            .JyoCD = IncMidByte(bytBuf, p, 2)   '' 競馬場コード
            .Kaiji = IncMidByte(bytBuf, p, 2)   '' 開催回[第N回]
            .Nichiji = IncMidByte(bytBuf, p, 2) '' 開催日目[N日目]
            .RaceNum = IncMidByte(bytBuf, p, 2) '' レース番号
        End With ' id
        With .HappyoTime
            .Month = IncMidByte(bytBuf, p, 2)   '' 月
            .Day = IncMidByte(bytBuf, p, 2)     '' 日
            .Hour = IncMidByte(bytBuf, p, 2)    '' 時
            .Minute = IncMidByte(bytBuf, p, 2)  '' 分
        End With ' HappyoTime
        .TorokuTosu = IncMidByte(bytBuf, p, 2)  '' 登録頭数
        .SyussoTosu = IncMidByte(bytBuf, p, 2)  '' 出走頭数
        .UmarenFlag = IncMidByte(bytBuf, p, 1)  '' 発売フラグ　馬連
        For i = 0 To 152
            With .OddsUmarenInfo(i)
                .Kumi = IncMidByte(bytBuf, p, 4)        '' 組番
                .Odds = IncMidByte(bytBuf, p, 6)        '' オッズ
                .Ninki = IncMidByte(bytBuf, p, 3)       '' 人気順
            End With ' OddsUmarenInfo
        Next i
        .TotalHyosuUmaren = IncMidByte(bytBuf, p, 11)   '' 馬連票数合計
        .CRLF = IncMidByte(bytBuf, p, 2)        '' レコード区切り
    End With

    End Sub


'
'   機能: 構造体にデータをセットする－オッズ（ワイド）
'
'   備考: なし
'
    Public Sub SetDataFromByte_O3(bytBuf() As Byte, ByRef mBuf As JV_O3_ODDS_WIDE)
    Dim i As Integer                                '' ループカウンタ
    Dim j As Integer                                '' ループカウンタ
    Dim k As Integer                                '' ループカウンタ
    Dim p As Long                                   '' 切り分け開始位置
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMidByte(bytBuf, p, 2)      '' レコード種別
            .DataKubun = IncMidByte(bytBuf, p, 1)       '' データ区分
            With .MakeDate
                .Year = IncMidByte(bytBuf, p, 4)        '' 年
                .Month = IncMidByte(bytBuf, p, 2)       '' 月
                .Day = IncMidByte(bytBuf, p, 2) '' 日
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMidByte(bytBuf, p, 4)    '' 開催年
            .MonthDay = IncMidByte(bytBuf, p, 4)        '' 開催月日
            .JyoCD = IncMidByte(bytBuf, p, 2)   '' 競馬場コード
            .Kaiji = IncMidByte(bytBuf, p, 2)   '' 開催回[第N回]
            .Nichiji = IncMidByte(bytBuf, p, 2) '' 開催日目[N日目]
            .RaceNum = IncMidByte(bytBuf, p, 2) '' レース番号
        End With ' id
        With .HappyoTime
            .Month = IncMidByte(bytBuf, p, 2)   '' 月
            .Day = IncMidByte(bytBuf, p, 2)     '' 日
            .Hour = IncMidByte(bytBuf, p, 2)    '' 時
            .Minute = IncMidByte(bytBuf, p, 2)  '' 分
        End With ' HappyoTime
        .TorokuTosu = IncMidByte(bytBuf, p, 2)  '' 登録頭数
        .SyussoTosu = IncMidByte(bytBuf, p, 2)  '' 出走頭数
        .WideFlag = IncMidByte(bytBuf, p, 1)    '' 発売フラグ　ワイド
        For i = 0 To 152
            With .OddsWideInfo(i)
                .Kumi = IncMidByte(bytBuf, p, 4)        '' 組番
                .OddsLow = IncMidByte(bytBuf, p, 5)     '' 最低オッズ
                .OddsHigh = IncMidByte(bytBuf, p, 5)    '' 最高オッズ
                .Ninki = IncMidByte(bytBuf, p, 3)       '' 人気順
            End With ' OddsWideInfo
        Next i
        .TotalHyosuWide = IncMidByte(bytBuf, p, 11)     '' ワイド票数合計
        .CRLF = IncMidByte(bytBuf, p, 2)        '' レコード区切り
    End With

    End Sub


'
'   機能: 構造体にデータをセットする－オッズ（馬単）
'
'   備考: なし
'
    Public Sub SetDataFromByte_O4(bytBuf() As Byte, ByRef mBuf As JV_O4_ODDS_UMATAN)
    Dim i As Integer                                '' ループカウンタ
    Dim j As Integer                                '' ループカウンタ
    Dim k As Integer                                '' ループカウンタ
    Dim p As Long                                   '' 切り分け開始位置
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMidByte(bytBuf, p, 2)      '' レコード種別
            .DataKubun = IncMidByte(bytBuf, p, 1)       '' データ区分
            With .MakeDate
                .Year = IncMidByte(bytBuf, p, 4)        '' 年
                .Month = IncMidByte(bytBuf, p, 2)       '' 月
                .Day = IncMidByte(bytBuf, p, 2)         '' 日
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMidByte(bytBuf, p, 4)            '' 開催年
            .MonthDay = IncMidByte(bytBuf, p, 4)        '' 開催月日
            .JyoCD = IncMidByte(bytBuf, p, 2)           '' 競馬場コード
            .Kaiji = IncMidByte(bytBuf, p, 2)           '' 開催回[第N回]
            .Nichiji = IncMidByte(bytBuf, p, 2)         '' 開催日目[N日目]
            .RaceNum = IncMidByte(bytBuf, p, 2)         '' レース番号
        End With ' id
        With .HappyoTime
            .Month = IncMidByte(bytBuf, p, 2)           '' 月
            .Day = IncMidByte(bytBuf, p, 2)             '' 日
            .Hour = IncMidByte(bytBuf, p, 2)            '' 時
            .Minute = IncMidByte(bytBuf, p, 2)          '' 分
        End With ' HappyoTime
        .TorokuTosu = IncMidByte(bytBuf, p, 2)          '' 登録頭数
        .SyussoTosu = IncMidByte(bytBuf, p, 2)          '' 出走頭数
        .UmatanFlag = IncMidByte(bytBuf, p, 1)          '' 発売フラグ　馬単
        For i = 0 To 305
            With .OddsUmatanInfo(i)
                .Kumi = IncMidByte(bytBuf, p, 4)        '' 組番
                .Odds = IncMidByte(bytBuf, p, 6)        '' オッズ
                .Ninki = IncMidByte(bytBuf, p, 3)       '' 人気順
            End With ' OddsUmatanInfo
        Next i
        .TotalHyosuUmatan = IncMidByte(bytBuf, p, 11)   '' 馬単票数合計
        .CRLF = IncMidByte(bytBuf, p, 2)                '' レコード区切り
    End With

    End Sub


'
'   機能: 構造体にデータをセットする－オッズ（３連複）
'
'   備考: なし
'
    Public Sub SetDataFromByte_O5(bytBuf() As Byte, ByRef mBuf As JV_O5_ODDS_SANREN)
    Dim i As Integer                                '' ループカウンタ
    Dim j As Integer                                '' ループカウンタ
    Dim k As Integer                                '' ループカウンタ
    Dim p As Long                                   '' 切り分け開始位置
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMidByte(bytBuf, p, 2)      '' レコード種別
            .DataKubun = IncMidByte(bytBuf, p, 1)       '' データ区分
            With .MakeDate
                .Year = IncMidByte(bytBuf, p, 4)        '' 年
                .Month = IncMidByte(bytBuf, p, 2)       '' 月
                .Day = IncMidByte(bytBuf, p, 2)         '' 日
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMidByte(bytBuf, p, 4)            '' 開催年
            .MonthDay = IncMidByte(bytBuf, p, 4)        '' 開催月日
            .JyoCD = IncMidByte(bytBuf, p, 2)           '' 競馬場コード
            .Kaiji = IncMidByte(bytBuf, p, 2)               '' 開催回[第N回]
            .Nichiji = IncMidByte(bytBuf, p, 2)         '' 開催日目[N日目]
            .RaceNum = IncMidByte(bytBuf, p, 2)         '' レース番号
        End With ' id
        With .HappyoTime
            .Month = IncMidByte(bytBuf, p, 2)           '' 月
            .Day = IncMidByte(bytBuf, p, 2)             '' 日
            .Hour = IncMidByte(bytBuf, p, 2)            '' 時
            .Minute = IncMidByte(bytBuf, p, 2)          '' 分
        End With ' HappyoTime
        .TorokuTosu = IncMidByte(bytBuf, p, 2)          '' 登録頭数
        .SyussoTosu = IncMidByte(bytBuf, p, 2)          '' 出走頭数
        .SanrenpukuFlag = IncMidByte(bytBuf, p, 1)      '' 発売フラグ　3連複
        For i = 0 To 815
            With .OddsSanrenInfo(i)
                .Kumi = IncMidByte(bytBuf, p, 6)        '' 組番
                .Odds = IncMidByte(bytBuf, p, 6)        '' オッズ
                .Ninki = IncMidByte(bytBuf, p, 3)       '' 人気順
            End With ' OddsSanrenInfo
        Next i
        .TotalHyosuSanrenpuku = IncMidByte(bytBuf, p, 11)       '' 3連複票数合計
        .CRLF = IncMidByte(bytBuf, p, 2)        '' レコード区切り
    End With
   
    End Sub


'
'   機能: 構造体にデータをセットする－オッズ（３連単）
'
'   備考: なし
'
    Public Sub SetDataFromByte_O6(bytBuf() As Byte, ByRef mBuf As JV_O6_ODDS_SANRENTAN)
    Dim i As Integer                                '' ループカウンター
    Dim p As Long                                   '' 切り分け開始位置
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMidByte(bytBuf, p, 2)        '' レコード種別
            .DataKubun = IncMidByte(bytBuf, p, 1)         '' データ区分
            With .MakeDate
                .Year = IncMidByte(bytBuf, p, 4)          '' 年
                .Month = IncMidByte(bytBuf, p, 2)         '' 月
                .Day = IncMidByte(bytBuf, p, 2)           '' 日
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMidByte(bytBuf, p, 4)              '' 開催年
            .MonthDay = IncMidByte(bytBuf, p, 4)          '' 開催月日
            .JyoCD = IncMidByte(bytBuf, p, 2)             '' 競馬場コード
            .Kaiji = IncMidByte(bytBuf, p, 2)             '' 開催回[第N回]
            .Nichiji = IncMidByte(bytBuf, p, 2)           '' 開催日目[N日目]
            .RaceNum = IncMidByte(bytBuf, p, 2)           '' レース番号
        End With ' id
        With .HappyoTime
            .Month = IncMidByte(bytBuf, p, 2)             '' 月
            .Day = IncMidByte(bytBuf, p, 2)               '' 日
            .Hour = IncMidByte(bytBuf, p, 2)              '' 時
            .Minute = IncMidByte(bytBuf, p, 2)            '' 分
        End With ' HappyoTime
        .TorokuTosu = IncMidByte(bytBuf, p, 2)            '' 登録頭数
        .SyussoTosu = IncMidByte(bytBuf, p, 2)            '' 出走頭数
        .SanrentanFlag = IncMidByte(bytBuf, p, 1)         '' 発売フラグ　3連単
        For i = 0 To 4895
            With .OddsSanrentanInfo(i)
                .Kumi = IncMidByte(bytBuf, p, 6)          '' 組番
                .Odds = IncMidByte(bytBuf, p, 7)          '' オッズ
                .Ninki = IncMidByte(bytBuf, p, 4)         '' 人気順
            End With
        Next i
        .TotalHyosuSanrentan = IncMidByte(bytBuf, p, 11)  '' 3連単票数合計
        .CRLF = IncMidByte(bytBuf, p, 2)                  '' レコード区切り
    End With
    
    End Sub

    
'
'   機能: 構造体にデータをセットする－競走馬マスタ
'
'   備考: なし
'
    Public Sub SetDataFromByte_UM(bytBuf() As Byte, ByRef mBuf As JV_UM_UMA)
    Dim i As Integer                                '' ループカウンタ
    Dim j As Integer                                '' ループカウンタ
    Dim k As Integer                                '' ループカウンタ
    Dim p As Long                                   '' 切り分け開始位置
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMidByte(bytBuf, p, 2)      '' レコード種別
            .DataKubun = IncMidByte(bytBuf, p, 1)       '' データ区分
            With .MakeDate
                .Year = IncMidByte(bytBuf, p, 4)        '' 年
                .Month = IncMidByte(bytBuf, p, 2)       '' 月
                .Day = IncMidByte(bytBuf, p, 2)         '' 日
            End With ' MakeDate
        End With ' head
        .KettoNum = IncMidByte(bytBuf, p, 10)           '' 血統登録番号
        .DelKubun = IncMidByte(bytBuf, p, 1)            '' 競走馬抹消区分
        With .RegDate
            .Year = IncMidByte(bytBuf, p, 4)            '' 年
            .Month = IncMidByte(bytBuf, p, 2)           '' 月
            .Day = IncMidByte(bytBuf, p, 2)             '' 日
        End With ' RegDate
        With .DelDate
            .Year = IncMidByte(bytBuf, p, 4)            '' 年
            .Month = IncMidByte(bytBuf, p, 2)           '' 月
            .Day = IncMidByte(bytBuf, p, 2)             '' 日
        End With ' DelDate
        With .BirthDate
            .Year = IncMidByte(bytBuf, p, 4)            '' 年
            .Month = IncMidByte(bytBuf, p, 2)           '' 月
            .Day = IncMidByte(bytBuf, p, 2)             '' 日
        End With ' BirthDate
        .BAMEI = IncMidByte(bytBuf, p, 36)              '' 馬名
        .BameiKana = IncMidByte(bytBuf, p, 36)          '' 馬名半角カナ
        .BameiEng = IncMidByte(bytBuf, p, 80)           '' 馬名欧字
        .UmaKigoCD = IncMidByte(bytBuf, p, 2)           '' 馬記号コード
        .SexCD = IncMidByte(bytBuf, p, 1)               '' 性別コード
        .HinsyuCD = IncMidByte(bytBuf, p, 1)            '' 品種コード
        .KeiroCD = IncMidByte(bytBuf, p, 2)             '' 毛色コード
        For i = 0 To 13
            With .Ketto3Info(i)
                .HansyokuNum = IncMidByte(bytBuf, p, 8) '' 繁殖登録番号
                .BAMEI = IncMidByte(bytBuf, p, 36)      '' 馬名
            End With ' Ketto3Info
        Next i
        .TozaiCD = IncMidByte(bytBuf, p, 1)             '' 東西所属コード
        .ChokyosiCode = IncMidByte(bytBuf, p, 5)        '' 調教師コード
        .ChokyosiRyakusyo = IncMidByte(bytBuf, p, 8)    '' 調教師名略称
        .Syotai = IncMidByte(bytBuf, p, 20)             '' 招待地域名
        .BreederCode = IncMidByte(bytBuf, p, 6)         '' 生産者コード
        .BreederName = IncMidByte(bytBuf, p, 70)        '' 生産者名
        .SanchiName = IncMidByte(bytBuf, p, 20)         '' 産地名
        .BanusiCode = IncMidByte(bytBuf, p, 6)          '' 馬主コード
        .BanusiName = IncMidByte(bytBuf, p, 64)         '' 馬主名
        .RuikeiHonsyoHeiti = IncMidByte(bytBuf, p, 9)   '' 平地本賞金累計
        .RuikeiHonsyoSyogai = IncMidByte(bytBuf, p, 9)  '' 障害本賞金累計
        .RuikeiFukaHeichi = IncMidByte(bytBuf, p, 9)    '' 平地付加賞金累計
        .RuikeiFukaSyogai = IncMidByte(bytBuf, p, 9)    '' 障害付加賞金累計
        .RuikeiSyutokuHeichi = IncMidByte(bytBuf, p, 9) '' 平地収得賞金累計
        .RuikeiSyutokuSyogai = IncMidByte(bytBuf, p, 9) '' 障害収得賞金累計
        With .ChakuSogo
            For j = 0 To 5
                .Chakukaisu(j) = IncMidByte(bytBuf, p, 3)
            Next j
        End With ' ChakuSogo
        With .ChakuChuo
            For j = 0 To 5
                .Chakukaisu(j) = IncMidByte(bytBuf, p, 3)
            Next j
        End With ' ChakuChuo
        For i = 0 To 6
            With .ChakuKaisuBa(i)
                For j = 0 To 5
                    .Chakukaisu(j) = IncMidByte(bytBuf, p, 3)
                Next j
            End With ' ChakuKaisuBa
        Next i
        For i = 0 To 11
            With .ChakuKaisuJyotai(i)
                For j = 0 To 5
                    .Chakukaisu(j) = IncMidByte(bytBuf, p, 3)
                Next j
            End With ' ChakuKaisuJyotai
        Next i
        For i = 0 To 5
            With .ChakuKaisuKyori(i)
                For j = 0 To 5
                    .Chakukaisu(j) = IncMidByte(bytBuf, p, 3)
                Next j
            End With ' ChakuKaisuKyoriu
        Next i
        For i = 0 To 3
            .Kyakusitu(i) = IncMidByte(bytBuf, p, 3)    '' 脚質傾向
        Next i
        .RaceCount = IncMidByte(bytBuf, p, 3)           '' 登録レース数
        .CRLF = IncMidByte(bytBuf, p, 2)                '' レコード区切り
    End With

    End Sub


'
'   機能: 構造体にデータをセットする－騎手マスタ
'
'   備考: なし
'
    Public Sub SetDataFromByte_KS(bytBuf() As Byte, ByRef mBuf As JV_KS_KISYU)
    Dim i As Integer                                '' ループカウンタ
    Dim j As Integer                                '' ループカウンタ
    Dim k As Integer                                '' ループカウンタ
    Dim p As Long                                   '' 切り分け開始位置
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMidByte(bytBuf, p, 2)      '' レコード種別
            .DataKubun = IncMidByte(bytBuf, p, 1)       '' データ区分
            With .MakeDate
                .Year = IncMidByte(bytBuf, p, 4)        '' 年
                .Month = IncMidByte(bytBuf, p, 2)       '' 月
                .Day = IncMidByte(bytBuf, p, 2) '' 日
            End With ' MakeDate
        End With ' head
        .KisyuCode = IncMidByte(bytBuf, p, 5)   '' 騎手コード
        .DelKubun = IncMidByte(bytBuf, p, 1)    '' 騎手抹消区分
        With .IssueDate
            .Year = IncMidByte(bytBuf, p, 4)    '' 年
            .Month = IncMidByte(bytBuf, p, 2)   '' 月
            .Day = IncMidByte(bytBuf, p, 2)     '' 日
        End With ' IssueDate
        With .DelDate
            .Year = IncMidByte(bytBuf, p, 4)    '' 年
            .Month = IncMidByte(bytBuf, p, 2)   '' 月
            .Day = IncMidByte(bytBuf, p, 2)     '' 日
        End With ' DelDate
        With .BirthDate
            .Year = IncMidByte(bytBuf, p, 4)    '' 年
            .Month = IncMidByte(bytBuf, p, 2)   '' 月
            .Day = IncMidByte(bytBuf, p, 2)     '' 日
        End With ' BirthDate
        .KisyuName = IncMidByte(bytBuf, p, 34)  '' 騎手名漢字
        .Reserved = IncMidByte(bytBuf, p, 34)   '' 予備
        .KisyuNameKana = IncMidByte(bytBuf, p, 30)      '' 騎手名半角カナ
        .KisyuRyakusyo = IncMidByte(bytBuf, p, 8)       '' 騎手名略称
        .KisyuNameEng = IncMidByte(bytBuf, p, 80)       '' 騎手名欧字
        .SexCD = IncMidByte(bytBuf, p, 1)       '' 性別区分
        .SikakuCD = IncMidByte(bytBuf, p, 1)    '' 騎乗資格コード
        .MinaraiCD = IncMidByte(bytBuf, p, 1)   '' 騎手見習コード
        .TozaiCD = IncMidByte(bytBuf, p, 1)     '' 騎手東西所属コード
        .Syotai = IncMidByte(bytBuf, p, 20)     '' 招待地域名
        .ChokyosiCode = IncMidByte(bytBuf, p, 5)        '' 所属調教師コード
        .ChokyosiRyakusyo = IncMidByte(bytBuf, p, 8)    '' 所属調教師名略称
        For i = 0 To 1
            With .HatuKiJyo(i)
                With .Hatukijyoid
                    .Year = IncMidByte(bytBuf, p, 4)    '' 開催年
                    .MonthDay = IncMidByte(bytBuf, p, 4)        '' 開催月日
                    .JyoCD = IncMidByte(bytBuf, p, 2)   '' 競馬場コード
                    .Kaiji = IncMidByte(bytBuf, p, 2)   '' 開催回[第N回]
                    .Nichiji = IncMidByte(bytBuf, p, 2) '' 開催日目[N日目]
                    .RaceNum = IncMidByte(bytBuf, p, 2) '' レース番号
                End With ' Hatukijyoid
                .SyussoTosu = IncMidByte(bytBuf, p, 2)  '' 出走頭数
                .KettoNum = IncMidByte(bytBuf, p, 10)   '' 血統登録番号
                .BAMEI = IncMidByte(bytBuf, p, 36)      '' 馬名
                .KakuteiJyuni = IncMidByte(bytBuf, p, 2)        '' 確定着順
                .IJyoCD = IncMidByte(bytBuf, p, 1)      '' 異常区分コード
            End With ' HatuKiJyo
        Next i
        For i = 0 To 1
            With .HatuSyori(i)
                With .Hatusyoriid
                    .Year = IncMidByte(bytBuf, p, 4)    '' 開催年
                    .MonthDay = IncMidByte(bytBuf, p, 4)        '' 開催月日
                    .JyoCD = IncMidByte(bytBuf, p, 2)   '' 競馬場コード
                    .Kaiji = IncMidByte(bytBuf, p, 2)   '' 開催回[第N回]
                    .Nichiji = IncMidByte(bytBuf, p, 2) '' 開催日目[N日目]
                    .RaceNum = IncMidByte(bytBuf, p, 2) '' レース番号
                End With ' Hatusyoriid
                .SyussoTosu = IncMidByte(bytBuf, p, 2)  '' 出走頭数
                .KettoNum = IncMidByte(bytBuf, p, 10)   '' 血統登録番号
                .BAMEI = IncMidByte(bytBuf, p, 36)      '' 馬名
            End With ' HatuSyori
        Next i
        For i = 0 To 2
            With .SaikinJyusyo(i)
                With .SaikinJyusyoid
                    .Year = IncMidByte(bytBuf, p, 4)    '' 開催年
                    .MonthDay = IncMidByte(bytBuf, p, 4)        '' 開催月日
                    .JyoCD = IncMidByte(bytBuf, p, 2)   '' 競馬場コード
                    .Kaiji = IncMidByte(bytBuf, p, 2)   '' 開催回[第N回]
                    .Nichiji = IncMidByte(bytBuf, p, 2) '' 開催日目[N日目]
                    .RaceNum = IncMidByte(bytBuf, p, 2) '' レース番号
                End With ' SaikinJyusyoid
                .Hondai = IncMidByte(bytBuf, p, 60)     '' 競走名本題
                .Ryakusyo10 = IncMidByte(bytBuf, p, 20) '' 競走名略称10字
                .Ryakusyo6 = IncMidByte(bytBuf, p, 12)  '' 競走名略称6字
                .Ryakusyo3 = IncMidByte(bytBuf, p, 6)   '' 競走名略称3字
                .GradeCD = IncMidByte(bytBuf, p, 1)     '' グレードコード
                .SyussoTosu = IncMidByte(bytBuf, p, 2)  '' 出走頭数
                .KettoNum = IncMidByte(bytBuf, p, 10)   '' 血統登録番号
                .BAMEI = IncMidByte(bytBuf, p, 36)      '' 馬名
            End With ' SaikinJyusyo
        Next i
        For i = 0 To 2
            With .HonZenRuikei(i)
                .SetYear = IncMidByte(bytBuf, p, 4)     '' 設定年
                .HonSyokinHeichi = IncMidByte(bytBuf, p, 10)    '' 平地本賞金合計
                .HonSyokinSyogai = IncMidByte(bytBuf, p, 10)    '' 障害本賞金合計
                .FukaSyokinHeichi = IncMidByte(bytBuf, p, 10)   '' 平地付加賞金合計
                .FukaSyokinSyogai = IncMidByte(bytBuf, p, 10)   '' 障害付加賞金合計
                With .ChakuKaisuHeichi
                    For k = 0 To 5
                        .Chakukaisu(k) = IncMidByte(bytBuf, p, 6)
                    Next k
                End With ' ChakuKaisuHeichi
                With .ChakuKaisuSyogai
                    For k = 0 To 5
                        .Chakukaisu(k) = IncMidByte(bytBuf, p, 6)
                    Next k
                End With ' ChakuKaisuSyogai
                For j = 0 To 19
                    With .ChakuKaisuJyo(j)
                        For k = 0 To 5
                            .Chakukaisu(k) = IncMidByte(bytBuf, p, 6)
                        Next k
                    End With ' ChakuKaisuJyo
                Next j
                For j = 0 To 5
                    With .ChakuKaisuKyori(j)
                        For k = 0 To 5
                            .Chakukaisu(k) = IncMidByte(bytBuf, p, 6)
                        Next k
                    End With ' ChakuKaisuKyori
                Next j
            End With ' HonZenRuikei
        Next i
        .CRLF = IncMidByte(bytBuf, p, 2)        '' レコード区切り
    End With

    End Sub


'
'   機能: 構造体にデータをセットする－調教師マスタ
'
'   備考: なし
'
    Public Sub SetDataFromByte_CH(bytBuf() As Byte, ByRef mBuf As JV_CH_CHOKYOSI)
    Dim i As Integer                                '' ループカウンタ
    Dim j As Integer                                '' ループカウンタ
    Dim k As Integer                                '' ループカウンタ
    Dim p As Long                                   '' 切り分け開始位置
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMidByte(bytBuf, p, 2)              '' レコード種別
            .DataKubun = IncMidByte(bytBuf, p, 1)               '' データ区分
            With .MakeDate
                .Year = IncMidByte(bytBuf, p, 4)                '' 年
                .Month = IncMidByte(bytBuf, p, 2)               '' 月
                .Day = IncMidByte(bytBuf, p, 2)                 '' 日
            End With ' MakeDate
        End With ' head
        .ChokyosiCode = IncMidByte(bytBuf, p, 5)                '' 調教師コード
        .DelKubun = IncMidByte(bytBuf, p, 1)                    '' 調教師抹消区分
        With .IssueDate
            .Year = IncMidByte(bytBuf, p, 4)                    '' 年
            .Month = IncMidByte(bytBuf, p, 2)                   '' 月
            .Day = IncMidByte(bytBuf, p, 2)                     '' 日
        End With ' IssueDate
        With .DelDate
            .Year = IncMidByte(bytBuf, p, 4)                    '' 年
            .Month = IncMidByte(bytBuf, p, 2)                   '' 月
            .Day = IncMidByte(bytBuf, p, 2)                     '' 日
        End With ' DelDate
        With .BirthDate
            .Year = IncMidByte(bytBuf, p, 4)                    '' 年
            .Month = IncMidByte(bytBuf, p, 2)                   '' 月
            .Day = IncMidByte(bytBuf, p, 2)                     '' 日
        End With ' BirthDate
        .ChokyosiName = IncMidByte(bytBuf, p, 34)               '' 調教師名漢字
        .ChokyosiNameKana = IncMidByte(bytBuf, p, 30)           '' 調教師名半角カナ
        .ChokyosiRyakusyo = IncMidByte(bytBuf, p, 8)            '' 調教師名略称
        .ChokyosiNameEng = IncMidByte(bytBuf, p, 80)            '' 調教師名欧字
        .SexCD = IncMidByte(bytBuf, p, 1)                       '' 性別区分
        .TozaiCD = IncMidByte(bytBuf, p, 1)                     '' 調教師東西所属コード
        .Syotai = IncMidByte(bytBuf, p, 20)                     '' 招待地域名
        For i = 0 To 2
            With .SaikinJyusyo(i)
                With .SaikinJyusyoid
                    .Year = IncMidByte(bytBuf, p, 4)            '' 開催年
                    .MonthDay = IncMidByte(bytBuf, p, 4)        '' 開催月日
                    .JyoCD = IncMidByte(bytBuf, p, 2)           '' 競馬場コード
                    .Kaiji = IncMidByte(bytBuf, p, 2)           '' 開催回[第N回]
                    .Nichiji = IncMidByte(bytBuf, p, 2)         '' 開催日目[N日目]
                    .RaceNum = IncMidByte(bytBuf, p, 2)         '' レース番号
                End With ' SaikinJyusyoid
                .Hondai = IncMidByte(bytBuf, p, 60)             '' 競走名本題
                .Ryakusyo10 = IncMidByte(bytBuf, p, 20)         '' 競走名略称10字
                .Ryakusyo6 = IncMidByte(bytBuf, p, 12)          '' 競走名略称6字
                .Ryakusyo3 = IncMidByte(bytBuf, p, 6)           '' 競走名略称3字
                .GradeCD = IncMidByte(bytBuf, p, 1)             '' グレードコード
                .SyussoTosu = IncMidByte(bytBuf, p, 2)          '' 出走頭数
                .KettoNum = IncMidByte(bytBuf, p, 10)           '' 血統登録番号
                .BAMEI = IncMidByte(bytBuf, p, 36)              '' 馬名
            End With ' SaikinJyusyo
        Next i
        For i = 0 To 2
            With .HonZenRuikei(i)
                .SetYear = IncMidByte(bytBuf, p, 4)             '' 設定年
                .HonSyokinHeichi = IncMidByte(bytBuf, p, 10)    '' 平地本賞金合計
                .HonSyokinSyogai = IncMidByte(bytBuf, p, 10)    '' 障害本賞金合計
                .FukaSyokinHeichi = IncMidByte(bytBuf, p, 10)   '' 平地付加賞金合計
                .FukaSyokinSyogai = IncMidByte(bytBuf, p, 10)   '' 障害付加賞金合計
                With .ChakuKaisuHeichi
                    For k = 0 To 5
                        .Chakukaisu(k) = IncMidByte(bytBuf, p, 6)
                    Next k
                End With ' ChakuKaisuHeichi
                With .ChakuKaisuSyogai
                    For k = 0 To 5
                        .Chakukaisu(k) = IncMidByte(bytBuf, p, 6)
                    Next k
                End With ' ChakuKaisuSyogai
                For j = 0 To 19
                    With .ChakuKaisuJyo(j)
                        For k = 0 To 5
                            .Chakukaisu(k) = IncMidByte(bytBuf, p, 6)
                        Next k
                    End With ' ChakuKaisuJyo
                Next j
                For j = 0 To 5
                    With .ChakuKaisuKyori(j)
                        For k = 0 To 5
                            .Chakukaisu(k) = IncMidByte(bytBuf, p, 6)
                        Next k
                    End With ' ChakuKaisuKyori
                Next j
            End With ' HonZenRuikei
        Next i
        .CRLF = IncMidByte(bytBuf, p, 2)        '' レコード区切り
    End With

    End Sub


'
'   機能: 構造体にデータをセットする－生産者マスタ
'
'   備考: なし
'
    Public Sub SetDataFromByte_BR(bytBuf() As Byte, ByRef mBuf As JV_BR_BREEDER)
    Dim i As Integer                                '' ループカウンタ
    Dim j As Integer                                '' ループカウンタ
    Dim k As Integer                                '' ループカウンタ
    Dim p As Long                                   '' 切り分け開始位置
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMidByte(bytBuf, p, 2)              '' レコード種別
            .DataKubun = IncMidByte(bytBuf, p, 1)               '' データ区分
            With .MakeDate
                .Year = IncMidByte(bytBuf, p, 4)                '' 年
                .Month = IncMidByte(bytBuf, p, 2)               '' 月
                .Day = IncMidByte(bytBuf, p, 2)                 '' 日
            End With ' MakeDate
        End With ' head
        .BreederCode = IncMidByte(bytBuf, p, 6)                 '' 生産者コード
        .BreederName_Co = IncMidByte(bytBuf, p, 70)             '' 生産者名(法人格有）
        .BreederName = IncMidByte(bytBuf, p, 70)                '' 生産者名(法人格無）
        .BreederNameKana = IncMidByte(bytBuf, p, 70)            '' 生産者名半角カナ
        .BreederNameEng = IncMidByte(bytBuf, p, 168)            '' 生産者名欧字
        .Address = IncMidByte(bytBuf, p, 20)                    '' 生産者住所自治省名
        For i = 0 To 1
            With .HonRuikei(i)
                .SetYear = IncMidByte(bytBuf, p, 4)             '' 設定年
                .HonSyokinTotal = IncMidByte(bytBuf, p, 10)     '' 本賞金合計
                .Fukasyokin = IncMidByte(bytBuf, p, 10)         '' 付加賞金合計
                For j = 0 To 5
                    .Chakukaisu(j) = IncMidByte(bytBuf, p, 6)   '' 着回数
                Next j
            End With ' HonRuikei
        Next i
        .CRLF = IncMidByte(bytBuf, p, 2)                        '' レコード区切り
    End With

    End Sub


'
'   機能: 構造体にデータをセットする－馬主マスタ
'
'   備考: なし
'
    Public Sub SetDataFromByte_BN(bytBuf() As Byte, ByRef mBuf As JV_BN_BANUSI)
    Dim i As Integer                                '' ループカウンタ
    Dim j As Integer                                '' ループカウンタ
    Dim k As Integer                                '' ループカウンタ
    Dim p As Long                                   '' 切り分け開始位置
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMidByte(bytBuf, p, 2)              '' レコード種別
            .DataKubun = IncMidByte(bytBuf, p, 1)               '' データ区分
            With .MakeDate
                .Year = IncMidByte(bytBuf, p, 4)                '' 年
                .Month = IncMidByte(bytBuf, p, 2)               '' 月
                .Day = IncMidByte(bytBuf, p, 2)                 '' 日
            End With ' MakeDate
        End With ' head
        .BanusiCode = IncMidByte(bytBuf, p, 6)                  '' 馬主コード
        .BanusiName_Co = IncMidByte(bytBuf, p, 64)              '' 馬主名（法人格有）
        .BanusiName = IncMidByte(bytBuf, p, 64)                 '' 馬主名（法人格無）
        .BanusiNameKana = IncMidByte(bytBuf, p, 50)             '' 馬主名半角カナ
        .BanusiNameEng = IncMidByte(bytBuf, p, 100)             '' 馬主名欧字
        .Fukusyoku = IncMidByte(bytBuf, p, 60)                  '' 服色標示
        For i = 0 To 1
            With .HonRuikei(i)
                .SetYear = IncMidByte(bytBuf, p, 4)             '' 設定年
                .HonSyokinTotal = IncMidByte(bytBuf, p, 10)     '' 本賞金合計
                .Fukasyokin = IncMidByte(bytBuf, p, 10)         '' 付加賞金合計
                For j = 0 To 5
                    .Chakukaisu(j) = IncMidByte(bytBuf, p, 6)   '' 着回数
                Next j
            End With ' HonRuikei
        Next i
        .CRLF = IncMidByte(bytBuf, p, 2)                        '' レコード区切り
    End With

    End Sub

'
'   機能: 構造体にデータをセットする－繁殖馬マスタ
'
'   備考: なし
'
    Public Sub SetDataFromByte_HN(bytBuf() As Byte, ByRef mBuf As JV_HN_HANSYOKU)
    Dim i As Integer                                '' ループカウンタ
    Dim j As Integer                                '' ループカウンタ
    Dim k As Integer                                '' ループカウンタ
    Dim p As Long                                   '' 切り分け開始位置
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMidByte(bytBuf, p, 2)      '' レコード種別
            .DataKubun = IncMidByte(bytBuf, p, 1)       '' データ区分
            With .MakeDate
                .Year = IncMidByte(bytBuf, p, 4)        '' 年
                .Month = IncMidByte(bytBuf, p, 2)       '' 月
                .Day = IncMidByte(bytBuf, p, 2)         '' 日
            End With ' MakeDate
        End With ' head
        .HansyokuNum = IncMidByte(bytBuf, p, 8)         '' 繁殖登録番号
        .Reserved = IncMidByte(bytBuf, p, 8)            '' 予備
        .KettoNum = IncMidByte(bytBuf, p, 10)           '' 血統登録番号
        .DelKubun = IncMidByte(bytBuf, p, 1)            '' 繁殖馬抹消区分
        .BAMEI = IncMidByte(bytBuf, p, 36)              '' 馬名
        .BameiKana = IncMidByte(bytBuf, p, 40)          '' 馬名半角カナ
        .BameiEng = IncMidByte(bytBuf, p, 80)           '' 馬名欧字
        .BirthYear = IncMidByte(bytBuf, p, 4)           '' 生年
        .SexCD = IncMidByte(bytBuf, p, 1)               '' 性別コード
        .HinsyuCD = IncMidByte(bytBuf, p, 1)            '' 品種コード
        .KeiroCD = IncMidByte(bytBuf, p, 2)             '' 毛色コード
        .HansyokuMochiKubun = IncMidByte(bytBuf, p, 1)  '' 繁殖馬持込区分
        .ImportYear = IncMidByte(bytBuf, p, 4)          '' 輸入年
        .SanchiName = IncMidByte(bytBuf, p, 20)         '' 産地名
        .HansyokuFNum = IncMidByte(bytBuf, p, 8)        '' 父馬繁殖登録番号
        .HansyokuMNum = IncMidByte(bytBuf, p, 8)        '' 母馬繁殖登録番号
        .CRLF = IncMidByte(bytBuf, p, 2)                '' レコード区切り
    End With

    End Sub


'
'   機能: 構造体にデータをセットする－産駒マスタ
'
'   備考: なし
'
    Public Sub SetDataFromByte_SK(bytBuf() As Byte, ByRef mBuf As JV_SK_SANKU)
    Dim i As Integer                                '' ループカウンタ
    Dim j As Integer                                '' ループカウンタ
    Dim k As Integer                                '' ループカウンタ
    Dim p As Long                                   '' 切り分け開始位置
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMidByte(bytBuf, p, 2)      '' レコード種別
            .DataKubun = IncMidByte(bytBuf, p, 1)       '' データ区分
            With .MakeDate
                .Year = IncMidByte(bytBuf, p, 4)        '' 年
                .Month = IncMidByte(bytBuf, p, 2)       '' 月
                .Day = IncMidByte(bytBuf, p, 2)         '' 日
            End With ' MakeDate
        End With ' head
        .KettoNum = IncMidByte(bytBuf, p, 10)           '' 血統登録番号
        With .BirthDate
            .Year = IncMidByte(bytBuf, p, 4)            '' 年
            .Month = IncMidByte(bytBuf, p, 2)           '' 月
            .Day = IncMidByte(bytBuf, p, 2)             '' 日
        End With ' BirthDate
        .SexCD = IncMidByte(bytBuf, p, 1)               '' 性別コード
        .HinsyuCD = IncMidByte(bytBuf, p, 1)            '' 品種コード
        .KeiroCD = IncMidByte(bytBuf, p, 2)             '' 毛色コード
        .SankuMochiKubun = IncMidByte(bytBuf, p, 1)     '' 産駒持込区分
        .ImportYear = IncMidByte(bytBuf, p, 4)          '' 輸入年
        .BreederCode = IncMidByte(bytBuf, p, 6)         '' 生産者コード
        .SanchiName = IncMidByte(bytBuf, p, 20)         '' 産地名
        For i = 0 To 13
            .HansyokuNum(i) = IncMidByte(bytBuf, p, 8)  '' 3代血統
        Next i
    End With

    End Sub

'
'   機能: 構造体にデータをセットする－レコードマスタ
'
'   備考: なし
'
    Public Sub SetDataFromByte_RC(bytBuf() As Byte, ByRef mBuf As JV_RC_RECORD)
    Dim i As Integer                                '' ループカウンタ
    Dim j As Integer                                '' ループカウンタ
    Dim k As Integer                                '' ループカウンタ
    Dim p As Long                                   '' 切り分け開始位置
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMidByte(bytBuf, p, 2)              '' レコード種別
            .DataKubun = IncMidByte(bytBuf, p, 1)               '' データ区分
            With .MakeDate
                .Year = IncMidByte(bytBuf, p, 4)                '' 年
                .Month = IncMidByte(bytBuf, p, 2)               '' 月
                .Day = IncMidByte(bytBuf, p, 2)                 '' 日
            End With ' MakeDate
        End With ' head
        .RecInfoKubun = IncMidByte(bytBuf, p, 1)                '' レコード識別区分
        With .id
            .Year = IncMidByte(bytBuf, p, 4)                    '' 開催年
            .MonthDay = IncMidByte(bytBuf, p, 4)                '' 開催月日
            .JyoCD = IncMidByte(bytBuf, p, 2)                   '' 競馬場コード
            .Kaiji = IncMidByte(bytBuf, p, 2)                   '' 開催回[第N回]
            .Nichiji = IncMidByte(bytBuf, p, 2)                 '' 開催日目[N日目]
            .RaceNum = IncMidByte(bytBuf, p, 2)                 '' レース番号
        End With ' id
        .TokuNum = IncMidByte(bytBuf, p, 4)                     '' 特別競走番号
        .Hondai = IncMidByte(bytBuf, p, 60)                     '' 競走名本題
        .GradeCD = IncMidByte(bytBuf, p, 1)                     '' グレードコード
        .SyubetuCD = IncMidByte(bytBuf, p, 2)                   '' 競走種別コード
        .KYORI = IncMidByte(bytBuf, p, 4)                       '' 距離
        .TrackCD = IncMidByte(bytBuf, p, 2)                     '' トラックコード
        .RecKubun = IncMidByte(bytBuf, p, 1)                    '' レコード区分
        .RecTime = IncMidByte(bytBuf, p, 4)                     '' レコードタイム
        With .TenkoBaba
            .TenkoCD = IncMidByte(bytBuf, p, 1)                 '' 天候コード
            .SibaBabaCD = IncMidByte(bytBuf, p, 1)              '' 芝馬場状態コード
            .DirtBabaCD = IncMidByte(bytBuf, p, 1)              '' ダート馬場状態コード
        End With ' TenkoBaba
        For i = 0 To 2
            With .RecUmaInfo(i)
                .KettoNum = IncMidByte(bytBuf, p, 10)           '' 血統登録番号
                .BAMEI = IncMidByte(bytBuf, p, 36)              '' 馬名
                .UmaKigoCD = IncMidByte(bytBuf, p, 2)           '' 馬記号コード
                .SexCD = IncMidByte(bytBuf, p, 1)               '' 性別コード
                .ChokyosiCode = IncMidByte(bytBuf, p, 5)        '' 調教師コード
                .ChokyosiName = IncMidByte(bytBuf, p, 34)       '' 調教師名
                .Futan = IncMidByte(bytBuf, p, 3)               '' 負担重量
                .KisyuCode = IncMidByte(bytBuf, p, 5)           '' 騎手コード
                .KisyuName = IncMidByte(bytBuf, p, 34)          '' 騎手名
            End With ' RecUmaInfo
        Next i
        .CRLF = IncMidByte(bytBuf, p, 2)                        '' レコード区切り
    End With

    End Sub


'
'   機能: 構造体にデータをセットする－坂路調教
'
'   備考: なし
'
    Public Sub SetDataFromByte_HC(lBuf() As Byte, ByRef mBuf As JV_HC_HANRO)
    Dim i As Integer                                '' ループカウンタ
    Dim j As Integer                                '' ループカウンタ
    Dim k As Integer                                '' ループカウンタ
    Dim p As Long                                   '' 切り分け開始位置
    
    With mBuf
        With .head
            .RecordSpec = Mid$(lBuf, 1, 2)  '' レコード種別
            .DataKubun = Mid$(lBuf, 3, 1)   '' データ区分
            With .MakeDate
                .Year = Mid$(lBuf, 4, 4)    '' 年
                .Month = Mid$(lBuf, 8, 2)   '' 月
                .Day = Mid$(lBuf, 10, 2)     '' 日
            End With ' MakeDate
        End With ' head
        .TresenKubun = Mid$(lBuf, 12, 1)     '' トレセン区分
        With .ChokyoDate
            .Year = Mid$(lBuf, 13, 4)        '' 年
            .Month = Mid$(lBuf, 17, 2)       '' 月
            .Day = Mid$(lBuf, 19, 2)         '' 日
        End With ' ChokyoDate
        .ChokyoTime = Mid$(lBuf, 21, 4)      '' 調教時刻
        .KettoNum = Mid$(lBuf, 25, 10)       '' 血統登録番号
        .HaronTime4 = Mid$(lBuf, 35, 4)      '' 4ハロンタイム合計(800M-0M)
        .LapTime4 = Mid$(lBuf, 39, 3)        '' ラップタイム(800M-600M)
        .HaronTime3 = Mid$(lBuf, 42, 4)      '' 3ハロンタイム合計(600M-0M)
        .LapTime3 = Mid$(lBuf, 46, 3)        '' ラップタイム(600M-400M)
        .HaronTime2 = Mid$(lBuf, 49, 4)      '' 2ハロンタイム合計(400M-0M)
        .LapTime2 = Mid$(lBuf, 53, 3)        '' ラップタイム(400M-200M)
        .LapTime1 = Mid$(lBuf, 56, 3)        '' ラップタイム(200M-0M)
        .CRLF = Mid$(lBuf, 59, 2)            '' レコード区切り
    End With

  End Sub


'
'   機能: 構造体にデータをセットする－馬体重
'
'   備考: なし
'
    Public Sub SetDataFromByte_WH(bytBuf() As Byte, ByRef mBuf As JV_WH_BATAIJYU)
    Dim i As Integer                                '' ループカウンタ
    Dim j As Integer                                '' ループカウンタ
    Dim k As Integer                                '' ループカウンタ
    Dim p As Long                                   '' 切り分け開始位置
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMidByte(bytBuf, p, 2)      '' レコード種別
            .DataKubun = IncMidByte(bytBuf, p, 1)       '' データ区分
            With .MakeDate
                .Year = IncMidByte(bytBuf, p, 4)        '' 年
                .Month = IncMidByte(bytBuf, p, 2)       '' 月
                .Day = IncMidByte(bytBuf, p, 2)         '' 日
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMidByte(bytBuf, p, 4)            '' 開催年
            .MonthDay = IncMidByte(bytBuf, p, 4)        '' 開催月日
            .JyoCD = IncMidByte(bytBuf, p, 2)           '' 競馬場コード
            .Kaiji = IncMidByte(bytBuf, p, 2)           '' 開催回[第N回]
            .Nichiji = IncMidByte(bytBuf, p, 2)         '' 開催日目[N日目]
            .RaceNum = IncMidByte(bytBuf, p, 2)         '' レース番号
        End With ' id
        With .HappyoTime
            .Month = IncMidByte(bytBuf, p, 2)           '' 月
            .Day = IncMidByte(bytBuf, p, 2)             '' 日
            .Hour = IncMidByte(bytBuf, p, 2)            '' 時
            .Minute = IncMidByte(bytBuf, p, 2)          '' 分
        End With ' HappyoTime
        For i = 0 To 17
            With .BataijyuInfo(i)
                .Umaban = IncMidByte(bytBuf, p, 2)      '' 馬番
                .BAMEI = IncMidByte(bytBuf, p, 36)      '' 馬名
                .BaTaijyu = IncMidByte(bytBuf, p, 3)    '' 馬体重
                .ZogenFugo = IncMidByte(bytBuf, p, 1)   '' 増減符号
                .ZogenSa = IncMidByte(bytBuf, p, 3)     '' 増減差
            End With ' BataijyuInfo
        Next i
        .CRLF = IncMidByte(bytBuf, p, 2)                '' レコード区切り
    End With

    End Sub


'
'   機能: 構造体にデータをセットする－天候馬場状態
'
'   備考: なし
'
    Public Sub SetDataFromByte_WE(bytBuf() As Byte, ByRef mBuf As JV_WE_WEATHER)
    Dim i As Integer                                '' ループカウンタ
    Dim j As Integer                                '' ループカウンタ
    Dim k As Integer                                '' ループカウンタ
    Dim p As Long                                   '' 切り分け開始位置
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMidByte(bytBuf, p, 2)      '' レコード種別
            .DataKubun = IncMidByte(bytBuf, p, 1)       '' データ区分
            With .MakeDate
                .Year = IncMidByte(bytBuf, p, 4)        '' 年
                .Month = IncMidByte(bytBuf, p, 2)       '' 月
                .Day = IncMidByte(bytBuf, p, 2)         '' 日
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMidByte(bytBuf, p, 4)            '' 開催年
            .MonthDay = IncMidByte(bytBuf, p, 4)        '' 開催月日
            .JyoCD = IncMidByte(bytBuf, p, 2)           '' 競馬場コード
            .Kaiji = IncMidByte(bytBuf, p, 2)           '' 開催回[第N回]
            .Nichiji = IncMidByte(bytBuf, p, 2)         '' 開催日目[N日目]
        End With ' id
        With .HappyoTime
            .Month = IncMidByte(bytBuf, p, 2)           '' 月
            .Day = IncMidByte(bytBuf, p, 2)             '' 日
            .Hour = IncMidByte(bytBuf, p, 2)            '' 時
            .Minute = IncMidByte(bytBuf, p, 2)          '' 分
        End With ' HappyoTime
        .HenkoID = IncMidByte(bytBuf, p, 1)             '' 変更識別
        With .TenkoBaba
            .TenkoCD = IncMidByte(bytBuf, p, 1)         '' 天候コード
            .SibaBabaCD = IncMidByte(bytBuf, p, 1)      '' 芝馬場状態コード
            .DirtBabaCD = IncMidByte(bytBuf, p, 1)      '' ダート馬場状態コード
        End With ' TenkoBaba
        With .TenkoBabaBefore
            .TenkoCD = IncMidByte(bytBuf, p, 1)         '' 天候コード
            .SibaBabaCD = IncMidByte(bytBuf, p, 1)      '' 芝馬場状態コード
            .DirtBabaCD = IncMidByte(bytBuf, p, 1)      '' ダート馬場状態コード
        End With ' TenkoBabaBefore
        .CRLF = IncMidByte(bytBuf, p, 2)                '' レコード区切り
    End With

    End Sub


'
'   機能: 構造体にデータをセットする－出走取消・競争除外
'
'   備考: なし
'
    Public Sub SetDataFromByte_AV(bytBuf() As Byte, ByRef mBuf As JV_AV_INFO)
    Dim i As Integer                                '' ループカウンタ
    Dim j As Integer                                '' ループカウンタ
    Dim k As Integer                                '' ループカウンタ
    Dim p As Long                                   '' 切り分け開始位置
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMidByte(bytBuf, p, 2)      '' レコード種別
            .DataKubun = IncMidByte(bytBuf, p, 1)       '' データ区分
            With .MakeDate
                .Year = IncMidByte(bytBuf, p, 4)        '' 年
                .Month = IncMidByte(bytBuf, p, 2)       '' 月
                .Day = IncMidByte(bytBuf, p, 2)         '' 日
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMidByte(bytBuf, p, 4)            '' 開催年
            .MonthDay = IncMidByte(bytBuf, p, 4)        '' 開催月日
            .JyoCD = IncMidByte(bytBuf, p, 2)           '' 競馬場コード
            .Kaiji = IncMidByte(bytBuf, p, 2)           '' 開催回[第N回]
            .Nichiji = IncMidByte(bytBuf, p, 2)         '' 開催日目[N日目]
            .RaceNum = IncMidByte(bytBuf, p, 2)         '' レース番号
        End With ' id
        With .HappyoTime
            .Month = IncMidByte(bytBuf, p, 2)           '' 月
            .Day = IncMidByte(bytBuf, p, 2)             '' 日
            .Hour = IncMidByte(bytBuf, p, 2)            '' 時
            .Minute = IncMidByte(bytBuf, p, 2)          '' 分
        End With ' HappyoTime
        .Umaban = IncMidByte(bytBuf, p, 2)              '' 馬番
        .BAMEI = IncMidByte(bytBuf, p, 36)              '' 馬名
        .JiyuKubun = IncMidByte(bytBuf, p, 3)           '' 事由区分
        .CRLF = IncMidByte(bytBuf, p, 2)                '' レコード区切り
    End With

    End Sub

'
'   機能: 構造体にデータをセットする－騎手変更
'
'   備考: なし
'
    Public Sub SetDataFromByte_JC(bytBuf() As Byte, ByRef mBuf As JV_JC_INFO)
    Dim i As Integer                                '' ループカウンタ
    Dim j As Integer                                '' ループカウンタ
    Dim k As Integer                                '' ループカウンタ
    Dim p As Long                                   '' 切り分け開始位置
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMidByte(bytBuf, p, 2)  '' レコード種別
            .DataKubun = IncMidByte(bytBuf, p, 1)   '' データ区分
            With .MakeDate
                .Year = IncMidByte(bytBuf, p, 4)    '' 年
                .Month = IncMidByte(bytBuf, p, 2)   '' 月
                .Day = IncMidByte(bytBuf, p, 2)     '' 日
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMidByte(bytBuf, p, 4)        '' 開催年
            .MonthDay = IncMidByte(bytBuf, p, 4)        '' 開催月日
            .JyoCD = IncMidByte(bytBuf, p, 2)       '' 競馬場コード
            .Kaiji = IncMidByte(bytBuf, p, 2)       '' 開催回[第N回]
            .Nichiji = IncMidByte(bytBuf, p, 2)     '' 開催日目[N日目]
            .RaceNum = IncMidByte(bytBuf, p, 2)     '' レース番号
        End With ' id
        With .HappyoTime
            .Month = IncMidByte(bytBuf, p, 2)       '' 月
            .Day = IncMidByte(bytBuf, p, 2)         '' 日
            .Hour = IncMidByte(bytBuf, p, 2)        '' 時
            .Minute = IncMidByte(bytBuf, p, 2)      '' 分
        End With ' HappyoTime
        .Umaban = IncMidByte(bytBuf, p, 2)          '' 馬番
        .BAMEI = IncMidByte(bytBuf, p, 36)          '' 馬名
        With .JCInfoAfter
            .Futan = IncMidByte(bytBuf, p, 3)       '' 負担重量
            .KisyuCode = IncMidByte(bytBuf, p, 5)   '' 騎手コード
            .KisyuName = IncMidByte(bytBuf, p, 34)  '' 騎手名
            .MinaraiCD = IncMidByte(bytBuf, p, 1)   '' 騎手見習コード
        End With ' JCInfoAfter
        With .JCInfoBefore
            .Futan = IncMidByte(bytBuf, p, 3)       '' 負担重量
            .KisyuCode = IncMidByte(bytBuf, p, 5)   '' 騎手コード
            .KisyuName = IncMidByte(bytBuf, p, 34)  '' 騎手名
            .MinaraiCD = IncMidByte(bytBuf, p, 1)   '' 騎手見習コード
        End With ' JCInfoBefore
        .CRLF = IncMidByte(bytBuf, p, 2)            '' レコード区切り
    End With

    End Sub

'
'   機能: 構造体にデータをセットする－データマイニング予想
'
'   備考: なし
'
    Public Sub SetDataFromByte_DM(bytBuf() As Byte, ByRef mBuf As JV_DM_INFO)
    Dim i As Integer                                '' ループカウンタ
    Dim j As Integer                                '' ループカウンタ
    Dim k As Integer                                '' ループカウンタ
    Dim p As Long                                   '' 切り分け開始位置
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMidByte(bytBuf, p, 2)  '' レコード種別
            .DataKubun = IncMidByte(bytBuf, p, 1)   '' データ区分
            With .MakeDate
                .Year = IncMidByte(bytBuf, p, 4)    '' 年
                .Month = IncMidByte(bytBuf, p, 2)   '' 月
                .Day = IncMidByte(bytBuf, p, 2)     '' 日
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMidByte(bytBuf, p, 4)        '' 開催年
            .MonthDay = IncMidByte(bytBuf, p, 4)    '' 開催月日
            .JyoCD = IncMidByte(bytBuf, p, 2)       '' 競馬場コード
            .Kaiji = IncMidByte(bytBuf, p, 2)       '' 開催回[第N回]
            .Nichiji = IncMidByte(bytBuf, p, 2)     '' 開催日目[N日目]
            .RaceNum = IncMidByte(bytBuf, p, 2)     '' レース番号
        End With ' id
        With .MakeHM
            .Hour = IncMidByte(bytBuf, p, 2)        '' 時
            .Minute = IncMidByte(bytBuf, p, 2)      '' 分
        End With ' MakeHM
        For i = 0 To 17
            With .DMInfo(i)
                .Umaban = IncMidByte(bytBuf, p, 2)  '' 馬番
                .DMTime = IncMidByte(bytBuf, p, 5)  '' 予想走破タイム
                .DMGosaP = IncMidByte(bytBuf, p, 4) '' 予想誤差(信頼度)＋
                .DMGosaM = IncMidByte(bytBuf, p, 4) '' 予想誤差(信頼度)－
            End With ' DMInfo
        Next i
        .CRLF = IncMidByte(bytBuf, p, 2)            '' レコード区切り
    End With

    End Sub


'
'   機能: 構造体にデータをセットする－開催スケジュール
'
'   備考: なし
'
    Public Sub SetDataFromByte_YS(bytBuf() As Byte, ByRef mBuf As JV_YS_SCHEDULE)
    Dim i As Integer                                '' ループカウンタ
    Dim j As Integer                                '' ループカウンタ
    Dim k As Integer                                '' ループカウンタ
    Dim p As Long                                   '' 切り分け開始位置
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMidByte(bytBuf, p, 2)      '' レコード種別
            .DataKubun = IncMidByte(bytBuf, p, 1)       '' データ区分
            With .MakeDate
                .Year = IncMidByte(bytBuf, p, 4)        '' 年
                .Month = IncMidByte(bytBuf, p, 2)       '' 月
                .Day = IncMidByte(bytBuf, p, 2)         '' 日
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMidByte(bytBuf, p, 4)            '' 開催年
            .MonthDay = IncMidByte(bytBuf, p, 4)        '' 開催月日
            .JyoCD = IncMidByte(bytBuf, p, 2)           '' 競馬場コード
            .Kaiji = IncMidByte(bytBuf, p, 2)           '' 開催回[第N回]
            .Nichiji = IncMidByte(bytBuf, p, 2)         '' 開催日目[N日目]
        End With ' id
        .YoubiCD = IncMidByte(bytBuf, p, 1)             '' 曜日コード
        For i = 0 To 2
            With .JyusyoInfo(i)
                .TokuNum = IncMidByte(bytBuf, p, 4)     '' 特別競走番号
                .Hondai = IncMidByte(bytBuf, p, 60)     '' 競走名本題
                .Ryakusyo10 = IncMidByte(bytBuf, p, 20) '' 競走名略称10字
                .Ryakusyo6 = IncMidByte(bytBuf, p, 12)  '' 競走名略称6字
                .Ryakusyo3 = IncMidByte(bytBuf, p, 6)   '' 競走名略称3字
                .Nkai = IncMidByte(bytBuf, p, 3)        '' 重賞回次[第N回]
                .GradeCD = IncMidByte(bytBuf, p, 1)     '' グレードコード
                .SyubetuCD = IncMidByte(bytBuf, p, 2)   '' 競走種別コード
                .KigoCD = IncMidByte(bytBuf, p, 3)      '' 競走記号コード
                .JyuryoCD = IncMidByte(bytBuf, p, 1)    '' 重量種別コード
                .KYORI = IncMidByte(bytBuf, p, 4)       '' 距離
                .TrackCD = IncMidByte(bytBuf, p, 2)     '' トラックコード
            End With ' JyusyoInfo
        Next i
        .CRLF = IncMidByte(bytBuf, p, 2)                '' レコード区切り
    End With

    End Sub
    
    
'
'   機能: 構造体にデータをセットする－発走時刻変更
'
'   備考: なし
'
    Public Sub SetDataFromByte_TC(bytBuf() As Byte, ByRef mBuf As JV_TC_HASSOU)

        Dim p As Long
    
        p = 1
        With mBuf
            With .head
                .RecordSpec = IncMidByte(bytBuf, p, 2)      '' レコード種別
                .DataKubun = IncMidByte(bytBuf, p, 1)       '' データ区分
                With .MakeDate
                    .Year = IncMidByte(bytBuf, p, 4)        '' 年
                    .Month = IncMidByte(bytBuf, p, 2)       '' 月
                    .Day = IncMidByte(bytBuf, p, 2)         '' 日
                End With
            End With
            With .id
                .Year = IncMidByte(bytBuf, p, 4)            '' 開催年
                .MonthDay = IncMidByte(bytBuf, p, 4)        '' 開催月日
                .JyoCD = IncMidByte(bytBuf, p, 2)           '' 競馬場コード
                .Kaiji = IncMidByte(bytBuf, p, 2)           '' 開催回[第N回]
                .Nichiji = IncMidByte(bytBuf, p, 2)         '' 開催日目[N日目]
                .RaceNum = IncMidByte(bytBuf, p, 2)         '' レース番号
            End With
            With .HappyoTime                                '' 発表月日時分
                .Month = IncMidByte(bytBuf, p, 2)
                .Day = IncMidByte(bytBuf, p, 2)
                .Hour = IncMidByte(bytBuf, p, 2)
                .Minute = IncMidByte(bytBuf, p, 2)
            End With
            .AtoHassoTime.Hour = IncMidByte(bytBuf, p, 2)            '' 変更後時
            .AtoHassoTime.Minute = IncMidByte(bytBuf, p, 2)          '' 変更後分
            .MaeHassoTime.Hour = IncMidByte(bytBuf, p, 2)            '' 変更前時
            .MaeHassoTime.Minute = IncMidByte(bytBuf, p, 2)          '' 変更前分
            .CRLF = IncMidByte(bytBuf, p, 2)                '' レコード区切り
        End With
        
    End Sub
    
'
'   機能: 構造体にデータをセットする－コース変更
'
'   備考: なし
'
    Public Sub SetDataFromByte_CC(bytBuf() As Byte, ByRef mBuf As JV_CC_COURSE)
    
        Dim p As Long
        
        p = 1
        With mBuf
            With .head
                .RecordSpec = IncMidByte(bytBuf, p, 2)      '' レコード種別
                .DataKubun = IncMidByte(bytBuf, p, 1)       '' データ区分
                With .MakeDate
                    .Year = IncMidByte(bytBuf, p, 4)        '' 年
                    .Month = IncMidByte(bytBuf, p, 2)       '' 月
                    .Day = IncMidByte(bytBuf, p, 2)         '' 日
                End With
            End With
            With .id
                .Year = IncMidByte(bytBuf, p, 4)            '' 開催年
                .MonthDay = IncMidByte(bytBuf, p, 4)        '' 開催月日
                .JyoCD = IncMidByte(bytBuf, p, 2)           '' 競馬場コード
                .Kaiji = IncMidByte(bytBuf, p, 2)           '' 開催回[第N回]
                .Nichiji = IncMidByte(bytBuf, p, 2)         '' 開催日目[N日目]
                .RaceNum = IncMidByte(bytBuf, p, 2)         '' レース番号
            End With
            With .HappyoTime                                '' 発表月日時分
                .Month = IncMidByte(bytBuf, p, 2)
                .Day = IncMidByte(bytBuf, p, 2)
                .Hour = IncMidByte(bytBuf, p, 2)
                .Minute = IncMidByte(bytBuf, p, 2)
            End With
            .AtoKyori = IncMidByte(bytBuf, p, 4)            '' 変更後距離
            .AtoTrackCD = IncMidByte(bytBuf, p, 2)          '' 変更後トラックコード
            .MaeKyori = IncMidByte(bytBuf, p, 4)            '' 変更前距離
            .MaeTrackCD = IncMidByte(bytBuf, p, 2)          '' 変更前トラックコード
            .JiyuKubun = IncMidByte(bytBuf, p, 1)           '' 事由コード
            .CRLF = IncMidByte(bytBuf, p, 2)                '' レコード区切り
        End With
        
    End Sub
