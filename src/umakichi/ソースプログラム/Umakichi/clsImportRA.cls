VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsImportRA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Import JV_RA_RACE "
'
'   JVData "RA" データベース登録クラス
'

Option Explicit
Option Compare Binary
Implements clsIImport

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部変数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private mBuf As JV_RA_RACE

Private mCon_RACE As ADODB.Connection
Private mCon_UMA_RACE_A As ADODB.Connection
Private mCon_UMA_RACE_B As ADODB.Connection

Private mRS_RACE As ADODB.Recordset


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部関数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: クラス初期化イベント
'
'   備考: なし
'
Private Sub Class_Initialize()
On Error GoTo Errorhandler
    Dim strCon As String
    ' コネクションのインスタンス生成
    Set mCon_RACE = New ADODB.Connection
    Set mCon_UMA_RACE_A = New ADODB.Connection
    Set mCon_UMA_RACE_B = New ADODB.Connection

    ' レコードセットのインスタンス生成
    Set mRS_RACE = New ADODB.Recordset
    
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   機能: コネクション、レコードセットをオープンする
'
'   備考: なし
'
Private Sub clsIImport_OpenDB()
On Error GoTo Errorhandler        ' コネクションオープン
    Dim strCon As String
    
    
    ' コネクションオープン
    strCon = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & gApp.R_DBPath & "\"
    mCon_RACE.Open strCon & "subRACE.mdb"
    mCon_UMA_RACE_A.Open strCon & "subUMA_RACE_A.mdb"
    mCon_UMA_RACE_B.Open strCon & "subUMA_RACE_B.mdb"

    ' レコードセットオープン
    With mRS_RACE
        .CursorLocation = adUseServer
        .Index = "PrimaryKey"
        .Open "RACE", mCon_RACE, adOpenKeyset, adLockOptimistic, adCmdTableDirect
        If Not (.EOF Or .BOF) Then
            .MoveFirst
        End If
    End With


    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   機能: コネクション、レコードセットをクローズする
'
'   備考: なし
'
Private Sub clsIImport_CloseDB()
On Error GoTo Errorhandler
    mRS_RACE.Close

    
    mCon_RACE.Close
    mCon_UMA_RACE_A.Close
    mCon_UMA_RACE_B.Close

    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   機能: JVReadの返す１行をデータベースに登録する
'
'   備考: DBに追加を試み、失敗したら更新を試みる
'
Private Function clsIImport_Add(lBuf() As Byte) As Boolean
On Error GoTo Errorhandler

    ' データ区分を確認して、処理を決める
    If lBuf(2) = ASCII_ZERO Then
        Call DeleteRecord(StrConv(lBuf, vbUnicode))
        clsIImport_Add = True
    Else
        ' 木曜出馬表のデータを削除する
        If lBuf(2) >= ASCII_TWO And lBuf(2) <= ASCII_SEVEN Then
            Call DeleteSEMokuyouRecord(StrConv(lBuf, vbUnicode))
        End If

        ' 一行データを構造体に切り分ける
        Call SetDataFromByte_RA(lBuf, mBuf)                                                                 '' 構造体に代入する

        ' データ登録
        If Not InsertDB() Then
            ' 新規挿入に失敗したら更新を試みる
            clsIImport_Add = UpdateDB()
        Else
            clsIImport_Add = True
        End If

    End If

    Exit Function

Errorhandler:
    gApp.ErrLog
    clsIImport_Add = False
End Function


'
'   機能: レコードを削除する
'
'   備考: なし
'
Private Sub DeleteRecord(lBuf As String)
    Dim strSQL      As String
    Dim Year        As String
    Dim MonthDay    As String
    Dim JyoCD       As String
    Dim Kaiji       As String
    Dim Nichiji     As String
    Dim RaceNum     As String
    
    Year = Mid$(lBuf, 12, 4)                   '' 開催年
    MonthDay = Mid$(lBuf, 16, 4)               '' 開催月日
    JyoCD = Mid$(lBuf, 20, 2)                  '' 競馬場コード
    Kaiji = Mid$(lBuf, 22, 2)                  '' 開催回[第N回]
    Nichiji = Mid$(lBuf, 24, 2)                '' 開催日目[N日目]
    RaceNum = Mid$(lBuf, 26, 2)                '' レース番号

    strSQL = "DELETE * FROM RACE"
    strSQL = strSQL & " WHERE [Year]   ='" & Year & "'"
    strSQL = strSQL & " AND [MonthDay] ='" & MonthDay & "'"
    strSQL = strSQL & " AND [JyoCD]    ='" & JyoCD & "'"
    strSQL = strSQL & " AND [Kaiji]    ='" & Kaiji & "'"
    strSQL = strSQL & " AND [Nichiji]  ='" & Nichiji & "'"
    strSQL = strSQL & " AND [RaceNum]  ='" & RaceNum & "'"
    
    mCon_RACE.Execute strSQL, , adExecuteNoRecords
    
End Sub


'
'   機能: SEのレコードを削除する
'
'   備考: なし
'
Private Sub DeleteSEMokuyouRecord(lBuf As String)
    Dim strSQL      As String
    Dim Year        As String
    Dim MonthDay    As String
    Dim JyoCD       As String
    Dim Kaiji       As String
    Dim Nichiji     As String
    Dim RaceNum     As String
    
    Year = Mid$(lBuf, 12, 4)                   '' 開催年
    MonthDay = Mid$(lBuf, 16, 4)               '' 開催月日
    JyoCD = Mid$(lBuf, 20, 2)                  '' 競馬場コード
    Kaiji = Mid$(lBuf, 22, 2)                  '' 開催回[第N回]
    Nichiji = Mid$(lBuf, 24, 2)                '' 開催日目[N日目]
    RaceNum = Mid$(lBuf, 26, 2)                '' レース番号

    strSQL = "DELETE * FROM UMA_RACE_A"
    strSQL = strSQL & " WHERE [Year]   ='" & Year & "'"
    strSQL = strSQL & " AND [MonthDay] ='" & MonthDay & "'"
    strSQL = strSQL & " AND [JyoCD]    ='" & JyoCD & "'"
    strSQL = strSQL & " AND [Kaiji]    ='" & Kaiji & "'"
    strSQL = strSQL & " AND [Nichiji]  ='" & Nichiji & "'"
    strSQL = strSQL & " AND [RaceNum]  ='" & RaceNum & "'"
    strSQL = strSQL & " AND [DataKubun]  ='1'"
    
    mCon_UMA_RACE_A.Execute strSQL, , adExecuteNoRecords
    
    strSQL = "DELETE * FROM UMA_RACE_B"
    strSQL = strSQL & " WHERE [B_Year]   ='" & Year & "'"
    strSQL = strSQL & " AND [B_MonthDay] ='" & MonthDay & "'"
    strSQL = strSQL & " AND [B_JyoCD]    ='" & JyoCD & "'"
    strSQL = strSQL & " AND [B_RaceNum]  ='" & RaceNum & "'"
    strSQL = strSQL & " AND [DataKubun]  ='1'"
    
    mCon_UMA_RACE_B.Execute strSQL, , adExecuteNoRecords

End Sub


'
'   機能: データベースに挿入する
'
'   備考: なし
'
Private Function InsertDB() As Boolean
On Error GoTo Errorhandler
    Dim i As Integer                                                                            '' ループカウンタ
    Dim j As Integer                                                                            '' ループカウンタ
    Dim k As Integer                                                                            '' ループカウンタ
    Dim rs As ADODB.Recordset                                                                   '' SQL文
    
    mCon_RACE.BeginTrans

    
    Set rs = mRS_RACE
    With mBuf.id
        rs.AddNew Array("Year", "MonthDay", "JyoCD", "Kaiji", "Nichiji", "RaceNum"), Array(.Year, .MonthDay, .JyoCD, .Kaiji, .Nichiji, .RaceNum)
    End With
    With mBuf
        With .head
            rs("RecordSpec") = .RecordSpec                                                      '' レコード種別
            rs("DataKubun") = .DataKubun                                                        '' データ区分
            With .MakeDate
                rs("MakeDate") = .Year & .Month & .Day                                          '' 年月日
            End With ' MakeDate
        End With ' head
        With .RaceInfo
            rs("YoubiCD") = .YoubiCD                                                            '' 曜日コード
            rs("TokuNum") = .TokuNum                                                            '' 特別競走番号
            rs("Hondai") = .Hondai                                                              '' 競走名本題
            rs("Fukudai") = .Fukudai                                                            '' 競走名副題
            rs("Kakko") = .Kakko                                                                '' 競走名カッコ内
            rs("HondaiEng") = .HondaiEng                                                        '' 競走名本題欧字
            rs("FukudaiEng") = .FukudaiEng                                                      '' 競走名副題欧字
            rs("KakkoEng") = .KakkoEng                                                          '' 競走名カッコ内欧字
            rs("Ryakusyo10") = .Ryakusyo10                                                      '' 競走名略称１０字
            rs("Ryakusyo6") = .Ryakusyo6                                                        '' 競走名略称６字
            rs("Ryakusyo3") = .Ryakusyo3                                                        '' 競走名略称３字
            rs("Kubun") = .Kubun                                                                '' 競走名区分
            rs("Nkai") = .Nkai                                                                  '' 重賞回次[第N回]
        End With ' RaceInfo
        rs("GradeCD") = .GradeCD                                                                '' グレードコード
        rs("GradeCDBefore") = .GradeCDBefore                                                    '' 変更前グレードコード
        With .JyokenInfo
            rs("SyubetuCD") = .SyubetuCD                                                        '' 競走種別コード
            rs("KigoCD") = .KigoCD                                                              '' 競走記号コード
            rs("JyuryoCD") = .JyuryoCD                                                          '' 重量種別コード
            For j = 0 To 4
                rs("JyokenCD" & j + 1) = .JyokenCD(j)                                           '' 競走条件コード
            Next j
        End With ' JyokenInfo
        rs("JyokenName") = .JyokenName                                                          '' 競走条件名称
        rs("Kyori") = .KYORI                                                                    '' 距離
        rs("KyoriBefore") = .KyoriBefore                                                        '' 変更前距離
        rs("TrackCD") = .TrackCD                                                                '' トラックコード
        rs("TrackCDBefore") = .TrackCDBefore                                                    '' 変更前トラックコード
        rs("CourseKubunCD") = .CourseKubunCD                                                    '' コース区分
        rs("CourseKubunCDBefore") = .CourseKubunCDBefore                                        '' 変更前コース区分
        For i = 0 To 6
            rs("Honsyokin" & i + 1) = .Honsyokin(i)                                             '' 本賞金
        Next i
        For i = 0 To 4
            rs("HonsyokinBefore" & i + 1) = .HonsyokinBefore(i)                                 '' 変更前本賞金
        Next i
        For i = 0 To 4
            rs("Fukasyokin" & i + 1) = .Fukasyokin(i)                                           '' 付加賞金
        Next i
        For i = 0 To 2
            rs("FukasyokinBefore" & i + 1) = .FukasyokinBefore(i)                               '' 変更前付加賞金
        Next i
        rs("HassoTime") = .HassoTime                                                            '' 発走時刻
        rs("HassoTimeBefore") = .HassoTimeBefore                                                '' 変更前発走時刻
        rs("TorokuTosu") = .TorokuTosu                                                          '' 登録頭数
        rs("SyussoTosu") = .SyussoTosu                                                          '' 出走頭数
        rs("NyusenTosu") = .NyusenTosu                                                          '' 入線頭数
        With .TenkoBaba
            rs("TenkoCD") = .TenkoCD                                                            '' 天候コード
            rs("SibaBabaCD") = .SibaBabaCD                                                      '' 芝馬場状態コード
            rs("DirtBabaCD") = .DirtBabaCD                                                      '' ダート馬場状態コード
        End With ' TenkoBaba
        For i = 0 To 24
            rs("LapTime" & i + 1) = .LapTime(i)                                                 '' ラップタイム
        Next i
        rs("SyogaiMileTime") = .SyogaiMileTime                                                  '' 障害マイルタイム
        rs("HaronTimeS3") = .HaronTimeS3                                                        '' 前３ハロンタイム
        rs("HaronTimeS4") = .HaronTimeS4                                                        '' 前４ハロンタイム
        rs("HaronTimeL3") = .HaronTimeL3                                                        '' 後３ハロンタイム
        rs("HaronTimeL4") = .HaronTimeL4                                                        '' 後４ハロンタイム
        For i = 0 To 3
            With .CornerInfo(i)
                rs("Corner" & i + 1) = .Corner                                                  '' コーナー
                rs("Syukaisu" & i + 1) = .Syukaisu                                              '' 周回数
                rs("Jyuni" & i + 1) = .Jyuni                                                    '' 各通過順位
            End With ' CornerInfo
        Next i
        rs("RecordUpKubun") = .RecordUpKubun                                                    '' レコード更新区分
    End With

    rs.Update ' RACE
    
    mCon_RACE.CommitTrans

    Set rs = Nothing
    
    InsertDB = True
    Exit Function
    
Errorhandler:
    If Err.Number <> -2147217887 Then
        gApp.ErrLog
    End If
    rs.CancelUpdate
    
    mCon_RACE.RollbackTrans

    InsertDB = False
End Function


'
'   機能: データベースを更新する
'
'   備考: なし
'
Private Function UpdateDB() As Boolean
On Error GoTo Errorhandler
    Dim rs As ADODB.Recordset
    Dim i As Integer                                                                            '' ループカウンタ
    Dim j As Integer                                                                            '' ループカウンタ
    Dim k As Integer                                                                            '' ループカウンタ
    Dim strSQL As String                                                                        '' SQL文
    
    Set rs = mRS_RACE

    With mBuf.id
        rs.Seek Array(.Year, .MonthDay, .JyoCD, .Kaiji, .Nichiji, .RaceNum)
    End With

    With mBuf.head.MakeDate
        If rs("Makedate") > .Year & .Month & .Day Then
            UpdateDB = True
            Exit Function
        End If
    End With

    
    mCon_RACE.BeginTrans

    With mBuf
        With .head
            rs("RecordSpec") = .RecordSpec                                                      '' レコード種別
            rs("DataKubun") = .DataKubun                                                        '' データ区分
            With .MakeDate
                rs("MakeDate") = .Year & .Month & .Day                                          '' 年月日
            End With ' MakeDate
        End With ' head
        With .RaceInfo
            rs("YoubiCD") = .YoubiCD                                                            '' 曜日コード
            rs("TokuNum") = .TokuNum                                                            '' 特別競走番号
            rs("Hondai") = .Hondai                                                              '' 競走名本題
            rs("Fukudai") = .Fukudai                                                            '' 競走名副題
            rs("Kakko") = .Kakko                                                                '' 競走名カッコ内
            rs("HondaiEng") = .HondaiEng                                                        '' 競走名本題欧字
            rs("FukudaiEng") = .FukudaiEng                                                      '' 競走名副題欧字
            rs("KakkoEng") = .KakkoEng                                                          '' 競走名カッコ内欧字
            rs("Ryakusyo10") = .Ryakusyo10                                                      '' 競走名略称１０字
            rs("Ryakusyo6") = .Ryakusyo6                                                        '' 競走名略称６字
            rs("Ryakusyo3") = .Ryakusyo3                                                        '' 競走名略称３字
            rs("Kubun") = .Kubun                                                                '' 競走名区分
            rs("Nkai") = .Nkai                                                                  '' 重賞回次[第N回]
        End With ' RaceInfo
        rs("GradeCD") = .GradeCD                                                                '' グレードコード
        rs("GradeCDBefore") = .GradeCDBefore                                                    '' 変更前グレードコード
        With .JyokenInfo
            rs("SyubetuCD") = .SyubetuCD                                                        '' 競走種別コード
            rs("KigoCD") = .KigoCD                                                              '' 競走記号コード
            rs("JyuryoCD") = .JyuryoCD                                                          '' 重量種別コード
            For j = 0 To 4
                rs("JyokenCD" & j + 1) = .JyokenCD(j)                                           '' 競走条件コード
            Next j
        End With ' JyokenInfo
        rs("JyokenName") = .JyokenName                                                          '' 競走条件名称
        rs("Kyori") = .KYORI                                                                    '' 距離
        rs("KyoriBefore") = .KyoriBefore                                                        '' 変更前距離
        rs("TrackCD") = .TrackCD                                                                '' トラックコード
        rs("TrackCDBefore") = .TrackCDBefore                                                    '' 変更前トラックコード
        rs("CourseKubunCD") = .CourseKubunCD                                                    '' コース区分
        rs("CourseKubunCDBefore") = .CourseKubunCDBefore                                        '' 変更前コース区分
        For i = 0 To 6
            rs("Honsyokin" & i + 1) = .Honsyokin(i)                                             '' 本賞金
        Next i
        For i = 0 To 4
            rs("HonsyokinBefore" & i + 1) = .HonsyokinBefore(i)                                 '' 変更前本賞金
        Next i
        For i = 0 To 4
            rs("Fukasyokin" & i + 1) = .Fukasyokin(i)                                           '' 付加賞金
        Next i
        For i = 0 To 2
            rs("FukasyokinBefore" & i + 1) = .FukasyokinBefore(i)                               '' 変更前付加賞金
        Next i
        rs("HassoTime") = .HassoTime                                                            '' 発走時刻
        rs("HassoTimeBefore") = .HassoTimeBefore                                                '' 変更前発走時刻
        rs("TorokuTosu") = .TorokuTosu                                                          '' 登録頭数
        rs("SyussoTosu") = .SyussoTosu                                                          '' 出走頭数
        rs("NyusenTosu") = .NyusenTosu                                                          '' 入線頭数
        With .TenkoBaba
            rs("TenkoCD") = .TenkoCD                                                            '' 天候コード
            rs("SibaBabaCD") = .SibaBabaCD                                                      '' 芝馬場状態コード
            rs("DirtBabaCD") = .DirtBabaCD                                                      '' ダート馬場状態コード
        End With ' TenkoBaba
        For i = 0 To 24
            rs("LapTime" & i + 1) = .LapTime(i)                                                 '' ラップタイム
        Next i
        rs("SyogaiMileTime") = .SyogaiMileTime                                                  '' 障害マイルタイム
        rs("HaronTimeS3") = .HaronTimeS3                                                        '' 前３ハロンタイム
        rs("HaronTimeS4") = .HaronTimeS4                                                        '' 前４ハロンタイム
        rs("HaronTimeL3") = .HaronTimeL3                                                        '' 後３ハロンタイム
        rs("HaronTimeL4") = .HaronTimeL4                                                        '' 後４ハロンタイム
        For i = 0 To 3
            With .CornerInfo(i)
                rs("Corner" & i + 1) = .Corner                                                  '' コーナー
                rs("Syukaisu" & i + 1) = .Syukaisu                                              '' 周回数
                rs("Jyuni" & i + 1) = .Jyuni                                                    '' 各通過順位
            End With ' CornerInfo
        Next i
        rs("RecordUpKubun") = .RecordUpKubun                                                    '' レコード更新区分
    End With

    rs.Update
    
    mCon_RACE.CommitTrans

    Set rs = Nothing
    
    UpdateDB = True
    Exit Function
    
Errorhandler:
    gApp.ErrLog
    
    mCon_RACE.RollbackTrans

    UpdateDB = False
End Function

