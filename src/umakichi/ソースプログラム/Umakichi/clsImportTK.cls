VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsImportTK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Import JV_TK_TOKUUMA "
'
'   JVData "TK" データベース登録クラス
'

Option Explicit
Option Compare Binary
Implements clsIImport

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部変数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private mCon_TOKU As ADODB.Connection
Private mCon_TOKU_RACE As ADODB.Connection

Private mRS_TOKU As ADODB.Recordset
Private mRS_TOKU_RACE As ADODB.Recordset

Private mBuf As JV_TK_TOKUUMA


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部関数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: クラス初期化イベント
'
'   備考: なし
'
Private Sub Class_Initialize()
On Error GoTo ErrorHandler
    Dim strCon As String
    Dim rs As ADODB.Recordset
    
    ' コネクションのインスタンス生成
    Set mCon_TOKU = New ADODB.Connection
    Set mCon_TOKU_RACE = New ADODB.Connection

    ' レコードセットのインスタンス生成
    Set mRS_TOKU = New ADODB.Recordset
    Set mRS_TOKU_RACE = New ADODB.Recordset
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: コネクション、レコードセットをオープンする
'
'   備考: なし
'
Private Sub clsIImport_OpenDB()
On Error GoTo ErrorHandler        ' コネクションオープン
    Dim strCon As String
    

        ' コネクションオープン
    strCon = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & gApp.R_DBPath & "\"
    mCon_TOKU.Open strCon & "subTOKU.mdb"
    mCon_TOKU_RACE.Open strCon & "subTOKU_RACE.mdb"

    ' レコードセットオープン
    With mRS_TOKU
        .CursorLocation = adUseServer
        .Index = "PrimaryKey"
        .Open "TOKU", mCon_TOKU, adOpenKeyset, adLockOptimistic
        If Not (.EOF Or .BOF) Then
            .MoveFirst
        End If
    End With
    
    With mRS_TOKU_RACE
        .CursorLocation = adUseServer
        .Index = "PrimaryKey"
        .Open "TOKU_RACE", mCon_TOKU_RACE, adOpenKeyset, adLockOptimistic, adCmdTableDirect
        If Not (.EOF Or .BOF) Then
            .MoveFirst
        End If
    End With

    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: コネクション、レコードセットをクローズする
'
'   備考: なし
'
Private Sub clsIImport_CloseDB()
On Error GoTo ErrorHandler

    mRS_TOKU.Close
    mRS_TOKU_RACE.Close

    mCon_TOKU.Close
    mCon_TOKU_RACE.Close

    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: JVReadの返す１行をデータベースに登録する
'
'   備考: DBに追加を試み、失敗したら更新を試みる
'
Private Function clsIImport_Add(lBuf() As Byte) As Boolean
On Error GoTo ErrorHandler

    If lBuf(2) = ASCII_ZERO Then
        Call DeleteRecord(StrConv(lBuf, vbUnicode))
        clsIImport_Add = True
    Else
        Call SetDataFromByte_TK(lBuf, mBuf)     '' 構造体に代入する

        ' データ登録
        If Not InsertDB() Then
            ' 新規挿入に失敗したら更新を試みる
            clsIImport_Add = UpdateDB()
        Else
            clsIImport_Add = True
        End If

    End If

    Exit Function

ErrorHandler:
    gApp.ErrLog
    clsIImport_Add = False
End Function


'
'   機能: レコードを削除する
'
'   備考: なし
'
Private Sub DeleteRecord(lBuf As String)
On Error GoTo ErrorHandler
    Dim strSQL      As String
    Dim Year        As String
    Dim MonthDay    As String
    Dim JyoCD       As String
    Dim Kaiji       As String
    Dim Nichiji     As String
    Dim RaceNum     As String
    
    mCon_TOKU.BeginTrans
    mCon_TOKU_RACE.BeginTrans
    
    Year = Mid$(lBuf, 12, 4)                    '' 開催年
    MonthDay = Mid$(lBuf, 16, 4)                '' 開催月日
    JyoCD = Mid$(lBuf, 20, 2)                   '' 競馬場コード
    Kaiji = Mid$(lBuf, 22, 2)                   '' 開催回[第N回]
    Nichiji = Mid$(lBuf, 24, 2)                 '' 開催日目[N日目]
    RaceNum = Mid$(lBuf, 26, 2)                 '' レース番号

    strSQL = "DELETE * FROM TOKU"
    strSQL = strSQL & " WHERE [Year]   ='" & Year & "'"
    strSQL = strSQL & " AND [MonthDay] ='" & MonthDay & "'"
    strSQL = strSQL & " AND [JyoCD]    ='" & JyoCD & "'"
    strSQL = strSQL & " AND [Kaiji]    ='" & Kaiji & "'"
    strSQL = strSQL & " AND [Nichiji]  ='" & Nichiji & "'"
    strSQL = strSQL & " AND [RaceNum]  ='" & RaceNum & "'"
    mCon_TOKU.Execute strSQL, , adExecuteNoRecords
    
    strSQL = "DELETE * FROM TOKU_RACE"
    strSQL = strSQL & " WHERE [Year]   ='" & Year & "'"
    strSQL = strSQL & " AND [MonthDay] ='" & MonthDay & "'"
    strSQL = strSQL & " AND [JyoCD]    ='" & JyoCD & "'"
    strSQL = strSQL & " AND [Kaiji]    ='" & Kaiji & "'"
    strSQL = strSQL & " AND [Nichiji]  ='" & Nichiji & "'"
    strSQL = strSQL & " AND [RaceNum]  ='" & RaceNum & "'"
    mCon_TOKU_RACE.Execute strSQL, , adExecuteNoRecords
    
    mCon_TOKU.CommitTrans
    mCon_TOKU_RACE.CommitTrans
        
    Exit Sub
ErrorHandler:
    mCon_TOKU.RollbackTrans
    mCon_TOKU_RACE.RollbackTrans
End Sub


'
'   機能: データベースに挿入する
'
'   備考: なし
'
Private Function InsertDB() As Boolean
On Error GoTo ErrorHandler
    Dim i As Integer                                                                  '' ループカウンタ
    Dim j As Integer                                                                  '' ループカウンタ
    Dim k As Integer                                                                  '' ループカウンタ
    Dim rs As ADODB.Recordset                                                         '' SQL文
    
    mCon_TOKU.BeginTrans
    mCon_TOKU_RACE.BeginTrans

    
    
    Set rs = mRS_TOKU_RACE
    ' ヘッダ部分
    With mBuf.id
        rs.AddNew Array("Year", "MonthDay", "JyoCD", "Kaiji", "Nichiji", "RaceNum"), Array(.Year, .MonthDay, .JyoCD, .Kaiji, .Nichiji, .RaceNum)
    End With
    
    With mBuf
        With .head
            rs("RecordSpec") = .RecordSpec                                            '' レコード種別
            rs("DataKubun") = .DataKubun                                              '' データ区分
            With .MakeDate
                rs("MakeDate") = .Year & .Month & .Day                                '' 年
            End With ' MakeDate
        End With ' head
        With .RaceInfo
            rs("YoubiCD") = .YoubiCD                                                  '' 曜日コード
            rs("TokuNum") = .TokuNum                                                  '' 特別競走番号
            rs("Hondai") = .Hondai                                                    '' 競走名本題
            rs("Fukudai") = .Fukudai                                                  '' 競走名副題
            rs("Kakko") = .Kakko                                                      '' 競走名カッコ内
            rs("HondaiEng") = .HondaiEng                                              '' 競走名本題欧字
            rs("FukudaiEng") = .FukudaiEng                                            '' 競走名副題欧字
            rs("KakkoEng") = .KakkoEng                                                '' 競走名カッコ内欧字
            rs("Ryakusyo10") = .Ryakusyo10                                            '' 競走名略称１０字
            rs("Ryakusyo6") = .Ryakusyo6                                              '' 競走名略称６字
            rs("Ryakusyo3") = .Ryakusyo3                                              '' 競走名略称３字
            rs("Kubun") = .Kubun                                                      '' 競走名区分
            rs("Nkai") = .Nkai                                                        '' 重賞回次[第N回]
        End With ' RaceInfo
        rs("GradeCD") = .GradeCD                                                      '' グレードコード
        With .JyokenInfo
            rs("SyubetuCD") = .SyubetuCD                                              '' 競走種別コード
            rs("KigoCD") = .KigoCD                                                    '' 競走記号コード
            rs("JyuryoCD") = .JyuryoCD                                                '' 重量種別コード
            For j = 0 To 4
                rs("JyokenCD" & j + 1) = .JyokenCD(j)                                 '' 競走条件コード
            Next j
        End With ' JyokenInfo
        rs("Kyori") = .KYORI                                                          '' 距離
        rs("TrackCD") = .TrackCD                                                      '' トラックコード
        rs("CourseKubunCD") = .CourseKubunCD                                          '' コース区分
        With .HandiDate
            rs("HandiDate") = .Year & .Month & .Day                                   '' HandiDate
        End With ' HandiDate
        rs("TorokuTosu") = .TorokuTosu                                                '' 登録頭数
        
    End With
    
    rs.Update ' TOKU_RACE


    Set rs = mRS_TOKU
    For i = 0 To 299
        If mBuf.TokuUmaInfo(i).num <> "   " Then
            With mBuf.id
                rs.AddNew Array("Year", "MonthDay", "JyoCD", "Kaiji", "Nichiji", "RaceNum", "Num"), Array(.Year, .MonthDay, .JyoCD, .Kaiji, .Nichiji, .RaceNum, mBuf.TokuUmaInfo(i).num)
            End With
            With mBuf
                With .TokuUmaInfo(i)
                    rs("Num") = .num                                                '' 連番
                    rs("KettoNum") = .KettoNum                                      '' 血統登録番号
                    rs("Bamei") = .BAMEI                                            '' 馬名
                    rs("UmaKigoCD") = .UmaKigoCD                                    '' 馬記号コード
                    rs("SexCD") = .SexCD                                            '' 性別コード
                    rs("TozaiCD") = .TozaiCD                                        '' 調教師東西所属コード
                    rs("ChokyosiCode") = .ChokyosiCode                              '' 調教師コード
                    rs("ChokyosiRyakusyo") = .ChokyosiRyakusyo                      '' 調教師名略称
                    rs("Futan") = .Futan                                            '' 負担重量
                    rs("Koryu") = .Koryu                                            '' 交流区分
                End With ' TokuUmaInfo
            End With
    
            rs.Update
        End If
    Next i

    
    mCon_TOKU.CommitTrans
    mCon_TOKU_RACE.CommitTrans

    Set rs = Nothing
    
    InsertDB = True
    Exit Function
    
ErrorHandler:
    If Err.Number <> -2147217887 Then
        gApp.ErrLog
    End If
    rs.CancelUpdate
    
    mCon_TOKU.RollbackTrans
    mCon_TOKU_RACE.RollbackTrans

    InsertDB = False
End Function


'
'   機能: データベースを更新する
'
'   備考: なし
'
Private Function UpdateDB() As Boolean
On Error GoTo ErrorHandler
    Dim rs As ADODB.Recordset
    Dim i As Integer                                                                  '' ループカウンタ
    Dim j As Integer                                                                  '' ループカウンタ
    Dim k As Integer                                                                  '' ループカウンタ
    Dim strSQL As String '' SQL文

    Set rs = mRS_TOKU_RACE

    With mBuf.id
        rs.Seek Array(.Year, .MonthDay, .JyoCD, .Kaiji, .Nichiji, .RaceNum)
    End With

    With mBuf.head.MakeDate
        If rs("Makedate") > .Year & .Month & .Day Then
            UpdateDB = True
            Exit Function
        End If
    End With

    
    mCon_TOKU.BeginTrans
    mCon_TOKU_RACE.BeginTrans


    ' ヘッダ部分
    With mBuf
        With .head
            rs("RecordSpec") = .RecordSpec                                            '' レコード種別
            rs("DataKubun") = .DataKubun                                              '' データ区分
            With .MakeDate
                rs("MakeDate") = .Year & .Month & .Day                                '' 年
            End With ' MakeDate
        End With ' head
        With .RaceInfo
            rs("YoubiCD") = .YoubiCD                                                  '' 曜日コード
            rs("TokuNum") = .TokuNum                                                  '' 特別競走番号
            rs("Hondai") = .Hondai                                                    '' 競走名本題
            rs("Fukudai") = .Fukudai                                                  '' 競走名副題
            rs("Kakko") = .Kakko                                                      '' 競走名カッコ内
            rs("HondaiEng") = .HondaiEng                                              '' 競走名本題欧字
            rs("FukudaiEng") = .FukudaiEng                                            '' 競走名副題欧字
            rs("KakkoEng") = .KakkoEng                                                '' 競走名カッコ内欧字
            rs("Ryakusyo10") = .Ryakusyo10                                            '' 競走名略称１０字
            rs("Ryakusyo6") = .Ryakusyo6                                              '' 競走名略称６字
            rs("Ryakusyo3") = .Ryakusyo3                                              '' 競走名略称３字
            rs("Kubun") = .Kubun                                                      '' 競走名区分
            rs("Nkai") = .Nkai                                                        '' 重賞回次[第N回]
        End With ' RaceInfo
        rs("GradeCD") = .GradeCD                                                      '' グレードコード
        With .JyokenInfo
            rs("SyubetuCD") = .SyubetuCD                                              '' 競走種別コード
            rs("KigoCD") = .KigoCD                                                    '' 競走記号コード
            rs("JyuryoCD") = .JyuryoCD                                                '' 重量種別コード
            For j = 0 To 4
                rs("JyokenCD" & j + 1) = .JyokenCD(j)                                 '' 競走条件コード
            Next j
        End With ' JyokenInfo
        rs("Kyori") = .KYORI                                                          '' 距離
        rs("TrackCD") = .TrackCD                                                      '' トラックコード
        rs("CourseKubunCD") = .CourseKubunCD                                          '' コース区分
        With .HandiDate
            rs("HandiDate") = .Year & .Month & .Day                                   '' HandiDate
        End With ' HandiDate
        rs("TorokuTosu") = .TorokuTosu                                                '' 登録頭数

    End With

    rs.Update



    ' 馬毎部分
    ' 馬毎情報はレコード数が可変するため、既存のデータを全て削除してから登録しなおします。

    With mBuf.id
        strSQL = "DELETE FROM TOKU"
        strSQL = strSQL & " WHERE [Year]='" & .Year & "'"
        strSQL = strSQL & " AND MonthDay='" & .MonthDay & "'"
        strSQL = strSQL & " AND JyoCD='" & .JyoCD & "'"
        strSQL = strSQL & " AND Kaiji='" & .Kaiji & "'"
        strSQL = strSQL & " AND Nichiji='" & .Nichiji & "'"
        strSQL = strSQL & " AND RaceNum='" & .RaceNum & "'"
        mCon_TOKU.Execute strSQL, , adExecuteNoRecords
    End With

    
    Set rs = mRS_TOKU

    ' 全部再登録する
    For i = 0 To 299
        If mBuf.TokuUmaInfo(i).num <> "   " Then
            With mBuf.id
                rs.AddNew Array("Year", "MonthDay", "JyoCD", "Kaiji", "Nichiji", "RaceNum", "Num"), Array(.Year, .MonthDay, .JyoCD, .Kaiji, .Nichiji, .RaceNum, mBuf.TokuUmaInfo(i).num)
            End With
            With mBuf
                With .TokuUmaInfo(i)
                    rs("Num") = .num                                                '' 連番
                    rs("KettoNum") = .KettoNum                                      '' 血統登録番号
                    rs("Bamei") = .BAMEI                                            '' 馬名
                    rs("UmaKigoCD") = .UmaKigoCD                                    '' 馬記号コード
                    rs("SexCD") = .SexCD                                            '' 性別コード
                    rs("TozaiCD") = .TozaiCD                                        '' 調教師東西所属コード
                    rs("ChokyosiCode") = .ChokyosiCode                              '' 調教師コード
                    rs("ChokyosiRyakusyo") = .ChokyosiRyakusyo                      '' 調教師名略称
                    rs("Futan") = .Futan                                            '' 負担重量
                    rs("Koryu") = .Koryu                                            '' 交流区分
                End With ' TokuUmaInfo
            End With
    
            rs.Update
        End If
    Next i

    
    mCon_TOKU.CommitTrans
    mCon_TOKU_RACE.CommitTrans


    Set rs = Nothing

    UpdateDB = True
    Exit Function

ErrorHandler:
    rs.CancelUpdate
    
    mCon_TOKU.RollbackTrans
    mCon_TOKU_RACE.RollbackTrans

    UpdateDB = False
End Function

