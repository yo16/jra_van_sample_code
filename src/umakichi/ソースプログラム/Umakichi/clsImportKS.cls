VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsImportKS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Import JV_KS_KISYU "
'
'   JVData "KS" データベース登録クラス
'

Option Explicit
Option Compare Binary
Implements clsIImport

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部変数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private mBuf As JV_KS_KISYU

Private mCon_KISHU As ADODB.Connection
Private mCon_KISHU_SEISEKI As ADODB.Connection

Private mRS_KISHU As ADODB.Recordset
Private mRS_KISHU_SEISEKI As ADODB.Recordset


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
    ' コネクションのインスタンス生成
    Set mCon_KISHU = New ADODB.Connection
    Set mCon_KISHU_SEISEKI = New ADODB.Connection

    ' レコードセットのインスタンス生成
    Set mRS_KISHU = New ADODB.Recordset
    Set mRS_KISHU_SEISEKI = New ADODB.Recordset
    
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
    mCon_KISHU.Open strCon & "subKISHU.mdb"
    mCon_KISHU_SEISEKI.Open strCon & "subKISHU_SEISEKI.mdb"

    ' レコードセットオープン
    With mRS_KISHU
        .CursorLocation = adUseServer
        .Index = "PrimaryKey"
        .Open "KISHU", mCon_KISHU, adOpenKeyset, adLockOptimistic, adCmdTableDirect
        If Not (.EOF Or .BOF) Then
            .MoveFirst
        End If
    End With
    With mRS_KISHU_SEISEKI
        .CursorLocation = adUseServer
        .Index = "PrimaryKey"
        .Open "KISHU_SEISEKI", mCon_KISHU_SEISEKI, adOpenKeyset, adLockOptimistic, adCmdTableDirect
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
    mRS_KISHU.Close
    mRS_KISHU_SEISEKI.Close

    mCon_KISHU.Close
    mCon_KISHU_SEISEKI.Close

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
        Call SetDataFromByte_KS(lBuf, mBuf)     '' 構造体に代入する

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
    Dim KisyuCode   As String
    
    mCon_KISHU.BeginTrans
    mCon_KISHU_SEISEKI.BeginTrans
    
    KisyuCode = Mid$(lBuf, 12, 5)                  ''

    strSQL = "DELETE * FROM KISHU"
    strSQL = strSQL & " WHERE [KisyuCode]   ='" & KisyuCode & "'"
    mCon_KISHU.Execute strSQL, , adExecuteNoRecords
    
    strSQL = "DELETE * FROM KISHU_SEISEKI"
    strSQL = strSQL & " WHERE [KisyuCode]   ='" & KisyuCode & "'"
    mCon_KISHU_SEISEKI.Execute strSQL, , adExecuteNoRecords
    
    mCon_KISHU.CommitTrans
    mCon_KISHU_SEISEKI.CommitTrans
        
    Exit Sub
ErrorHandler:
    mCon_KISHU.RollbackTrans
    mCon_KISHU_SEISEKI.RollbackTrans
End Sub


'
'   機能: データベースに挿入する
'
'   備考: なし
'
Private Function InsertDB() As Boolean
On Error GoTo ErrorHandler
    Dim i As Integer                                                                                  '' ループカウンタ
    Dim j As Integer                                                                                  '' ループカウンタ
    Dim k As Integer                                                                                  '' ループカウンタ
    Dim rs As ADODB.Recordset                                                                         '' SQL文

    mCon_KISHU.BeginTrans
    mCon_KISHU_SEISEKI.BeginTrans


                                                                                                      '' KISHU(騎手マスタ)
    Set rs = mRS_KISHU
    rs.AddNew Array("KisyuCode"), Array(mBuf.KisyuCode)
    With mBuf
        With .head
            rs("RecordSpec") = .RecordSpec                                                            '' レコード種別
            rs("DataKubun") = .DataKubun                                                              '' データ区分
            With .MakeDate
                rs("MakeDate") = .Year & .Month & .Day                                                '' 年月日
            End With ' MakeDate
        End With ' head
        rs("DelKubun") = .DelKubun                                                                    '' 騎手抹消区分
        With .IssueDate
                rs("IssueDate") = .Year & .Month & .Day                                               '' 年月日
        End With ' IssueDate
        With .DelDate
                rs("DelDate") = .Year & .Month & .Day                                                 '' 年月日
        End With ' DelDate
        With .BirthDate
                rs("BirthDate") = .Year & .Month & .Day                                               '' 年月日
        End With ' BirthDate
        rs("KisyuName") = .KisyuName                                                                  '' 騎手名漢字
        rs("reserved") = .Reserved                                                                    '' 予備
        rs("KisyuNameKana") = .KisyuNameKana                                                          '' 騎手名半角カナ
        rs("KisyuRyakusyo") = .KisyuRyakusyo                                                          '' 騎手名略称
        rs("KisyuNameEng") = .KisyuNameEng                                                            '' 騎手名欧字
        rs("SexCD") = .SexCD                                                                          '' 性別区分
        rs("SikakuCD") = .SikakuCD                                                                    '' 騎乗資格コード
        rs("MinaraiCD") = .MinaraiCD                                                                  '' 騎手見習コード
        rs("TozaiCD") = .TozaiCD                                                                      '' 騎手東西所属コード
        rs("Syotai") = .Syotai                                                                        '' 招待地域名
        rs("ChokyosiCode") = .ChokyosiCode                                                            '' 所属調教師コード
        rs("ChokyosiRyakusyo") = .ChokyosiRyakusyo                                                    '' 所属調教師名略称
        For i = 0 To 1
            With .HatuKiJyo(i)
                With .Hatukijyoid
                    rs("HatuKiJyo" & i + 1 & "Hatukijyoid") = .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum
                End With ' Hatukijyoid
                rs("HatuKiJyo" & i + 1 & "SyussoTosu") = .SyussoTosu                                  '' 出走頭数
                rs("HatuKiJyo" & i + 1 & "KettoNum") = .KettoNum                                      '' 血統登録番号
                rs("HatuKiJyo" & i + 1 & "Bamei") = .BAMEI                                            '' 馬名
                rs("HatuKiJyo" & i + 1 & "KakuteiJyuni") = .KakuteiJyuni                              '' 確定着順
                rs("HatuKiJyo" & i + 1 & "IJyoCD") = .IJyoCD                                          '' 異常区分コード
            End With ' HatuKiJyo
        Next i
        For i = 0 To 1
            With .HatuSyori(i)
                With .Hatusyoriid
                    rs("HatuSyori" & i + 1 & "Hatusyoriid") = .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum
                                                                                                      '' 開催年 開催月日 競馬場コード 開催回[第N回] 開催日目[N日目] レース番号
                End With ' Hatusyoriid
                rs("HatuSyori" & i + 1 & "SyussoTosu") = .SyussoTosu                                  '' 出走頭数
                rs("HatuSyori" & i + 1 & "KettoNum") = .KettoNum                                      '' 血統登録番号
                rs("HatuSyori" & i + 1 & "Bamei") = .BAMEI                                            '' 馬名
            End With ' HatuSyori
        Next i
        For i = 0 To 2
            With .SaikinJyusyo(i)
                With .SaikinJyusyoid
                    rs("SaikinJyusyo" & i + 1 & "SaikinJyusyoid") = .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum
                                                                                                      '' 開催年 開催月日 競馬場コード 開催回[第N回] 開催日目[N日目] レース番号
                End With ' SaikinJyusyoid
                rs("SaikinJyusyo" & i + 1 & "Hondai") = .Hondai                                       '' 競走名本題
                rs("SaikinJyusyo" & i + 1 & "Ryakusyo10") = .Ryakusyo10                               '' 競走名略称10字
                rs("SaikinJyusyo" & i + 1 & "Ryakusyo6") = .Ryakusyo6                                 '' 競走名略称6字
                rs("SaikinJyusyo" & i + 1 & "Ryakusyo3") = .Ryakusyo3                                 '' 競走名略称3字
                rs("SaikinJyusyo" & i + 1 & "GradeCD") = .GradeCD                                     '' グレードコード
                rs("SaikinJyusyo" & i + 1 & "SyussoTosu") = .SyussoTosu                               '' 出走頭数
                rs("SaikinJyusyo" & i + 1 & "KettoNum") = .KettoNum                                   '' 血統登録番号
                rs("SaikinJyusyo" & i + 1 & "Bamei") = .BAMEI                                         '' 馬名
            End With ' SaikinJyusyo
        Next i


    End With ' mBuf

    rs.Update ' KISHU


                                                                                                      '' KISHU_SEISEKI (騎手マスタ_成績)
    For i = 0 To 2
        With mBuf
            Set rs = mRS_KISHU_SEISEKI
            rs.AddNew Array("KisyuCode", "Num"), Array(.KisyuCode, i)
            With .HonZenRuikei(i)
                rs("SetYear") = .SetYear                                                              '' 設定年
                rs("HonSyokinHeichi") = .HonSyokinHeichi                                              '' 平地本賞金合計
                rs("HonSyokinSyogai") = .HonSyokinSyogai                                              '' 障害本賞金合計
                rs("FukaSyokinHeichi") = .FukaSyokinHeichi                                            '' 平地付加賞金合計
                rs("FukaSyokinSyogai") = .FukaSyokinSyogai                                            '' 障害付加賞金合計
                With .ChakuKaisuHeichi
                    For k = 0 To 5
                        rs("HeichiChakukaisu" & k + 1) = .Chakukaisu(k)
                    Next k
                End With ' ChakuKaisuHeichi
                With .ChakuKaisuSyogai
                    For k = 0 To 5
                        rs("SyogaiChakukaisu" & k + 1) = .Chakukaisu(k)
                    Next k
                End With ' ChakuKaisuSyogai
                For j = 0 To 5
                    With .ChakuKaisuKyori(j)
                        For k = 0 To 5
                            rs("Kyori" & j + 1 & "Chakukaisu" & k + 1) = .Chakukaisu(k)
                        Next k
                    End With ' ChakuKaisuKyori
                Next j
                For j = 0 To 19
                    With .ChakuKaisuJyo(j)
                        For k = 0 To 5
                            rs("Jyo" & j + 1 & "Chakukaisu" & k + 1) = .Chakukaisu(k)
                        Next k
                    End With ' ChakuKaisuJyo
                Next j
            End With ' HonZenRuikei

        End With ' mBuf

        rs.Update ' KISHU_SEISEKI
    Next i

    mCon_KISHU.CommitTrans
    mCon_KISHU_SEISEKI.CommitTrans

    Set rs = Nothing

    InsertDB = True
    Exit Function

ErrorHandler:
    If Err.Number <> -2147217887 Then
        gApp.ErrLog
    End If
    rs.CancelUpdate
    
    mCon_KISHU.RollbackTrans
    mCon_KISHU_SEISEKI.RollbackTrans

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
    Dim i As Integer                                                                                  '' ループカウンタ
    Dim j As Integer                                                                                  '' ループカウンタ
    Dim k As Integer                                                                                  '' ループカウンタ
    Dim strSQL As String                                                                              '' SQL文

    Set rs = mRS_KISHU

    rs.Seek Array(mBuf.KisyuCode)

    With mBuf.head.MakeDate
        If rs("Makedate") > .Year & .Month & .Day Then
            UpdateDB = True
            Exit Function
        End If
    End With

    
    mCon_KISHU.BeginTrans
    mCon_KISHU_SEISEKI.BeginTrans

    With mBuf
        With .head
            rs("RecordSpec") = .RecordSpec                                                            '' レコード種別
            rs("DataKubun") = .DataKubun                                                              '' データ区分
            With .MakeDate
                rs("MakeDate") = .Year & .Month & .Day                                                '' 年月日
            End With ' MakeDate
        End With ' head
        rs("DelKubun") = .DelKubun                                                                    '' 騎手抹消区分
        With .IssueDate
                rs("IssueDate") = .Year & .Month & .Day                                               '' 年月日
        End With ' IssueDate
        With .DelDate
                rs("DelDate") = .Year & .Month & .Day                                                 '' 年月日
        End With ' DelDate
        With .BirthDate
                rs("BirthDate") = .Year & .Month & .Day                                               '' 年月日
        End With ' BirthDate
        rs("KisyuName") = .KisyuName                                                                  '' 騎手名漢字
        rs("reserved") = .Reserved                                                                    '' 予備
        rs("KisyuNameKana") = .KisyuNameKana                                                          '' 騎手名半角カナ
        rs("KisyuRyakusyo") = .KisyuRyakusyo                                                          '' 騎手名略称
        rs("KisyuNameEng") = .KisyuNameEng                                                            '' 騎手名欧字
        rs("SexCD") = .SexCD                                                                          '' 性別区分
        rs("SikakuCD") = .SikakuCD                                                                    '' 騎乗資格コード
        rs("MinaraiCD") = .MinaraiCD                                                                  '' 騎手見習コード
        rs("TozaiCD") = .TozaiCD                                                                      '' 騎手東西所属コード
        rs("Syotai") = .Syotai                                                                        '' 招待地域名
        rs("ChokyosiCode") = .ChokyosiCode                                                            '' 所属調教師コード
        rs("ChokyosiRyakusyo") = .ChokyosiRyakusyo                                                    '' 所属調教師名略称
        For i = 0 To 1
            With .HatuKiJyo(i)
                With .Hatukijyoid
                    rs("HatuKiJyo" & i + 1 & "Hatukijyoid") = .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum
                End With ' Hatukijyoid
                rs("HatuKiJyo" & i + 1 & "SyussoTosu") = .SyussoTosu                                  '' 出走頭数
                rs("HatuKiJyo" & i + 1 & "KettoNum") = .KettoNum                                      '' 血統登録番号
                rs("HatuKiJyo" & i + 1 & "Bamei") = .BAMEI                                            '' 馬名
                rs("HatuKiJyo" & i + 1 & "KakuteiJyuni") = .KakuteiJyuni                              '' 確定着順
                rs("HatuKiJyo" & i + 1 & "IJyoCD") = .IJyoCD                                          '' 異常区分コード
            End With ' HatuKiJyo
        Next i
        For i = 0 To 1
            With .HatuSyori(i)
                With .Hatusyoriid
                    rs("HatuSyori" & i + 1 & "Hatusyoriid") = .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum
                                                                                                      '' 開催年 開催月日 競馬場コード 開催回[第N回] 開催日目[N日目] レース番号
                End With ' Hatusyoriid
                rs("HatuSyori" & i + 1 & "SyussoTosu") = .SyussoTosu                                  '' 出走頭数
                rs("HatuSyori" & i + 1 & "KettoNum") = .KettoNum                                      '' 血統登録番号
                rs("HatuSyori" & i + 1 & "Bamei") = .BAMEI                                            '' 馬名
            End With ' HatuSyori
        Next i
        For i = 0 To 2
            With .SaikinJyusyo(i)
                With .SaikinJyusyoid
                    rs("SaikinJyusyo" & i + 1 & "SaikinJyusyoid") = .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum
                                                                                                      '' 開催年 開催月日 競馬場コード 開催回[第N回] 開催日目[N日目] レース番号
                End With ' SaikinJyusyoid
                rs("SaikinJyusyo" & i + 1 & "Hondai") = .Hondai                                       '' 競走名本題
                rs("SaikinJyusyo" & i + 1 & "Ryakusyo10") = .Ryakusyo10                               '' 競走名略称10字
                rs("SaikinJyusyo" & i + 1 & "Ryakusyo6") = .Ryakusyo6                                 '' 競走名略称6字
                rs("SaikinJyusyo" & i + 1 & "Ryakusyo3") = .Ryakusyo3                                 '' 競走名略称3字
                rs("SaikinJyusyo" & i + 1 & "GradeCD") = .GradeCD                                     '' グレードコード
                rs("SaikinJyusyo" & i + 1 & "SyussoTosu") = .SyussoTosu                               '' 出走頭数
                rs("SaikinJyusyo" & i + 1 & "KettoNum") = .KettoNum                                   '' 血統登録番号
                rs("SaikinJyusyo" & i + 1 & "Bamei") = .BAMEI                                         '' 馬名
            End With ' SaikinJyusyo
        Next i


    End With ' mBuf

    rs.Update ' KISHU



    ' KISHU_SEISEKI (騎手マスタ_成績)
    Set rs = mRS_KISHU_SEISEKI
    For i = 0 To 2
        With mBuf
            ' Call SafeSeek(rs, Array("KisyuCode", "Num"), Array(mBuf.KisyuCode, CStr(i))
            rs.Seek Array(mBuf.KisyuCode, CStr(i))
            
            With .HonZenRuikei(i)
                rs("SetYear") = .SetYear                                                              '' 設定年
                rs("HonSyokinHeichi") = .HonSyokinHeichi                                              '' 平地本賞金合計
                rs("HonSyokinSyogai") = .HonSyokinSyogai                                              '' 障害本賞金合計
                rs("FukaSyokinHeichi") = .FukaSyokinHeichi                                            '' 平地付加賞金合計
                rs("FukaSyokinSyogai") = .FukaSyokinSyogai                                            '' 障害付加賞金合計
                With .ChakuKaisuHeichi
                    For k = 0 To 5
                        rs("HeichiChakukaisu" & k + 1) = .Chakukaisu(k)
                    Next k
                End With ' ChakuKaisuHeichi
                With .ChakuKaisuSyogai
                    For k = 0 To 5
                        rs("SyogaiChakukaisu" & k + 1) = .Chakukaisu(k)
                    Next k
                End With ' ChakuKaisuSyogai
                For j = 0 To 5
                    With .ChakuKaisuKyori(j)
                        For k = 0 To 5
                            rs("Kyori" & j + 1 & "Chakukaisu" & k + 1) = .Chakukaisu(k)
                        Next k
                    End With ' ChakuKaisuKyori
                Next j
                For j = 0 To 19
                    With .ChakuKaisuJyo(j)
                        For k = 0 To 5
                            rs("Jyo" & j + 1 & "Chakukaisu" & k + 1) = .Chakukaisu(k)
                        Next k
                    End With ' ChakuKaisuJyo
                Next j
            End With ' HonZenRuikei

        End With ' mBuf

        rs.Update ' KISHU_SEISEKI
    Next i

    mCon_KISHU.CommitTrans
    mCon_KISHU_SEISEKI.CommitTrans

    Set rs = Nothing

    UpdateDB = True
    Exit Function

ErrorHandler:
    gApp.ErrLog
    
    mCon_KISHU.RollbackTrans
    mCon_KISHU_SEISEKI.RollbackTrans

    UpdateDB = False
End Function
