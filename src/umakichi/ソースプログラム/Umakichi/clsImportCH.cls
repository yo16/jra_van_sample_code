VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsImportCH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Import JV_CH_CHOKYOSI "
'
'   JVData "CH" データベース登録クラス
'

Option Explicit
Option Compare Binary
Implements clsIImport

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部変数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private mBuf As JV_CH_CHOKYOSI

Private mCon_CHOKYO As ADODB.Connection
Private mCon_CHOKYO_SEISEKI As ADODB.Connection

Private mRS_CHOKYO As ADODB.Recordset
Private mRS_CHOKYO_SEISEKI As ADODB.Recordset


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
    Set mCon_CHOKYO = New ADODB.Connection
    Set mCon_CHOKYO_SEISEKI = New ADODB.Connection

    ' レコードセットのインスタンス生成
    Set mRS_CHOKYO = New ADODB.Recordset
    Set mRS_CHOKYO_SEISEKI = New ADODB.Recordset
    
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
    mCon_CHOKYO.Open strCon & "subCHOKYO.mdb"
    mCon_CHOKYO_SEISEKI.Open strCon & "subCHOKYO_SEISEKI.mdb"

    ' レコードセットオープン
    With mRS_CHOKYO
        .CursorLocation = adUseServer
        .Index = "PrimaryKey"
        .Open "CHOKYO", mCon_CHOKYO, adOpenKeyset, adLockOptimistic, adCmdTableDirect
        If Not (.EOF Or .BOF) Then
            .MoveFirst
        End If
    End With
    With mRS_CHOKYO_SEISEKI
        .CursorLocation = adUseServer
        .Index = "PrimaryKey"
        .Open "CHOKYO_SEISEKI", mCon_CHOKYO_SEISEKI, adOpenKeyset, adLockOptimistic, adCmdTableDirect
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
    mRS_CHOKYO.Close
    mRS_CHOKYO_SEISEKI.Close

    mCon_CHOKYO.Close
    mCon_CHOKYO_SEISEKI.Close

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
        Call SetDataFromByte_CH(lBuf, mBuf)     '' 構造体に代入する

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
    Dim ChokyosiCode As String

    mCon_CHOKYO.BeginTrans
    mCon_CHOKYO_SEISEKI.BeginTrans
    
    ChokyosiCode = Mid$(lBuf, 12, 5)                  ''

    strSQL = "DELETE * FROM CHOKYO"
    strSQL = strSQL & " WHERE [ChokyosiCode]   ='" & ChokyosiCode & "'"
    mCon_CHOKYO.Execute strSQL, , adExecuteNoRecords
    
    strSQL = "DELETE * FROM CHOKYO_SEISEKI"
    strSQL = strSQL & " WHERE [ChokyosiCode]   ='" & ChokyosiCode & "'"
    mCon_CHOKYO_SEISEKI.Execute strSQL, , adExecuteNoRecords
    
    mCon_CHOKYO.CommitTrans
    mCon_CHOKYO_SEISEKI.CommitTrans
        
    Exit Sub
ErrorHandler:
    mCon_CHOKYO.RollbackTrans
    mCon_CHOKYO_SEISEKI.RollbackTrans
End Sub


'
'   機能: データベースに挿入する
'
'   備考: なし
'
Private Function InsertDB() As Boolean
On Error GoTo ErrorHandler
    Dim i As Integer                                                                                 '' ループカウンタ
    Dim j As Integer                                                                                 '' ループカウンタ
    Dim k As Integer                                                                                 '' ループカウンタ
    Dim rs As ADODB.Recordset                                                                        '' SQL文
    
    mCon_CHOKYO.BeginTrans
    mCon_CHOKYO_SEISEKI.BeginTrans

    
    ' ヘッダ部分
    Set rs = mRS_CHOKYO
    rs.AddNew Array("ChokyosiCode"), Array(mBuf.ChokyosiCode)
    With mBuf
        With .head
            rs("RecordSpec") = .RecordSpec                                                           '' レコード種別
            rs("DataKubun") = .DataKubun                                                             '' データ区分
            With .MakeDate
                rs("MakeDate") = .Year & .Month & .Day                                               '' 年月日
            End With ' MakeDate
        End With ' head
        rs("DelKubun") = .DelKubun                                                                   '' 調教師抹消区分
        With .IssueDate
                rs("IssueDate") = .Year & .Month & .Day                                              '' 年月日
        End With ' IssueDate
        With .DelDate
                rs("DelDate") = .Year & .Month & .Day                                                '' 年月日
        End With ' DelDate
        With .BirthDate
                rs("BirthDate") = .Year & .Month & .Day                                              '' 年月日
        End With ' BirthDate
        rs("ChokyosiName") = .ChokyosiName                                                           '' 調教師名漢字
        rs("ChokyosiNameKana") = .ChokyosiNameKana                                                   '' 調教師名半角カナ
        rs("ChokyosiRyakusyo") = .ChokyosiRyakusyo                                                   '' 調教師名略称
        rs("ChokyosiNameEng") = .ChokyosiNameEng                                                     '' 調教師名欧字
        rs("SexCD") = .SexCD                                                                         '' 性別区分
        rs("TozaiCD") = .TozaiCD                                                                     '' 調教師東西所属コード
        rs("Syotai") = .Syotai                                                                       '' 招待地域名
        For i = 0 To 2
            With .SaikinJyusyo(i)
                With .SaikinJyusyoid
                    rs("SaikinJyusyo" & i + 1 & "SaikinJyusyoid") = .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum
                                                                                                     '' 開催年 開催月日 競馬場コード 開催回[第N回] 開催日目[N日目] レース番号
                End With ' SaikinJyusyoid
                rs("SaikinJyusyo" & i + 1 & "Hondai") = .Hondai                                      '' 競走名本題
                rs("SaikinJyusyo" & i + 1 & "Ryakusyo10") = .Ryakusyo10                              '' 競走名略称10字
                rs("SaikinJyusyo" & i + 1 & "Ryakusyo6") = .Ryakusyo6                                '' 競走名略称6字
                rs("SaikinJyusyo" & i + 1 & "Ryakusyo3") = .Ryakusyo3                                '' 競走名略称3字
                rs("SaikinJyusyo" & i + 1 & "GradeCD") = .GradeCD                                    '' グレードコード
                rs("SaikinJyusyo" & i + 1 & "SyussoTosu") = .SyussoTosu                              '' 出走頭数
                rs("SaikinJyusyo" & i + 1 & "KettoNum") = .KettoNum                                  '' 血統登録番号
                rs("SaikinJyusyo" & i + 1 & "Bamei") = .BAMEI                                        '' 馬名
            End With ' SaikinJyusyo
        Next i
        
        
    End With ' mBuf
        
    rs.Update ' CHOKYO
        
        
    ' 成績部分
    For i = 0 To 2
        With mBuf
            Set rs = mRS_CHOKYO_SEISEKI
            rs.AddNew Array("ChokyosiCode", "Num"), Array(mBuf.ChokyosiCode, i)
            With .HonZenRuikei(i)
                rs("SetYear") = .SetYear                                                             '' 設定年
                rs("HonSyokinHeichi") = .HonSyokinHeichi                                             '' 平地本賞金合計
                rs("HonSyokinSyogai") = .HonSyokinSyogai                                             '' 障害本賞金合計
                rs("FukaSyokinHeichi") = .FukaSyokinHeichi                                           '' 平地付加賞金合計
                rs("FukaSyokinSyogai") = .FukaSyokinSyogai                                           '' 障害付加賞金合計
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
            End With ' HonZenRuikei
            With .HonZenRuikei(i)
                For j = 0 To 19
                    With .ChakuKaisuJyo(j)
                        For k = 0 To 5
                            rs("Jyo" & j + 1 & "Chakukaisu" & k + 1) = .Chakukaisu(k)
                        Next k
                    End With ' ChakuKaisuJyo
                Next j
            End With ' HonZenRuikei
                
        End With ' mBuf
        
        rs.Update ' CHOKYO_SEISEKI
    Next i
    
    mCon_CHOKYO.CommitTrans
    mCon_CHOKYO_SEISEKI.CommitTrans

    Set rs = Nothing
    
    InsertDB = True
    Exit Function
    
ErrorHandler:
    If Err.Number <> -2147217887 Then
        gApp.ErrLog
    End If
    rs.CancelUpdate
    
    mCon_CHOKYO.RollbackTrans
    mCon_CHOKYO_SEISEKI.RollbackTrans

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
    Dim i As Integer                                                                                 '' ループカウンタ
    Dim j As Integer                                                                                 '' ループカウンタ
    Dim k As Integer                                                                                 '' ループカウンタ
    Dim strSQL As String                                                                             '' SQL文
    
    Set rs = mRS_CHOKYO

    rs.Seek Array(mBuf.ChokyosiCode)

    With mBuf.head.MakeDate
        If rs("Makedate") > .Year & .Month & .Day Then
            UpdateDB = True
            Exit Function
        End If
    End With

    
    mCon_CHOKYO.BeginTrans
    mCon_CHOKYO_SEISEKI.BeginTrans

    With mBuf
        With .head
            rs("RecordSpec") = .RecordSpec                                                           '' レコード種別
            rs("DataKubun") = .DataKubun                                                             '' データ区分
            With .MakeDate
                rs("MakeDate") = .Year & .Month & .Day                                               '' 年月日
            End With ' MakeDate
        End With ' head
        rs("DelKubun") = .DelKubun                                                                   '' 調教師抹消区分
        With .IssueDate
                rs("IssueDate") = .Year & .Month & .Day                                              '' 年月日
        End With ' IssueDate
        With .DelDate
                rs("DelDate") = .Year & .Month & .Day                                                '' 年月日
        End With ' DelDate
        With .BirthDate
                rs("BirthDate") = .Year & .Month & .Day                                              '' 年月日
        End With ' BirthDate
        rs("ChokyosiName") = .ChokyosiName                                                           '' 調教師名漢字
        rs("ChokyosiNameKana") = .ChokyosiNameKana                                                   '' 調教師名半角カナ
        rs("ChokyosiRyakusyo") = .ChokyosiRyakusyo                                                   '' 調教師名略称
        rs("ChokyosiNameEng") = .ChokyosiNameEng                                                     '' 調教師名欧字
        rs("SexCD") = .SexCD                                                                         '' 性別区分
        rs("TozaiCD") = .TozaiCD                                                                     '' 調教師東西所属コード
        rs("Syotai") = .Syotai                                                                       '' 招待地域名
        For i = 0 To 2
            With .SaikinJyusyo(i)
                With .SaikinJyusyoid
                    rs("SaikinJyusyo" & i + 1 & "SaikinJyusyoid") = .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum
                                                                                                     '' 開催年 開催月日 競馬場コード 開催回[第N回] 開催日目[N日目] レース番号
                End With ' SaikinJyusyoid
                rs("SaikinJyusyo" & i + 1 & "Hondai") = .Hondai                                      '' 競走名本題
                rs("SaikinJyusyo" & i + 1 & "Ryakusyo10") = .Ryakusyo10                              '' 競走名略称10字
                rs("SaikinJyusyo" & i + 1 & "Ryakusyo6") = .Ryakusyo6                                '' 競走名略称6字
                rs("SaikinJyusyo" & i + 1 & "Ryakusyo3") = .Ryakusyo3                                '' 競走名略称3字
                rs("SaikinJyusyo" & i + 1 & "GradeCD") = .GradeCD                                    '' グレードコード
                rs("SaikinJyusyo" & i + 1 & "SyussoTosu") = .SyussoTosu                              '' 出走頭数
                rs("SaikinJyusyo" & i + 1 & "KettoNum") = .KettoNum                                  '' 血統登録番号
                rs("SaikinJyusyo" & i + 1 & "Bamei") = .BAMEI                                        '' 馬名
            End With ' SaikinJyusyo
        Next i
        
        
    End With ' mBuf
        
    rs.Update ' CHOKYO
        
        
    ' 成績部分
    Set rs = mRS_CHOKYO_SEISEKI
    For i = 0 To 2
        With mBuf
            ' Call SafeSeek(rs, Array("ChokyosiCode", "Num"), Array(mBuf.ChokyosiCode, CStr(i))
            rs.Seek Array(mBuf.ChokyosiCode, CStr(i))
            With .HonZenRuikei(i)
                rs("SetYear") = .SetYear                                                             '' 設定年
                rs("HonSyokinHeichi") = .HonSyokinHeichi                                             '' 平地本賞金合計
                rs("HonSyokinSyogai") = .HonSyokinSyogai                                             '' 障害本賞金合計
                rs("FukaSyokinHeichi") = .FukaSyokinHeichi                                           '' 平地付加賞金合計
                rs("FukaSyokinSyogai") = .FukaSyokinSyogai                                           '' 障害付加賞金合計
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
            End With ' HonZenRuikei
            With .HonZenRuikei(i)
                For j = 0 To 19
                    With .ChakuKaisuJyo(j)
                        For k = 0 To 5
                            rs("Jyo" & j + 1 & "Chakukaisu" & k + 1) = .Chakukaisu(k)
                        Next k
                    End With ' ChakuKaisuJyo
                Next j
            End With ' HonZenRuikei
                
        End With ' mBuf
        
        rs.Update ' CHOKYO_SEISEKI
    Next i
    
    mCon_CHOKYO.CommitTrans
    mCon_CHOKYO_SEISEKI.CommitTrans

    Set rs = Nothing
    
    UpdateDB = True
    Exit Function
    
ErrorHandler:
    gApp.ErrLog
    
    mCon_CHOKYO.RollbackTrans
    mCon_CHOKYO_SEISEKI.RollbackTrans

    UpdateDB = False
End Function
