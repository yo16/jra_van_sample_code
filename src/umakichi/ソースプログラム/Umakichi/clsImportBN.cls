VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsImportBN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Import JV_BN_BANUSI "
'
'   JVData "BN" データベース登録クラス
'

Option Explicit
Option Compare Binary
Implements clsIImport

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部変数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private mBuf As JV_BN_BANUSI

Private mCon_BANUSI As ADODB.Connection

Private mRS_BANUSI As ADODB.Recordset


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
    Set mCon_BANUSI = New ADODB.Connection

    ' レコードセットのインスタンス生成
    Set mRS_BANUSI = New ADODB.Recordset
    
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
    mCon_BANUSI.Open strCon & "subBANUSI.mdb"

    ' レコードセットオープン
    With mRS_BANUSI
        .CursorLocation = adUseServer
        .Index = "PrimaryKey"
        .Open "BANUSI", mCon_BANUSI, adOpenKeyset, adLockOptimistic, adCmdTableDirect
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
    mRS_BANUSI.Close
    
    mCon_BANUSI.Close

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

        Call SetDataFromByte_BN(lBuf, mBuf)     '' 構造体に代入する

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
    Dim BanusiCode As String

    mCon_BANUSI.BeginTrans
    
    BanusiCode = Mid$(lBuf, 12, 6)                   ''

    strSQL = "DELETE * FROM BANUSI"
    strSQL = strSQL & " WHERE [BanusiCode]   ='" & BanusiCode & "'"
    
    mCon_BANUSI.Execute strSQL, , adExecuteNoRecords
    
    mCon_BANUSI.CommitTrans
        
    Exit Sub
ErrorHandler:
    mCon_BANUSI.RollbackTrans
End Sub


'
'   機能: データベースに挿入する
'
'   備考: なし
'
Private Function InsertDB() As Boolean
On Error GoTo ErrorHandler
    Dim i As Integer                                                                        '' ループカウンタ
    Dim j As Integer                                                                        '' ループカウンタ
    Dim k As Integer                                                                        '' ループカウンタ
    Dim rs As ADODB.Recordset                                                               '' SQL文
        
    mCon_BANUSI.BeginTrans

        
    Set rs = mRS_BANUSI
    rs.AddNew Array("BanusiCode"), Array(mBuf.BanusiCode)
    With mBuf
        With .head
            rs("RecordSpec") = .RecordSpec                                                  '' レコード種別
            rs("DataKubun") = .DataKubun                                                    '' データ区分
            With .MakeDate
                rs("MakeDate") = .Year & .Month & .Day                                      '' 年月日
            End With ' MakeDate
        End With ' head
        rs("BanusiName_Co") = .BanusiName_Co                                                '' 馬主名_法人格有
        rs("BanusiName") = .BanusiName                                                      '' 馬主名_法人格無
        rs("BanusiNameKana") = .BanusiNameKana                                              '' 馬主名半角カナ
        rs("BanusiNameEng") = .BanusiNameEng                                                '' 馬主名欧字
        rs("Fukusyoku") = .Fukusyoku                                                        '' 服色標示
        With .HonRuikei(0)
            rs("H_SetYear") = .SetYear                                                      '' 設定年
            rs("H_HonSyokinTotal") = .HonSyokinTotal                                        '' 本賞金合計
            rs("H_Fukasyokin") = .Fukasyokin                                                '' 付加賞金合計
            For j = 0 To 5
                rs("H_Chakukaisu" & j + 1) = .Chakukaisu(j)                                 '' 着回数
            Next j
        End With ' HonRuikei(0)
        With .HonRuikei(1)
            rs("R_SetYear") = .SetYear                                                      '' 設定年
            rs("R_HonSyokinTotal") = .HonSyokinTotal                                        '' 本賞金合計
            rs("R_Fukasyokin") = .Fukasyokin                                                '' 付加賞金合計
            For j = 0 To 5
                rs("R_Chakukaisu" & j + 1) = .Chakukaisu(j)                                 '' 着回数
            Next j
        End With ' HonRuikei(1)
        
    
    End With ' mBuf
    rs.Update ' BANUSI
    
    mCon_BANUSI.CommitTrans

    Set rs = Nothing
    
    InsertDB = True
    Exit Function

ErrorHandler:
    If Err.Number <> -2147217887 Then
        gApp.ErrLog
    End If
    rs.CancelUpdate
    
    mCon_BANUSI.RollbackTrans

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
    Dim i As Integer                                                                        '' ループカウンタ
    Dim j As Integer                                                                        '' ループカウンタ
    Dim k As Integer                                                                        '' ループカウンタ
    Dim strSQL As String                                                                    '' SQL文
        
    Set rs = mRS_BANUSI

    rs.Seek Array(mBuf.BanusiCode)
    
    
    With mBuf.head.MakeDate
        If rs("Makedate") > .Year & .Month & .Day Then
            UpdateDB = True
            Exit Function
        End If
    End With

    
    mCon_BANUSI.BeginTrans

    With mBuf
        With .head
            rs("RecordSpec") = .RecordSpec                                                  '' レコード種別
            rs("DataKubun") = .DataKubun                                                    '' データ区分
            With .MakeDate
                rs("MakeDate") = .Year & .Month & .Day                                      '' 年月日
            End With ' MakeDate
        End With ' head
        rs("BanusiName_Co") = .BanusiName_Co                                                '' 馬主名_法人格有
        rs("BanusiName") = .BanusiName                                                      '' 馬主名_法人格無
        rs("BanusiNameKana") = .BanusiNameKana                                              '' 馬主名半角カナ
        rs("BanusiNameEng") = .BanusiNameEng                                                '' 馬主名欧字
        rs("Fukusyoku") = .Fukusyoku                                                        '' 服色標示
        With .HonRuikei(0)
            rs("H_SetYear") = .SetYear                                                      '' 設定年
            rs("H_HonSyokinTotal") = .HonSyokinTotal                                        '' 本賞金合計
            rs("H_Fukasyokin") = .Fukasyokin                                                '' 付加賞金合計
            For j = 0 To 5
                rs("H_Chakukaisu" & j + 1) = .Chakukaisu(j)                                 '' 着回数
            Next j
        End With ' HonRuikei(0)
        With .HonRuikei(1)
            rs("R_SetYear") = .SetYear                                                      '' 設定年
            rs("R_HonSyokinTotal") = .HonSyokinTotal                                        '' 本賞金合計
            rs("R_Fukasyokin") = .Fukasyokin                                                '' 付加賞金合計
            For j = 0 To 5
                rs("R_Chakukaisu" & j + 1) = .Chakukaisu(j)                                 '' 着回数
            Next j
        End With ' HonRuikei(1)
        
    
    End With ' mBuf
    rs.Update ' BANUSI
    
    mCon_BANUSI.CommitTrans

    Set rs = Nothing
    
    UpdateDB = True
    Exit Function

ErrorHandler:
    gApp.ErrLog
    
    mCon_BANUSI.RollbackTrans

    UpdateDB = False
End Function



'
'   機能: データベースに追加する
'
'   備考: なし
'
Private Function ImportDB() As Boolean
On Error GoTo ErrorHandler
    Dim i As Integer                                                                        '' ループカウンタ
    Dim j As Integer                                                                        '' ループカウンタ
    Dim k As Integer                                                                        '' ループカウンタ
    Dim rs As ADODB.Recordset                                                               '' SQL文
        
    Set rs = mRS_BANUSI
    
    mCon_BANUSI.BeginTrans
        
    rs.Seek Array(mBuf.BanusiCode)
    
    If Not rs.EOF Then
        With mBuf.head.MakeDate
            If rs("Makedate") > .Year & .Month & .Day Then
                ImportDB = True
                mCon_BANUSI.RollbackTrans
                Exit Function
            End If
        End With
    Else
        rs.AddNew
        rs("BanusiCode") = mBuf.BanusiCode
    End If
    
    With mBuf
        With .head
            rs("RecordSpec") = .RecordSpec                                                  '' レコード種別
            rs("DataKubun") = .DataKubun                                                    '' データ区分
            With .MakeDate
                rs("MakeDate") = .Year & .Month & .Day                                      '' 年月日
            End With ' MakeDate
        End With ' head
        rs("BanusiName_Co") = .BanusiName_Co                                                '' 馬主名_法人格有
        rs("BanusiName") = .BanusiName                                                      '' 馬主名_法人格無
        rs("BanusiNameKana") = .BanusiNameKana                                              '' 馬主名半角カナ
        rs("BanusiNameEng") = .BanusiNameEng                                                '' 馬主名欧字
        rs("Fukusyoku") = .Fukusyoku                                                        '' 服色標示
        With .HonRuikei(0)
            rs("H_SetYear") = .SetYear                                                      '' 設定年
            rs("H_HonSyokinTotal") = .HonSyokinTotal                                        '' 本賞金合計
            rs("H_Fukasyokin") = .Fukasyokin                                                '' 付加賞金合計
            For j = 0 To 5
                rs("H_Chakukaisu" & j + 1) = .Chakukaisu(j)                                 '' 着回数
            Next j
        End With ' HonRuikei(0)
        With .HonRuikei(1)
            rs("R_SetYear") = .SetYear                                                      '' 設定年
            rs("R_HonSyokinTotal") = .HonSyokinTotal                                        '' 本賞金合計
            rs("R_Fukasyokin") = .Fukasyokin                                                '' 付加賞金合計
            For j = 0 To 5
                rs("R_Chakukaisu" & j + 1) = .Chakukaisu(j)                                 '' 着回数
            Next j
        End With ' HonRuikei(1)
        
    End With ' mBuf
    
    rs.Update ' BANUSI
    
    mCon_BANUSI.CommitTrans

    Set rs = Nothing
    
    ImportDB = True
    Exit Function

ErrorHandler:
    gApp.ErrLog
    rs.CancelUpdate
    mCon_BANUSI.RollbackTrans

    ImportDB = False
End Function

