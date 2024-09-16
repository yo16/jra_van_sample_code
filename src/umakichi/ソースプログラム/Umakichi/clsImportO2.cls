VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsImportO2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Import JV_O2_ODDS_UMAREN "
'
'   JVData "O2" データベース登録クラス
'

Option Explicit
Option Compare Binary
Implements clsIImport

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部変数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private mBuf As JV_O2_ODDS_UMAREN

Private mCon_ODDS_UMAREN As ADODB.Connection

Private mRS_ODDS_UMAREN As ADODB.Recordset


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
    Set mCon_ODDS_UMAREN = New ADODB.Connection

    ' レコードセットのインスタンス生成
    Set mRS_ODDS_UMAREN = New ADODB.Recordset
    
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
    mCon_ODDS_UMAREN.Open strCon & "subODDS_UMAREN.mdb"

    ' レコードセットオープン
    With mRS_ODDS_UMAREN
        .CursorLocation = adUseServer
        .Index = "PrimaryKey"
        .Open "ODDS_UMAREN", mCon_ODDS_UMAREN, adOpenKeyset, adLockOptimistic, adCmdTableDirect
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
    mRS_ODDS_UMAREN.Close

    mCon_ODDS_UMAREN.Close

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

        Call SetDataFromByte_O2(lBuf, mBuf)     '' 構造体に代入する

        If Not InsertDB() Then
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

    strSQL = "DELETE * FROM ODDS_UMAREN"
    strSQL = strSQL & " WHERE [Year]   ='" & Year & "'"
    strSQL = strSQL & " AND [MonthDay] ='" & MonthDay & "'"
    strSQL = strSQL & " AND [JyoCD]    ='" & JyoCD & "'"
    strSQL = strSQL & " AND [Kaiji]    ='" & Kaiji & "'"
    strSQL = strSQL & " AND [Nichiji]  ='" & Nichiji & "'"
    strSQL = strSQL & " AND [RaceNum]  ='" & RaceNum & "'"
    
    mCon_ODDS_UMAREN.Execute strSQL, , adExecuteNoRecords
    
End Sub


'
'   機能: データベースに挿入する
'
'   備考: なし
'
Private Function InsertDB() As Boolean
On Error GoTo ErrorHandler
    Dim i As Integer                                                                 '' ループカウンタ
    Dim j As Integer                                                                 '' ループカウンタ
    Dim k As Integer                                                                 '' ループカウンタ
    Dim rs As ADODB.Recordset                                                        '' SQL文
    Dim strData As String
    
    mCon_ODDS_UMAREN.BeginTrans

    
    ' ODDS_UMAREN_HEAD (オッズ_馬連_ヘッダ)
    Set rs = mRS_ODDS_UMAREN
    With mBuf.id
        rs.AddNew Array("Year", "MonthDay", "JyoCD", "Kaiji", "Nichiji", "RaceNum"), Array(.Year, .MonthDay, .JyoCD, .Kaiji, .Nichiji, .RaceNum)
    End With
    With mBuf
        With .head
            rs("RecordSpec") = .RecordSpec                                           '' レコード種別
            rs("DataKubun") = .DataKubun                                             '' データ区分
            With .MakeDate
                rs("MakeDate") = .Year & .Month & .Day                               '' 年月日
            End With ' MakeDate
        End With ' head
        With .HappyoTime
            rs("HappyoTime") = .Month & .Day & .Hour & .Minute
        End With ' HappyoTime
        rs("TorokuTosu") = .TorokuTosu                                               '' 登録頭数
        rs("SyussoTosu") = .SyussoTosu                                               '' 出走頭数
        rs("UmarenFlag") = .UmarenFlag                                               '' 発売フラグ　馬連
        rs("TotalHyosuUmaren") = .TotalHyosuUmaren                                   '' 馬連票数合計
    End With
    
    
    ' ODDS_UMAREN (オッズ_馬連)
    
    If mBuf.UmarenFlag <> "0" Then '（0:発売なし 1:発売前取消 3:発売後取消 7:発売あり）
        For i = 0 To 152
            If mBuf.OddsUmarenInfo(i).Kumi <> "    " Then
                With mBuf
                    With .OddsUmarenInfo(i)
                        strData = strData & .Kumi
                        strData = strData & .Odds
                        strData = strData & .Ninki
                    End With ' OddsUmarenInfo
                End With
            End If
        Next i
        rs("DATA") = strData
    End If
    
    rs.Update ' ODDS_UMAREN

    mCon_ODDS_UMAREN.CommitTrans

    
    Set rs = Nothing
    
    InsertDB = True
    Exit Function

ErrorHandler:
    If Err.Number <> -2147217887 Then
        gApp.ErrLog
    End If
    rs.CancelUpdate
    
    mCon_ODDS_UMAREN.RollbackTrans

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
    Dim i As Integer                                                                 '' ループカウンタ
    Dim j As Integer                                                                 '' ループカウンタ
    Dim k As Integer                                                                 '' ループカウンタ
    Dim strSQL As String                                                             '' SQL文
    Dim strData As String
    
    Set rs = mRS_ODDS_UMAREN
    
    With mBuf.id
        rs.Seek Array(.Year, .MonthDay, .JyoCD, .Kaiji, .Nichiji, .RaceNum)
    End With

    With mBuf.head.MakeDate
        If rs("Makedate") > .Year & .Month & .Day Then
            UpdateDB = True
            Exit Function
        End If
    End With

    
    mCon_ODDS_UMAREN.BeginTrans

    With mBuf
        With .head
            rs("RecordSpec") = .RecordSpec                                           '' レコード種別
            rs("DataKubun") = .DataKubun                                             '' データ区分
            With .MakeDate
                rs("MakeDate") = .Year & .Month & .Day                               '' 年月日
            End With ' MakeDate
        End With ' head
        With .HappyoTime
            rs("HappyoTime") = .Month & .Day & .Hour & .Minute
        End With ' HappyoTime
        rs("TorokuTosu") = .TorokuTosu                                               '' 登録頭数
        rs("SyussoTosu") = .SyussoTosu                                               '' 出走頭数
        rs("UmarenFlag") = .UmarenFlag                                               '' 発売フラグ　馬連
        rs("TotalHyosuUmaren") = .TotalHyosuUmaren                                   '' 馬連票数合計
    End With

    ' ODDS_UMAREN (オッズ_馬連)
    
    If mBuf.UmarenFlag <> "0" Then '（0:発売なし 1:発売前取消 3:発売後取消 7:発売あり）
        For i = 0 To 152
            If mBuf.OddsUmarenInfo(i).Kumi <> "    " Then
                With mBuf
                    With .OddsUmarenInfo(i)
                        strData = strData & .Kumi
                        strData = strData & .Odds
                        strData = strData & .Ninki
                    End With ' OddsUmarenInfo
                End With
            End If
        Next i
        rs("DATA") = strData
    End If
    
    rs.Update

    mCon_ODDS_UMAREN.CommitTrans

    
    Set rs = Nothing

    UpdateDB = True
    Exit Function

ErrorHandler:
    gApp.ErrLog
    
    mCon_ODDS_UMAREN.RollbackTrans

    UpdateDB = False
End Function


