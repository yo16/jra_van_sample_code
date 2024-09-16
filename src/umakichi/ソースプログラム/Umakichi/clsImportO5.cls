VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsImportO5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Import JV_O5_ODDS_SANREN "
'
'   JVData "O5" データベース登録クラス
'

Option Explicit
Option Compare Binary
Implements clsIImport

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部変数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private mBuf As JV_O5_ODDS_SANREN

Private mCon_ODDS_SANREN(0 To 9) As ADODB.Connection

Private mRS_ODDS_SANREN(0 To 9) As ADODB.Recordset


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
    Dim i As Integer
    
    ' コネクションのインスタンス生成
    For i = 0 To 9
        Set mCon_ODDS_SANREN(i) = New ADODB.Connection
    
        ' レコードセットのインスタンス生成
        Set mRS_ODDS_SANREN(i) = New ADODB.Recordset
    Next i
    
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
    Dim i As Long
        
            ' コネクションオープン
    For i = 0 To 9
        strCon = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & gApp.R_DBPath & "\"
        mCon_ODDS_SANREN(i).Open strCon & "subODDS_SANREN" & i & ".mdb"
        
        ' レコードセットオープン
        With mRS_ODDS_SANREN(i)
            .CursorLocation = adUseServer
            .Index = "PrimaryKey"
            .Open "ODDS_SANREN", mCon_ODDS_SANREN(i), adOpenKeyset, adLockOptimistic, adCmdTableDirect
            If Not (.EOF Or .BOF) Then
                .MoveFirst
            End If
        End With
    Next i

    Exit Sub
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   機能: コネクション、レコードセットをクローズする
'
'   備考: なし
'
Private Sub clsIImport_CloseDB()
On Error GoTo ErrorHandler
    Dim i As Integer

    For i = 0 To 9
        mRS_ODDS_SANREN(i).Close
    
        mCon_ODDS_SANREN(i).Close
        
    Next i
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

        Call SetDataFromByte_O5(lBuf, mBuf)     '' 構造体に代入する

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
    Dim intMDBIndex As Integer                      '' 分割MDBのインデックス
    
    Year = Mid$(lBuf, 12, 4)                   '' 開催年
    MonthDay = Mid$(lBuf, 16, 4)               '' 開催月日
    JyoCD = Mid$(lBuf, 20, 2)                  '' 競馬場コード
    Kaiji = Mid$(lBuf, 22, 2)                  '' 開催回[第N回]
    Nichiji = Mid$(lBuf, 24, 2)                '' 開催日目[N日目]
    RaceNum = Mid$(lBuf, 26, 2)                '' レース番号

    ' 分割MDBのインデックス番号を場所コードから得る
    If JyoCD >= "01" And JyoCD <= "10" Then
        intMDBIndex = val(mBuf.id.JyoCD) - 1
    Else
        intMDBIndex = 0
    End If
    
    strSQL = "DELETE * FROM ODDS_SANREN"
    strSQL = strSQL & " WHERE [Year]   ='" & Year & "'"
    strSQL = strSQL & " AND [MonthDay] ='" & MonthDay & "'"
    strSQL = strSQL & " AND [JyoCD]    ='" & JyoCD & "'"
    strSQL = strSQL & " AND [Kaiji]    ='" & Kaiji & "'"
    strSQL = strSQL & " AND [Nichiji]  ='" & Nichiji & "'"
    strSQL = strSQL & " AND [RaceNum]  ='" & RaceNum & "'"
    
    mCon_ODDS_SANREN(intMDBIndex).Execute strSQL, , adExecuteNoRecords
    
End Sub


'
'   機能: データベースに挿入する
'
'   備考: なし
'
Private Function InsertDB() As Boolean
On Error GoTo ErrorHandler
    Dim i As Integer                                                                '' ループカウンタ
    Dim j As Integer                                                                '' ループカウンタ
    Dim k As Integer                                                                '' ループカウンタ
    Dim rs As ADODB.Recordset                                                       '' SQL文
    Dim intMDBIndex As Integer                                                      '' 分割MDBのインデックス
    Dim strData As String           '' オッズデータ
    
    ' 分割MDBのインデックス番号を場所コードから得る
    If mBuf.id.JyoCD >= "01" And mBuf.id.JyoCD <= "10" Then
        intMDBIndex = val(mBuf.id.JyoCD) - 1
    Else
        intMDBIndex = 0
    End If

    mCon_ODDS_SANREN(intMDBIndex).BeginTrans
    
    ' ODDS_SANREN_HEAD (オッズ_3連複_ヘッダ)
    Set rs = mRS_ODDS_SANREN(intMDBIndex)
    With mBuf.id
        rs.AddNew Array("Year", "MonthDay", "JyoCD", "Kaiji", "Nichiji", "RaceNum"), Array(.Year, .MonthDay, .JyoCD, .Kaiji, .Nichiji, .RaceNum)
    End With
    With mBuf
        With .head
            rs("RecordSpec") = .RecordSpec                                          '' レコード種別
            rs("DataKubun") = .DataKubun                                            '' データ区分
            With .MakeDate
                rs("MakeDate") = .Year & .Month & .Day                              '' 年月日
            End With ' MakeDate
        End With ' head
        With .HappyoTime
            rs("HappyoTime") = .Month & .Day & .Hour & .Minute
        End With ' HappyoTime
        rs("TorokuTosu") = .TorokuTosu                                              '' 登録頭数
        rs("SyussoTosu") = .SyussoTosu                                              '' 出走頭数
        rs("SanrenFlag") = .SanrenpukuFlag                                          '' 発売フラグ　3連複
        rs("TotalHyosuSanren") = .TotalHyosuSanrenpuku                          '' 3連複票数合計
    End With
    
    ' ODDS_SANREN (オッズ_3連複)
    strData = ""
    If mBuf.SanrenpukuFlag <> "0" Then '（0:発売なし 1:発売前取消 3:発売後取消 7:発売あり）
        For i = 0 To 815
            If mBuf.OddsSanrenInfo(i).Kumi <> "      " Then
                With mBuf
                    With .OddsSanrenInfo(i)
                        strData = strData & .Kumi
                        strData = strData & .Odds
                        strData = strData & .Ninki
                    End With ' OddsSanrenInfo
                End With ' mbuf
                
            End If
        Next i
        rs("DATA") = strData
    End If
        
    rs.Update ' ODDS_SANREN

    mCon_ODDS_SANREN(intMDBIndex).CommitTrans
    
    Set rs = Nothing
    
    InsertDB = True
    Exit Function
    
ErrorHandler:
    If Err.Number <> -2147217887 Then
        gApp.ErrLog
    End If
    rs.CancelUpdate
    
    mCon_ODDS_SANREN(intMDBIndex).RollbackTrans

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
    Dim i As Integer                                                                '' ループカウンタ
    Dim j As Integer                                                                '' ループカウンタ
    Dim k As Integer                                                                '' ループカウンタ
    Dim strSQL As String                                                            '' SQL文
    Dim intMDBIndex As Integer                                                      '' 分割MDBのインデックス
    Dim strData As String
    
    ' 分割MDBのインデックス番号を場所コードから得る
    If mBuf.id.JyoCD >= "01" And mBuf.id.JyoCD <= "10" Then
        intMDBIndex = val(mBuf.id.JyoCD) - 1
    Else
        intMDBIndex = 0
    End If

    Set rs = mRS_ODDS_SANREN(intMDBIndex)

    With mBuf.id
        rs.Seek Array(.Year, .MonthDay, .JyoCD, .Kaiji, .Nichiji, .RaceNum)
    End With

    With mBuf.head.MakeDate
        If rs("Makedate") > .Year & .Month & .Day Then
            UpdateDB = True
            Exit Function
        End If
    End With

    mCon_ODDS_SANREN(intMDBIndex).BeginTrans

    
    ' ヘッダ部分
    With mBuf
        With .head
            rs("RecordSpec") = .RecordSpec                                          '' レコード種別
            rs("DataKubun") = .DataKubun                                            '' データ区分
            With .MakeDate
                rs("MakeDate") = .Year & .Month & .Day                              '' 年月日
            End With ' MakeDate
        End With ' head
        With .HappyoTime
            rs("HappyoTime") = .Month & .Day & .Hour & .Minute
        End With ' HappyoTime
        rs("TorokuTosu") = .TorokuTosu                                              '' 登録頭数
        rs("SyussoTosu") = .SyussoTosu                                              '' 出走頭数
        rs("SanrenFlag") = .SanrenpukuFlag                                          '' 発売フラグ　3連複
        rs("TotalHyosuSanren") = .TotalHyosuSanrenpuku                              '' 3連複票数合計
    End With
    
    ' ODDS_SANREN (オッズ_３連複)
    
    If mBuf.SanrenpukuFlag <> "0" Then '（0:発売なし 1:発売前取消 3:発売後取消 7:発売あり）
        For i = 0 To 815
            If mBuf.OddsSanrenInfo(i).Kumi <> "      " Then
                With mBuf
                    With .OddsSanrenInfo(i)
                        strData = strData & .Kumi
                        strData = strData & .Odds
                        strData = strData & .Ninki
                    End With ' OddsSanrenInfo
                End With ' mbuf
            End If
        Next i
        rs("DATA") = strData
    End If
    
    rs.Update ' ODDS_SANREN
        
    mCon_ODDS_SANREN(intMDBIndex).CommitTrans

    Set rs = Nothing
    
    UpdateDB = True
    Exit Function
    
ErrorHandler:
    gApp.ErrLog
    
    
    mCon_ODDS_SANREN(intMDBIndex).RollbackTrans

    UpdateDB = False
End Function




