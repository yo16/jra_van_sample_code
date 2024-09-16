VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsImportO1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Import JV_O1_ODDS_TANFUKUWAKU "
'
'   JVData "O1" データベース登録クラス
'

Option Explicit
Option Compare Binary
Implements clsIImport

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部関数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private mBuf As JV_O1_ODDS_TANFUKUWAKU

Private mCon_ODDS_TANPUKUWAKU As ADODB.Connection

Private mRS_ODDS_TANPUKUWAKU As ADODB.Recordset



'
'   機能: クラス初期化イベント
'
'   備考: なし
'
Private Sub Class_Initialize()
On Error GoTo ErrorHandler
    Dim strCon As String

    ' コネクションのインスタンス生成
    Set mCon_ODDS_TANPUKUWAKU = New ADODB.Connection

    ' レコードセットのインスタンス生成
    Set mRS_ODDS_TANPUKUWAKU = New ADODB.Recordset
    
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
    mCon_ODDS_TANPUKUWAKU.Open strCon & "subODDS_TANPUKUWAKU.mdb"

    ' レコードセットオープン
    With mRS_ODDS_TANPUKUWAKU
        .CursorLocation = adUseServer
        .Index = "PrimaryKey"
        .Open "ODDS_TANPUKUWAKU", mCon_ODDS_TANPUKUWAKU, adOpenKeyset, adLockOptimistic, adCmdTableDirect
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
    mRS_ODDS_TANPUKUWAKU.Close
    
    mCon_ODDS_TANPUKUWAKU.Close
    
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

        Call SetDataFromByte_O1(lBuf, mBuf)     '' 構造体に代入する

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

    strSQL = "DELETE * FROM RACE"
    strSQL = strSQL & " WHERE [Year]   ='" & Year & "'"
    strSQL = strSQL & " AND [MonthDay] ='" & MonthDay & "'"
    strSQL = strSQL & " AND [JyoCD]    ='" & JyoCD & "'"
    strSQL = strSQL & " AND [Kaiji]    ='" & Kaiji & "'"
    strSQL = strSQL & " AND [Nichiji]  ='" & Nichiji & "'"
    strSQL = strSQL & " AND [RaceNum]  ='" & RaceNum & "'"
    
    mCon_ODDS_TANPUKUWAKU.Execute strSQL, , adExecuteNoRecords
    
End Sub




'
'   機能: データベースに挿入する
'
'   備考: なし
'
Private Function InsertDB() As Boolean
On Error GoTo ErrorHandler
    Dim rs As ADODB.Recordset
    Dim i As Integer                                                                 '' ループカウンタ
    Dim j As Integer                                                                 '' ループカウンタ
    Dim k As Integer                                                                 '' ループカウンタ
    Dim strData As String
    Dim strData2 As String
    Dim strData3 As String
    mCon_ODDS_TANPUKUWAKU.BeginTrans
    
    
    ' ヘッダ部分
    Set rs = mRS_ODDS_TANPUKUWAKU
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
        rs("TansyoFlag") = .TansyoFlag                                               '' 発売フラグ
        rs("FukusyoFlag") = .FukusyoFlag                                             '' 発売フラグ
        rs("WakurenFlag") = .WakurenFlag                                             '' 発売フラグ　枠連
        rs("FukuChakuBaraiKey") = .FukuChakuBaraiKey                                 '' 複勝着払キー
        rs("TotalHyosuTansyo") = .TotalHyosuTansyo                                   '' 単勝票数合計
        rs("TotalHyosuFukusyo") = .TotalHyosuFukusyo                                 '' 複勝票数合計
        rs("TotalHyosuWakuren") = .TotalHyosuWakuren                                 '' 枠連票数合計
    End With
    
    
    ' ODDS_TANPUKU (オッズ_単複)
    
    If mBuf.TansyoFlag <> "0" Then '（0:発売なし 1:発売前取消 3:発売後取消 7:発売あり）
        For i = 0 To 27
            With mBuf.OddsTansyoInfo(i)
                If .Umaban <> "  " Then
                    strData = strData & .Umaban
                    strData = strData & .Odds
                    strData = strData & .Ninki
                End If
            End With ' mBuf.OddsTansyoInfo
        Next i
        rs("TanDATA") = strData
    End If
    If mBuf.FukusyoFlag <> "0" Then '（0:発売なし 1:発売前取消 3:発売後取消 7:発売あり）
        For i = 0 To 27
            With mBuf.OddsFukusyoInfo(i)
                If .Umaban <> "  " Then
                    strData2 = strData2 & .Umaban
                    strData2 = strData2 & .OddsLow
                    strData2 = strData2 & .OddsHigh
                    strData2 = strData2 & .Ninki
                End If
            End With ' mBuf.OddsFukusyoInfo
        Next i
        rs("FukuDATA") = strData2
    End If
    If mBuf.WakurenFlag <> "0" Then '（0:発売なし 1:発売前取消 3:発売後取消 7:発売あり）
        For i = 0 To 35
            With mBuf.OddsWakurenInfo(i)
                If .Kumi <> "  " Then
                    strData3 = strData3 & .Kumi
                    strData3 = strData3 & .Odds
                    strData3 = strData3 & .Ninki
                End If
            End With ' mBuf.OddsWakurenInfo
        Next i
        rs("WakuDATA") = strData3
    End If
    
    rs.Update
    
    mCon_ODDS_TANPUKUWAKU.CommitTrans
    
    InsertDB = True
    Exit Function

ErrorHandler:
    If Err.Number <> -2147217887 Then
        gApp.ErrLog
    End If
    rs.CancelUpdate
    
    mCon_ODDS_TANPUKUWAKU.RollbackTrans

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
    Dim strData2 As String
    Dim strData3 As String
    
    Set rs = mRS_ODDS_TANPUKUWAKU

    With mBuf.id
        rs.Seek Array(.Year, .MonthDay, .JyoCD, .Kaiji, .Nichiji, .RaceNum)
    End With

    With mBuf.head.MakeDate
        If rs("Makedate") > .Year & .Month & .Day Then
            UpdateDB = True
            Exit Function
        End If
    End With

    
    mCon_ODDS_TANPUKUWAKU.BeginTrans
    
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
        rs("TansyoFlag") = .TansyoFlag                                               '' 発売フラグ
        rs("FukusyoFlag") = .FukusyoFlag                                             '' 発売フラグ
        rs("WakurenFlag") = .WakurenFlag                                             '' 発売フラグ　枠連
        rs("FukuChakuBaraiKey") = .FukuChakuBaraiKey                                 '' 複勝着払キー
        rs("TotalHyosuTansyo") = .TotalHyosuTansyo                                   '' 単勝票数合計
        rs("TotalHyosuFukusyo") = .TotalHyosuFukusyo                                 '' 複勝票数合計
        rs("TotalHyosuWakuren") = .TotalHyosuWakuren                                 '' 枠連票数合計
    End With
    
    
    ' ODDS_TANPUKU (オッズ_単複)
    
    If mBuf.TansyoFlag <> "0" Then '（0:発売なし 1:発売前取消 3:発売後取消 7:発売あり）
        For i = 0 To 27
            With mBuf.OddsTansyoInfo(i)
                If .Umaban <> "  " Then
                    strData = strData & .Umaban
                    strData = strData & .Odds
                    strData = strData & .Ninki
                End If
            End With ' mBuf.OddsTansyoInfo
        Next i
        rs("TanDATA") = strData
    End If
    If mBuf.FukusyoFlag <> "0" Then '（0:発売なし 1:発売前取消 3:発売後取消 7:発売あり）
        For i = 0 To 27
            With mBuf.OddsFukusyoInfo(i)
                If .Umaban <> "  " Then
                    strData2 = strData2 & .Umaban
                    strData2 = strData2 & .OddsLow
                    strData2 = strData2 & .OddsHigh
                    strData2 = strData2 & .Ninki
                End If
            End With ' mBuf.OddsFukusyoInfo
        Next i
        rs("FukuDATA") = strData2
    End If
    If mBuf.WakurenFlag <> "0" Then '（0:発売なし 1:発売前取消 3:発売後取消 7:発売あり）
        For i = 0 To 35
            With mBuf.OddsWakurenInfo(i)
                If .Kumi <> "  " Then
                    strData3 = strData3 & .Kumi
                    strData3 = strData3 & .Odds
                    strData3 = strData3 & .Ninki
                End If
            End With ' mBuf.OddsWakurenInfo
        Next i
        rs("WakuDATA") = strData3
    End If
    
    rs.Update
    
    mCon_ODDS_TANPUKUWAKU.CommitTrans

    
    Set rs = Nothing

    UpdateDB = True
    Exit Function

ErrorHandler:
    gApp.ErrLog
    
    mCon_ODDS_TANPUKUWAKU.RollbackTrans

    UpdateDB = False
End Function


