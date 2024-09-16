VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsImportH1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Import JV_H1_HYOSU_ZENKAKE "
'
'   JVData "H1" データベース登録クラス
'

Option Explicit
Option Compare Binary
Implements clsIImport

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部変数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private mBuf As JV_H1_HYOSU_ZENKAKE

Private mCon_HYOSU As ADODB.Connection
Private mCon_HYOSU_TANPUKU As ADODB.Connection
Private mCon_HYOSU_WAKU As ADODB.Connection
Private mCon_HYOSU_UMAREN As ADODB.Connection
Private mCon_HYOSU_WIDE As ADODB.Connection
Private mCon_HYOSU_UMATAN(0 To 9) As ADODB.Connection
Private mCon_HYOSU_SANREN(0 To 9) As ADODB.Connection

Private mRS_HYOSU As ADODB.Recordset
Private mRS_HYOSU_TANPUKU As ADODB.Recordset
Private mRS_HYOSU_WAKU As ADODB.Recordset
Private mRS_HYOSU_UMAREN As ADODB.Recordset
Private mRS_HYOSU_WIDE As ADODB.Recordset
Private mRS_HYOSU_UMATAN(0 To 9) As ADODB.Recordset
Private mRS_HYOSU_SANREN(0 To 9) As ADODB.Recordset


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
    Dim i As Integer
    Dim strCon As String
    
    ' コネクションのインスタンス生成
    Set mCon_HYOSU = New ADODB.Connection
    Set mCon_HYOSU_TANPUKU = New ADODB.Connection
    Set mCon_HYOSU_WAKU = New ADODB.Connection
    Set mCon_HYOSU_UMAREN = New ADODB.Connection
    Set mCon_HYOSU_WIDE = New ADODB.Connection
    For i = 0 To 9
        Set mCon_HYOSU_UMATAN(i) = New ADODB.Connection
        Set mCon_HYOSU_SANREN(i) = New ADODB.Connection
    Next i
    
    ' レコードセットのインスタンス生成
    Set mRS_HYOSU = New ADODB.Recordset
    Set mRS_HYOSU_TANPUKU = New ADODB.Recordset
    Set mRS_HYOSU_WAKU = New ADODB.Recordset
    Set mRS_HYOSU_UMAREN = New ADODB.Recordset
    Set mRS_HYOSU_WIDE = New ADODB.Recordset
    For i = 0 To 9
        Set mRS_HYOSU_UMATAN(i) = New ADODB.Recordset
        Set mRS_HYOSU_SANREN(i) = New ADODB.Recordset
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
    strCon = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & gApp.R_DBPath & "\"
    mCon_HYOSU.Open strCon & "subHYOSU.mdb"
    mCon_HYOSU_TANPUKU.Open strCon & "subHYOSU_TANPUKU.mdb"
    mCon_HYOSU_WAKU.Open strCon & "subHYOSU_WAKU.mdb"
    mCon_HYOSU_UMAREN.Open strCon & "subHYOSU_UMAREN.mdb"
    mCon_HYOSU_WIDE.Open strCon & "subHYOSU_WIDE.mdb"
    For i = 0 To 9
        mCon_HYOSU_UMATAN(i).Open strCon & "subHYOSU_UMATAN" & i & ".mdb"
        mCon_HYOSU_SANREN(i).Open strCon & "subHYOSU_SANREN" & i & ".mdb"
    Next i
     
    ' レコードセットオープン
    With mRS_HYOSU
        .CursorLocation = adUseServer
        .Index = "PrimaryKey"
        .Open "HYOSU", mCon_HYOSU, adOpenKeyset, adLockOptimistic, adCmdTableDirect
        If Not (.EOF Or .BOF) Then
            .MoveFirst
        End If
    End With
    With mRS_HYOSU_TANPUKU
        .CursorLocation = adUseServer
        .Index = "PrimaryKey"
        .Open "HYOSU_TANPUKU", mCon_HYOSU_TANPUKU, adOpenKeyset, adLockOptimistic, adCmdTableDirect
        If Not (.EOF Or .BOF) Then
            .MoveFirst
        End If
    End With
    With mRS_HYOSU_WAKU
        .CursorLocation = adUseServer
        .Index = "PrimaryKey"
        .Open "HYOSU_WAKU", mCon_HYOSU_WAKU, adOpenKeyset, adLockOptimistic, adCmdTableDirect
        If Not (.EOF Or .BOF) Then
            .MoveFirst
        End If
    End With
    With mRS_HYOSU_UMAREN
        .CursorLocation = adUseServer
        .Index = "PrimaryKey"
        .Open "HYOSU_UMAREN", mCon_HYOSU_UMAREN, adOpenKeyset, adLockOptimistic, adCmdTableDirect
        If Not (.EOF Or .BOF) Then
            .MoveFirst
        End If
    End With
    With mRS_HYOSU_WIDE
        .CursorLocation = adUseServer
        .Index = "PrimaryKey"
        .Open "HYOSU_WIDE", mCon_HYOSU_WIDE, adOpenKeyset, adLockOptimistic, adCmdTableDirect
        If Not (.EOF Or .BOF) Then
            .MoveFirst
        End If
    End With
    For i = 0 To 9
        With mRS_HYOSU_UMATAN(i)
            .CursorLocation = adUseServer
            .Index = "PrimaryKey"
            .Open "HYOSU_UMATAN", mCon_HYOSU_UMATAN(i), adOpenKeyset, adLockOptimistic, adCmdTableDirect
            If Not (.EOF Or .BOF) Then
                .MoveFirst
            End If
        End With
        With mRS_HYOSU_SANREN(i)
            .CursorLocation = adUseServer
            .Index = "PrimaryKey"
            .Open "HYOSU_SANREN", mCon_HYOSU_SANREN(i), adOpenKeyset, adLockOptimistic, adCmdTableDirect
            If Not (.EOF Or .BOF) Then
                .MoveFirst
            End If
        End With
    Next i
    
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
    Dim i As Integer

    ' レコードセットを閉じる
    mRS_HYOSU.Close
    mRS_HYOSU_TANPUKU.Close
    mRS_HYOSU_WAKU.Close
    mRS_HYOSU_UMAREN.Close
    mRS_HYOSU_WIDE.Close
    For i = 0 To 9
        mRS_HYOSU_UMATAN(i).Close
        mRS_HYOSU_SANREN(i).Close
    Next i
    
    ' コネクションを閉じる
    mCon_HYOSU.Close
    mCon_HYOSU_TANPUKU.Close
    mCon_HYOSU_WAKU.Close
    mCon_HYOSU_UMAREN.Close
    mCon_HYOSU_WIDE.Close
    For i = 0 To 9
        mCon_HYOSU_UMATAN(i).Close
        mCon_HYOSU_SANREN(i).Close
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

        ' 構造体に代入する
        Call SetDataFromByte_H1(lBuf, mBuf)

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
    
    strSQL = "DELETE * FROM RACE"
    strSQL = strSQL & " WHERE [Year]   ='" & Year & "'"
    strSQL = strSQL & " AND [MonthDay] ='" & MonthDay & "'"
    strSQL = strSQL & " AND [JyoCD]    ='" & JyoCD & "'"
    strSQL = strSQL & " AND [Kaiji]    ='" & Kaiji & "'"
    strSQL = strSQL & " AND [Nichiji]  ='" & Nichiji & "'"
    strSQL = strSQL & " AND [RaceNum]  ='" & RaceNum & "'"
    
            mCon_HYOSU.Execute strSQL, , adExecuteNoRecords
    mCon_HYOSU_TANPUKU.Execute strSQL, , adExecuteNoRecords
       mCon_HYOSU_WAKU.Execute strSQL, , adExecuteNoRecords
     mCon_HYOSU_UMAREN.Execute strSQL, , adExecuteNoRecords
       mCon_HYOSU_WIDE.Execute strSQL, , adExecuteNoRecords
     mCon_HYOSU_UMATAN(intMDBIndex).Execute strSQL, , adExecuteNoRecords
     mCon_HYOSU_SANREN(intMDBIndex).Execute strSQL, , adExecuteNoRecords
    
End Sub



'
'   機能: データベースに挿入する
'
'   備考: なし
'
Private Function InsertDB() As Boolean
On Error GoTo ErrorHandler
    Dim i As Integer                                                                    '' ループカウンタ
    Dim j As Integer                                                                    '' ループカウンタ
    Dim k As Integer                                                                    '' ループカウンタ
    Dim rs As ADODB.Recordset                                                           '' SQL文
    Dim intMDBIndex As Integer
    Dim strData As String
    Dim strData2 As String
    
    ' 分割MDBのインデックス番号を場所コードから得る
    If mBuf.id.JyoCD >= "00" And mBuf.id.JyoCD <= "10" Then
        intMDBIndex = val(mBuf.id.JyoCD) - 1
    Else
        intMDBIndex = 0
    End If

    mCon_HYOSU.BeginTrans
    mCon_HYOSU_TANPUKU.BeginTrans
    mCon_HYOSU_WAKU.BeginTrans
    mCon_HYOSU_UMAREN.BeginTrans
    mCon_HYOSU_WIDE.BeginTrans
    mCon_HYOSU_UMATAN(intMDBIndex).BeginTrans
    mCon_HYOSU_SANREN(intMDBIndex).BeginTrans
    
    ' HYOSU (票数)
    Set rs = mRS_HYOSU
    With mBuf.id
        rs.AddNew Array("Year", "MonthDay", "JyoCD", "Kaiji", "Nichiji", "RaceNum"), Array(.Year, .MonthDay, .JyoCD, .Kaiji, .Nichiji, .RaceNum)
    End With
    With mBuf
        With .head
            rs("RecordSpec") = .RecordSpec                                              '' レコード種別
            rs("DataKubun") = .DataKubun                                                '' データ区分
            With .MakeDate
                rs("MakeDate") = .Year & .Month & .Day                                  '' 年月日
            End With ' MakeDate
        End With ' head
        rs("TorokuTosu") = .TorokuTosu                                                  '' 登録頭数
        rs("SyussoTosu") = .SyussoTosu                                                  '' 出走頭数
        For i = 0 To 6
            rs("HatubaiFlag" & i + 1) = .HatubaiFlag(i)                                 '' 発売フラグ
        Next i
        rs("FukuChakuBaraiKey") = .FukuChakuBaraiKey                                    '' 複勝着払キー
        For i = 0 To 27
            rs("HenkanUma" & i + 1) = .HenkanUma(i)                                     '' 返還馬番情報(馬番01～28)
        Next i
        For i = 0 To 7
            rs("HenkanWaku" & i + 1) = .HenkanWaku(i)                                   '' 返還枠番情報(枠番1～8)
        Next i
        For i = 0 To 7
            rs("HenkanDoWaku" & i + 1) = .HenkanDoWaku(i)                               '' 返還同枠情報(枠番1～8)
        Next i
        For i = 0 To 13
            rs("HyoTotal" & i + 1) = .HyoTotal(i)                                       '' 票数合計
        Next i
        
        
    End With ' mBuf
        
    rs.Update ' HYOSU
        

    ' HYOSU_TANPUKU (票数_単複)
    Set rs = mRS_HYOSU_TANPUKU
    strData = ""
    strData2 = ""
    If mBuf.HatubaiFlag(0) <> "0" Then
        With mBuf.id
            rs.AddNew Array("Year", "MonthDay", "JyoCD", "Kaiji", "Nichiji", "RaceNum"), Array(.Year, .MonthDay, .JyoCD, .Kaiji, .Nichiji, .RaceNum)
        End With
        For i = 0 To 27
            If mBuf.HyoTansyo(i).Umaban <> "  " Then
                With mBuf.HyoTansyo(i)
                    strData = strData & .Umaban
                    strData = strData & .Hyo
                    strData = strData & .Ninki
                End With ' mBuf.HyoTansyo
            End If
            If mBuf.HyoFukusyo(i).Umaban <> "  " Then
                With mBuf.HyoFukusyo(i)
                    strData2 = strData2 & .Umaban
                    strData2 = strData2 & .Hyo
                    strData2 = strData2 & .Ninki
                End With ' mBuf.HyoFukusyo
            End If
        Next i
        rs("TanData") = strData
        rs("FukuData") = strData2
        rs.Update ' HYOSU_TANPUKU
    End If
    
    ' HYOSU_WAKU (票数_枠連)
    Set rs = mRS_HYOSU_WAKU
    strData = ""
    If mBuf.HatubaiFlag(2) <> "0" Then
        With mBuf.id
            rs.AddNew Array("Year", "MonthDay", "JyoCD", "Kaiji", "Nichiji", "RaceNum"), Array(.Year, .MonthDay, .JyoCD, .Kaiji, .Nichiji, .RaceNum)
        End With
        For i = 0 To 35
            If mBuf.HyoWakuren(i).Umaban <> "  " Then
                With mBuf
                    With .HyoWakuren(i)
                        strData = strData & .Umaban
                        strData = strData & .Hyo
                        strData = strData & .Ninki
                    End With ' HyoWakuren
                
                End With ' mBuf
            End If
        Next i
        rs("DATA") = strData
        rs.Update ' HYOSU_WAKU
    End If
    
    ' HYOSU_UMAREN (票数_馬連)
    Set rs = mRS_HYOSU_UMAREN
    strData = ""
    If mBuf.HatubaiFlag(3) <> "0" Then
        With mBuf.id
            rs.AddNew Array("Year", "MonthDay", "JyoCD", "Kaiji", "Nichiji", "RaceNum"), Array(.Year, .MonthDay, .JyoCD, .Kaiji, .Nichiji, .RaceNum)
        End With
        For i = 0 To 152
            If mBuf.HyoUmaren(i).Kumi <> "    " Then
                With mBuf
                    With .HyoUmaren(i)
                    End With ' HyoUmaren
                
                End With ' mBuf
            End If
        Next i
        rs("DATA") = strData
        rs.Update ' HYOSU_WAKU
    End If
    
    
    ' HYOSU_WIDE (票数_ワイド)
    Set rs = mRS_HYOSU_WIDE
    strData = ""
    If mBuf.HatubaiFlag(4) <> "0" Then
        With mBuf.id
            rs.AddNew Array("Year", "MonthDay", "JyoCD", "Kaiji", "Nichiji", "RaceNum"), Array(.Year, .MonthDay, .JyoCD, .Kaiji, .Nichiji, .RaceNum)
        End With
        For i = 0 To 152
            If mBuf.HyoWide(i).Kumi <> "    " Then
                With mBuf
                    With .HyoWide(i)
                        strData = strData & .Kumi
                        strData = strData & .Hyo
                        strData = strData & .Ninki
                    End With ' HyoWide
                End With ' mBuf
            End If
        Next i
        rs("DATA") = strData
        rs.Update ' HYOSU_UMARENWIDE
    End If
    
    ' HYOSU_UMATAN (票数_馬単)
    Set rs = mRS_HYOSU_UMATAN(intMDBIndex)
    strData = ""
    If mBuf.HatubaiFlag(5) <> "0" Then
        With mBuf.id
            rs.AddNew Array("Year", "MonthDay", "JyoCD", "Kaiji", "Nichiji", "RaceNum"), Array(.Year, .MonthDay, .JyoCD, .Kaiji, .Nichiji, .RaceNum)
        End With
        For i = 0 To 305
            If mBuf.HyoUmatan(i).Kumi <> "    " Then
                With mBuf.HyoUmatan(i)
                    strData = strData & .Kumi
                    strData = strData & .Hyo
                    strData = strData & .Ninki
                End With ' mBuf.HyoUmatan(i)
            End If
        Next i
        rs("DATA") = strData
        rs.Update ' HYOSU_UMATAN
    End If

    ' HYOSU_SANREN (票数_三連)
    Set rs = mRS_HYOSU_SANREN(intMDBIndex)
    strData = ""
    If mBuf.HatubaiFlag(6) <> "0" Then
        With mBuf.id
            rs.AddNew Array("Year", "MonthDay", "JyoCD", "Kaiji", "Nichiji", "RaceNum"), Array(.Year, .MonthDay, .JyoCD, .Kaiji, .Nichiji, .RaceNum)
        End With
        For i = 0 To 815
            If mBuf.HyoSanrenpuku(i).Kumi <> "      " Then
                With mBuf.HyoSanrenpuku(i)
                        strData = strData & .Kumi
                        strData = strData & .Hyo
                        strData = strData & .Ninki
                End With ' mBuf.HyoSanrenpuku(i)
            End If
        Next i
        rs("DATA") = strData
        rs.Update ' HYOSU_SANREN
    End If
    
    mCon_HYOSU.CommitTrans
    mCon_HYOSU_TANPUKU.CommitTrans
    mCon_HYOSU_WAKU.CommitTrans
    mCon_HYOSU_UMAREN.CommitTrans
    mCon_HYOSU_WIDE.CommitTrans
    mCon_HYOSU_UMATAN(intMDBIndex).CommitTrans
    mCon_HYOSU_SANREN(intMDBIndex).CommitTrans
    
    Set rs = Nothing
    
    InsertDB = True
    Exit Function
    
ErrorHandler:
    If Err.Number <> -2147217887 Then
        gApp.ErrLog
    End If
    rs.CancelUpdate
    
    mCon_HYOSU.RollbackTrans
    mCon_HYOSU_TANPUKU.RollbackTrans
    mCon_HYOSU_WAKU.RollbackTrans
    mCon_HYOSU_UMAREN.RollbackTrans
    mCon_HYOSU_WIDE.RollbackTrans
    mCon_HYOSU_UMATAN(intMDBIndex).RollbackTrans
    mCon_HYOSU_SANREN(intMDBIndex).RollbackTrans

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
    Dim i As Integer                                                                    '' ループカウンタ
    Dim j As Integer                                                                    '' ループカウンタ
    Dim k As Integer                                                                    '' ループカウンタ
    Dim strSQL As String                                                                '' SQL文
    Dim intMDBIndex As Integer
    Dim id As Long
    Dim strData As String
    Dim strData2 As String
    
    ' 分割MDBのインデックス番号を場所コードから得る
    If mBuf.id.JyoCD >= "00" And mBuf.id.JyoCD <= "10" Then
        intMDBIndex = val(mBuf.id.JyoCD) - 1
    Else
        intMDBIndex = 0
    End If
    
    Set rs = mRS_HYOSU

    With mBuf.id
        rs.Seek Array(.Year, .MonthDay, .JyoCD, .Kaiji, .Nichiji, .RaceNum)
    End With

    With mBuf.head.MakeDate
        If rs("Makedate") > .Year & .Month & .Day Then
            UpdateDB = True
            Exit Function
        End If
    End With

    mCon_HYOSU.BeginTrans
    mCon_HYOSU_TANPUKU.BeginTrans
    mCon_HYOSU_WAKU.BeginTrans
    mCon_HYOSU_UMAREN.BeginTrans
    mCon_HYOSU_WIDE.BeginTrans
    mCon_HYOSU_UMATAN(intMDBIndex).BeginTrans
    mCon_HYOSU_SANREN(intMDBIndex).BeginTrans

    With mBuf
        With .head
            rs("RecordSpec") = .RecordSpec                                              '' レコード種別
            rs("DataKubun") = .DataKubun                                                '' データ区分
            With .MakeDate
                rs("MakeDate") = .Year & .Month & .Day                                  '' 年月日
            End With ' MakeDate
        End With ' head
        rs("TorokuTosu") = .TorokuTosu                                                  '' 登録頭数
        rs("SyussoTosu") = .SyussoTosu                                                  '' 出走頭数
        For i = 0 To 6
            rs("HatubaiFlag" & i + 1) = .HatubaiFlag(i)                                 '' 発売フラグ
        Next i
        rs("FukuChakuBaraiKey") = .FukuChakuBaraiKey                                    '' 複勝着払キー
        For i = 0 To 27
            rs("HenkanUma" & i + 1) = .HenkanUma(i)                                     '' 返還馬番情報(馬番01～28)
        Next i
        For i = 0 To 7
            rs("HenkanWaku" & i + 1) = .HenkanWaku(i)                                   '' 返還枠番情報(枠番1～8)
        Next i
        For i = 0 To 7
            rs("HenkanDoWaku" & i + 1) = .HenkanDoWaku(i)                               '' 返還同枠情報(枠番1～8)
        Next i
        For i = 0 To 13
            rs("HyoTotal" & i + 1) = .HyoTotal(i)                                       '' 票数合計
        Next i
        
        
    End With ' mBuf
        
    rs.Update ' HYOSU
        
    ' HYOSU_TANPUKU (票数_単複)
    Set rs = mRS_HYOSU_TANPUKU
    
    '削除
    With mBuf.id
        strSQL = "DELETE FROM HYOSU_TANPUKU"
        strSQL = strSQL & " WHERE [Year]='" & .Year & "'"
        strSQL = strSQL & " AND [MonthDay]='" & .MonthDay & "'"
        strSQL = strSQL & " AND [JyoCD]='" & .JyoCD & "'"
        strSQL = strSQL & " AND [Kaiji]='" & .Kaiji & "'"
        strSQL = strSQL & " AND [Nichiji]='" & .Nichiji & "'"
        strSQL = strSQL & " AND [RaceNum]='" & .RaceNum & "'"
    End With
    mCon_HYOSU_TANPUKU.Execute strSQL, , adExecuteNoRecords

    strData = ""
    strData2 = ""
    If mBuf.HatubaiFlag(0) <> "0" Then
        With mBuf.id
            rs.AddNew Array("Year", "MonthDay", "JyoCD", "Kaiji", "Nichiji", "RaceNum"), Array(.Year, .MonthDay, .JyoCD, .Kaiji, .Nichiji, .RaceNum)
        End With
        For i = 0 To 27
            If mBuf.HyoTansyo(i).Umaban <> "  " Then
                With mBuf.HyoTansyo(i)
                    strData = strData & .Umaban
                    strData = strData & .Hyo
                    strData = strData & .Ninki
                End With ' mBuf.HyoTansyo
            End If
            If mBuf.HyoFukusyo(i).Umaban <> "  " Then
                With mBuf.HyoFukusyo(i)
                    strData2 = strData2 & .Umaban
                    strData2 = strData2 & .Hyo
                    strData2 = strData2 & .Ninki
                End With ' mBuf.HyoFukusyo
            End If
        Next i
        rs("TanData") = strData
        rs("FukuData") = strData2
        rs.Update ' HYOSU_TANPUKU
    End If
    
    
    ' HYOSU_WAKU (票数_枠連)
    Set rs = mRS_HYOSU_WAKU
    
    '削除
    With mBuf.id
        strSQL = "DELETE FROM HYOSU_WAKU"
        strSQL = strSQL & " WHERE [Year]='" & .Year & "'"
        strSQL = strSQL & " AND [MonthDay]='" & .MonthDay & "'"
        strSQL = strSQL & " AND [JyoCD]='" & .JyoCD & "'"
        strSQL = strSQL & " AND [Kaiji]='" & .Kaiji & "'"
        strSQL = strSQL & " AND [Nichiji]='" & .Nichiji & "'"
        strSQL = strSQL & " AND [RaceNum]='" & .RaceNum & "'"
    End With
    mCon_HYOSU_WAKU.Execute strSQL, , adExecuteNoRecords
    
    strData = ""
    If mBuf.HatubaiFlag(2) <> "0" Then
        With mBuf.id
            rs.AddNew Array("Year", "MonthDay", "JyoCD", "Kaiji", "Nichiji", "RaceNum"), Array(.Year, .MonthDay, .JyoCD, .Kaiji, .Nichiji, .RaceNum)
        End With
        For i = 0 To 35
            If mBuf.HyoWakuren(i).Umaban <> "  " Then
                With mBuf
                    With .HyoWakuren(i)
                        strData = strData & .Umaban
                        strData = strData & .Hyo
                        strData = strData & .Ninki
                    End With ' HyoWakuren
                
                End With ' mBuf
            End If
        Next i
        rs("DATA") = strData
        rs.Update ' HYOSU_WAKU
    End If
    
    ' HYOSU_UMAREN (票数_馬連)
    Set rs = mRS_HYOSU_UMAREN
    
    '削除
    With mBuf.id
        strSQL = "DELETE FROM HYOSU_UMAREN"
        strSQL = strSQL & " WHERE [Year]='" & .Year & "'"
        strSQL = strSQL & " AND [MonthDay]='" & .MonthDay & "'"
        strSQL = strSQL & " AND [JyoCD]='" & .JyoCD & "'"
        strSQL = strSQL & " AND [Kaiji]='" & .Kaiji & "'"
        strSQL = strSQL & " AND [Nichiji]='" & .Nichiji & "'"
        strSQL = strSQL & " AND [RaceNum]='" & .RaceNum & "'"
    End With
    mCon_HYOSU_UMAREN.Execute strSQL, , adExecuteNoRecords
     
    strData = ""
    If mBuf.HatubaiFlag(3) <> "0" Then
        With mBuf.id
            rs.AddNew Array("Year", "MonthDay", "JyoCD", "Kaiji", "Nichiji", "RaceNum"), Array(.Year, .MonthDay, .JyoCD, .Kaiji, .Nichiji, .RaceNum)
        End With
        For i = 0 To 152
            If mBuf.HyoUmaren(i).Kumi <> "    " Then
                With mBuf
                    With .HyoUmaren(i)
                    End With ' HyoUmaren
                
                End With ' mBuf
            End If
        Next i
        rs("DATA") = strData
        rs.Update ' HYOSU_WAKU
    End If
    
    
    
    ' HYOSU_WIDE (票数_ワイド)
    Set rs = mRS_HYOSU_WIDE
    
    '削除
    With mBuf.id
        strSQL = "DELETE FROM HYOSU_WIDE"
        strSQL = strSQL & " WHERE [Year]='" & .Year & "'"
        strSQL = strSQL & " AND [MonthDay]='" & .MonthDay & "'"
        strSQL = strSQL & " AND [JyoCD]='" & .JyoCD & "'"
        strSQL = strSQL & " AND [Kaiji]='" & .Kaiji & "'"
        strSQL = strSQL & " AND [Nichiji]='" & .Nichiji & "'"
        strSQL = strSQL & " AND [RaceNum]='" & .RaceNum & "'"
    End With
    mCon_HYOSU_WIDE.Execute strSQL, , adExecuteNoRecords
     
    strData = ""
    If mBuf.HatubaiFlag(4) <> "0" Then
        With mBuf.id
            rs.AddNew Array("Year", "MonthDay", "JyoCD", "Kaiji", "Nichiji", "RaceNum"), Array(.Year, .MonthDay, .JyoCD, .Kaiji, .Nichiji, .RaceNum)
        End With
        For i = 0 To 152
            If mBuf.HyoWide(i).Kumi <> "    " Then
                With mBuf
                    With .HyoWide(i)
                        strData = strData & .Kumi
                        strData = strData & .Hyo
                        strData = strData & .Ninki
                    End With ' HyoWide
                End With ' mBuf
            End If
        Next i
        rs("DATA") = strData
        rs.Update ' HYOSU_UMARENWIDE
    End If
    
    ' HYOSU_UMATAN (票数_馬単)
    Set rs = mRS_HYOSU_UMATAN(intMDBIndex)
    
    '削除
    With mBuf.id
        strSQL = "DELETE FROM HYOSU_UMATAN"
        strSQL = strSQL & " WHERE [Year]='" & .Year & "'"
        strSQL = strSQL & " AND [MonthDay]='" & .MonthDay & "'"
        strSQL = strSQL & " AND [JyoCD]='" & .JyoCD & "'"
        strSQL = strSQL & " AND [Kaiji]='" & .Kaiji & "'"
        strSQL = strSQL & " AND [Nichiji]='" & .Nichiji & "'"
        strSQL = strSQL & " AND [RaceNum]='" & .RaceNum & "'"
    End With
    mCon_HYOSU_UMATAN(intMDBIndex).Execute strSQL, , adExecuteNoRecords
     
    strData = ""
    If mBuf.HatubaiFlag(5) <> "0" Then
        With mBuf.id
            rs.AddNew Array("Year", "MonthDay", "JyoCD", "Kaiji", "Nichiji", "RaceNum"), Array(.Year, .MonthDay, .JyoCD, .Kaiji, .Nichiji, .RaceNum)
        End With
        For i = 0 To 305
            If mBuf.HyoUmatan(i).Kumi <> "    " Then
                With mBuf.HyoUmatan(i)
                    strData = strData & .Kumi
                    strData = strData & .Hyo
                    strData = strData & .Ninki
                End With ' mBuf.HyoUmatan(i)
            End If
        Next i
        rs("DATA") = strData
        rs.Update ' HYOSU_UMATAN
    End If
    
    ' HYOSU_SANREN (票数_三連)
    
    Set rs = mRS_HYOSU_SANREN(intMDBIndex)
    
    '削除
    With mBuf.id
        strSQL = "DELETE FROM HYOSU_SANREN"
        strSQL = strSQL & " WHERE [Year]='" & .Year & "'"
        strSQL = strSQL & " AND [MonthDay]='" & .MonthDay & "'"
        strSQL = strSQL & " AND [JyoCD]='" & .JyoCD & "'"
        strSQL = strSQL & " AND [Kaiji]='" & .Kaiji & "'"
        strSQL = strSQL & " AND [Nichiji]='" & .Nichiji & "'"
        strSQL = strSQL & " AND [RaceNum]='" & .RaceNum & "'"
    End With
    mCon_HYOSU_SANREN(intMDBIndex).Execute strSQL, , adExecuteNoRecords
     
    strData = ""
    If mBuf.HatubaiFlag(6) <> "0" Then
        With mBuf.id
            rs.AddNew Array("Year", "MonthDay", "JyoCD", "Kaiji", "Nichiji", "RaceNum"), Array(.Year, .MonthDay, .JyoCD, .Kaiji, .Nichiji, .RaceNum)
        End With
        For i = 0 To 815
            If mBuf.HyoSanrenpuku(i).Kumi <> "      " Then
                With mBuf.HyoSanrenpuku(i)
                        strData = strData & .Kumi
                        strData = strData & .Hyo
                        strData = strData & .Ninki
                End With ' mBuf.HyoSanrenpuku(i)
            End If
        Next i
        rs("DATA") = strData
        rs.Update ' HYOSU_SANREN
    End If
    
    mCon_HYOSU.CommitTrans
    mCon_HYOSU_TANPUKU.CommitTrans
    mCon_HYOSU_WAKU.CommitTrans
    mCon_HYOSU_UMAREN.CommitTrans
    mCon_HYOSU_WIDE.CommitTrans
    mCon_HYOSU_UMATAN(intMDBIndex).CommitTrans
    mCon_HYOSU_SANREN(intMDBIndex).CommitTrans
    
    Set rs = Nothing
    
    UpdateDB = True
    Exit Function
    
ErrorHandler:
    gApp.ErrLog
    
    
    mCon_HYOSU.RollbackTrans
    mCon_HYOSU_TANPUKU.RollbackTrans
    mCon_HYOSU_WAKU.RollbackTrans
    mCon_HYOSU_UMAREN.RollbackTrans
    mCon_HYOSU_WIDE.RollbackTrans
    mCon_HYOSU_UMATAN(intMDBIndex).RollbackTrans
    mCon_HYOSU_SANREN(intMDBIndex).RollbackTrans
    
    UpdateDB = False
End Function



