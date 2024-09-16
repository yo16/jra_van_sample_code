VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsImportDM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Import JV_DM_INFO "
'
'   JVData "DM" データベース登録クラス
'

Option Explicit
Option Compare Binary
Implements clsIImport

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部変数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private mBuf As JV_DM_INFO

Private mCon_MINING As ADODB.Connection

Private mRS_MINING As ADODB.Recordset


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
    Set mCon_MINING = New ADODB.Connection

    ' レコードセットのインスタンス生成
    Set mRS_MINING = New ADODB.Recordset
    
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
    mCon_MINING.Open strCon & "subMINING.mdb"

    ' レコードセットオープン
    With mRS_MINING
        .CursorLocation = adUseServer
        .Index = "PrimaryKey"
        .Open "MINING", mCon_MINING, adOpenKeyset, adLockOptimistic, adCmdTableDirect
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
    mRS_MINING.Close

    mCon_MINING.Close

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
        Call SetDataFromByte_DM(lBuf, mBuf)     '' 構造体に代入する

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
    
    mCon_MINING.BeginTrans
    
    Year = Mid$(lBuf, 12, 4)                    '' 開催年
    MonthDay = Mid$(lBuf, 16, 4)                '' 開催月日
    JyoCD = Mid$(lBuf, 20, 2)                   '' 競馬場コード
    Kaiji = Mid$(lBuf, 22, 2)                   '' 開催回[第N回]
    Nichiji = Mid$(lBuf, 24, 2)                 '' 開催日目[N日目]
    RaceNum = Mid$(lBuf, 26, 2)                 '' レース番号

    strSQL = "DELETE * FROM MINING"
    strSQL = strSQL & " WHERE [Year]   ='" & Year & "'"
    strSQL = strSQL & " AND [MonthDay] ='" & MonthDay & "'"
    strSQL = strSQL & " AND [JyoCD]    ='" & JyoCD & "'"
    strSQL = strSQL & " AND [Kaiji]    ='" & Kaiji & "'"
    strSQL = strSQL & " AND [Nichiji]  ='" & Nichiji & "'"
    strSQL = strSQL & " AND [RaceNum]  ='" & RaceNum & "'"
    
    mCon_MINING.Execute strSQL, , adExecuteNoRecords
    
    mCon_MINING.CommitTrans
        
    Exit Sub
ErrorHandler:
    mCon_MINING.RollbackTrans
End Sub


'
'   機能: データベースに挿入する
'
'   備考: なし
'
Private Function InsertDB() As Boolean
On Error GoTo ErrorHandler
    Dim i As Integer                                                               '' ループカウンタ
    Dim j As Integer                                                               '' ループカウンタ
    Dim k As Integer                                                               '' ループカウンタ
    Dim rs As ADODB.Recordset                                                      '' SQL文
    
    mCon_MINING.BeginTrans

    
    Set rs = mRS_MINING
    With mBuf.id
        rs.AddNew Array("Year", "MonthDay", "JyoCD", "Kaiji", "Nichiji", "RaceNum"), Array(.Year, .MonthDay, .JyoCD, .Kaiji, .Nichiji, .RaceNum)
    End With
    With mBuf
        With .head
            rs("RecordSpec") = .RecordSpec                                         '' レコード種別
            rs("DataKubun") = .DataKubun                                           '' データ区分
            With .MakeDate
                rs("MakeDate") = .Year & .Month & .Day                             '' 年月日
            End With ' MakeDate
        End With ' head
        With .MakeHM
            rs("MakeHM") = .Hour & .Minute                                         '' 時分
        End With ' MakeHM
        For i = 0 To 17
            With .DMInfo(i)
                rs("Umaban" & i + 1) = .Umaban                                           '' 馬番
                rs("DMTime" & i + 1) = .DMTime                                           '' 予想走破タイム
                rs("DMGosaP" & i + 1) = .DMGosaP                                         '' 予想誤差(信頼度)＋
                rs("DMGosaM" & i + 1) = .DMGosaM                                         '' 予想誤差(信頼度)－
            End With ' DMInfo
        Next i
        
        
    End With ' mBuf
    
    rs.Update ' MINING
    
    mCon_MINING.CommitTrans

    
    InsertDB = True
    Exit Function

ErrorHandler:
    rs.CancelUpdate
    
    mCon_MINING.RollbackTrans

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
    Dim i As Integer                                                               '' ループカウンタ
    Dim j As Integer                                                               '' ループカウンタ
    Dim k As Integer                                                               '' ループカウンタ
    Dim strSQL As String                                                           '' SQL文
    
    Set rs = mRS_MINING

    With mBuf.id
        rs.Seek Array(.Year, .MonthDay, .JyoCD, .Kaiji, .Nichiji, .RaceNum)
    End With

    With mBuf.head.MakeDate
        If rs("Makedate") > .Year & .Month & .Day Then
            UpdateDB = True
            Exit Function
        End If
    End With

    
    mCon_MINING.BeginTrans

    With mBuf
        With .head
            rs("RecordSpec") = .RecordSpec                                         '' レコード種別
            rs("DataKubun") = .DataKubun                                           '' データ区分
            With .MakeDate
                rs("MakeDate") = .Year & .Month & .Day                             '' 年月日
            End With ' MakeDate
        End With ' head
        With .MakeHM
            rs("MakeHM") = .Hour & .Minute                                         '' 時分
        End With ' MakeHM
        For i = 0 To 17
            With .DMInfo(i)
                rs("Umaban" & i + 1) = .Umaban                                           '' 馬番
                rs("DMTime" & i + 1) = .DMTime                                           '' 予想走破タイム
                rs("DMGosaP" & i + 1) = .DMGosaP                                         '' 予想誤差(信頼度)＋
                rs("DMGosaM" & i + 1) = .DMGosaM                                         '' 予想誤差(信頼度)－
            End With ' DMInfo
        Next i
    End With ' mBuf
    
    rs.Update
    
    mCon_MINING.CommitTrans

    
    UpdateDB = True
    Exit Function

ErrorHandler:
    gApp.ErrLog
    
    mCon_MINING.RollbackTrans

    UpdateDB = False
End Function
