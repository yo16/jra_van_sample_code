VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsImportTC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'   JVData "TC" データベース登録クラス
'

Option Explicit
Option Compare Binary
Implements clsIImport

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部変数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private mBuf As JV_TC_HASSOU

Private mCon_HASSOU_CHANGE As ADODB.Connection

Private mRS_HASSOU_CHANGE As ADODB.Recordset


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

    ' コネクションのインスタンス生成
    Set mCon_HASSOU_CHANGE = New ADODB.Connection

    ' レコードセットのインスタンス生成
    Set mRS_HASSOU_CHANGE = New ADODB.Recordset

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
On Error GoTo ErrorHandler
    Dim strCon As String

    ' コネクションオープン
    strCon = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & gApp.R_DBPath & "\"
    mCon_HASSOU_CHANGE.Open strCon & "subHASSOU_CHANGE.mdb"
    ' レコードセットオープン
    With mRS_HASSOU_CHANGE
        .CursorLocation = adUseServer
        .Index = "PrimaryKey"
        .Open "HASSOU_CHANGE", mCon_HASSOU_CHANGE, adOpenKeyset, adLockOptimistic, adCmdTableDirect
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
    mRS_HASSOU_CHANGE.Close

    mCon_HASSOU_CHANGE.Close

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
        Call SetDataFromByte_TC(lBuf, mBuf)     ''構造体に代入する

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

    mCon_HASSOU_CHANGE.BeginTrans

    Year = Mid$(lBuf, 12, 4)            '' 開催年
    MonthDay = Mid$(lBuf, 16, 4)        '' 開催月日
    JyoCD = Mid$(lBuf, 20, 2)           '' 競馬場コード
    Kaiji = Mid$(lBuf, 22, 2)           '' 開催回[第N回]
    Nichiji = Mid$(lBuf, 24, 2)         '' 開催日目[N日目]
    RaceNum = Mid$(lBuf, 26, 2)         '' レース番号
    
    strSQL = "DELETE * FROM HASSOU_CHANGE"
    strSQL = strSQL & " WHERE [Year]    ='" & Year & "'"
    strSQL = strSQL & " AND [MonthDay]  ='" & MonthDay & "'"
    strSQL = strSQL & " AND [JyoCD]     ='" & JyoCD & "'"
    strSQL = strSQL & " AND [Kaiji]     ='" & Kaiji & "'"
    strSQL = strSQL & " AND [Nichiji]   ='" & Nichiji & "'"
    strSQL = strSQL & " AND [RaceNum]   ='" & RaceNum & "';"

    mCon_HASSOU_CHANGE.Execute strSQL, , adExecuteNoRecords

    mCon_HASSOU_CHANGE.CommitTrans

    Exit Sub
ErrorHandler:
    mCon_HASSOU_CHANGE.RollbackTrans
End Sub


'
'   機能: データベースに挿入する
'
'   備考: なし
'
Private Function InsertDB() As Boolean
On Error GoTo ErrorHandler
    Dim rs As ADODB.Recordset

    mCon_HASSOU_CHANGE.BeginTrans

    Set rs = mRS_HASSOU_CHANGE

    With mBuf.id
        rs.AddNew Array("Year", "MonthDay", "JyoCD", "Kaiji", "Nichiji", "RaceNum"), Array(.Year, .MonthDay, .JyoCD, .Kaiji, .Nichiji, .RaceNum)
    End With

    With mBuf
        With .head
            rs("RecordSpec") = .RecordSpec
            rs("DataKubun") = .DataKubun
            With .MakeDate
                rs("MakeDate") = .Year & .Month & .Day
            End With
        End With
        With .HappyoTime
            rs("HappyoTime") = .Month & .Day & .Hour & .Minute
        End With
        With .AtoHassoTime
            rs("AtoHassoTime") = .Hour & .Minute
        End With
        With .MaeHassoTime
            rs("MaeHassoTime") = .Hour & .Minute
        End With
    End With

    rs.Update

    mCon_HASSOU_CHANGE.CommitTrans

    InsertDB = True
    Exit Function

ErrorHandler:
    rs.CancelUpdate

    mCon_HASSOU_CHANGE.RollbackTrans
End Function


'
'   機能: データベースを更新する
'
'   備考: なし
'
Private Function UpdateDB() As Boolean
On Error GoTo ErrorHandler
    Dim rs As ADODB.Recordset
    Dim strSQL As String

    Set rs = mRS_HASSOU_CHANGE

    With mBuf.id
        rs.Seek Array(.Year, .MonthDay, .JyoCD, .Kaiji, .Nichiji, .RaceNum)
    End With

    With mBuf.head.MakeDate
        If rs("MakeDate") > .Year & .Month & .Day Then
            UpdateDB = True
            Exit Function
        End If
    End With

    mCon_HASSOU_CHANGE.BeginTrans

    With mBuf
        With .head
            rs("RecordSpec") = .RecordSpec
            rs("DataKubun") = .DataKubun
            With .MakeDate
                rs("MakeDate") = .Year & .Month & .Day
            End With
        End With
        With .HappyoTime
            rs("HappyoTime") = .Month & .Day & .Hour & .Minute
        End With
        With .AtoHassoTime
            rs("AtoHassoTime") = .Hour & .Minute
        End With
        With .MaeHassoTime
            rs("MaeHassoTime") = .Hour & .Minute
        End With
    End With

    rs.Update

    mCon_HASSOU_CHANGE.CommitTrans

    UpdateDB = True
    Exit Function

ErrorHandler:
    gApp.ErrLog

    mCon_HASSOU_CHANGE.RollbackTrans

    UpdateDB = False
End Function


