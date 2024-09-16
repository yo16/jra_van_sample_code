VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsImportHC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Import JV_HC_HANRO "
'
'   JVData "HC" データベース登録クラス
'

Option Explicit
Option Compare Binary
Implements clsIImport

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部変数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private mBuf As JV_HC_HANRO

Private mCon_HANRO As ADODB.Connection

Private mRS_HANRO As ADODB.Recordset


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
    Set mCon_HANRO = New ADODB.Connection

    ' レコードセットのインスタンス生成
    Set mRS_HANRO = New ADODB.Recordset
    
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
    mCon_HANRO.Open strCon & "subHANRO.mdb"

    ' レコードセットオープン
    With mRS_HANRO
        .CursorLocation = adUseServer
        .Index = "PrimaryKey"
        .Open "HANRO", mCon_HANRO, adOpenKeyset, adLockOptimistic, adCmdTableDirect
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
    mRS_HANRO.Close
'
    mCon_HANRO.Close

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

        ' データ登録
        If Not InsertDB(StrConv(lBuf, vbUnicode)) Then
            ' 新規挿入に失敗したら更新を試みる
            clsIImport_Add = UpdateDB(StrConv(lBuf, vbUnicode))
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
Private Sub DeleteRecord(ByRef lBuf As String)
On Error GoTo ErrorHandler
    Dim strSQL      As String
    Dim TresenKubun As String
    Dim ChokyoDate  As String
    Dim ChokyoTime  As String
    Dim KettoNum    As String
    
    mCon_HANRO.BeginTrans
    
    TresenKubun = Mid$(lBuf, 12, 1)                    ''
    ChokyoDate = Mid$(lBuf, 13, 8)                ''
    ChokyoTime = Mid$(lBuf, 21, 4)                    ''
    KettoNum = Mid$(lBuf, 25, 10)                   ''

    strSQL = "DELETE * FROM HANRO"
    strSQL = strSQL & " WHERE [TresenKubun] ='" & TresenKubun & "'"
    strSQL = strSQL & " AND [ChokyoDate]    ='" & ChokyoDate & "'"
    strSQL = strSQL & " AND [ChokyoTime]    ='" & ChokyoTime & "'"
    strSQL = strSQL & " AND [KettoNum]      ='" & KettoNum & "'"
    
    mCon_HANRO.Execute strSQL, , adExecuteNoRecords
    
    mCon_HANRO.CommitTrans
        
    Exit Sub
ErrorHandler:
    mCon_HANRO.RollbackTrans
End Sub


'
'   機能: データベースに挿入する
'
'   備考: なし
'
Private Function InsertDB(ByRef lBuf As String) As Boolean
On Error GoTo ErrorHandler
    Dim i As Integer                                                               '' ループカウンタ
    Dim j As Integer                                                               '' ループカウンタ
    Dim k As Integer                                                               '' ループカウンタ
    Dim rs As ADODB.Recordset                                                      '' SQL文
    
    mCon_HANRO.BeginTrans

    Set rs = mRS_HANRO
    
    rs.AddNew Array("TresenKubun", "ChokyoDate", "ChokyoTime", "KettoNum"), Array(Mid$(lBuf, 12, 1), Mid$(lBuf, 13, 8), Mid$(lBuf, 21, 4), Mid$(lBuf, 25, 10))
    
    rs("RecordSpec") = Mid$(lBuf, 1, 2)                                      '' レコード種別
    rs("DataKubun") = Mid$(lBuf, 3, 1)                                       '' データ区分
    rs("MakeDate") = Mid$(lBuf, 4, 8)                                        '' 年月日
    rs("HaronTime4") = Mid$(lBuf, 35, 4)                                     '' 4ハロンタイム合計(800M-0M)
    rs("LapTime4") = Mid$(lBuf, 39, 3)                                       '' ラップタイム(800M-600M)
    rs("HaronTime3") = Mid$(lBuf, 42, 4)                                     '' 3ハロンタイム合計(600M-0M)
    rs("LapTime3") = Mid$(lBuf, 46, 3)                                       '' ラップタイム(600M-400M)
    rs("HaronTime2") = Mid$(lBuf, 49, 4)                                     '' 2ハロンタイム合計(400M-0M)
    rs("LapTime2") = Mid$(lBuf, 53, 3)                                       '' ラップタイム(400M-200M)
    rs("LapTime1") = Mid$(lBuf, 56, 3)                                       '' ラップタイム(200M-0M)
            
    rs.Update ' HANRO
    
    mCon_HANRO.CommitTrans

    Set rs = Nothing
    
    InsertDB = True
    Exit Function

ErrorHandler:
    If Err.Number <> -2147217887 Then
        gApp.ErrLog
    End If
    rs.CancelUpdate
    
    mCon_HANRO.RollbackTrans

    InsertDB = False
End Function


'
'   機能: データベースを更新する
'
'   備考: なし
'
Private Function UpdateDB(ByRef lBuf As String) As Boolean
On Error GoTo ErrorHandler
    Dim rs As ADODB.Recordset
    Dim i As Integer                                                               '' ループカウンタ
    Dim j As Integer                                                               '' ループカウンタ
    Dim k As Integer                                                               '' ループカウンタ
    Dim strSQL As String                                                           '' SQL文
    
    Set rs = mRS_HANRO

    rs.Seek Array(Mid$(lBuf, 12, 1), Mid$(lBuf, 13, 8), Mid$(lBuf, 21, 4), Mid$(lBuf, 25, 10))

    If rs("Makedate") > Mid$(lBuf, 4, 8) Then
        UpdateDB = True
        Exit Function
    End If

    rs("RecordSpec") = Mid$(lBuf, 1, 2)                                      '' レコード種別
    rs("DataKubun") = Mid$(lBuf, 3, 1)                                       '' データ区分
    rs("MakeDate") = Mid$(lBuf, 4, 8)                                        '' 年月日
    rs("HaronTime4") = Mid$(lBuf, 35, 4)                                     '' 4ハロンタイム合計(800M-0M)
    rs("LapTime4") = Mid$(lBuf, 39, 3)                                       '' ラップタイム(800M-600M)
    rs("HaronTime3") = Mid$(lBuf, 42, 4)                                     '' 3ハロンタイム合計(600M-0M)
    rs("LapTime3") = Mid$(lBuf, 46, 3)                                       '' ラップタイム(600M-400M)
    rs("HaronTime2") = Mid$(lBuf, 49, 4)                                     '' 2ハロンタイム合計(400M-0M)
    rs("LapTime2") = Mid$(lBuf, 53, 3)                                       '' ラップタイム(400M-200M)
    rs("LapTime1") = Mid$(lBuf, 56, 3)                                       '' ラップタイム(200M-0M)
    
    rs.Update ' HANRO
    
    Set rs = Nothing
    
    UpdateDB = True
    Exit Function

ErrorHandler:
    gApp.ErrLog
    
    UpdateDB = False
End Function
