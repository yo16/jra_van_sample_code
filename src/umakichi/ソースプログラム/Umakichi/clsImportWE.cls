VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsImportWE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Import JV_WE_WEATHER "
'
'   JVData "WE" データベース登録クラス
'

Option Explicit
Option Compare Binary
Implements clsIImport

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部変数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private mBuf As JV_WE_WEATHER

Private mCon_TENKO_BABA As ADODB.Connection

Private mRS_TENKO_BABA As ADODB.Recordset


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
    Set mCon_TENKO_BABA = New ADODB.Connection

    ' レコードセットのインスタンス生成
    Set mRS_TENKO_BABA = New ADODB.Recordset
    
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
    mCon_TENKO_BABA.Open strCon & "subTENKO_BABA.mdb"

    ' レコードセットオープン
    With mRS_TENKO_BABA
        .CursorLocation = adUseServer
        .Index = "PrimaryKey"
        .Open "TENKO_BABA", mCon_TENKO_BABA, adOpenKeyset, adLockOptimistic, adCmdTableDirect
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
    mRS_TENKO_BABA.Close

    mCon_TENKO_BABA.Close

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
        Call SetDataFromByte_WE(lBuf, mBuf)     '' 構造体に代入する

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
    Dim HappyoTime  As String
    Dim HenkoID     As String
    
    mCon_TENKO_BABA.BeginTrans
    
    Year = Mid$(lBuf, 12, 4)                    '' 開催年
    MonthDay = Mid$(lBuf, 16, 4)                '' 開催月日
    JyoCD = Mid$(lBuf, 20, 2)                   '' 競馬場コード
    Kaiji = Mid$(lBuf, 22, 2)                   '' 開催回[第N回]
    Nichiji = Mid$(lBuf, 24, 2)                 '' 開催日目[N日目]
    HappyoTime = Mid$(lBuf, 26, 8)                   ''
    HenkoID = Mid$(lBuf, 34, 1)                   ''

    strSQL = "DELETE * FROM TENKO_BABA"
    strSQL = strSQL & " WHERE [Year]    ='" & Year & "'"
    strSQL = strSQL & " AND [MonthDay]  ='" & MonthDay & "'"
    strSQL = strSQL & " AND [JyoCD]     ='" & JyoCD & "'"
    strSQL = strSQL & " AND [Kaiji]     ='" & Kaiji & "'"
    strSQL = strSQL & " AND [Nichiji]   ='" & Nichiji & "'"
    strSQL = strSQL & " AND [HappyoTime]='" & HappyoTime & "'"
    strSQL = strSQL & " AND [HenkoID]   ='" & HenkoID & "'"
    
    mCon_TENKO_BABA.Execute strSQL, , adExecuteNoRecords
    
    mCon_TENKO_BABA.CommitTrans
        
    Exit Sub
ErrorHandler:
    mCon_TENKO_BABA.RollbackTrans
End Sub


'
'   機能: データベースに挿入する
'
'   備考: なし
'
Private Function InsertDB() As Boolean
On Error GoTo ErrorHandler
    Dim i As Integer                                                            '' ループカウンタ
    Dim j As Integer                                                            '' ループカウンタ
    Dim k As Integer                                                            '' ループカウンタ
    Dim rs As ADODB.Recordset                                                   '' SQL文
    
    
    mCon_TENKO_BABA.BeginTrans

    Set rs = mRS_TENKO_BABA
    
    With mBuf.HappyoTime
        rs.AddNew Array("Year", "MonthDay", "JyoCD", "Kaiji", "Nichiji", "HappyoTime", "HenkoID"), Array(mBuf.id.Year, mBuf.id.MonthDay, mBuf.id.JyoCD, mBuf.id.Kaiji, mBuf.id.Nichiji, .Month & .Day & .Hour & .Minute, mBuf.HenkoID)
    End With
    
    With mBuf
        With .head
            rs("RecordSpec") = .RecordSpec                                      '' レコード種別
            rs("DataKubun") = .DataKubun                                        '' データ区分
            With .MakeDate
                rs("MakeDate") = .Year & .Month & .Day                              '' 年月日
            End With ' MakeDate
        End With ' head
        With .TenkoBaba
            rs("AtoTenkoCD") = .TenkoCD                                            '' 天候コード
            rs("AtoSibaBabaCD") = .SibaBabaCD                                      '' 芝馬場状態コード
            rs("AtoDirtBabaCD") = .DirtBabaCD                                      '' ダート馬場状態コード
        End With ' TenkoBaba
        With .TenkoBabaBefore
            rs("MaeTenkoCD") = .TenkoCD                                            '' 天候コード
            rs("MaeSibaBabaCD") = .SibaBabaCD                                      '' 芝馬場状態コード
            rs("MaeDirtBabaCD") = .DirtBabaCD                                      '' ダート馬場状態コード
        End With ' TenkoBabaBefore
    End With
    rs.Update ' TENKO_BABA
    
    mCon_TENKO_BABA.CommitTrans

    
    InsertDB = True
    Exit Function

ErrorHandler:
    rs.CancelUpdate
    
    mCon_TENKO_BABA.RollbackTrans

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
    Dim i As Integer                                                            '' ループカウンタ
    Dim j As Integer                                                            '' ループカウンタ
    Dim k As Integer                                                            '' ループカウンタ
    Dim strSQL As String                                                        '' SQL文
    
    Set rs = mRS_TENKO_BABA

    With mBuf.HappyoTime
        rs.Seek Array(mBuf.id.Year, mBuf.id.MonthDay, mBuf.id.JyoCD, mBuf.id.Kaiji, mBuf.id.Nichiji, .Month & .Day & .Hour & .Minute, mBuf.HenkoID)
    End With

    With mBuf.head.MakeDate
        If rs("Makedate") > .Year & .Month & .Day Then
            UpdateDB = True
            Exit Function
        End If
    End With

    
    mCon_TENKO_BABA.BeginTrans

    With mBuf
        With .head
            rs("RecordSpec") = .RecordSpec                                      '' レコード種別
            rs("DataKubun") = .DataKubun                                        '' データ区分
            With .MakeDate
                rs("MakeDate") = .Year & .Month & .Day                              '' 年月日
            End With ' MakeDate
        End With ' head
        With .id
            rs("Year") = .Year                                                  '' 開催年
            rs("MonthDay") = .MonthDay                                          '' 開催月日
            rs("JyoCD") = .JyoCD                                                '' 競馬場コード
            rs("Kaiji") = .Kaiji                                                '' 開催回[第N回]
            rs("Nichiji") = .Nichiji                                            '' 開催日目[N日目]
        End With ' id
        With .HappyoTime
            rs("HappyoTime") = .Month & .Day & .Hour & .Minute
        End With ' HappyoTime
        rs("HenkoID") = .HenkoID                                                '' 変更識別
        With .TenkoBaba
            rs("AtoTenkoCD") = .TenkoCD                                            '' 天候コード
            rs("AtoSibaBabaCD") = .SibaBabaCD                                      '' 芝馬場状態コード
            rs("AtoDirtBabaCD") = .DirtBabaCD                                      '' ダート馬場状態コード
        End With ' TenkoBaba
        With .TenkoBabaBefore
            rs("MaeTenkoCD") = .TenkoCD                                            '' 天候コード
            rs("MaeSibaBabaCD") = .SibaBabaCD                                      '' 芝馬場状態コード
            rs("MaeDirtBabaCD") = .DirtBabaCD                                      '' ダート馬場状態コード
        End With ' TenkoBabaBefore
    End With
    rs.Update
    
    mCon_TENKO_BABA.CommitTrans

    
    Set rs = Nothing
    
    UpdateDB = True
    Exit Function

ErrorHandler:
    gApp.ErrLog
    
    mCon_TENKO_BABA.RollbackTrans

    UpdateDB = False
End Function
