VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsImportHN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Import JV_HN_HANSYOKU "
'
'   JVData "HN" データベース登録クラス
'

Option Explicit
Option Compare Binary
Implements clsIImport

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部変数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private mBuf As JV_HN_HANSYOKU

Private mCon_HANSYOKU As ADODB.Connection

Private mRS_HANSYOKU As ADODB.Recordset


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
    Set mCon_HANSYOKU = New ADODB.Connection

    ' レコードセットのインスタンス生成
    Set mRS_HANSYOKU = New ADODB.Recordset
    
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
    mCon_HANSYOKU.Open strCon & "subHANSYOKU.mdb"

    ' レコードセットオープン
    With mRS_HANSYOKU
        .CursorLocation = adUseServer
        .Index = "PrimaryKey"
        .Open "HANSYOKU", mCon_HANSYOKU, adOpenKeyset, adLockOptimistic, adCmdTableDirect
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
    mRS_HANSYOKU.Close

    mCon_HANSYOKU.Close

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
        Call SetDataFromByte_HN(lBuf, mBuf)     '' 構造体に代入する

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
    Dim HansyokuNum As String

    mCon_HANSYOKU.BeginTrans
    
    HansyokuNum = Mid$(lBuf, 12, 5)                 ''

    strSQL = "DELETE * FROM HANSYOKU"
    strSQL = strSQL & " WHERE [HansyokuNum]   ='" & HansyokuNum & "'"
    
    mCon_HANSYOKU.Execute strSQL, , adExecuteNoRecords
    
    mCon_HANSYOKU.CommitTrans
        
    Exit Sub
ErrorHandler:
    mCon_HANSYOKU.RollbackTrans
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
        
    mCon_HANSYOKU.BeginTrans

    
    Set rs = mRS_HANSYOKU
    rs.AddNew Array("HansyokuNum"), Array(mBuf.HansyokuNum)
    With mBuf
        With .head
            rs("RecordSpec") = .RecordSpec                                         '' レコード種別
            rs("DataKubun") = .DataKubun                                           '' データ区分
            With .MakeDate
                rs("MakeDate") = .Year & .Month & .Day                             '' 年月日
            End With ' MakeDate
        End With ' head
        rs("reserved") = .Reserved                                                 '' 予備
        rs("KettoNum") = .KettoNum                                                 '' 血統登録番号
        rs("DelKubun") = .DelKubun                                                 '' 繁殖馬抹消区分
        rs("Bamei") = .BAMEI                                                       '' 馬名
        rs("BameiKana") = .BameiKana                                               '' 馬名半角カナ
        rs("BameiEng") = .BameiEng                                                 '' 馬名欧字
        rs("BirthYear") = .BirthYear                                               '' 生年
        rs("SexCD") = .SexCD                                                       '' 性別コード
        rs("HinsyuCD") = .HinsyuCD                                                 '' 品種コード
        rs("KeiroCD") = .KeiroCD                                                   '' 毛色コード
        rs("HansyokuMochiKubun") = .HansyokuMochiKubun                             '' 繁殖馬持込区分
        rs("ImportYear") = .ImportYear                                             '' 輸入年
        rs("SanchiName") = .SanchiName                                             '' 産地名
        rs("HansyokuFNum") = .HansyokuFNum                                         '' 父馬繁殖登録番号
        rs("HansyokuMNum") = .HansyokuMNum                                         '' 母馬繁殖登録番号
        
        
    End With
        
    rs.Update ' HANSYOKU
        
    mCon_HANSYOKU.CommitTrans

    Set rs = Nothing
    
    InsertDB = True
    Exit Function

ErrorHandler:
    If Err.Number <> -2147217887 Then
        gApp.ErrLog
    End If
    rs.CancelUpdate
    
    mCon_HANSYOKU.RollbackTrans

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
        
    Set rs = mRS_HANSYOKU

    rs.Seek Array(mBuf.HansyokuNum)

    With mBuf.head.MakeDate
        If rs("Makedate") > .Year & .Month & .Day Then
            UpdateDB = True
            Exit Function
        End If
    End With

    
    mCon_HANSYOKU.BeginTrans

    With mBuf
        With .head
            rs("RecordSpec") = .RecordSpec                                         '' レコード種別
            rs("DataKubun") = .DataKubun                                           '' データ区分
            With .MakeDate
                rs("MakeDate") = .Year & .Month & .Day                             '' 年月日
            End With ' MakeDate
        End With ' head
        rs("reserved") = .Reserved                                                 '' 予備
        rs("KettoNum") = .KettoNum                                                 '' 血統登録番号
        rs("DelKubun") = .DelKubun                                                 '' 繁殖馬抹消区分
        rs("Bamei") = .BAMEI                                                       '' 馬名
        rs("BameiKana") = .BameiKana                                               '' 馬名半角カナ
        rs("BameiEng") = .BameiEng                                                 '' 馬名欧字
        rs("BirthYear") = .BirthYear                                               '' 生年
        rs("SexCD") = .SexCD                                                       '' 性別コード
        rs("HinsyuCD") = .HinsyuCD                                                 '' 品種コード
        rs("KeiroCD") = .KeiroCD                                                   '' 毛色コード
        rs("HansyokuMochiKubun") = .HansyokuMochiKubun                             '' 繁殖馬持込区分
        rs("ImportYear") = .ImportYear                                             '' 輸入年
        rs("SanchiName") = .SanchiName                                             '' 産地名
        rs("HansyokuFNum") = .HansyokuFNum                                         '' 父馬繁殖登録番号
        rs("HansyokuMNum") = .HansyokuMNum                                         '' 母馬繁殖登録番号
        
        
    End With
        
    rs.Update ' HANSYOKU
    
    mCon_HANSYOKU.CommitTrans

    Set rs = Nothing
    
    UpdateDB = True
    Exit Function

ErrorHandler:
    gApp.ErrLog
    
    mCon_HANSYOKU.RollbackTrans

    UpdateDB = False
End Function
