VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsImportSK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Import JV_SK_SANKU "
'
'   JVData "SK" データベース登録クラス
'

Option Explicit
Option Compare Binary
Implements clsIImport

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部変数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private mBuf As JV_SK_SANKU

Private mCon_SANKU As ADODB.Connection

Private mRS_SANKU As ADODB.Recordset


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
    Set mCon_SANKU = New ADODB.Connection

    ' レコードセットのインスタンス生成
    Set mRS_SANKU = New ADODB.Recordset
    
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
    mCon_SANKU.Open strCon & "subSANKU.mdb"

    ' レコードセットオープン
    With mRS_SANKU
        .CursorLocation = adUseServer
        .Index = "PrimaryKey"
        .Open "SANKU", mCon_SANKU, adOpenKeyset, adLockOptimistic, adCmdTableDirect
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
    mRS_SANKU.Close

    mCon_SANKU.Close

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
        Call SetDataFromByte_SK(lBuf, mBuf)     '' 構造体に代入する

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
    Dim KettoNum    As String
    
    mCon_SANKU.BeginTrans
    
    KettoNum = Mid$(lBuf, 12, 10)                  ''

    strSQL = "DELETE * FROM SANKU"
    strSQL = strSQL & " WHERE [KettoNum]   ='" & KettoNum & "'"
    
    mCon_SANKU.Execute strSQL, , adExecuteNoRecords
    
    mCon_SANKU.CommitTrans
        
    Exit Sub
ErrorHandler:
    mCon_SANKU.RollbackTrans
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
    
    
    mCon_SANKU.BeginTrans

    Set rs = mRS_SANKU
    rs.AddNew Array("KettoNum"), Array(mBuf.KettoNum)
    With mBuf
        With .head
            rs("RecordSpec") = .RecordSpec                                      '' レコード種別
            rs("DataKubun") = .DataKubun                                        '' データ区分
            With .MakeDate
                rs("MakeDate") = .Year & .Month & .Day                          '' 年月日
            End With ' MakeDate
        End With ' head
        With .BirthDate
            rs("BirthDate") = .Year & .Month & .Day                             '' 年月日
        End With ' BirthDate
        rs("SexCD") = .SexCD                                                    '' 性別コード
        rs("HinsyuCD") = .HinsyuCD                                              '' 品種コード
        rs("KeiroCD") = .KeiroCD                                                '' 毛色コード
        rs("SankuMochiKubun") = .SankuMochiKubun                                '' 産駒持込区分
        rs("ImportYear") = .ImportYear                                          '' 輸入年
        rs("BreederCode") = .BreederCode                                        '' 生産者コード
        rs("SanchiName") = .SanchiName                                          '' 産地名
        rs("FNum") = .HansyokuNum(0)                                            '' 父繁殖登録番号
        rs("MNum") = .HansyokuNum(1)                                            '' 母繁殖登録番号
        rs("FFNum") = .HansyokuNum(2)                                           '' 父父繁殖登録番号
        rs("FMNum") = .HansyokuNum(3)                                           '' 父母繁殖登録番号
        rs("MFNum") = .HansyokuNum(4)                                           '' 母父繁殖登録番号
        rs("MMNum") = .HansyokuNum(5)                                           '' 母母繁殖登録番号
        rs("FFFNum") = .HansyokuNum(6)                                          '' 父父父繁殖登録番号
        rs("FFMNum") = .HansyokuNum(7)                                          '' 父父母繁殖登録番号
        rs("FMFNum") = .HansyokuNum(8)                                          '' 父母父繁殖登録番号
        rs("FMMNum") = .HansyokuNum(9)                                          '' 父母母繁殖登録番号
        rs("MFFNum") = .HansyokuNum(10)                                         '' 母父父繁殖登録番号
        rs("MFMNum") = .HansyokuNum(11)                                         '' 母父母繁殖登録番号
        rs("MMFNum") = .HansyokuNum(12)                                         '' 母母父繁殖登録番号
        rs("MMMNum") = .HansyokuNum(13)                                         '' 母母母繁殖登録番号
    End With
    
    rs.Update ' SANKU
    
    mCon_SANKU.CommitTrans

    Set rs = Nothing
    
    InsertDB = True
    Exit Function

ErrorHandler:
    If Err.Number <> -2147217887 Then
        gApp.ErrLog
    End If
    rs.CancelUpdate
    
    mCon_SANKU.RollbackTrans

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
    
    Set rs = mRS_SANKU

    rs.Seek Array(mBuf.KettoNum)

    With mBuf.head.MakeDate
        If rs("Makedate") > .Year & .Month & .Day Then
            UpdateDB = True
            Exit Function
        End If
    End With

    
    mCon_SANKU.BeginTrans

    With mBuf
        With .head
            rs("RecordSpec") = .RecordSpec                                      '' レコード種別
            rs("DataKubun") = .DataKubun                                        '' データ区分
            With .MakeDate
                rs("MakeDate") = .Year & .Month & .Day                          '' 年月日
            End With ' MakeDate
        End With ' head
        With .BirthDate
            rs("BirthDate") = .Year & .Month & .Day                             '' 年月日
        End With ' BirthDate
        rs("SexCD") = .SexCD                                                    '' 性別コード
        rs("HinsyuCD") = .HinsyuCD                                              '' 品種コード
        rs("KeiroCD") = .KeiroCD                                                '' 毛色コード
        rs("SankuMochiKubun") = .SankuMochiKubun                                '' 産駒持込区分
        rs("ImportYear") = .ImportYear                                          '' 輸入年
        rs("BreederCode") = .BreederCode                                        '' 生産者コード
        rs("SanchiName") = .SanchiName                                          '' 産地名
        rs("FNum") = .HansyokuNum(0)                                            '' 父繁殖登録番号
        rs("MNum") = .HansyokuNum(1)                                            '' 母繁殖登録番号
        rs("FFNum") = .HansyokuNum(2)                                           '' 父父繁殖登録番号
        rs("FMNum") = .HansyokuNum(3)                                           '' 父母繁殖登録番号
        rs("MFNum") = .HansyokuNum(4)                                           '' 母父繁殖登録番号
        rs("MMNum") = .HansyokuNum(5)                                           '' 母母繁殖登録番号
        rs("FFFNum") = .HansyokuNum(6)                                          '' 父父父繁殖登録番号
        rs("FFMNum") = .HansyokuNum(7)                                          '' 父父母繁殖登録番号
        rs("FMFNum") = .HansyokuNum(8)                                          '' 父母父繁殖登録番号
        rs("FMMNum") = .HansyokuNum(9)                                          '' 父母母繁殖登録番号
        rs("MFFNum") = .HansyokuNum(10)                                         '' 母父父繁殖登録番号
        rs("MFMNum") = .HansyokuNum(11)                                         '' 母父母繁殖登録番号
        rs("MMFNum") = .HansyokuNum(12)                                         '' 母母父繁殖登録番号
        rs("MMMNum") = .HansyokuNum(13)                                         '' 母母母繁殖登録番号
    End With
    
    rs.Update
    
    mCon_SANKU.CommitTrans

    Set rs = Nothing
    
    UpdateDB = True
    Exit Function

ErrorHandler:
    gApp.ErrLog
    
    mCon_SANKU.RollbackTrans

    UpdateDB = False
End Function
