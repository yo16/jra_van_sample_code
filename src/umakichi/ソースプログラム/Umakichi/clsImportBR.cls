VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsImportBR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Import JV_BR_BREEDER "
'
'   JVData "BR" データベース登録クラス
'

Option Explicit
Option Compare Binary
Implements clsIImport

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部変数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private mBuf As JV_BR_BREEDER

Private mCon_SEISAN As ADODB.Connection

Private mRS_SEISAN As ADODB.Recordset



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
    Set mCon_SEISAN = New ADODB.Connection

    ' レコードセットのインスタンス生成
    Set mRS_SEISAN = New ADODB.Recordset
    
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
    mCon_SEISAN.Open strCon & "subSEISAN.mdb"

    ' レコードセットオープン
    With mRS_SEISAN
        .CursorLocation = adUseServer
        .Index = "PrimaryKey"
        .Open "SEISAN", mCon_SEISAN, adOpenKeyset, adLockOptimistic, adCmdTableDirect
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
    mRS_SEISAN.Close

    mCon_SEISAN.Close

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
        Call SetDataFromByte_BR(lBuf, mBuf)     '' 構造体に代入する

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
    Dim BreederCode As String

    mCon_SEISAN.BeginTrans
    
    BreederCode = Mid$(lBuf, 12, 6)                  ''

    strSQL = "DELETE * FROM SEISAN"
    strSQL = strSQL & " WHERE [BreederCode]   ='" & BreederCode & "'"
    
    mCon_SEISAN.Execute strSQL, , adExecuteNoRecords
    
    mCon_SEISAN.CommitTrans
        
    Exit Sub
ErrorHandler:
    mCon_SEISAN.RollbackTrans
End Sub


'
'   機能: データベースに挿入する
'
'   備考: なし
'
Private Function InsertDB() As Boolean
On Error GoTo ErrorHandler
    Dim i As Integer                                                                      '' ループカウンタ
    Dim j As Integer                                                                      '' ループカウンタ
    Dim k As Integer                                                                      '' ループカウンタ
    Dim rs As ADODB.Recordset                                                             '' SQL文
    
    Set rs = mRS_SEISAN
    
    mCon_SEISAN.BeginTrans
    
    rs.AddNew Array("BreederCode"), Array(mBuf.BreederCode)
    With mBuf
        With .head
            rs("RecordSpec") = .RecordSpec                                                '' レコード種別
            rs("DataKubun") = .DataKubun                                                  '' データ区分
            With .MakeDate
                rs("MakeDate") = .Year & .Month & .Day                                    '' 年月日
            End With ' MakeDate
        End With ' head
        rs("BreederName_Co") = .BreederName_Co                                            '' 生産者名_法人格有
        rs("BreederName") = .BreederName                                                  '' 生産者名_法人格無
        rs("BreederNameKana") = .BreederNameKana                                          '' 生産者名半角カナ
        rs("BreederNameEng") = .BreederNameEng                                            '' 生産者名欧字
        rs("Address") = .Address                                                          '' 生産者住所自治省名
        With .HonRuikei(0)
            rs("H_SetYear") = .SetYear                                                    '' 設定年
            rs("H_HonSyokinTotal") = .HonSyokinTotal                                      '' 本賞金合計
            rs("H_Fukasyokin") = .Fukasyokin                                              '' 付加賞金合計
            For j = 0 To 5
                rs("H_Chakukaisu" & j + 1) = .Chakukaisu(j)                               '' 着回数
            Next j
        End With ' HonRuikei(0)
        With .HonRuikei(1)
            rs("R_SetYear") = .SetYear                                                    '' 設定年
            rs("R_HonSyokinTotal") = .HonSyokinTotal                                      '' 本賞金合計
            rs("R_Fukasyokin") = .Fukasyokin                                              '' 付加賞金合計
            For j = 0 To 5
                rs("R_Chakukaisu" & j + 1) = .Chakukaisu(j)                               '' 着回数
            Next j
        End With ' HonRuikei(1)
        
       
    End With ' mBuf
    
    rs.Update ' SEISAN
    
    mCon_SEISAN.CommitTrans

    Set rs = Nothing
    
    InsertDB = True
    Exit Function
    
ErrorHandler:
    If Err.Number <> -2147217887 Then
        gApp.ErrLog
    End If
    rs.CancelUpdate
    
    mCon_SEISAN.RollbackTrans

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
    Dim i As Integer                                                                      '' ループカウンタ
    Dim j As Integer                                                                      '' ループカウンタ
    Dim k As Integer                                                                      '' ループカウンタ
    Dim strSQL As String                                                                  '' SQL文
    


    Set rs = mRS_SEISAN
    
    rs.Seek Array(mBuf.BreederCode)

    With mBuf.head.MakeDate
        If rs("Makedate") > .Year & .Month & .Day Then
            UpdateDB = True
            Exit Function
        End If
    End With

    mCon_SEISAN.BeginTrans
    
    With mBuf
        With .head
            rs("RecordSpec") = .RecordSpec                                                '' レコード種別
            rs("DataKubun") = .DataKubun                                                  '' データ区分
            With .MakeDate
                rs("MakeDate") = .Year & .Month & .Day                                    '' 年月日
            End With ' MakeDate
        End With ' head
        rs("BreederName_Co") = .BreederName_Co                                            '' 生産者名_法人格有
        rs("BreederName") = .BreederName                                                  '' 生産者名_法人格無
        rs("BreederNameKana") = .BreederNameKana                                          '' 生産者名半角カナ
        rs("BreederNameEng") = .BreederNameEng                                            '' 生産者名欧字
        rs("Address") = .Address                                                          '' 生産者住所自治省名
        With .HonRuikei(0)
            rs("H_SetYear") = .SetYear                                                    '' 設定年
            rs("H_HonSyokinTotal") = .HonSyokinTotal                                      '' 本賞金合計
            rs("H_Fukasyokin") = .Fukasyokin                                              '' 付加賞金合計
            For j = 0 To 5
                rs("H_Chakukaisu" & j + 1) = .Chakukaisu(j)                               '' 着回数
            Next j
        End With ' HonRuikei(0)
        With .HonRuikei(1)
            rs("R_SetYear") = .SetYear                                                    '' 設定年
            rs("R_HonSyokinTotal") = .HonSyokinTotal                                      '' 本賞金合計
            rs("R_Fukasyokin") = .Fukasyokin                                              '' 付加賞金合計
            For j = 0 To 5
                rs("R_Chakukaisu" & j + 1) = .Chakukaisu(j)                               '' 着回数
            Next j
        End With ' HonRuikei(1)
        
       
    End With ' mBuf
    
    rs.Update
    
    mCon_SEISAN.CommitTrans

    
    UpdateDB = True
    Exit Function
    
ErrorHandler:
    gApp.ErrLog
    
    mCon_SEISAN.RollbackTrans

    UpdateDB = False
End Function
