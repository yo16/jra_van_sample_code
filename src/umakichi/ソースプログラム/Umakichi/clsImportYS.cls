VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsImportYS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Import JV_YS_SCHEDULE "
'
'   JVData "YS" データベース登録クラス
'

Option Explicit
Option Compare Binary
Implements clsIImport

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部変数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private mBuf As JV_YS_SCHEDULE

Private mCon_SCHEDULE As ADODB.Connection

Private mRS_SCHEDULE As ADODB.Recordset


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
    Set mCon_SCHEDULE = New ADODB.Connection

    ' レコードセットのインスタンス生成
    Set mRS_SCHEDULE = New ADODB.Recordset
    
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
    mCon_SCHEDULE.Open strCon & "subSCHEDULE.mdb"

    ' レコードセットオープン
    With mRS_SCHEDULE
        .CursorLocation = adUseServer
        .Index = "PrimaryKey"
        .Open "SCHEDULE", mCon_SCHEDULE, adOpenKeyset, adLockOptimistic, adCmdTableDirect
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
    
    mRS_SCHEDULE.Close

    mCon_SCHEDULE.Close

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
        Call SetDataFromByte_YS(lBuf, mBuf)     '' 構造体に代入する

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
    
    mCon_SCHEDULE.BeginTrans
    
    Year = Mid$(lBuf, 12, 4)                    '' 開催年
    MonthDay = Mid$(lBuf, 16, 4)                '' 開催月日
    JyoCD = Mid$(lBuf, 20, 2)                   '' 競馬場コード
    Kaiji = Mid$(lBuf, 22, 2)                   '' 開催回[第N回]
    Nichiji = Mid$(lBuf, 24, 2)                 '' 開催日目[N日目]

    strSQL = "DELETE * FROM SCHEDULE"
    strSQL = strSQL & " WHERE [Year]   ='" & Year & "'"
    strSQL = strSQL & " AND [MonthDay] ='" & MonthDay & "'"
    strSQL = strSQL & " AND [JyoCD]    ='" & JyoCD & "'"
    strSQL = strSQL & " AND [Kaiji]    ='" & Kaiji & "'"
    strSQL = strSQL & " AND [Nichiji]  ='" & Nichiji & "'"
    
    mCon_SCHEDULE.Execute strSQL, , adExecuteNoRecords
    
    mCon_SCHEDULE.CommitTrans
        
    Exit Sub
ErrorHandler:
    mCon_SCHEDULE.RollbackTrans
End Sub



'
'   機能: データベースに挿入する
'
'   備考: なし
'
Private Function InsertDB() As Boolean
On Error GoTo ErrorHandler
    Dim i As Integer                                                                  '' ループカウンタ
    Dim j As Integer                                                                  '' ループカウンタ
    Dim k As Integer                                                                  '' ループカウンタ
    Dim rs As ADODB.Recordset                                                         '' SQL文
    
    
    mCon_SCHEDULE.BeginTrans

    Set rs = mRS_SCHEDULE
    With mBuf.id
        rs.AddNew Array("Year", "MonthDay", "JyoCD", "Kaiji", "Nichiji"), Array(.Year, .MonthDay, .JyoCD, .Kaiji, .Nichiji)
    End With
    With mBuf
        With .head
            rs("RecordSpec") = .RecordSpec                                            '' レコード種別
            rs("DataKubun") = .DataKubun                                              '' データ区分
            With .MakeDate
                rs("MakeDate") = .Year & .Month & .Day                                '' 年月日
            End With ' MakeDate
        End With ' head
        rs("YoubiCD") = .YoubiCD                                                      '' 曜日コード
        With .JyusyoInfo(0)
            rs("Jyusyo1TokuNum") = .TokuNum                                           '' 特別競走番号
            rs("Jyusyo1Hondai") = .Hondai                                             '' 競走名本題
            rs("Jyusyo1Ryakusyo10") = .Ryakusyo10                                     '' 競走名略称10字
            rs("Jyusyo1Ryakusyo6") = .Ryakusyo6                                       '' 競走名略称6字
            rs("Jyusyo1Ryakusyo3") = .Ryakusyo3                                       '' 競走名略称3字
            rs("Jyusyo1Nkai") = .Nkai                                                 '' 重賞回次[第N回]
            rs("Jyusyo1GradeCD") = .GradeCD                                           '' グレードコード
            rs("Jyusyo1SyubetuCD") = .SyubetuCD                                       '' 競走種別コード
            rs("Jyusyo1KigoCD") = .KigoCD                                             '' 競走記号コード
            rs("Jyusyo1JyuryoCD") = .JyuryoCD                                         '' 重量種別コード
            rs("Jyusyo1Kyori") = .KYORI                                               '' 距離
            rs("Jyusyo1TrackCD") = .TrackCD                                           '' トラックコード
        End With ' JyusyoInfo(0)
        With .JyusyoInfo(1)
            rs("Jyusyo2TokuNum") = .TokuNum                                           '' 特別競走番号
            rs("Jyusyo2Hondai") = .Hondai                                             '' 競走名本題
            rs("Jyusyo2Ryakusyo10") = .Ryakusyo10                                     '' 競走名略称10字
            rs("Jyusyo2Ryakusyo6") = .Ryakusyo6                                       '' 競走名略称6字
            rs("Jyusyo2Ryakusyo3") = .Ryakusyo3                                       '' 競走名略称3字
            rs("Jyusyo2Nkai") = .Nkai                                                 '' 重賞回次[第N回]
            rs("Jyusyo2GradeCD") = .GradeCD                                           '' グレードコード
            rs("Jyusyo2SyubetuCD") = .SyubetuCD                                       '' 競走種別コード
            rs("Jyusyo2KigoCD") = .KigoCD                                             '' 競走記号コード
            rs("Jyusyo2JyuryoCD") = .JyuryoCD                                         '' 重量種別コード
            rs("Jyusyo2Kyori") = .KYORI                                               '' 距離
            rs("Jyusyo2TrackCD") = .TrackCD                                           '' トラックコード
        End With ' JyusyoInfo(1)
        With .JyusyoInfo(2)
            rs("Jyusyo3TokuNum") = .TokuNum                                           '' 特別競走番号
            rs("Jyusyo3Hondai") = .Hondai                                             '' 競走名本題
            rs("Jyusyo3Ryakusyo10") = .Ryakusyo10                                     '' 競走名略称10字
            rs("Jyusyo3Ryakusyo6") = .Ryakusyo6                                       '' 競走名略称6字
            rs("Jyusyo3Ryakusyo3") = .Ryakusyo3                                       '' 競走名略称3字
            rs("Jyusyo3Nkai") = .Nkai                                                 '' 重賞回次[第N回]
            rs("Jyusyo3GradeCD") = .GradeCD                                           '' グレードコード
            rs("Jyusyo3SyubetuCD") = .SyubetuCD                                       '' 競走種別コード
            rs("Jyusyo3KigoCD") = .KigoCD                                             '' 競走記号コード
            rs("Jyusyo3JyuryoCD") = .JyuryoCD                                         '' 重量種別コード
            rs("Jyusyo3Kyori") = .KYORI                                               '' 距離
            rs("Jyusyo3TrackCD") = .TrackCD                                           '' トラックコード
        End With ' JyusyoInfo(2)
    End With
    
    rs.Update ' SCHEDULE
    
    mCon_SCHEDULE.CommitTrans

    Set rs = Nothing
    
    InsertDB = True
    Exit Function
    
ErrorHandler:
    If Err.Number <> -2147217887 Then
        gApp.ErrLog
    End If
    rs.CancelUpdate
    
    mCon_SCHEDULE.RollbackTrans

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
    Dim i As Integer                                                                  '' ループカウンタ
    Dim j As Integer                                                                  '' ループカウンタ
    Dim k As Integer                                                                  '' ループカウンタ
    Dim strSQL As String                                                              '' SQL文
    
    Set rs = mRS_SCHEDULE

    With mBuf.id
        rs.Seek Array(.Year, .MonthDay, .JyoCD, .Kaiji, .Nichiji)
    End With

    With mBuf.head.MakeDate
        If rs("Makedate") > .Year & .Month & .Day Then
            UpdateDB = True
            Exit Function
        End If
    End With

    
    mCon_SCHEDULE.BeginTrans

    
    With mBuf
        With .head
            rs("RecordSpec") = .RecordSpec                                            '' レコード種別
            rs("DataKubun") = .DataKubun                                              '' データ区分
            With .MakeDate
                rs("MakeDate") = .Year & .Month & .Day                                '' 年月日
            End With ' MakeDate
        End With ' head
        rs("YoubiCD") = .YoubiCD                                                      '' 曜日コード
        With .JyusyoInfo(0)
            rs("Jyusyo1TokuNum") = .TokuNum                                           '' 特別競走番号
            rs("Jyusyo1Hondai") = .Hondai                                             '' 競走名本題
            rs("Jyusyo1Ryakusyo10") = .Ryakusyo10                                     '' 競走名略称10字
            rs("Jyusyo1Ryakusyo6") = .Ryakusyo6                                       '' 競走名略称6字
            rs("Jyusyo1Ryakusyo3") = .Ryakusyo3                                       '' 競走名略称3字
            rs("Jyusyo1Nkai") = .Nkai                                                 '' 重賞回次[第N回]
            rs("Jyusyo1GradeCD") = .GradeCD                                           '' グレードコード
            rs("Jyusyo1SyubetuCD") = .SyubetuCD                                       '' 競走種別コード
            rs("Jyusyo1KigoCD") = .KigoCD                                             '' 競走記号コード
            rs("Jyusyo1JyuryoCD") = .JyuryoCD                                         '' 重量種別コード
            rs("Jyusyo1Kyori") = .KYORI                                               '' 距離
            rs("Jyusyo1TrackCD") = .TrackCD                                           '' トラックコード
        End With ' JyusyoInfo(0)
        With .JyusyoInfo(1)
            rs("Jyusyo2TokuNum") = .TokuNum                                           '' 特別競走番号
            rs("Jyusyo2Hondai") = .Hondai                                             '' 競走名本題
            rs("Jyusyo2Ryakusyo10") = .Ryakusyo10                                     '' 競走名略称10字
            rs("Jyusyo2Ryakusyo6") = .Ryakusyo6                                       '' 競走名略称6字
            rs("Jyusyo2Ryakusyo3") = .Ryakusyo3                                       '' 競走名略称3字
            rs("Jyusyo2Nkai") = .Nkai                                                 '' 重賞回次[第N回]
            rs("Jyusyo2GradeCD") = .GradeCD                                           '' グレードコード
            rs("Jyusyo2SyubetuCD") = .SyubetuCD                                       '' 競走種別コード
            rs("Jyusyo2KigoCD") = .KigoCD                                             '' 競走記号コード
            rs("Jyusyo2JyuryoCD") = .JyuryoCD                                         '' 重量種別コード
            rs("Jyusyo2Kyori") = .KYORI                                               '' 距離
            rs("Jyusyo2TrackCD") = .TrackCD                                           '' トラックコード
        End With ' JyusyoInfo(1)
        With .JyusyoInfo(2)
            rs("Jyusyo3TokuNum") = .TokuNum                                           '' 特別競走番号
            rs("Jyusyo3Hondai") = .Hondai                                             '' 競走名本題
            rs("Jyusyo3Ryakusyo10") = .Ryakusyo10                                     '' 競走名略称10字
            rs("Jyusyo3Ryakusyo6") = .Ryakusyo6                                       '' 競走名略称6字
            rs("Jyusyo3Ryakusyo3") = .Ryakusyo3                                       '' 競走名略称3字
            rs("Jyusyo3Nkai") = .Nkai                                                 '' 重賞回次[第N回]
            rs("Jyusyo3GradeCD") = .GradeCD                                           '' グレードコード
            rs("Jyusyo3SyubetuCD") = .SyubetuCD                                       '' 競走種別コード
            rs("Jyusyo3KigoCD") = .KigoCD                                             '' 競走記号コード
            rs("Jyusyo3JyuryoCD") = .JyuryoCD                                         '' 重量種別コード
            rs("Jyusyo3Kyori") = .KYORI                                               '' 距離
            rs("Jyusyo3TrackCD") = .TrackCD                                           '' トラックコード
        End With ' JyusyoInfo(2)
    End With
    
    rs.Update
    
    mCon_SCHEDULE.CommitTrans

    Set rs = Nothing
    
    UpdateDB = True
    Exit Function
    
ErrorHandler:
    gApp.ErrLog
    
    mCon_SCHEDULE.RollbackTrans

    UpdateDB = False
End Function
