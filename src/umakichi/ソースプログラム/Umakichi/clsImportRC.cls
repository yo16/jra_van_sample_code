VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsImportRC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Import JV_RC_RECORD "
'
'   JVData "RC" データベース登録クラス
'

Option Explicit
Option Compare Binary
Implements clsIImport

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部変数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private mBuf As JV_RC_RECORD

Private mCon_RECORD As ADODB.Connection

Private mRS_RECORD As ADODB.Recordset


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
    Set mCon_RECORD = New ADODB.Connection

    ' レコードセットのインスタンス生成
    Set mRS_RECORD = New ADODB.Recordset
    
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
    mCon_RECORD.Open strCon & "subRECORD.mdb"

    ' レコードセットオープン
    With mRS_RECORD
        .CursorLocation = adUseServer
        .Index = "PrimaryKey"
        .Open "RECORD", mCon_RECORD, adOpenKeyset, adLockOptimistic, adCmdTableDirect
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

    mRS_RECORD.Close

    mCon_RECORD.Close

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
        Call DeleteRecord(lBuf)
        clsIImport_Add = True
    Else
        Call SetDataFromByte_RC(lBuf, mBuf) '' 構造体に代入する

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
Private Sub DeleteRecord(lBuf() As Byte)
On Error GoTo ErrorHandler
    Dim strSQL      As String
    Dim RecInfoKubun As String
    Dim Year        As String
    Dim MonthDay    As String
    Dim JyoCD       As String
    Dim Kaiji       As String
    Dim Nichiji     As String
    Dim RaceNum     As String
    Dim TokuNum     As String
    Dim SyubetuCD   As String
    Dim KYORI       As String
    Dim TrackCD     As String
    
    mCon_RECORD.BeginTrans
    
    RecInfoKubun = StrConv(MidB$(lBuf, 12, 1), vbUnicode)            ''
    Year = StrConv(MidB$(lBuf, 13, 4), vbUnicode)                    '' 開催年
    MonthDay = StrConv(MidB$(lBuf, 17, 4), vbUnicode)                '' 開催月日
    JyoCD = StrConv(MidB$(lBuf, 21, 2), vbUnicode)                   '' 競馬場コード
    Kaiji = StrConv(MidB$(lBuf, 23, 2), vbUnicode)                   '' 開催回[第N回]
    Nichiji = StrConv(MidB$(lBuf, 25, 2), vbUnicode)                 '' 開催日目[N日目]
    RaceNum = StrConv(MidB$(lBuf, 27, 2), vbUnicode)                 '' レース番号
    TokuNum = StrConv(MidB$(lBuf, 29, 4), vbUnicode)                 ''
    SyubetuCD = StrConv(MidB$(lBuf, 94, 2), vbUnicode)               ''
    KYORI = StrConv(MidB$(lBuf, 96, 4), vbUnicode)                   ''
    TrackCD = StrConv(MidB$(lBuf, 100, 2), vbUnicode)                ''

    strSQL = "DELETE * FROM RECORD"
    strSQL = strSQL & " WHERE [RecInfoKubun]='" & RecInfoKubun & "'"
    strSQL = strSQL & " AND [Year]     ='" & Year & "'"
    strSQL = strSQL & " AND [MonthDay] ='" & MonthDay & "'"
    strSQL = strSQL & " AND [JyoCD]    ='" & JyoCD & "'"
    strSQL = strSQL & " AND [Kaiji]    ='" & Kaiji & "'"
    strSQL = strSQL & " AND [Nichiji]  ='" & Nichiji & "'"
    strSQL = strSQL & " AND [RaceNum]  ='" & RaceNum & "'"
    strSQL = strSQL & " AND [TokuNum_SyubetuCD]='" & TokuNum & SyubetuCD & "'"
    strSQL = strSQL & " AND [Kyori]  ='" & KYORI & "'"
    strSQL = strSQL & " AND [TrackCD]  ='" & TrackCD & "'"
    
    mCon_RECORD.Execute strSQL, , adExecuteNoRecords
    
    mCon_RECORD.CommitTrans
            
    Exit Sub
ErrorHandler:
    mCon_RECORD.RollbackTrans
End Sub


'
'   機能: データベースに挿入する
'
'   備考: なし
'
Private Function InsertDB() As Boolean
On Error GoTo ErrorHandler
    Dim i As Integer                                                                               '' ループカウンタ
    Dim j As Integer                                                                               '' ループカウンタ
    Dim k As Integer                                                                               '' ループカウンタ
    Dim rs As ADODB.Recordset                                                                      '' SQL文
    
    
    mCon_RECORD.BeginTrans

    Set rs = mRS_RECORD
    With mBuf.id
        rs.AddNew Array("RecInfoKubun", "Year", "MonthDay", "JyoCD", "Kaiji", "Nichiji", "RaceNum", "TokuNum_SyubetuCD", "Kyori", "TrackCD"), Array(mBuf.RecInfoKubun, .Year, .MonthDay, .JyoCD, .Kaiji, .Nichiji, .RaceNum, mBuf.TokuNum & mBuf.SyubetuCD, mBuf.KYORI, mBuf.TrackCD)
    End With
    With mBuf
        With .head
            rs("RecordSpec") = .RecordSpec                                                         '' レコード種別
            rs("DataKubun") = .DataKubun                                                           '' データ区分
            With .MakeDate
                rs("MakeDate") = .Year & .Month & .Day                                             '' 年月日
            End With ' MakeDate
        End With ' head
        rs("RecInfoKubun") = .RecInfoKubun                                                         '' レコード識別区分
        rs("Hondai") = .Hondai                                                                     '' 競走名本題
        rs("GradeCD") = .GradeCD                                                                   '' グレードコード
        rs("RecKubun") = .RecKubun                                                                 '' レコード区分
        rs("RecTime") = .RecTime                                                                   '' レコードタイム
        With .TenkoBaba
            rs("TenkoCD") = .TenkoCD                                                               '' 天候コード
            rs("SibaBabaCD") = .SibaBabaCD                                                         '' 芝馬場状態コード
            rs("DirtBabaCD") = .DirtBabaCD                                                         '' ダート馬場状態コード
        End With ' TenkoBaba
        With .RecUmaInfo(0)
            rs("RecUmaKettoNum1") = .KettoNum                                                      '' 血統登録番号
            rs("RecUmaBamei1") = .BAMEI                                                            '' 馬名
            rs("RecUmaUmaKigoCD1") = .UmaKigoCD                                                    '' 馬記号コード
            rs("RecUmaSexCD1") = .SexCD                                                            '' 性別コード
            rs("RecUmaChokyosiCode1") = .ChokyosiCode                                              '' 調教師コード
            rs("RecUmaChokyosiName1") = .ChokyosiName                                              '' 調教師名
            rs("RecUmaFutan1") = .Futan                                                            '' 負担重量
            rs("RecUmaKisyuCode1") = .KisyuCode                                                    '' 騎手コード
            rs("RecUmaKisyuName1") = .KisyuName                                                    '' 騎手名
        End With ' RecUmaInfo
        With .RecUmaInfo(1)
            rs("RecUmaKettoNum2") = .KettoNum                                                      '' 血統登録番号
            rs("RecUmaBamei2") = .BAMEI                                                            '' 馬名
            rs("RecUmaUmaKigoCD2") = .UmaKigoCD                                                    '' 馬記号コード
            rs("RecUmaSexCD2") = .SexCD                                                            '' 性別コード
            rs("RecUmaChokyosiCode2") = .ChokyosiCode                                              '' 調教師コード
            rs("RecUmaChokyosiName2") = .ChokyosiName                                              '' 調教師名
            rs("RecUmaFutan2") = .Futan                                                            '' 負担重量
            rs("RecUmaKisyuCode2") = .KisyuCode                                                    '' 騎手コード
            rs("RecUmaKisyuName2") = .KisyuName                                                    '' 騎手名
        End With ' RecUmaInfo
        With .RecUmaInfo(2)
            rs("RecUmaKettoNum3") = .KettoNum                                                      '' 血統登録番号
            rs("RecUmaBamei3") = .BAMEI                                                            '' 馬名
            rs("RecUmaUmaKigoCD3") = .UmaKigoCD                                                    '' 馬記号コード
            rs("RecUmaSexCD3") = .SexCD                                                            '' 性別コード
            rs("RecUmaChokyosiCode3") = .ChokyosiCode                                              '' 調教師コード
            rs("RecUmaChokyosiName3") = .ChokyosiName                                              '' 調教師名
            rs("RecUmaFutan3") = .Futan                                                            '' 負担重量
            rs("RecUmaKisyuCode3") = .KisyuCode                                                    '' 騎手コード
            rs("RecUmaKisyuName3") = .KisyuName                                                    '' 騎手名
        End With ' RecUmaInfo
    End With
    
    rs.Update ' RECORD
    
    mCon_RECORD.CommitTrans

    Set rs = Nothing
    
    InsertDB = True
    Exit Function
    
ErrorHandler:
    If Err.Number <> -2147217887 Then
        gApp.ErrLog
    End If
    rs.CancelUpdate
    
    mCon_RECORD.RollbackTrans

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
    Dim i As Integer                                                                               '' ループカウンタ
    Dim j As Integer                                                                               '' ループカウンタ
    Dim k As Integer                                                                               '' ループカウンタ
    Dim strSQL As String                                                                           '' SQL文
    
    Set rs = mRS_RECORD

    With mBuf.id
        rs.Seek Array(mBuf.RecInfoKubun, .Year, .MonthDay, .JyoCD, .Kaiji, .Nichiji, .RaceNum, mBuf.TokuNum & mBuf.SyubetuCD, mBuf.KYORI, mBuf.TrackCD)
    End With

    With mBuf.head.MakeDate
        If rs("Makedate") > .Year & .Month & .Day Then
            UpdateDB = True
            Exit Function
        End If
    End With

    
    mCon_RECORD.BeginTrans

    
    With mBuf
        With .head
            rs("RecordSpec") = .RecordSpec                                                         '' レコード種別
            rs("DataKubun") = .DataKubun                                                           '' データ区分
            With .MakeDate
                rs("MakeDate") = .Year & .Month & .Day                                             '' 年月日
            End With ' MakeDate
        End With ' head
        rs("RecInfoKubun") = .RecInfoKubun                                                         '' レコード識別区分
        rs("Hondai") = .Hondai                                                                     '' 競走名本題
        rs("GradeCD") = .GradeCD                                                                   '' グレードコード
        rs("RecKubun") = .RecKubun                                                                 '' レコード区分
        rs("RecTime") = .RecTime                                                                   '' レコードタイム
        With .TenkoBaba
            rs("TenkoCD") = .TenkoCD                                                               '' 天候コード
            rs("SibaBabaCD") = .SibaBabaCD                                                         '' 芝馬場状態コード
            rs("DirtBabaCD") = .DirtBabaCD                                                         '' ダート馬場状態コード
        End With ' TenkoBaba
        With .RecUmaInfo(0)
            rs("RecUmaKettoNum1") = .KettoNum                                                      '' 血統登録番号
            rs("RecUmaBamei1") = .BAMEI                                                            '' 馬名
            rs("RecUmaUmaKigoCD1") = .UmaKigoCD                                                    '' 馬記号コード
            rs("RecUmaSexCD1") = .SexCD                                                            '' 性別コード
            rs("RecUmaChokyosiCode1") = .ChokyosiCode                                              '' 調教師コード
            rs("RecUmaChokyosiName1") = .ChokyosiName                                              '' 調教師名
            rs("RecUmaFutan1") = .Futan                                                            '' 負担重量
            rs("RecUmaKisyuCode1") = .KisyuCode                                                    '' 騎手コード
            rs("RecUmaKisyuName1") = .KisyuName                                                    '' 騎手名
        End With ' RecUmaInfo
        With .RecUmaInfo(1)
            rs("RecUmaKettoNum2") = .KettoNum                                                      '' 血統登録番号
            rs("RecUmaBamei2") = .BAMEI                                                            '' 馬名
            rs("RecUmaUmaKigoCD2") = .UmaKigoCD                                                    '' 馬記号コード
            rs("RecUmaSexCD2") = .SexCD                                                            '' 性別コード
            rs("RecUmaChokyosiCode2") = .ChokyosiCode                                              '' 調教師コード
            rs("RecUmaChokyosiName2") = .ChokyosiName                                              '' 調教師名
            rs("RecUmaFutan2") = .Futan                                                            '' 負担重量
            rs("RecUmaKisyuCode2") = .KisyuCode                                                    '' 騎手コード
            rs("RecUmaKisyuName2") = .KisyuName                                                    '' 騎手名
        End With ' RecUmaInfo
        With .RecUmaInfo(2)
            rs("RecUmaKettoNum3") = .KettoNum                                                      '' 血統登録番号
            rs("RecUmaBamei3") = .BAMEI                                                            '' 馬名
            rs("RecUmaUmaKigoCD3") = .UmaKigoCD                                                    '' 馬記号コード
            rs("RecUmaSexCD3") = .SexCD                                                            '' 性別コード
            rs("RecUmaChokyosiCode3") = .ChokyosiCode                                              '' 調教師コード
            rs("RecUmaChokyosiName3") = .ChokyosiName                                              '' 調教師名
            rs("RecUmaFutan3") = .Futan                                                            '' 負担重量
            rs("RecUmaKisyuCode3") = .KisyuCode                                                    '' 騎手コード
            rs("RecUmaKisyuName3") = .KisyuName                                                    '' 騎手名
        End With ' RecUmaInfo
    End With
    
    rs.Update
    
    mCon_RECORD.CommitTrans

    
    UpdateDB = True
    Exit Function
    
ErrorHandler:
    gApp.ErrLog
    
    mCon_RECORD.RollbackTrans

    UpdateDB = False
End Function
