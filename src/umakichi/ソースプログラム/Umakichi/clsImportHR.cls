VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsImportHR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Import JV_HR_PAY "
'
'   JVData "HR" データベース登録クラス
'

Option Explicit
Option Compare Binary
Implements clsIImport

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部変数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private mBuf As JV_HR_PAY

Private mCon_HARAI As ADODB.Connection

Private mRS_HARAI As ADODB.Recordset


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
    Set mCon_HARAI = New ADODB.Connection

    ' レコードセットのインスタンス生成
    Set mRS_HARAI = New ADODB.Recordset
    
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
    mCon_HARAI.Open strCon & "subHARAI.mdb"

    ' レコードセットオープン
    With mRS_HARAI
        .CursorLocation = adUseServer
        .Index = "PrimaryKey"
        .Open "HARAI", mCon_HARAI, adOpenKeyset, adLockOptimistic, adCmdTableDirect
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
    mRS_HARAI.Close

    mCon_HARAI.Close

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
        Call SetDataFromByte_HR(lBuf, mBuf)     '' 構造体に代入する

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
    
    mCon_HARAI.BeginTrans
    
    Year = Mid$(lBuf, 12, 4)                    '' 開催年
    MonthDay = Mid$(lBuf, 16, 4)                '' 開催月日
    JyoCD = Mid$(lBuf, 20, 2)                   '' 競馬場コード
    Kaiji = Mid$(lBuf, 22, 2)                   '' 開催回[第N回]
    Nichiji = Mid$(lBuf, 24, 2)                 '' 開催日目[N日目]
    RaceNum = Mid$(lBuf, 26, 2)                 '' レース番号

    strSQL = "DELETE * FROM HARAI"
    strSQL = strSQL & " WHERE [Year]   ='" & Year & "'"
    strSQL = strSQL & " AND [MonthDay] ='" & MonthDay & "'"
    strSQL = strSQL & " AND [JyoCD]    ='" & JyoCD & "'"
    strSQL = strSQL & " AND [Kaiji]    ='" & Kaiji & "'"
    strSQL = strSQL & " AND [Nichiji]  ='" & Nichiji & "'"
    strSQL = strSQL & " AND [RaceNum]  ='" & RaceNum & "'"
    
    mCon_HARAI.Execute strSQL, , adExecuteNoRecords
    
    mCon_HARAI.CommitTrans
        
    Exit Sub
ErrorHandler:
    mCon_HARAI.RollbackTrans
End Sub


'
'   機能: データベースに挿入する
'
'   備考: なし
'
Private Function InsertDB() As Boolean
On Error GoTo ErrorHandler
    Dim i As Integer                                                                          '' ループカウンタ
    Dim j As Integer                                                                          '' ループカウンタ
    Dim k As Integer                                                                          '' ループカウンタ
    Dim rs As ADODB.Recordset                                                                 '' SQL文
    
    mCon_HARAI.BeginTrans

    
    Set rs = mRS_HARAI
    With mBuf.id
        rs.AddNew Array("Year", "MonthDay", "JyoCD", "Kaiji", "Nichiji", "RaceNum"), Array(.Year, .MonthDay, .JyoCD, .Kaiji, .Nichiji, .RaceNum)
    End With
    With mBuf
        With .head
            rs("RecordSpec") = .RecordSpec                                                    '' レコード種別
            rs("DataKubun") = .DataKubun                                                      '' データ区分
            With .MakeDate
                rs("MakeDate") = .Year & .Month & .Day                                        '' 年月日
            End With ' MakeDate
        End With ' head
        rs("TorokuTosu") = .TorokuTosu                                                        '' 登録頭数
        rs("SyussoTosu") = .SyussoTosu                                                        '' 出走頭数
        For i = 0 To 8
            rs("FuseirituFlag" & i + 1) = .FuseirituFlag(i)                                   '' 不成立フラグ
        Next i
        For i = 0 To 8
            rs("TokubaraiFlag" & i + 1) = .TokubaraiFlag(i)                                   '' 特払フラグ
        Next i
        For i = 0 To 8
            rs("HenkanFlag" & i + 1) = .HenkanFlag(i)                                         '' 返還フラグ
        Next i
        For i = 0 To 27
            rs("HenkanUma" & i + 1) = .HenkanUma(i)                                           '' 返還馬番情報(馬番01～28)
        Next i
        For i = 0 To 7
            rs("HenkanWaku" & i + 1) = .HenkanWaku(i)                                         '' 返還枠番情報(枠番1～8)
        Next i
        For i = 0 To 7
            rs("HenkanDoWaku" & i + 1) = .HenkanDoWaku(i)                                     '' 返還同枠情報(枠番1～8)
        Next i
        For i = 0 To 2
            With .PayTansyo(i)
                rs("PayTansyoUmaban" & i + 1) = .Umaban                                       '' 馬番
                rs("PayTansyoPay" & i + 1) = .Pay                                             '' 払戻金
                rs("PayTansyoNinki" & i + 1) = .Ninki                                         '' 人気順
            End With ' PayTansyo
        Next i
        For i = 0 To 4
            With .PayFukusyo(i)
                rs("PayFukusyoUmaban" & i + 1) = .Umaban                                      '' 馬番
                rs("PayFukusyoPay" & i + 1) = .Pay                                            '' 払戻金
                rs("PayFukusyoNinki" & i + 1) = .Ninki                                        '' 人気順
            End With ' PayFukusyo
        Next i
        For i = 0 To 2
            With .PayWakuren(i)
                rs("PayWakurenKumiban" & i + 1) = .Umaban                                      '' 馬番
                rs("PayWakurenPay" & i + 1) = .Pay                                            '' 払戻金
                rs("PayWakurenNinki" & i + 1) = .Ninki                                        '' 人気順
            End With ' PayWakuren
        Next i
        For i = 0 To 2
            With .PayUmaren(i)
                rs("PayUmarenKumiban" & i + 1) = .Kumi                                         '' 組番
                rs("PayUmarenPay" & i + 1) = .Pay                                             '' 払戻金
                rs("PayUmarenNinki" & i + 1) = .Ninki                                         '' 人気順
            End With ' PayUmaren
        Next i
        For i = 0 To 6
            With .PayWide(i)
                rs("PayWideKumiban" & i + 1) = .Kumi                                           '' 組番
                rs("PayWidePay" & i + 1) = .Pay                                               '' 払戻金
                rs("PayWideNinki" & i + 1) = .Ninki                                           '' 人気順
            End With ' PayWide
        Next i
        For i = 0 To 2
            With .PayReserved1(i)
                rs("PayReserved1Kumiban" & i + 1) = .Kumi                                      '' 組番
                rs("PayReserved1Pay" & i + 1) = .Pay                                          '' 払戻金
                rs("PayReserved1Ninki" & i + 1) = .Ninki                                      '' 人気順
            End With ' PayReserved1
        Next i
        For i = 0 To 5
            With .PayUmatan(i)
                rs("PayUmatanKumiban" & i + 1) = .Kumi                                         '' 組番
                rs("PayUmatanPay" & i + 1) = .Pay                                             '' 払戻金
                rs("PayUmatanNinki" & i + 1) = .Ninki                                         '' 人気順
            End With ' PayUmatan
        Next i
        For i = 0 To 2
            With .PaySanrenpuku(i)
                rs("PaySanrenpukuKumiban" & i + 1) = .Kumi                                     '' 組番
                rs("PaySanrenpukuPay" & i + 1) = .Pay                                         '' 払戻金
                rs("PaySanrenpukuNinki" & i + 1) = .Ninki                                     '' 人気順
            End With ' PaySanrenpuku
        Next i
        For i = 0 To 5
            With .PaySanrentan(i)
                rs("PaySanrentanKumiban" & i + 1) = .Kumi                                      '' 組番
                rs("PaySanrentanPay" & i + 1) = .Pay                                          '' 払戻金
                rs("PaySanrentanNinki" & i + 1) = .Ninki                                      '' 人気順
            End With ' PaySanrentan
        Next i
    End With
    
    rs.Update ' HARAI
    
    mCon_HARAI.CommitTrans

    Set rs = Nothing
    
    InsertDB = True
    Exit Function
    
ErrorHandler:
    If Err.Number <> -2147217887 Then
        gApp.ErrLog
    End If
    rs.CancelUpdate
    
    mCon_HARAI.RollbackTrans

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
    Dim i As Integer                                                                          '' ループカウンタ
    Dim j As Integer                                                                          '' ループカウンタ
    Dim k As Integer                                                                          '' ループカウンタ
    Dim strSQL As String                                                                      '' SQL文
    
    Set rs = mRS_HARAI

    With mBuf.id
        rs.Seek Array(.Year, .MonthDay, .JyoCD, .Kaiji, .Nichiji, .RaceNum)
    End With

    With mBuf.head.MakeDate
        If rs("Makedate") > .Year & .Month & .Day Then
            UpdateDB = True
            Exit Function
        End If
    End With

    
    mCon_HARAI.BeginTrans

    
    With mBuf
        With .head
            rs("RecordSpec") = .RecordSpec                                                    '' レコード種別
            rs("DataKubun") = .DataKubun                                                      '' データ区分
            With .MakeDate
                rs("MakeDate") = .Year & .Month & .Day                                        '' 年月日
            End With ' MakeDate
        End With ' head
        rs("TorokuTosu") = .TorokuTosu                                                        '' 登録頭数
        rs("SyussoTosu") = .SyussoTosu                                                        '' 出走頭数
        For i = 0 To 8
            rs("FuseirituFlag" & i + 1) = .FuseirituFlag(i)                                   '' 不成立フラグ
        Next i
        For i = 0 To 8
            rs("TokubaraiFlag" & i + 1) = .TokubaraiFlag(i)                                   '' 特払フラグ
        Next i
        For i = 0 To 8
            rs("HenkanFlag" & i + 1) = .HenkanFlag(i)                                         '' 返還フラグ
        Next i
        For i = 0 To 27
            rs("HenkanUma" & i + 1) = .HenkanUma(i)                                           '' 返還馬番情報(馬番01～28)
        Next i
        For i = 0 To 7
            rs("HenkanWaku" & i + 1) = .HenkanWaku(i)                                         '' 返還枠番情報(枠番1～8)
        Next i
        For i = 0 To 7
            rs("HenkanDoWaku" & i + 1) = .HenkanDoWaku(i)                                     '' 返還同枠情報(枠番1～8)
        Next i
        For i = 0 To 2
            With .PayTansyo(i)
                rs("PayTansyoUmaban" & i + 1) = .Umaban                                       '' 馬番
                rs("PayTansyoPay" & i + 1) = .Pay                                             '' 払戻金
                rs("PayTansyoNinki" & i + 1) = .Ninki                                         '' 人気順
            End With ' PayTansyo
        Next i
        For i = 0 To 4
            With .PayFukusyo(i)
                rs("PayFukusyoUmaban" & i + 1) = .Umaban                                      '' 馬番
                rs("PayFukusyoPay" & i + 1) = .Pay                                            '' 払戻金
                rs("PayFukusyoNinki" & i + 1) = .Ninki                                        '' 人気順
            End With ' PayFukusyo
        Next i
        For i = 0 To 2
            With .PayWakuren(i)
                rs("PayWakurenKumiban" & i + 1) = .Umaban                                      '' 馬番
                rs("PayWakurenPay" & i + 1) = .Pay                                            '' 払戻金
                rs("PayWakurenNinki" & i + 1) = .Ninki                                        '' 人気順
            End With ' PayWakuren
        Next i
        For i = 0 To 2
            With .PayUmaren(i)
                rs("PayUmarenKumiban" & i + 1) = .Kumi                                         '' 組番
                rs("PayUmarenPay" & i + 1) = .Pay                                             '' 払戻金
                rs("PayUmarenNinki" & i + 1) = .Ninki                                         '' 人気順
            End With ' PayUmaren
        Next i
        For i = 0 To 6
            With .PayWide(i)
                rs("PayWideKumiban" & i + 1) = .Kumi                                           '' 組番
                rs("PayWidePay" & i + 1) = .Pay                                               '' 払戻金
                rs("PayWideNinki" & i + 1) = .Ninki                                           '' 人気順
            End With ' PayWide
        Next i
        For i = 0 To 2
            With .PayReserved1(i)
                rs("PayReserved1Kumiban" & i + 1) = .Kumi                                      '' 組番
                rs("PayReserved1Pay" & i + 1) = .Pay                                          '' 払戻金
                rs("PayReserved1Ninki" & i + 1) = .Ninki                                      '' 人気順
            End With ' PayReserved1
        Next i
        For i = 0 To 5
            With .PayUmatan(i)
                rs("PayUmatanKumiban" & i + 1) = .Kumi                                         '' 組番
                rs("PayUmatanPay" & i + 1) = .Pay                                             '' 払戻金
                rs("PayUmatanNinki" & i + 1) = .Ninki                                         '' 人気順
            End With ' PayUmatan
        Next i
        For i = 0 To 2
            With .PaySanrenpuku(i)
                rs("PaySanrenpukuKumiban" & i + 1) = .Kumi                                     '' 組番
                rs("PaySanrenpukuPay" & i + 1) = .Pay                                         '' 払戻金
                rs("PaySanrenpukuNinki" & i + 1) = .Ninki                                     '' 人気順
            End With ' PaySanrenpuku
        Next i
        For i = 0 To 5
            With .PaySanrentan(i)
                rs("PaySanrentanKumiban" & i + 1) = .Kumi                                      '' 組番
                rs("PaySanrentanPay" & i + 1) = .Pay                                          '' 払戻金
                rs("PaySanrentanNinki" & i + 1) = .Ninki                                      '' 人気順
            End With ' PaySanrentan
        Next i
    End With
    
    rs.Update ' HARAI
    
    
    mCon_HARAI.CommitTrans

    Set rs = Nothing
    
    UpdateDB = True
    Exit Function
    
ErrorHandler:
    gApp.ErrLog
    
    mCon_HARAI.RollbackTrans

    UpdateDB = False
End Function
