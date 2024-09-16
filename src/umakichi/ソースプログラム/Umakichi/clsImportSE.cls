VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsImportSE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Import JV_SE_RACE_UMA "
'
'   JVData "SE" データベース登録クラス
'
'   馬毎レース情報は、データ容量が多いため、MDBを10分割します。
'

Option Explicit
Option Compare Binary
Implements clsIImport

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部変数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private mBuf As JV_SE_RACE_UMA                      '' SE構造体

Private mCon_UMA_RACE_A As ADODB.Connection   '' ADOコネクション
Private mCon_UMA_RACE_B As ADODB.Connection   '' ADOコネクション

Private mRS_UMA_RACE_A As ADODB.Recordset     '' ADOレコードセット
Private mRS_UMA_RACE_B As ADODB.Recordset     '' ADOレコードセット


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
    Set mCon_UMA_RACE_A = New ADODB.Connection
    Set mCon_UMA_RACE_B = New ADODB.Connection

    ' レコードセットのインスタンス生成
    Set mRS_UMA_RACE_A = New ADODB.Recordset
    Set mRS_UMA_RACE_B = New ADODB.Recordset
    
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
    mCon_UMA_RACE_A.Open strCon & "subUMA_RACE_A.mdb"
    mCon_UMA_RACE_B.Open strCon & "subUMA_RACE_B.mdb"

    ' レコードセットオープン
    With mRS_UMA_RACE_A
        .CursorLocation = adUseServer
        .Index = "PrimaryKey"
        .Open "UMA_RACE_A", mCon_UMA_RACE_A, adOpenKeyset, adLockOptimistic, adCmdTableDirect
        If Not (.EOF Or .BOF) Then
            .MoveFirst
        End If
    End With

    With mRS_UMA_RACE_B
        .CursorLocation = adUseServer
        .Index = "PrimaryKey"
        .Open "UMA_RACE_B", mCon_UMA_RACE_B, adOpenKeyset, adLockOptimistic, adCmdTableDirect
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
    
    mRS_UMA_RACE_B.Close
    mRS_UMA_RACE_A.Close
    mCon_UMA_RACE_A.Close
    mCon_UMA_RACE_B.Close

    Exit Sub
ErrorHandler:
   gApp.ErrLog
   Resume Next
End Sub


'
'   機能: JVReadの返す１行をデータベースに登録する
'
'   備考: DBに追加を試みる
'
Private Function clsIImport_Add(lBuf() As Byte) As Boolean
On Error GoTo ErrorHandler

    If lBuf(2) = ASCII_ZERO Then
        Call DeleteRecord(StrConv(lBuf, vbUnicode))
        clsIImport_Add = True
    Else

        Call SetDataFromByte_SE(lBuf, mBuf)                                                     '' 構造体に代入する

        clsIImport_Add = ImportDB()
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
    Dim Umaban      As String
    Dim KettoNum    As String
    
    mCon_UMA_RACE_A.BeginTrans
    mCon_UMA_RACE_B.BeginTrans
    
    Year = Mid$(lBuf, 12, 4)                    '' 開催年
    MonthDay = Mid$(lBuf, 16, 4)                '' 開催月日
    JyoCD = Mid$(lBuf, 20, 2)                   '' 競馬場コード
    Kaiji = Mid$(lBuf, 22, 2)                   '' 開催回[第N回]
    Nichiji = Mid$(lBuf, 24, 2)                 '' 開催日目[N日目]
    RaceNum = Mid$(lBuf, 26, 2)                 '' レース番号
    Umaban = Mid$(lBuf, 29, 2)                  '' 馬番
    KettoNum = Mid$(lBuf, 31, 10)               '' 血統登録番号

    strSQL = "DELETE * FROM UMA_RACE_A"
    strSQL = strSQL & " WHERE [Year]   ='" & Year & "'"
    strSQL = strSQL & " AND [MonthDay] ='" & MonthDay & "'"
    strSQL = strSQL & " AND [JyoCD]    ='" & JyoCD & "'"
    strSQL = strSQL & " AND [Kaiji]    ='" & Kaiji & "'"
    strSQL = strSQL & " AND [Nichiji]  ='" & Nichiji & "'"
    strSQL = strSQL & " AND [RaceNum]  ='" & RaceNum & "'"
    strSQL = strSQL & " AND [Umaban]   ='" & Umaban & "'"
    strSQL = strSQL & " AND [KettoNum] ='" & KettoNum & "'"
    
    mCon_UMA_RACE_A.Execute strSQL, , adExecuteNoRecords
    
    strSQL = "DELETE * FROM UMA_RACE_B"
    strSQL = strSQL & " WHERE [B_Year]   ='" & Year & "'"
    strSQL = strSQL & " AND [B_MonthDay] ='" & MonthDay & "'"
    strSQL = strSQL & " AND [B_JyoCD]    ='" & JyoCD & "'"
    strSQL = strSQL & " AND [B_RaceNum]  ='" & RaceNum & "'"
    strSQL = strSQL & " AND [B_Umaban]   ='" & Umaban & "'"
    strSQL = strSQL & " AND [B_KettoNum] ='" & KettoNum & "'"
    
    mCon_UMA_RACE_B.Execute strSQL, , adExecuteNoRecords
    
    mCon_UMA_RACE_A.CommitTrans
    mCon_UMA_RACE_B.CommitTrans
        
    Exit Sub
ErrorHandler:
    mCon_UMA_RACE_A.RollbackTrans
    mCon_UMA_RACE_B.RollbackTrans
End Sub


'
'   機能: データベースに追加する
'
'   備考: なし
'
Private Function ImportDB() As Boolean
On Error GoTo ErrorHandler
    Dim i As Integer                                                                '' ループカウンタ
    Dim j As Integer                                                                '' ループカウンタ
    Dim k As Integer                                                                '' ループカウンタ
    Dim rs As ADODB.Recordset                                                       '' SQL文
    
    mCon_UMA_RACE_A.BeginTrans
    mCon_UMA_RACE_B.BeginTrans

    Set rs = mRS_UMA_RACE_A
    
    With mBuf.id
        Call rs.Seek(Array(.Year, .MonthDay, .JyoCD, .Kaiji, .Nichiji, .RaceNum, mBuf.Umaban, mBuf.KettoNum))
    End With
    
    If Not rs.EOF Then
        ' 既存のレコードがあれば、データ作成日を比較
        With mBuf.head.MakeDate
            If rs("Makedate") > .Year & .Month & .Day Then
                ' 古いデータは上書きしない
                ImportDB = True
                mCon_UMA_RACE_A.RollbackTrans
                mCon_UMA_RACE_B.RollbackTrans
                Exit Function
            End If
        End With
    Else
        ' 既存のレコードがなければ
        
        ' SE のみの事情により、 馬番が"0"の同レースがあれば、上書き
        With mBuf.id
            Call rs.Seek(Array(.Year, .MonthDay, .JyoCD, .Kaiji, .Nichiji, .RaceNum, "00", mBuf.KettoNum))
        End With
        If Not rs.EOF Then
            ' 既存のレコードがあれば、データ作成日を比較
            With mBuf.head.MakeDate
                If rs("Makedate") > .Year & .Month & .Day Then
                    ' 古いデータは上書きしない
                    ImportDB = True
                    mCon_UMA_RACE_A.RollbackTrans
                    mCon_UMA_RACE_B.RollbackTrans
                    Exit Function
                End If
            End With
        Else
            rs.AddNew
            With mBuf.id
                rs("Year") = .Year
                rs("MonthDay") = .MonthDay
                rs("JyoCD") = .JyoCD
                rs("Kaiji") = .Kaiji
                rs("Nichiji") = .Nichiji
                rs("RaceNum") = .RaceNum
            End With
            With mBuf
                rs("Umaban") = .Umaban
                rs("KettoNum") = .KettoNum
            End With
        End If
    End If
    
    With mBuf
        With .head
            rs("RecordSpec") = .RecordSpec                                          '' レコード種別
            rs("DataKubun") = .DataKubun                                            '' データ区分
            With .MakeDate
                rs("MakeDate") = .Year & .Month & .Day                              '' 年月日
            End With ' MakeDate
        End With ' head
        rs("Wakuban") = .Wakuban                                                    '' 枠番
        rs("Umaban") = .Umaban
        rs("Bamei") = .BAMEI                                                        '' 馬名
        rs("UmaKigoCD") = .UmaKigoCD                                                '' 馬記号コード
        rs("SexCD") = .SexCD                                                        '' 性別コード
        rs("HinsyuCD") = .HinsyuCD                                                  '' 品種コード
        rs("KeiroCD") = .KeiroCD                                                    '' 毛色コード
        rs("Barei") = .Barei                                                        '' 馬齢
        rs("TozaiCD") = .TozaiCD                                                    '' 東西所属コード
        rs("ChokyosiCode") = .ChokyosiCode                                          '' 調教師コード
        rs("ChokyosiRyakusyo") = .ChokyosiRyakusyo                                  '' 調教師名略称
        rs("BanusiCode") = .BanusiCode                                              '' 馬主コード
        rs("BanusiName") = .BanusiName                                              '' 馬主名
        rs("Fukusyoku") = .Fukusyoku                                                '' 服色標示
        rs("reserved1") = .reserved1                                                '' 予備
    End With ' mBuf
    
    rs.Update ' UMA_RACE
    
    Set rs = mRS_UMA_RACE_B
    
    With mBuf.id
        Call rs.Seek(Array(.Year, .MonthDay, .JyoCD, .RaceNum, mBuf.Umaban, mBuf.KettoNum))
    End With
    
    If rs.EOF Then
        ' 既存のレコードがなければ
        
        ' SE のみの事情により、 馬番が"0"の同レースがあれば、上書き
        With mBuf.id
            Call rs.Seek(Array(.Year, .MonthDay, .JyoCD, .RaceNum, "0", mBuf.KettoNum))
        End With
        If rs.EOF Then
            rs.AddNew
            With mBuf.id
                rs("B_Year") = .Year
                rs("B_MonthDay") = .MonthDay
                rs("B_JyoCD") = .JyoCD
                rs("B_RaceNum") = .RaceNum
            End With
            With mBuf
                rs("B_Umaban") = .Umaban
                rs("B_KettoNum") = .KettoNum
            End With
        End If
    End If

    With mBuf
        rs("DataKubun") = .head.DataKubun                                           '' データ区分
        rs("B_Umaban") = .Umaban
        rs("Futan") = .Futan                                                        '' 負担重量
        rs("FutanBefore") = .FutanBefore                                            '' 変更前負担重量
        rs("Blinker") = .BLINKER                                                    '' ブリンカー使用区分
        rs("reserved2") = .reserved2                                                '' 予備
        rs("KisyuCode") = .KisyuCode                                                '' 騎手コード
        rs("KisyuCodeBefore") = .KisyuCodeBefore                                    '' 変更前騎手コード
        rs("KisyuRyakusyo") = .KisyuRyakusyo                                        '' 騎手名略称
        rs("KisyuRyakusyoBefore") = .KisyuRyakusyoBefore                            '' 変更前騎手名略称
        rs("MinaraiCD") = .MinaraiCD                                                '' 騎手見習コード
        rs("MinaraiCDBefore") = .MinaraiCDBefore                                    '' 変更前騎手見習コード
        rs("BaTaijyu") = .BaTaijyu                                                  '' 馬体重
        rs("ZogenFugo") = .ZogenFugo                                                '' 増減符号
        rs("ZogenSa") = .ZogenSa                                                    '' 増減差
        rs("IJyoCD") = .IJyoCD                                                      '' 異常区分コード
        rs("NyusenJyuni") = .NyusenJyuni                                            '' 入線順位
        rs("KakuteiJyuni") = .KakuteiJyuni                                          '' 確定着順
        rs("DochakuKubun") = .DochakuKubun                                          '' 同着区分
        rs("DochakuTosu") = .DochakuTosu                                            '' 同着頭数
        rs("Time") = .TIME                                                          '' 走破タイム
        rs("ChakusaCD") = .ChakusaCD                                                '' 着差コード
        rs("ChakusaCDP") = .ChakusaCDP                                              '' +着差コード
        rs("ChakusaCDPP") = .ChakusaCDPP                                            '' ++着差コード
        rs("Jyuni1c") = .Jyuni1c                                                    '' 1コーナーでの順位
        rs("Jyuni2c") = .Jyuni2c                                                    '' 2コーナーでの順位
        rs("Jyuni3c") = .Jyuni3c                                                    '' 3コーナーでの順位
        rs("Jyuni4c") = .Jyuni4c                                                    '' 4コーナーでの順位
        rs("Odds") = .Odds                                                          '' 単勝オッズ
        rs("Ninki") = .Ninki                                                        '' 単勝人気順
        rs("Honsyokin") = .Honsyokin                                                '' 獲得本賞金
        rs("Fukasyokin") = .Fukasyokin                                              '' 獲得付加賞金
        rs("reserved3") = .reserved3                                                '' 予備
        rs("reserved4") = .reserved4                                                '' 予備
        rs("HaronTimeL4") = .HaronTimeL4                                            '' 後４ハロンタイム
        rs("HaronTimeL3") = .HaronTimeL3                                            '' 後３ハロンタイム
        With .ChakuUmaInfo(0)
            rs("KettoNum1") = .KettoNum                                             '' 血統登録番号（相手馬1）
            rs("Bamei1") = .BAMEI                                                   '' 馬名
        End With
        With .ChakuUmaInfo(1)
            rs("KettoNum2") = .KettoNum                                             '' 血統登録番号（相手馬2）
            rs("Bamei2") = .BAMEI                                                   '' 馬名
        End With
        With .ChakuUmaInfo(2)
            rs("KettoNum3") = .KettoNum                                             '' 血統登録番号（相手馬3）
            rs("Bamei3") = .BAMEI                                                   '' 馬名
        End With
        rs("TimeDiff") = .TimeDiff                                                  '' タイム差
        rs("RecordUpKubun") = .RecordUpKubun                                        '' レコード更新区分
        rs("DMKubun") = .DMKubun                                                    '' マイニング区分
        rs("DMTime") = .DMTime                                                      '' マイニング予想走破タイム
        rs("DMGosaP") = .DMGosaP                                                    '' 予測誤差(信頼度)＋
        rs("DMGosaM") = .DMGosaM                                                    '' 予測誤差(信頼度)－
        rs("DMJyuni") = .DMJyuni                                                    '' マイニング予想順位
        rs("KyakusituKubun") = .KyakusituKubun                                      '' 今回レース脚質判定
    End With
    
    rs.Update ' UMA_RACE
    
    mCon_UMA_RACE_A.CommitTrans
    mCon_UMA_RACE_B.CommitTrans

    Set rs = Nothing
    
    ImportDB = True
    Exit Function
    
ErrorHandler:
    gApp.ErrLog
    rs.CancelUpdate

    mCon_UMA_RACE_A.RollbackTrans
    mCon_UMA_RACE_B.RollbackTrans

    ImportDB = False
End Function

