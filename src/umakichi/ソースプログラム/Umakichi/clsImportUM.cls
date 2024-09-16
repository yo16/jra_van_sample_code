VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsImportUM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Import JV_UM_UMA "
'
'   JVData "UM" データベース登録クラス
'

Option Explicit
Option Compare Binary
Implements clsIImport


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部変数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private mBuf As JV_UM_UMA

Private mCon_UMA As ADODB.Connection

Private mRS_UMA As ADODB.Recordset


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
    Set mCon_UMA = New ADODB.Connection

    ' レコードセットのインスタンス生成
    Set mRS_UMA = New ADODB.Recordset
    
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
    mCon_UMA.Open strCon & "subUMA.mdb"

    ' レコードセットオープン
    With mRS_UMA
        .CursorLocation = adUseServer
        .Index = "PrimaryKey"
        .Open "UMA", mCon_UMA, adOpenKeyset, adLockOptimistic, adCmdTableDirect
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
    mRS_UMA.Close

    mCon_UMA.Close

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
        Call SetDataFromByte_UM(lBuf, mBuf)     '' 構造体に代入する

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
    
    mCon_UMA.BeginTrans
    
    KettoNum = Mid$(lBuf, 12, 10)                  ''

    strSQL = "DELETE * FROM UMA"
    strSQL = strSQL & " WHERE [KettoNum]   ='" & KettoNum & "'"
    
    mCon_UMA.Execute strSQL, , adExecuteNoRecords
    
    mCon_UMA.CommitTrans
        
    Exit Sub
ErrorHandler:
    mCon_UMA.RollbackTrans
End Sub


'
'   機能: データベースに挿入する
'
'   備考: なし
'
Private Function InsertDB() As Boolean
On Error GoTo ErrorHandler
    Dim i As Integer                                                                '' ループカウンタ
    Dim j As Integer                                                                '' ループカウンタ
    Dim k As Integer                                                                '' ループカウンタ
    Dim rs As ADODB.Recordset                                                       '' SQL文
    
    
    mCon_UMA.BeginTrans

    
    Set rs = mRS_UMA
    rs.AddNew Array("KettoNum"), Array(mBuf.KettoNum)
    
    With mBuf
        With .head
            rs("RecordSpec") = .RecordSpec                                          '' レコード種別
            rs("DataKubun") = .DataKubun                                            '' データ区分
            With .MakeDate
                rs("MakeDate") = .Year & .Month & .Day                              '' 年月日
            End With ' MakeDate
        End With ' head
        rs("KettoNum") = .KettoNum                                                  '' 血統登録番号
        rs("DelKubun") = .DelKubun                                                  '' 競走馬抹消区分
        With .RegDate
            rs("RegDate") = .Year & .Month & .Day                                   '' 年月日
        End With ' RegDate
        With .DelDate
            rs("DelDate") = .Year & .Month & .Day                                   '' 年月日
        End With ' DelDate
        With .BirthDate
            rs("BirthDate") = .Year & .Month & .Day                                 '' 年月日
        End With ' BirthDate
        rs("Bamei") = .BAMEI                                                        '' 馬名
        rs("BameiKana") = .BameiKana                                                '' 馬名半角カナ
        rs("BameiEng") = .BameiEng                                                  '' 馬名欧字
        rs("UmaKigoCD") = .UmaKigoCD                                                '' 馬記号コード
        rs("SexCD") = .SexCD                                                        '' 性別コード
        rs("HinsyuCD") = .HinsyuCD                                                  '' 品種コード
        rs("KeiroCD") = .KeiroCD                                                    '' 毛色コード
        For i = 0 To 13
            With .Ketto3Info(i)
                rs("Ketto3InfoHansyokuNum" & i + 1) = .HansyokuNum                  '' 繁殖登録番号
                rs("Ketto3InfoBamei" & i + 1) = .BAMEI                              '' 馬名
            End With ' Ketto3Info
        Next i
        rs("TozaiCD") = .TozaiCD                                                    '' 東西所属コード
        rs("ChokyosiCode") = .ChokyosiCode                                          '' 調教師コード
        rs("ChokyosiRyakusyo") = .ChokyosiRyakusyo                                  '' 調教師名略称
        rs("Syotai") = .Syotai                                                      '' 招待地域名
        rs("BreederCode") = .BreederCode                                            '' 生産者コード
        rs("BreederName") = .BreederName                                            '' 生産者名
        rs("SanchiName") = .SanchiName                                              '' 産地名
        rs("BanusiCode") = .BanusiCode                                              '' 馬主コード
        rs("BanusiName") = .BanusiName                                              '' 馬主名
        rs("RuikeiHonsyoHeiti") = .RuikeiHonsyoHeiti                                '' 平地本賞金累計
        rs("RuikeiHonsyoSyogai") = .RuikeiHonsyoSyogai                              '' 障害本賞金累計
        rs("RuikeiFukaHeichi") = .RuikeiFukaHeichi                                  '' 平地付加賞金累計
        rs("RuikeiFukaSyogai") = .RuikeiFukaSyogai                                  '' 障害付加賞金累計
        rs("RuikeiSyutokuHeichi") = .RuikeiSyutokuHeichi                            '' 平地収得賞金累計
        rs("RuikeiSyutokuSyogai") = .RuikeiSyutokuSyogai                            '' 障害収得賞金累計
        With .ChakuSogo
            For i = 0 To 5
                rs("SogoChakukaisu" & i + 1) = .Chakukaisu(i)
            Next i
        End With ' ChakuSogo
        With .ChakuChuo
            For i = 0 To 5
                rs("ChuoChakukaisu" & i + 1) = .Chakukaisu(i)
            Next i
        End With ' ChakuChuo
        For i = 0 To 6
            With .ChakuKaisuBa(i)
                For j = 0 To 5
                    rs("Ba" & i + 1 & "Chakukaisu" & j + 1) = .Chakukaisu(j)
                Next j
            End With ' ChakuKaisuBa
        Next i
        For i = 0 To 11
            With .ChakuKaisuJyotai(i)
                For j = 0 To 5
                    rs("Jyotai" & i + 1 & "Chakukaisu" & j + 1) = .Chakukaisu(j)
                Next j
            End With ' ChakuKaisuJyotai
        Next i
        For i = 0 To 5
            With .ChakuKaisuKyori(i)
                For j = 0 To 5
                    rs("Kyori" & i + 1 & "Chakukaisu" & j + 1) = .Chakukaisu(j)
                Next j
            End With ' ChakuKaisuKyori
        Next i
        For i = 0 To 3
            rs("Kyakusitu" & i + 1) = .Kyakusitu(i)                                 '' 脚質傾向
        Next i
        rs("RaceCount") = .RaceCount                                                '' 登録レース数
    End With
    
    rs.Update ' UMA
    
    
    mCon_UMA.CommitTrans

    Set rs = Nothing
    
    InsertDB = True
    Exit Function
    
ErrorHandler:
    If Err.Number <> -2147217887 Then
        gApp.ErrLog
    End If
    rs.CancelUpdate
    
    mCon_UMA.RollbackTrans

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
    Dim i As Integer                                                                '' ループカウンタ
    Dim j As Integer                                                                '' ループカウンタ
    Dim k As Integer                                                                '' ループカウンタ
    Dim strSQL As String                                                            '' SQL文
    
    Set rs = mRS_UMA

    rs.Seek Array(mBuf.KettoNum)

    With mBuf.head.MakeDate
        If rs("Makedate") > .Year & .Month & .Day Then
            UpdateDB = True
            Exit Function
        End If
    End With

    
    mCon_UMA.BeginTrans


    With mBuf
        With .head
            rs("RecordSpec") = .RecordSpec                                          '' レコード種別
            rs("DataKubun") = .DataKubun                                            '' データ区分
            With .MakeDate
                rs("MakeDate") = .Year & .Month & .Day                              '' 年月日
            End With ' MakeDate
        End With ' head
        rs("KettoNum") = .KettoNum                                                  '' 血統登録番号
        rs("DelKubun") = .DelKubun                                                  '' 競走馬抹消区分
        With .RegDate
            rs("RegDate") = .Year & .Month & .Day                                   '' 年月日
        End With ' RegDate
        With .DelDate
            rs("DelDate") = .Year & .Month & .Day                                   '' 年月日
        End With ' DelDate
        With .BirthDate
            rs("BirthDate") = .Year & .Month & .Day                                 '' 年月日
        End With ' BirthDate
        rs("Bamei") = .BAMEI                                                        '' 馬名
        rs("BameiKana") = .BameiKana                                                '' 馬名半角カナ
        rs("BameiEng") = .BameiEng                                                  '' 馬名欧字
        rs("UmaKigoCD") = .UmaKigoCD                                                '' 馬記号コード
        rs("SexCD") = .SexCD                                                        '' 性別コード
        rs("HinsyuCD") = .HinsyuCD                                                  '' 品種コード
        rs("KeiroCD") = .KeiroCD                                                    '' 毛色コード
        For i = 0 To 13
            With .Ketto3Info(i)
                rs("Ketto3InfoHansyokuNum" & i + 1) = .HansyokuNum                  '' 繁殖登録番号
                rs("Ketto3InfoBamei" & i + 1) = .BAMEI                              '' 馬名
            End With ' Ketto3Info
        Next i
        rs("TozaiCD") = .TozaiCD                                                    '' 東西所属コード
        rs("ChokyosiCode") = .ChokyosiCode                                          '' 調教師コード
        rs("ChokyosiRyakusyo") = .ChokyosiRyakusyo                                  '' 調教師名略称
        rs("Syotai") = .Syotai                                                      '' 招待地域名
        rs("BreederCode") = .BreederCode                                            '' 生産者コード
        rs("BreederName") = .BreederName                                            '' 生産者名
        rs("SanchiName") = .SanchiName                                              '' 産地名
        rs("BanusiCode") = .BanusiCode                                              '' 馬主コード
        rs("BanusiName") = .BanusiName                                              '' 馬主名
        rs("RuikeiHonsyoHeiti") = .RuikeiHonsyoHeiti                                '' 平地本賞金累計
        rs("RuikeiHonsyoSyogai") = .RuikeiHonsyoSyogai                              '' 障害本賞金累計
        rs("RuikeiFukaHeichi") = .RuikeiFukaHeichi                                  '' 平地付加賞金累計
        rs("RuikeiFukaSyogai") = .RuikeiFukaSyogai                                  '' 障害付加賞金累計
        rs("RuikeiSyutokuHeichi") = .RuikeiSyutokuHeichi                            '' 平地収得賞金累計
        rs("RuikeiSyutokuSyogai") = .RuikeiSyutokuSyogai                            '' 障害収得賞金累計
        With .ChakuSogo
            For i = 0 To 5
                rs("SogoChakukaisu" & i + 1) = .Chakukaisu(i)
            Next i
        End With ' ChakuSogo
        With .ChakuChuo
            For i = 0 To 5
                rs("ChuoChakukaisu" & i + 1) = .Chakukaisu(i)
            Next i
        End With ' ChakuChuo
        For i = 0 To 6
            With .ChakuKaisuBa(i)
                For j = 0 To 5
                    rs("Ba" & i + 1 & "Chakukaisu" & j + 1) = .Chakukaisu(j)
                Next j
            End With ' ChakuKaisuBa
        Next i
        For i = 0 To 11
            With .ChakuKaisuJyotai(i)
                For j = 0 To 5
                    rs("Jyotai" & i + 1 & "Chakukaisu" & j + 1) = .Chakukaisu(j)
                Next j
            End With ' ChakuKaisuJyotai
        Next i
        For i = 0 To 5
            With .ChakuKaisuKyori(i)
                For j = 0 To 5
                    rs("Kyori" & i + 1 & "Chakukaisu" & j + 1) = .Chakukaisu(j)
                Next j
            End With ' ChakuKaisuKyori
        Next i
        For i = 0 To 3
            rs("Kyakusitu" & i + 1) = .Kyakusitu(i)                                 '' 脚質傾向
        Next i
        rs("RaceCount") = .RaceCount                                                '' 登録レース数
    End With
    
    rs.Update
    
    
    mCon_UMA.CommitTrans

    Set rs = Nothing
    
    UpdateDB = True
    Exit Function
    
ErrorHandler:
    rs.CancelUpdate
    
    mCon_UMA.RollbackTrans

    UpdateDB = False
End Function
