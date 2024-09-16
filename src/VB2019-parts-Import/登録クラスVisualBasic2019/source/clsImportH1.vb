Option Strict Off
Option Explicit On
Option Compare Binary
Friend Class clsImportH1
	' @(h) clsReadH1.cls
	' @(s)
	' JVData "H1" データベース登録クラス
	'
	
	Private mBuf As JV_H1_HYOSU_ZENKAKE '' 票数（全賭式）構造体
	Private mRS1 As ADODB.Recordset
	Private mRS2 As ADODB.Recordset
	Private mRS3 As ADODB.Recordset
	Private mRS4 As ADODB.Recordset
	Private mRS5 As ADODB.Recordset
	Private mRS6 As ADODB.Recordset
	
	
	' @(f)
	'
	' 機能      : 初期処理
	'
	' 引き数    :
	'
	' 返り値    :
	'
	' 機能説明  :
	'

    Private Sub Class_Initialize_Renamed()
        On Error GoTo ErrorHandler

        Dim strSql As String ''SQL文

        'レコードセットオープン
        strSql = "SELECT * FROM HYOSU"
        mRS1 = New ADODB.Recordset()
        mRS1.Open(strSql, gCon, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockPessimistic)

        strSql = "SELECT * FROM HYOSU_TANPUKU"
        mRS2 = New ADODB.Recordset()
        mRS2.Open(strSql, gCon, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockPessimistic)

        strSql = "SELECT * FROM HYOSU_WAKU"
        mRS3 = New ADODB.Recordset()
        mRS3.Open(strSql, gCon, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockPessimistic)

        strSql = "SELECT * FROM HYOSU_UMARENWIDE"
        mRS4 = New ADODB.Recordset()
        mRS4.Open(strSql, gCon, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockPessimistic)

        strSql = "SELECT * FROM HYOSU_UMATAN"
        mRS5 = New ADODB.Recordset()
        mRS5.Open(strSql, gCon, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockPessimistic)

        strSql = "SELECT * FROM HYOSU_SANREN"
        mRS6 = New ADODB.Recordset()
        mRS6.Open(strSql, gCon, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockPessimistic)

ExitHandler:
        Exit Sub
ErrorHandler:
        System.Diagnostics.Debug.WriteLine(Err.Description)
        Resume ExitHandler
    End Sub
    Public Sub New()
        MyBase.New()
        Class_Initialize_Renamed()
    End Sub

    ' @(f)
    '
    ' 機能      : Closeのコーディング
    '
    ' 機能説明  : ガーベッジコレクションにCloseを呼ばせるとどこで呼ばれるか分からない為、
    '           　明示的に呼び出す必要がある。
    '
    Public Sub Close()
        'レコードセットクローズ
        mRS1.Close()

        mRS1 = Nothing
        mRS2.Close()

        mRS2 = Nothing
        mRS3.Close()

        mRS3 = Nothing
        mRS4.Close()

        mRS4 = Nothing
        mRS5.Close()

        mRS5 = Nothing
        mRS6.Close()

        mRS6 = Nothing


    End Sub

    ' @(f)
    '
    ' 機能      : 終了処理
    '
    ' 引き数    :
    '
    ' 返り値    :
    '
    ' 機能説明  : レコードセットクローズ
    '

    Private Sub Class_Terminate_Renamed()
        On Error GoTo ErrorHandler


ExitHandler:
        Exit Sub
ErrorHandler:
        System.Diagnostics.Debug.WriteLine(Err.Description)
        Resume ExitHandler
    End Sub
    Protected Overrides Sub Finalize()
        Class_Terminate_Renamed()
        MyBase.Finalize()
    End Sub


    ' @(f)
    '
    ' 機能      : Addプロシージャを呼ぶ
    '
    ' 引き数    : lBuf - JVData 識別子"H1" の１行
    '
    ' 返り値    : True - 成功, False - 失敗
    '
    ' 機能説明  : clsIReadインターフェイスAddプロシージャの実装
    '
    Public Function Add(ByRef strBuf As String, ByVal lngBufSize As Integer) As Boolean
        On Error GoTo ErrorHandler

        Dim strMakeDate As String '' 登録するデータの作成年月日

        '構造体にデータセット
        mBuf.SetData(strBuf)

        With mBuf.head.MakeDate
            strMakeDate = .Year & .Month & .Day
        End With

        'INSERT処理
        If Not InsertDB() Then
            'UPDATE処理（INSERTが失敗した場合）
            If Not UpdateDB(strMakeDate) Then System.Diagnostics.Debug.WriteLine("更新に失敗しました。" & Left(strBuf, 2))
        End If

        Add = True

ExitHandler:
        Exit Function
ErrorHandler:
        Add = False
        System.Diagnostics.Debug.WriteLine(Err.Description)
        Resume ExitHandler
    End Function




    ' @(f)
    '
    ' 機能      : データベースに追加する
    '
    ' 引き数    :
    '
    ' 返り値    : True - 成功, False - 失敗
    '
    ' 機能説明  :
    '
    Public Function InsertDB() As Boolean
        On Error GoTo ErrorHandler
        Dim i As Short '' ループカウンタ
        Dim j As Short '' ループカウンタ
        Dim k As Short '' ループカウンタ

        gCon.BeginTrans()
        System.Diagnostics.Debug.WriteLine("BeginTrans")

        mRS1.AddNew()

        ' HYOSU (票数)
        With mBuf
            With .head
                mRS1.Fields("RecordSpec").Value = .RecordSpec '' レコード種別
                mRS1.Fields("DataKubun").Value = .DataKubun '' データ区分
                With .MakeDate
                    mRS1.Fields("MakeDate").Value = .Year & .Month & .Day '' 年月日
                End With ' MakeDate
            End With ' head
            With .id
                mRS1.Fields("Year").Value = .Year '' 開催年
                mRS1.Fields("MonthDay").Value = .MonthDay '' 開催月日
                mRS1.Fields("JyoCD").Value = .JyoCD '' 競馬場コード
                mRS1.Fields("Kaiji").Value = .Kaiji '' 開催回第N回
                mRS1.Fields("Nichiji").Value = .Nichiji '' 開催日目N日目
                mRS1.Fields("RaceNum").Value = .RaceNum '' レース番号
            End With ' id
            mRS1.Fields("TorokuTosu").Value = .TorokuTosu '' 登録頭数
            mRS1.Fields("SyussoTosu").Value = .SyussoTosu '' 出走頭数
            For i = 0 To 6
                mRS1.Fields("HatubaiFlag" & i + 1).Value = .HatubaiFlag(i) '' 発売フラグ
            Next i
            mRS1.Fields("FukuChakuBaraiKey").Value = .FukuChakuBaraiKey '' 複勝着払キー
            For i = 0 To 27
                mRS1.Fields("HenkanUma" & i + 1).Value = .HenkanUma(i) '' 返還馬番情報(馬番01～28)
            Next i
            For i = 0 To 7
                mRS1.Fields("HenkanWaku" & i + 1).Value = .HenkanWaku(i) '' 返還枠番情報(枠番1～8)
            Next i
            For i = 0 To 7
                mRS1.Fields("HenkanDoWaku" & i + 1).Value = .HenkanDoWaku(i) '' 返還同枠情報(枠番1～8)
            Next i
            For i = 0 To 13
                mRS1.Fields("HyoTotal" & i + 1).Value = .HyoTotal(i) '' 票数合計
            Next i
        End With ' mBuf

        With mBuf.id
            System.Diagnostics.Debug.WriteLine("Insert HYOSU : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum)
        End With ' id

        mRS1.Update()


        ' HYOSU_TANPUKU (票数_単複)
            For i = 0 To 27
                If mBuf.HyoTansyo(i).Umaban <> "  " Then
                    mRS2.AddNew()
                    With mBuf
                        With .head
                            With .MakeDate
                                mRS2.Fields("MakeDate").Value = .Year & .Month & .Day '' 年月日
                            End With ' MakeDate
                        End With ' head
                        With .id
                            mRS2.Fields("Year").Value = .Year '' 開催年
                            mRS2.Fields("MonthDay").Value = .MonthDay '' 開催月日
                            mRS2.Fields("JyoCD").Value = .JyoCD '' 競馬場コード
                            mRS2.Fields("Kaiji").Value = .Kaiji '' 開催回第N回
                            mRS2.Fields("Nichiji").Value = .Nichiji '' 開催日目N日目
                            mRS2.Fields("RaceNum").Value = .RaceNum '' レース番号
                        End With ' id
                        With .HyoTansyo(i)
                            mRS2.Fields("Umaban").Value = .Umaban '' 馬番
                            mRS2.Fields("TanHyo").Value = .Hyo '' 票数
                            mRS2.Fields("TanNinki").Value = .Ninki '' 人気
                        End With ' HyoTansyo
                        With .HyoFukusyo(i)
                            mRS2.Fields("FukuHyo").Value = .Hyo '' 票数
                            mRS2.Fields("FukuNinki").Value = .Ninki '' 人気
                        End With ' HyoFukusyo
                    End With ' mBuf

                    mRS2.Update()

                    With mBuf.id
                        System.Diagnostics.Debug.WriteLine("Insert HYOSU_TANPUKU : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum & mBuf.HyoTansyo(i).Umaban)
                    End With ' id
                End If
            Next i

        ' HYOSU_WAKU (票数_枠連)
            For i = 0 To 35
                If mBuf.HyoWakuren(i).Umaban <> "  " Then
                    mRS3.AddNew()
                    With mBuf
                        With .head
                            With .MakeDate
                                mRS3.Fields("MakeDate").Value = .Year & .Month & .Day '' 年月日
                            End With ' MakeDate
                        End With ' head
                        With .id
                            mRS3.Fields("Year").Value = .Year '' 開催年
                            mRS3.Fields("MonthDay").Value = .MonthDay '' 開催月日
                            mRS3.Fields("JyoCD").Value = .JyoCD '' 競馬場コード
                            mRS3.Fields("Kaiji").Value = .Kaiji '' 開催回第N回
                            mRS3.Fields("Nichiji").Value = .Nichiji '' 開催日目N日目
                            mRS3.Fields("RaceNum").Value = .RaceNum '' レース番号
                        End With ' id
                        With .HyoWakuren(i)
                            mRS3.Fields("Kumi").Value = .Umaban '' 組番
                            mRS3.Fields("Hyo").Value = .Hyo '' 票数
                            mRS3.Fields("Ninki").Value = .Ninki '' 人気
                        End With ' HyoWakuren
                    End With ' mBuf

                    With mBuf.id
                        System.Diagnostics.Debug.WriteLine("Insert HYOSU_WAKU : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum & mBuf.HyoWakuren(i).Umaban)
                    End With ' id
                    mRS3.Update()
                End If
            Next i

        ' HYOSU_UMARENWIDE (票数_馬連・ワイド)
            For i = 0 To 152
                If mBuf.HyoUmaren(i).Kumi <> "    " Then
                    mRS4.AddNew()
                    With mBuf
                        With .head
                            With .MakeDate
                                mRS4.Fields("MakeDate").Value = .Year & .Month & .Day '' 年月日
                            End With ' MakeDate
                        End With ' head
                        With .id
                            mRS4.Fields("Year").Value = .Year '' 開催年
                            mRS4.Fields("MonthDay").Value = .MonthDay '' 開催月日
                            mRS4.Fields("JyoCD").Value = .JyoCD '' 競馬場コード
                            mRS4.Fields("Kaiji").Value = .Kaiji '' 開催回第N回
                            mRS4.Fields("Nichiji").Value = .Nichiji '' 開催日目N日目
                            mRS4.Fields("RaceNum").Value = .RaceNum '' レース番号
                        End With ' id
                        With .HyoUmaren(i)
                            mRS4.Fields("Kumi").Value = .Kumi '' 組番
                            mRS4.Fields("UmarenHyo").Value = .Hyo '' 票数
                            mRS4.Fields("UmarenNinki").Value = .Ninki '' 人気
                        End With ' HyoUmaren
                        With .HyoWide(i)
                            mRS4.Fields("WideHyo").Value = .Hyo '' 票数
                            mRS4.Fields("WideNinki").Value = .Ninki '' 人気
                        End With ' HyoWide
                    End With ' mBuf
                    With mBuf.id
                        System.Diagnostics.Debug.WriteLine("Insert HYOSU_UMARENWIDE : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum & mBuf.HyoUmaren(i).Kumi)
                    End With ' id
                    mRS4.Update()
                End If
            Next i

        ' HYOSU_UMATAN (票数_馬単)
            For i = 0 To 305
                If mBuf.HyoUmatan(i).Kumi <> "    " Then
                    mRS5.AddNew()
                    With mBuf
                        With .head
                            With .MakeDate
                                mRS5.Fields("MakeDate").Value = .Year & .Month & .Day '' 年月日
                            End With ' MakeDate
                        End With ' head
                        With .id
                            mRS5.Fields("Year").Value = .Year '' 開催年
                            mRS5.Fields("MonthDay").Value = .MonthDay '' 開催月日
                            mRS5.Fields("JyoCD").Value = .JyoCD '' 競馬場コード
                            mRS5.Fields("Kaiji").Value = .Kaiji '' 開催回第N回
                            mRS5.Fields("Nichiji").Value = .Nichiji '' 開催日目N日目
                            mRS5.Fields("RaceNum").Value = .RaceNum '' レース番号
                        End With ' id
                        With .HyoUmatan(i)
                            mRS5.Fields("Kumi").Value = .Kumi '' 組番
                            mRS5.Fields("Hyo").Value = .Hyo '' 票数
                            mRS5.Fields("Ninki").Value = .Ninki '' 人気
                        End With ' HyoUmatan
                    End With ' mBuf
                    With mBuf.id
                        System.Diagnostics.Debug.WriteLine("Insert HYOSU_UMATAN : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum & mBuf.HyoUmatan(i).Kumi)
                    End With ' id
                    mRS5.Update()
                End If
            Next i

        ' HYOSU_SANREN (票数_三連)
            For i = 0 To 815
                If mBuf.HyoSanrenpuku(i).Kumi <> "      " Then
                    mRS6.AddNew()
                    With mBuf
                        With .head
                            With .MakeDate
                                mRS6.Fields("MakeDate").Value = .Year & .Month & .Day '' 年月日
                            End With ' MakeDate
                        End With ' head
                        With .id
                            mRS6.Fields("Year").Value = .Year '' 開催年
                            mRS6.Fields("MonthDay").Value = .MonthDay '' 開催月日
                            mRS6.Fields("JyoCD").Value = .JyoCD '' 競馬場コード
                            mRS6.Fields("Kaiji").Value = .Kaiji '' 開催回第N回
                            mRS6.Fields("Nichiji").Value = .Nichiji '' 開催日目N日目
                            mRS6.Fields("RaceNum").Value = .RaceNum '' レース番号
                        End With ' id
                        With .HyoSanrenpuku(i)
                            mRS6.Fields("Kumi").Value = .Kumi '' 組番
                            mRS6.Fields("Hyo").Value = .Hyo '' 票数
                            mRS6.Fields("Ninki").Value = .Ninki '' 人気
                        End With ' HyoSanrenpuku
                    End With ' mBuf

                    With mBuf.id
                        System.Diagnostics.Debug.WriteLine("Insert HYOSU_SANREN : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum & mBuf.HyoSanrenpuku(i).Kumi)
                    End With ' id
                    mRS6.Update()
                End If
            Next i

        gCon.CommitTrans()
        System.Diagnostics.Debug.WriteLine("CommitTrans")

        InsertDB = True

ExitHandler:
        Exit Function
ErrorHandler:
        InsertDB = False
        mRS1.CancelUpdate()
        mRS2.CancelUpdate()
        mRS3.CancelUpdate()
        mRS4.CancelUpdate()
        mRS5.CancelUpdate()
        mRS6.CancelUpdate()
        gCon.RollbackTrans()
        System.Diagnostics.Debug.WriteLine(Err.Description)
        System.Diagnostics.Debug.WriteLine("RollbackTrans")
        Resume ExitHandler
    End Function

    ' @(f)
    '
    ' 機能      : データベースを更新する
    '
    ' 引き数    :
    '
    ' 返り値    : True - 成功, False - 失敗
    '
    ' 機能説明  :
    '
    Public Function UpdateDB(ByVal strMakeDate As String) As Boolean
        On Error GoTo ErrorHandler
        Dim i As Short '' ループカウンタ
        Dim j As Short '' ループカウンタ
        Dim k As Short '' ループカウンタ
        Dim strSql As String '' SQL文

        gCon.BeginTrans()
        System.Diagnostics.Debug.WriteLine("BeginTrans")

        ' HYOSU (票数)
        strSql = "UPDATE HYOSU SET "
        With mBuf
            With .head
                strSql = strSql & SS & "RecordSpec" & SE & "='" & Replace(.RecordSpec, "'", "''") & "'," '' レコード種別
                strSql = strSql & SS & "DataKubun" & SE & "='" & Replace(.DataKubun, "'", "''") & "'," '' データ区分
                strSql = strSql & SS & "MakeDate" & SE & "= '" & Replace(strMakeDate, "'", "''") & "'," '' 年月日
            End With ' head
            With .id
                strSql = strSql & SS & "Year" & SE & "='" & Replace(.Year, "'", "''") & "'," '' 開催年
                strSql = strSql & SS & "MonthDay" & SE & "='" & Replace(.MonthDay, "'", "''") & "'," '' 開催月日
                strSql = strSql & SS & "JyoCD" & SE & "='" & Replace(.JyoCD, "'", "''") & "'," '' 競馬場コード
                strSql = strSql & SS & "Kaiji" & SE & "='" & Replace(.Kaiji, "'", "''") & "'," '' 開催回第N回
                strSql = strSql & SS & "Nichiji" & SE & "='" & Replace(.Nichiji, "'", "''") & "'," '' 開催日目N日目
                strSql = strSql & SS & "RaceNum" & SE & "='" & Replace(.RaceNum, "'", "''") & "'," '' レース番号
            End With ' id
            strSql = strSql & SS & "TorokuTosu" & SE & "='" & Replace(.TorokuTosu, "'", "''") & "'," '' 登録頭数
            strSql = strSql & SS & "SyussoTosu" & SE & "='" & Replace(.SyussoTosu, "'", "''") & "'," '' 出走頭数
            For i = 0 To 6
                strSql = strSql & SS & "HatubaiFlag" & i + 1 & SE & "='" & Replace(.HatubaiFlag(i), "'", "''") & "'," '' 発売フラグ
            Next i
            strSql = strSql & SS & "FukuChakuBaraiKey" & SE & "='" & Replace(.FukuChakuBaraiKey, "'", "''") & "'," '' 複勝着払キー
            For i = 0 To 27
                strSql = strSql & SS & "HenkanUma" & i + 1 & SE & "='" & Replace(.HenkanUma(i), "'", "''") & "'," '' 返還馬番情報(馬番01～28)
            Next i
            For i = 0 To 7
                strSql = strSql & SS & "HenkanWaku" & i + 1 & SE & "='" & Replace(.HenkanWaku(i), "'", "''") & "'," '' 返還枠番情報(枠番1～8)
            Next i
            For i = 0 To 7
                strSql = strSql & SS & "HenkanDoWaku" & i + 1 & SE & "='" & Replace(.HenkanDoWaku(i), "'", "''") & "'," '' 返還同枠情報(枠番1～8)
            Next i
            For i = 0 To 13
                strSql = strSql & SS & "HyoTotal" & i + 1 & SE & "='" & Replace(.HyoTotal(i), "'", "''") & "'," '' 票数合計
            Next i

            strSql = Left(strSql, Len(strSql) - 1) ' Delete Last Comma

            With .id
                strSql = strSql & " WHERE " & SS & "Year" & SE & "='" & Replace(.Year, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "MonthDay" & SE & "='" & Replace(.MonthDay, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "JyoCD" & SE & "='" & Replace(.JyoCD, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "Kaiji" & SE & "='" & Replace(.Kaiji, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "Nichiji" & SE & "='" & Replace(.Nichiji, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "RaceNum" & SE & "='" & Replace(.RaceNum, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "MakeDate" & SE & " <= '" & Replace(strMakeDate, "'", "''") & "'"
            End With ' id
        End With ' mBuf

        With mBuf.id
            System.Diagnostics.Debug.WriteLine("UPDATE HYOSU : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum)
        End With ' id

        gCon.Execute(strSql)

        ' HYOSU_TANPUKU (票数_単複)
        For i = 0 To 27
            strSql = "UPDATE HYOSU_TANPUKU SET "
            With mBuf
                strSql = strSql & SS & "MakeDate" & SE & "= '" & Replace(strMakeDate, "'", "''") & "'," '' 年月日
                With .id
                    strSql = strSql & SS & "Year" & SE & "='" & Replace(.Year, "'", "''") & "'," '' 開催年
                    strSql = strSql & SS & "MonthDay" & SE & "='" & Replace(.MonthDay, "'", "''") & "'," '' 開催月日
                    strSql = strSql & SS & "JyoCD" & SE & "='" & Replace(.JyoCD, "'", "''") & "'," '' 競馬場コード
                    strSql = strSql & SS & "Kaiji" & SE & "='" & Replace(.Kaiji, "'", "''") & "'," '' 開催回第N回
                    strSql = strSql & SS & "Nichiji" & SE & "='" & Replace(.Nichiji, "'", "''") & "'," '' 開催日目N日目
                    strSql = strSql & SS & "RaceNum" & SE & "='" & Replace(.RaceNum, "'", "''") & "'," '' レース番号
                End With ' id
                With .HyoTansyo(i)
                    strSql = strSql & SS & "Umaban" & SE & "='" & Replace(.Umaban, "'", "''") & "'," '' 馬番
                    strSql = strSql & SS & "TanHyo" & SE & "='" & Replace(.Hyo, "'", "''") & "'," '' 票数
                    strSql = strSql & SS & "TanNinki" & SE & "='" & Replace(.Ninki, "'", "''") & "'," '' 人気
                End With ' HyoTansyo
                With .HyoFukusyo(i)
                    strSql = strSql & SS & "FukuHyo" & SE & "='" & Replace(.Hyo, "'", "''") & "'," '' 票数
                    strSql = strSql & SS & "FukuNinki" & SE & "='" & Replace(.Ninki, "'", "''") & "'," '' 人気
                End With ' HyoFukusyo

                strSql = Left(strSql, Len(strSql) - 1) ' Delete Last Comma

                With .id
                    strSql = strSql & " WHERE " & SS & "Year" & SE & "='" & Replace(.Year, "'", "''") & "'"
                    strSql = strSql & " AND " & SS & "MonthDay" & SE & "='" & Replace(.MonthDay, "'", "''") & "'"
                    strSql = strSql & " AND " & SS & "JyoCD" & SE & "='" & Replace(.JyoCD, "'", "''") & "'"
                    strSql = strSql & " AND " & SS & "Kaiji" & SE & "='" & Replace(.Kaiji, "'", "''") & "'"
                    strSql = strSql & " AND " & SS & "Nichiji" & SE & "='" & Replace(.Nichiji, "'", "''") & "'"
                    strSql = strSql & " AND " & SS & "RaceNum" & SE & "='" & Replace(.RaceNum, "'", "''") & "'"
                    strSql = strSql & " AND " & SS & "MakeDate" & SE & " <= '" & Replace(strMakeDate, "'", "''") & "'"
                End With
                strSql = strSql & " AND " & SS & "Umaban" & SE & "='" & Replace(.HyoTansyo(i).Umaban, "'", "''") & "'"
            End With ' mBuf

            With mBuf.id
                System.Diagnostics.Debug.WriteLine("UPDATE HYOSU_TANPUKU : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum & mBuf.HyoTansyo(i).Umaban)
            End With ' id

            gCon.Execute(strSql)
        Next i

        ' HYOSU_WAKU (票数_枠連)
        For i = 0 To 35
            strSql = "UPDATE HYOSU_WAKU SET "
            With mBuf
                strSql = strSql & SS & "MakeDate" & SE & "= '" & Replace(strMakeDate, "'", "''") & "'," '' 年月日
                With .id
                    strSql = strSql & SS & "Year" & SE & "='" & Replace(.Year, "'", "''") & "'," '' 開催年
                    strSql = strSql & SS & "MonthDay" & SE & "='" & Replace(.MonthDay, "'", "''") & "'," '' 開催月日
                    strSql = strSql & SS & "JyoCD" & SE & "='" & Replace(.JyoCD, "'", "''") & "'," '' 競馬場コード
                    strSql = strSql & SS & "Kaiji" & SE & "='" & Replace(.Kaiji, "'", "''") & "'," '' 開催回第N回
                    strSql = strSql & SS & "Nichiji" & SE & "='" & Replace(.Nichiji, "'", "''") & "'," '' 開催日目N日目
                    strSql = strSql & SS & "RaceNum" & SE & "='" & Replace(.RaceNum, "'", "''") & "'," '' レース番号
                End With ' id
                With .HyoWakuren(i)
                    strSql = strSql & SS & "Kumi" & SE & "='" & Replace(.Umaban, "'", "''") & "'," '' 馬番
                    strSql = strSql & SS & "Hyo" & SE & "='" & Replace(.Hyo, "'", "''") & "'," '' 票数
                    strSql = strSql & SS & "Ninki" & SE & "='" & Replace(.Ninki, "'", "''") & "'," '' 人気
                End With ' HyoWakuren

                strSql = Left(strSql, Len(strSql) - 1) ' Delete Last Comma

                With .id
                    strSql = strSql & " WHERE " & SS & "Year" & SE & "='" & Replace(.Year, "'", "''") & "'"
                    strSql = strSql & " AND " & SS & "MonthDay" & SE & "='" & Replace(.MonthDay, "'", "''") & "'"
                    strSql = strSql & " AND " & SS & "JyoCD" & SE & "='" & Replace(.JyoCD, "'", "''") & "'"
                    strSql = strSql & " AND " & SS & "Kaiji" & SE & "='" & Replace(.Kaiji, "'", "''") & "'"
                    strSql = strSql & " AND " & SS & "Nichiji" & SE & "='" & Replace(.Nichiji, "'", "''") & "'"
                    strSql = strSql & " AND " & SS & "RaceNum" & SE & "='" & Replace(.RaceNum, "'", "''") & "'"
                    strSql = strSql & " AND " & SS & "MakeDate" & SE & " <= '" & Replace(strMakeDate, "'", "''") & "'"
                End With
                strSql = strSql & " AND " & SS & "Kumi" & SE & "='" & Replace(.HyoWakuren(i).Umaban, "'", "''") & "'"
            End With ' mBuf

            With mBuf.id
                System.Diagnostics.Debug.WriteLine("UPDATE HYOSU_WAKU : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum & mBuf.HyoWakuren(i).Umaban)
            End With ' id

            gCon.Execute(strSql)
        Next i

        ' HYOSU_UMARENWIDE (票数_馬連・ワイド)
        For i = 0 To 152
            strSql = "UPDATE HYOSU_UMARENWIDE SET "
            With mBuf
                strSql = strSql & SS & "MakeDate" & SE & "= '" & Replace(strMakeDate, "'", "''") & "'," '' 年月日
                With .id
                    strSql = strSql & SS & "Year" & SE & "='" & Replace(.Year, "'", "''") & "'," '' 開催年
                    strSql = strSql & SS & "MonthDay" & SE & "='" & Replace(.MonthDay, "'", "''") & "'," '' 開催月日
                    strSql = strSql & SS & "JyoCD" & SE & "='" & Replace(.JyoCD, "'", "''") & "'," '' 競馬場コード
                    strSql = strSql & SS & "Kaiji" & SE & "='" & Replace(.Kaiji, "'", "''") & "'," '' 開催回第N回
                    strSql = strSql & SS & "Nichiji" & SE & "='" & Replace(.Nichiji, "'", "''") & "'," '' 開催日目N日目
                    strSql = strSql & SS & "RaceNum" & SE & "='" & Replace(.RaceNum, "'", "''") & "'," '' レース番号
                End With ' id
                With .HyoUmaren(i)
                    strSql = strSql & SS & "Kumi" & SE & "='" & Replace(.Kumi, "'", "''") & "'," '' 組番
                    strSql = strSql & SS & "UmarenHyo" & SE & "='" & Replace(.Hyo, "'", "''") & "'," '' 票数
                    strSql = strSql & SS & "UmarenNinki" & SE & "='" & Replace(.Ninki, "'", "''") & "'," '' 人気
                End With ' HyoUmaren
                With .HyoWide(i)
                    strSql = strSql & SS & "WideHyo" & SE & "='" & Replace(.Hyo, "'", "''") & "'," '' 票数
                    strSql = strSql & SS & "WideNinki" & SE & "='" & Replace(.Ninki, "'", "''") & "'," '' 人気
                End With ' HyoWide

                strSql = Left(strSql, Len(strSql) - 1) ' Delete Last Comma

                With .id
                    strSql = strSql & " WHERE " & SS & "Year" & SE & "='" & Replace(.Year, "'", "''") & "'"
                    strSql = strSql & " AND " & SS & "MonthDay" & SE & "='" & Replace(.MonthDay, "'", "''") & "'"
                    strSql = strSql & " AND " & SS & "JyoCD" & SE & "='" & Replace(.JyoCD, "'", "''") & "'"
                    strSql = strSql & " AND " & SS & "Kaiji" & SE & "='" & Replace(.Kaiji, "'", "''") & "'"
                    strSql = strSql & " AND " & SS & "Nichiji" & SE & "='" & Replace(.Nichiji, "'", "''") & "'"
                    strSql = strSql & " AND " & SS & "RaceNum" & SE & "='" & Replace(.RaceNum, "'", "''") & "'"
                    strSql = strSql & " AND " & SS & "MakeDate" & SE & " <= '" & Replace(strMakeDate, "'", "''") & "'"
                End With
                strSql = strSql & " AND " & SS & "Kumi" & SE & "='" & Replace(.HyoUmaren(i).Kumi, "'", "''") & "'"
            End With ' mBuf

            With mBuf.id
                System.Diagnostics.Debug.WriteLine("UPDATE HYOSU_UMARENWIDE : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum & mBuf.HyoUmaren(i).Kumi)
            End With ' id

            gCon.Execute(strSql)

        Next i

        ' HYOSU_UMATAN (票数_馬単)
        For i = 0 To 305
            strSql = "UPDATE HYOSU_UMATAN SET "
            With mBuf
                strSql = strSql & SS & "MakeDate" & SE & "= '" & Replace(strMakeDate, "'", "''") & "'," '' 年月日
                With .id
                    strSql = strSql & SS & "Year" & SE & "='" & Replace(.Year, "'", "''") & "'," '' 開催年
                    strSql = strSql & SS & "MonthDay" & SE & "='" & Replace(.MonthDay, "'", "''") & "'," '' 開催月日
                    strSql = strSql & SS & "JyoCD" & SE & "='" & Replace(.JyoCD, "'", "''") & "'," '' 競馬場コード
                    strSql = strSql & SS & "Kaiji" & SE & "='" & Replace(.Kaiji, "'", "''") & "'," '' 開催回第N回
                    strSql = strSql & SS & "Nichiji" & SE & "='" & Replace(.Nichiji, "'", "''") & "'," '' 開催日目N日目
                    strSql = strSql & SS & "RaceNum" & SE & "='" & Replace(.RaceNum, "'", "''") & "'," '' レース番号
                End With ' id
                With .HyoUmatan(i)
                    strSql = strSql & SS & "Kumi" & SE & "='" & Replace(.Kumi, "'", "''") & "'," '' 組番
                    strSql = strSql & SS & "Hyo" & SE & "='" & Replace(.Hyo, "'", "''") & "'," '' 票数
                    strSql = strSql & SS & "Ninki" & SE & "='" & Replace(.Ninki, "'", "''") & "'," '' 人気
                End With ' HyoUmatan

                strSql = Left(strSql, Len(strSql) - 1) ' Delete Last Comma

                With .id
                    strSql = strSql & " WHERE " & SS & "Year" & SE & "='" & Replace(.Year, "'", "''") & "'"
                    strSql = strSql & " AND " & SS & "MonthDay" & SE & "='" & Replace(.MonthDay, "'", "''") & "'"
                    strSql = strSql & " AND " & SS & "JyoCD" & SE & "='" & Replace(.JyoCD, "'", "''") & "'"
                    strSql = strSql & " AND " & SS & "Kaiji" & SE & "='" & Replace(.Kaiji, "'", "''") & "'"
                    strSql = strSql & " AND " & SS & "Nichiji" & SE & "='" & Replace(.Nichiji, "'", "''") & "'"
                    strSql = strSql & " AND " & SS & "RaceNum" & SE & "='" & Replace(.RaceNum, "'", "''") & "'"
                    strSql = strSql & " AND " & SS & "MakeDate" & SE & " <= '" & Replace(strMakeDate, "'", "''") & "'"
                End With
                strSql = strSql & " AND " & SS & "Kumi" & SE & "='" & Replace(.HyoUmatan(i).Kumi, "'", "''") & "'"
            End With ' mBuf

            With mBuf.id
                System.Diagnostics.Debug.WriteLine("UPDATE HYOSU_UMATAN : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum & mBuf.HyoUmatan(i).Kumi)
            End With ' id

            gCon.Execute(strSql)
        Next i

        ' HYOSU_SANREN (票数_三連)
        For i = 0 To 815
            strSql = "UPDATE HYOSU_SANREN SET "
            With mBuf
                strSql = strSql & SS & "MakeDate" & SE & "= '" & Replace(strMakeDate, "'", "''") & "'," '' 年月日
                With .id
                    strSql = strSql & SS & "Year" & SE & "='" & Replace(.Year, "'", "''") & "'," '' 開催年
                    strSql = strSql & SS & "MonthDay" & SE & "='" & Replace(.MonthDay, "'", "''") & "'," '' 開催月日
                    strSql = strSql & SS & "JyoCD" & SE & "='" & Replace(.JyoCD, "'", "''") & "'," '' 競馬場コード
                    strSql = strSql & SS & "Kaiji" & SE & "='" & Replace(.Kaiji, "'", "''") & "'," '' 開催回第N回
                    strSql = strSql & SS & "Nichiji" & SE & "='" & Replace(.Nichiji, "'", "''") & "'," '' 開催日目N日目
                    strSql = strSql & SS & "RaceNum" & SE & "='" & Replace(.RaceNum, "'", "''") & "'," '' レース番号
                End With ' id
                With .HyoSanrenpuku(i)
                    strSql = strSql & SS & "Kumi" & SE & "='" & Replace(.Kumi, "'", "''") & "'," '' 組番
                    strSql = strSql & SS & "Hyo" & SE & "='" & Replace(.Hyo, "'", "''") & "'," '' 票数
                    strSql = strSql & SS & "Ninki" & SE & "='" & Replace(.Ninki, "'", "''") & "'," '' 人気
                End With ' HyoSanrenpuku

                strSql = Left(strSql, Len(strSql) - 1) ' Delete Last Comma

                With .id
                    strSql = strSql & " WHERE " & SS & "Year" & SE & "='" & Replace(.Year, "'", "''") & "'"
                    strSql = strSql & " AND " & SS & "MonthDay" & SE & "='" & Replace(.MonthDay, "'", "''") & "'"
                    strSql = strSql & " AND " & SS & "JyoCD" & SE & "='" & Replace(.JyoCD, "'", "''") & "'"
                    strSql = strSql & " AND " & SS & "Kaiji" & SE & "='" & Replace(.Kaiji, "'", "''") & "'"
                    strSql = strSql & " AND " & SS & "Nichiji" & SE & "='" & Replace(.Nichiji, "'", "''") & "'"
                    strSql = strSql & " AND " & SS & "RaceNum" & SE & "='" & Replace(.RaceNum, "'", "''") & "'"
                    strSql = strSql & " AND " & SS & "MakeDate" & SE & " <= '" & Replace(strMakeDate, "'", "''") & "'"
                End With
                strSql = strSql & " AND " & SS & "Kumi" & SE & "='" & Replace(.HyoSanrenpuku(i).Kumi, "'", "''") & "'"
            End With ' mBuf

            With mBuf.id
                System.Diagnostics.Debug.WriteLine("UPDATE HYOSU_SANREN : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum & mBuf.HyoSanrenpuku(i).Kumi)
            End With ' id

            gCon.Execute(strSql)
        Next i

        gCon.CommitTrans()
        System.Diagnostics.Debug.WriteLine("CommitTrans")

        UpdateDB = True

ExitHandler:
        Exit Function
ErrorHandler:
        System.Diagnostics.Debug.WriteLine(Err.Description)
        System.Diagnostics.Debug.WriteLine("RollbackTrans")
        gCon.RollbackTrans()
        UpdateDB = False
        Resume ExitHandler
    End Function
End Class