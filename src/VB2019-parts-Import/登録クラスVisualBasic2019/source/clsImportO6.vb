Option Strict Off
Option Explicit On
Option Compare Binary
Friend Class clsImportO6
    ' @(h) clsReadO6.cls
    ' @(s)
    ' JVData "O6" データベース登録クラス
    '

    Private mBuf As JV_O6_ODDS_SANRENTAN ''オッズ（3連単）構造体
    Private mRS1 As ADODB.Recordset
    Private mRS2 As ADODB.Recordset

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
        strSql = "SELECT * FROM ODDS_SANRENTAN_HEAD"
        mRS1 = New ADODB.Recordset()
        mRS1.Open(strSql, gCon, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockPessimistic)

        'レコードセットオープン
        strSql = "SELECT * FROM ODDS_SANRENTAN"
        mRS2 = New ADODB.Recordset()
        mRS2.Open(strSql, gCon, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockPessimistic)

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
    ' 引き数    : lBuf - JVData 識別子"O5" の１行
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

        ' ODDS_SANRENTAN_HEAD (オッズ_3連単_ヘッダ)

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
            With .HappyoTime
                mRS1.Fields("HappyoTime").Value = .Month & .Day & .Hour & .Minute
            End With ' HappyoTime
            mRS1.Fields("TorokuTosu").Value = .TorokuTosu '' 登録頭数
            mRS1.Fields("SyussoTosu").Value = .SyussoTosu '' 出走頭数
            mRS1.Fields("SanrentanFlag").Value = .SanrentanFlag '' 発売フラグ　3連単
            mRS1.Fields("TotalHyosuSanrentan").Value = .TotalHyosuSanrentan '' 3連単票数合計
        End With

        With mBuf.id
            System.Diagnostics.Debug.WriteLine("Insert ODDS_SANRENTAN_HEAD : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum)
        End With ' id

        mRS1.Update()

        ' ODDS_SANRENTAN (オッズ_3連単)
        If mBuf.SanrentanFlag <> "0" Then
            For i = 0 To 4895
                If mBuf.OddsSanrentanInfo(i).Kumi <> "      " Then
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
                        With .OddsSanrentanInfo(i)
                            mRS2.Fields("Kumi").Value = .Kumi '' 組番
                            mRS2.Fields("Odds").Value = .Odds '' オッズ
                            mRS2.Fields("Ninki").Value = .Ninki '' 人気順
                        End With ' OddsSanrentanInfo
                    End With

                    With mBuf.id
                        System.Diagnostics.Debug.WriteLine("Insert ODDS_SANRENTAN : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum & mBuf.OddsSanrentanInfo(i).Kumi)
                    End With ' id
                    mRS2.Update()
                End If
            Next i
        End If

        gCon.CommitTrans()
        System.Diagnostics.Debug.WriteLine("CommitTrans")

        InsertDB = True

ExitHandler:
        Exit Function
ErrorHandler:
        InsertDB = False
        mRS1.CancelUpdate()
        mRS2.CancelUpdate()
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
    Public Function UpdateDB(ByRef strMakeDate As String) As Boolean
        On Error GoTo ErrorHandler
        Dim i As Short '' ループカウンタ
        Dim j As Short '' ループカウンタ
        Dim k As Short '' ループカウンタ
        Dim strSql As String '' SQL文

        gCon.BeginTrans()
        System.Diagnostics.Debug.WriteLine("BeginTrans")

        ' ODDS_SANRENTAN_HEAD (オッズ_3連単_ヘッダ)
        strSql = "UPDATE ODDS_SANRENTAN_HEAD SET "
        With mBuf
            With .head
                strSql = strSql & SS & "RecordSpec" & SE & "='" & Replace(.RecordSpec, "'", "''") & "'," '' レコード種別
                strSql = strSql & SS & "DataKubun" & SE & "='" & Replace(.DataKubun, "'", "''") & "'," '' データ区分
                strSql = strSql & SS & "MakeDate" & SE & "='" & Replace(strMakeDate, "'", "''") & "'," '' 年月日
            End With ' head
            With .id
                strSql = strSql & SS & "Year" & SE & "='" & Replace(.Year, "'", "''") & "'," '' 開催年
                strSql = strSql & SS & "MonthDay" & SE & "='" & Replace(.MonthDay, "'", "''") & "'," '' 開催月日
                strSql = strSql & SS & "JyoCD" & SE & "='" & Replace(.JyoCD, "'", "''") & "'," '' 競馬場コード
                strSql = strSql & SS & "Kaiji" & SE & "='" & Replace(.Kaiji, "'", "''") & "'," '' 開催回第N回
                strSql = strSql & SS & "Nichiji" & SE & "='" & Replace(.Nichiji, "'", "''") & "'," '' 開催日目N日目
                strSql = strSql & SS & "RaceNum" & SE & "='" & Replace(.RaceNum, "'", "''") & "'," '' レース番号
            End With ' id
            With .HappyoTime
                strSql = strSql & SS & "HappyoTime" & SE & "='" & Replace(.Month & .Day & .Hour & .Minute, "'", "''") & "',"
            End With ' HappyoTime
            strSql = strSql & SS & "TorokuTosu" & SE & "='" & Replace(.TorokuTosu, "'", "''") & "'," '' 登録頭数
            strSql = strSql & SS & "SyussoTosu" & SE & "='" & Replace(.SyussoTosu, "'", "''") & "'," '' 出走頭数
            strSql = strSql & SS & "SanrentanFlag" & SE & "='" & Replace(.SanrentanFlag, "'", "''") & "'," '' 発売フラグ　3連単
            strSql = strSql & SS & "TotalHyosuSanrentan" & SE & "='" & Replace(.TotalHyosuSanrentan, "'", "''") & "'" '' 3連単票数合計
            With .id
                strSql = strSql & " WHERE " & SS & "Year" & SE & "='" & Replace(.Year, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "MonthDay" & SE & "='" & Replace(.MonthDay, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "JyoCD" & SE & "='" & Replace(.JyoCD, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "Kaiji" & SE & "='" & Replace(.Kaiji, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "Nichiji" & SE & "='" & Replace(.Nichiji, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "RaceNum" & SE & "='" & Replace(.RaceNum, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "MakeDate" & SE & " <= '" & Replace(strMakeDate, "'", "''") & "'"
            End With
        End With

        With mBuf.id
            System.Diagnostics.Debug.WriteLine("UPDATE ODDS_SANRENTAN_HEAD : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum)
        End With ' id
        gCon.Execute(strSql)


        ' ODDS_SANRENTAN (オッズ_3連単)
        For i = 0 To 4895
            strSql = "UPDATE ODDS_SANRENTAN SET "
            With mBuf
                strSql = strSql & SS & "MakeDate" & SE & "='" & Replace(strMakeDate, "'", "''") & "'," '' 年月日
                With .id
                    strSql = strSql & SS & "Year" & SE & "='" & Replace(.Year, "'", "''") & "'," '' 開催年
                    strSql = strSql & SS & "MonthDay" & SE & "='" & Replace(.MonthDay, "'", "''") & "'," '' 開催月日
                    strSql = strSql & SS & "JyoCD" & SE & "='" & Replace(.JyoCD, "'", "''") & "'," '' 競馬場コード
                    strSql = strSql & SS & "Kaiji" & SE & "='" & Replace(.Kaiji, "'", "''") & "'," '' 開催回第N回
                    strSql = strSql & SS & "Nichiji" & SE & "='" & Replace(.Nichiji, "'", "''") & "'," '' 開催日目N日目
                    strSql = strSql & SS & "RaceNum" & SE & "='" & Replace(.RaceNum, "'", "''") & "'," '' レース番号
                End With ' id
                With .OddsSanrentanInfo(i)
                    strSql = strSql & SS & "Kumi" & SE & "='" & Replace(.Kumi, "'", "''") & "'," '' 組番
                    strSql = strSql & SS & "Odds" & SE & "='" & Replace(.Odds, "'", "''") & "'," '' オッズ
                    strSql = strSql & SS & "Ninki" & SE & "='" & Replace(.Ninki, "'", "''") & "'," '' 人気順
                End With ' OddsSanrentanInfo

                strSql = Left(strSql, Len(strSql) - 1) ' Delete Last Comma

                With .id
                    strSql = strSql & " WHERE " & SS & "Year" & SE & "='" & Replace(.Year, "'", "''") & "'"
                    strSql = strSql & " AND " & SS & "MonthDay" & SE & "='" & Replace(.MonthDay, "'", "''") & "'"
                    strSql = strSql & " AND " & SS & "JyoCD" & SE & "='" & Replace(.JyoCD, "'", "''") & "'"
                    strSql = strSql & " AND " & SS & "Kaiji" & SE & "='" & Replace(.Kaiji, "'", "''") & "'"
                    strSql = strSql & " AND " & SS & "Nichiji" & SE & "='" & Replace(.Nichiji, "'", "''") & "'"
                    strSql = strSql & " AND " & SS & "RaceNum" & SE & "='" & Replace(.RaceNum, "'", "''") & "'"
                    strSql = strSql & " AND " & SS & "Kumi" & SE & "='" & Replace(mBuf.OddsSanrentanInfo(i).Kumi, "'", "''") & "'"
                    strSql = strSql & " AND " & SS & "MakeDate" & SE & " <= '" & Replace(strMakeDate, "'", "''") & "'"
                End With ' id
            End With ' mbuf

            With mBuf.id
                System.Diagnostics.Debug.WriteLine("UPDATE ODDS_SANRENTAN : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum & mBuf.OddsSanrentanInfo(i).Kumi)
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