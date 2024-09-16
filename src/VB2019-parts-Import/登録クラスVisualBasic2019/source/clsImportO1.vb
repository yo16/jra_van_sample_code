Option Strict Off
Option Explicit On
Option Compare Binary
Friend Class clsImportO1
	' @(h) clsReadO1.cls
	' @(s)
	' JVData "O1" データベース登録クラス
	'
	
	Private mBuf As JV_O1_ODDS_TANFUKUWAKU '' オッズ（単複枠）構造体
	Private mRS1 As ADODB.Recordset
	Private mRS2 As ADODB.Recordset
	Private mRS3 As ADODB.Recordset
	
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
		strSql = "SELECT * FROM ODDS_TANPUKUWAKU_HEAD"
		mRS1 = New ADODB.Recordset
		mRS1.Open(strSql, gCon, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockPessimistic)
		
		strSql = "SELECT * FROM ODDS_TANPUKU"
		mRS2 = New ADODB.Recordset
		mRS2.Open(strSql, gCon, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockPessimistic)
		
		strSql = "SELECT * FROM ODDS_WAKU"
		mRS3 = New ADODB.Recordset
		mRS3.Open(strSql, gCon, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockPessimistic)
		
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
    ' 引き数    : lBuf - JVData 識別子"O1" の１行
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
        ' ヘッダ部分
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
                mRS1.Fields("HappyoTime").Value = .Month & .Day & .Hour & .Minute '' 月日時分
            End With ' HappyoTime
            mRS1.Fields("TorokuTosu").Value = .TorokuTosu '' 登録頭数
            mRS1.Fields("SyussoTosu").Value = .SyussoTosu '' 出走頭数
            mRS1.Fields("TansyoFlag").Value = .TansyoFlag '' 発売フラグ
            mRS1.Fields("FukusyoFlag").Value = .FukusyoFlag '' 発売フラグ
            mRS1.Fields("WakurenFlag").Value = .WakurenFlag '' 発売フラグ　枠連
            mRS1.Fields("FukuChakuBaraiKey").Value = .FukuChakuBaraiKey '' 複勝着払キー
            mRS1.Fields("TotalHyosuTansyo").Value = .TotalHyosuTansyo '' 単勝票数合計
            mRS1.Fields("TotalHyosuFukusyo").Value = .TotalHyosuFukusyo '' 複勝票数合計
            mRS1.Fields("TotalHyosuWakuren").Value = .TotalHyosuWakuren '' 枠連票数合計
        End With

        With mBuf.id
            System.Diagnostics.Debug.WriteLine("Insert ODDS_TANPUKUWAKU_HEAD : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum)
        End With ' id

        mRS1.Update()

        '' ODDS_TANPUKU (オッズ_単複)
        If mBuf.TansyoFlag <> "0" Then
            For i = 0 To 27
                If mBuf.OddsTansyoInfo(i).Umaban <> "  " Then
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
                        With .OddsTansyoInfo(i)
                            mRS2.Fields("Umaban").Value = .Umaban '' 馬番
                            mRS2.Fields("TanOdds").Value = .Odds '' オッズ
                            mRS2.Fields("TanNinki").Value = .Ninki '' 人気順
                        End With ' OddsTansyoInfo
                        With .OddsFukusyoInfo(i)
                            mRS2.Fields("FukuOddsLow").Value = .OddsLow '' 最低オッズ
                            mRS2.Fields("FukuOddsHigh").Value = .OddsHigh '' 最高オッズ
                            mRS2.Fields("FukuNinki").Value = .Ninki '' 人気順
                        End With ' OddsFukusyoInfo
                    End With ' mBuf

                    With mBuf.id
                        System.Diagnostics.Debug.WriteLine("Insert ODDS_TANPUKU : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum & mBuf.OddsTansyoInfo(i).Umaban)
                    End With ' id

                    mRS2.Update()

                End If
            Next i
        End If


        '' ODDS_WAKU (オッズ_枠連)
        If mBuf.TansyoFlag <> "0" Then
            For i = 0 To 35

                If mBuf.OddsWakurenInfo(i).Kumi <> "  " Then
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
                        With .OddsWakurenInfo(i)
                            mRS3.Fields("Kumi").Value = .Kumi '' 組
                            mRS3.Fields("Odds").Value = .Odds '' オッズ
                            mRS3.Fields("Ninki").Value = .Ninki '' 人気順
                        End With ' OddsWakurenInfo
                    End With ' mBuf

                    With mBuf.id
                        System.Diagnostics.Debug.WriteLine("Insert ODDS_WAKU : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum & mBuf.OddsWakurenInfo(i).Kumi)
                    End With ' id

                    mRS3.Update()
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
        gCon.RollbackTrans()
        mRS1.CancelUpdate()
        mRS2.CancelUpdate()
        mRS3.CancelUpdate()
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

        ' ヘッダ部分
        strSql = "UPDATE ODDS_TANPUKUWAKU_HEAD SET "
        With mBuf
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
            strSql = strSql & SS & "TansyoFlag" & SE & "='" & Replace(.TansyoFlag, "'", "''") & "'," '' 発売フラグ
            strSql = strSql & SS & "FukusyoFlag" & SE & "='" & Replace(.FukusyoFlag, "'", "''") & "'," '' 発売フラグ
            strSql = strSql & SS & "WakurenFlag" & SE & "='" & Replace(.WakurenFlag, "'", "''") & "'," '' 発売フラグ　枠連
            strSql = strSql & SS & "FukuChakuBaraiKey" & SE & "='" & Replace(.FukuChakuBaraiKey, "'", "''") & "'," '' 複勝着払キー
            strSql = strSql & SS & "TotalHyosuTansyo" & SE & "='" & Replace(.TotalHyosuTansyo, "'", "''") & "'," '' 単勝票数合計
            strSql = strSql & SS & "TotalHyosuFukusyo" & SE & "='" & Replace(.TotalHyosuFukusyo, "'", "''") & "'," '' 複勝票数合計
            strSql = strSql & SS & "TotalHyosuWakuren" & SE & "='" & Replace(.TotalHyosuWakuren, "'", "''") & "'," '' 枠連票数合計

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

            With mBuf.id
                System.Diagnostics.Debug.WriteLine("UPDATE ODDS_TANPUKUWAKU_HEAD : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum)
            End With ' id

            gCon.Execute(strSql)
        End With

        '' ODDS_TANPUKU (オッズ_単複)
        For i = 0 To 27
            strSql = "UPDATE ODDS_TANPUKU SET "
            With mBuf
                With .head
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
                With .OddsTansyoInfo(i)
                    strSql = strSql & SS & "Umaban" & SE & "='" & Replace(.Umaban, "'", "''") & "'," '' 馬番
                    strSql = strSql & SS & "TanOdds" & SE & "='" & Replace(.Odds, "'", "''") & "'," '' オッズ
                    strSql = strSql & SS & "TanNinki" & SE & "='" & Replace(.Ninki, "'", "''") & "'," '' 人気順
                End With ' OddsTansyoInfo
                With .OddsFukusyoInfo(i)
                    strSql = strSql & SS & "FukuOddsLow" & SE & "='" & Replace(.OddsLow, "'", "''") & "'," '' 最低オッズ
                    strSql = strSql & SS & "FukuOddsHigh" & SE & "='" & Replace(.OddsHigh, "'", "''") & "'," '' 最高オッズ
                    strSql = strSql & SS & "FukuNinki" & SE & "='" & Replace(.Ninki, "'", "''") & "'," '' 人気順
                End With ' OddsFukusyoInfo

                strSql = Left(strSql, Len(strSql) - 1) ' Delete Last Comma

                With .id
                    strSql = strSql & " WHERE " & SS & "Year" & SE & "='" & Replace(.Year, "'", "''") & "'"
                    strSql = strSql & " AND " & SS & "MonthDay" & SE & "='" & Replace(.MonthDay, "'", "''") & "'"
                    strSql = strSql & " AND " & SS & "JyoCD" & SE & "='" & Replace(.JyoCD, "'", "''") & "'"
                    strSql = strSql & " AND " & SS & "Kaiji" & SE & "='" & Replace(.Kaiji, "'", "''") & "'"
                    strSql = strSql & " AND " & SS & "Nichiji" & SE & "='" & Replace(.Nichiji, "'", "''") & "'"
                    strSql = strSql & " AND " & SS & "RaceNum" & SE & "='" & Replace(.RaceNum, "'", "''") & "'"
                    strSql = strSql & " AND " & SS & "Umaban" & SE & "='" & Replace(mBuf.OddsTansyoInfo(i).Umaban, "'", "''") & "'"
                    strSql = strSql & " AND " & SS & "MakeDate" & SE & " <= '" & Replace(strMakeDate, "'", "''") & "'"
                End With
            End With ' mBuf

            With mBuf.id
                System.Diagnostics.Debug.WriteLine("UPDATE ODDS_TANPUKU : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum & mBuf.OddsTansyoInfo(i).Umaban)
            End With ' id
            gCon.Execute(strSql)
        Next i

        '' ODDS_WAKU (オッズ_枠連)
        For i = 0 To 35
            strSql = "UPDATE ODDS_WAKU SET "
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
                With .OddsWakurenInfo(i)
                    strSql = strSql & SS & "Kumi" & SE & "='" & Replace(.Kumi, "'", "''") & "'," '' 組
                    strSql = strSql & SS & "Odds" & SE & "='" & Replace(.Odds, "'", "''") & "'," '' オッズ
                    strSql = strSql & SS & "Ninki" & SE & "='" & Replace(.Ninki, "'", "''") & "'," '' 人気順
                End With ' OddsWakurenInfo

                strSql = Left(strSql, Len(strSql) - 1) ' Delete Last Comma

                With .id
                    strSql = strSql & " WHERE " & SS & "Year" & SE & "='" & Replace(.Year, "'", "''") & "'"
                    strSql = strSql & " AND " & SS & "MonthDay" & SE & "='" & Replace(.MonthDay, "'", "''") & "'"
                    strSql = strSql & " AND " & SS & "JyoCD" & SE & "='" & Replace(.JyoCD, "'", "''") & "'"
                    strSql = strSql & " AND " & SS & "Kaiji" & SE & "='" & Replace(.Kaiji, "'", "''") & "'"
                    strSql = strSql & " AND " & SS & "Nichiji" & SE & "='" & Replace(.Nichiji, "'", "''") & "'"
                    strSql = strSql & " AND " & SS & "RaceNum" & SE & "='" & Replace(.RaceNum, "'", "''") & "'"
                    strSql = strSql & " AND " & SS & "Kumi" & SE & "='" & Replace(mBuf.OddsWakurenInfo(i).Kumi, "'", "''") & "'"
                    strSql = strSql & " AND " & SS & "MakeDate" & SE & " <= '" & Replace(strMakeDate, "'", "''") & "'"
                End With
            End With

            With mBuf.id
                System.Diagnostics.Debug.WriteLine("UPDATE ODDS_WAKU : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum & mBuf.OddsWakurenInfo(i).Kumi)
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