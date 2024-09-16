Option Strict Off
Option Explicit On
Option Compare Binary
Friend Class clsImportWF
	' @(h) clsReadWF.cls
	' @(s)
	' JVData "WF" データベース登録クラス
	'
	
    Private mBuf As JV_WF_INFO     '' 重勝式(WIN5)構造体

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
        strSql = "SELECT * FROM JYUSYOSIKI_HEAD"
        mRS1 = New ADODB.Recordset()
        mRS1.Open(strSql, gCon, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockPessimistic)

        strSql = "SELECT * FROM JYUSYOSIKI"
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
    ' 引き数    : lBuf - JVData 識別子"WF" の１行
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

        gCon.BeginTrans()
        System.Diagnostics.Debug.WriteLine("BeginTrans")

        mRS1.AddNew()

        With mBuf
            With .head
                mRS1.Fields("RecordSpec").Value = .RecordSpec             '' レコード種別
                mRS1.Fields("DataKubun").Value = .DataKubun               '' データ区分
                With .MakeDate
                    mRS1.Fields("MakeDate").Value = .Year & .Month & .Day '' 年月日
                End With ' MakeDate
            End With ' head
            With .KaisaiDate
                mRS1.Fields("Year").Value = .Year                         '' 開催年
                mRS1.Fields("MonthDay").Value = .Month & .Day             '' 開催月日
            End With ' KaisaiDate
            mRS1.Fields("reserved1").Value = .reserved1                   '' 予備1
            For i = 0 To 4
                With .WFRaceInfo(i)
                    mRS1.Fields("JyoCD" & i + 1).Value = .JyoCD           '' 競馬場コード
                    mRS1.Fields("Kaiji" & i + 1).Value = .Kaiji           '' 開催回
                    mRS1.Fields("Nichiji" & i + 1).Value = .Nichiji       '' 開催日目
                    mRS1.Fields("RaceNum" & i + 1).Value = .RaceNum       '' レース番号
                End With ' WFRaceInfo()
            Next i
            mRS1.Fields("reserved2").Value = .reserved2                   '' 予備2
            For i = 0 To 4
                With .WFYukoHyoInfo(i)
                    mRS1.Fields("YukoHyosu" & i + 1).Value = .Yuko_Hyo    '' 有効票数
                End With ' WFYukoHyoInfo()
            Next
            mRS1.Fields("HenkanFlag").Value = .HenkanFlag                 '' 返還フラグ
            mRS1.Fields("FuseirituFlag").Value = .FuseiritsuFlag          '' 不成立フラグ
            mRS1.Fields("TekichunashiFlag").Value = .TekichunashiFlag     '' 的中無フラグ
            mRS1.Fields("CarryoverSyoki").Value = .COShoki                '' キャリーオーバー金額初期
            mRS1.Fields("CarryoverZandaka").Value = .COZanDaka            '' キャリーオーバー金額残高

            With .KaisaiDate
                System.Diagnostics.Debug.WriteLine("Insert JYUSYOSIKI_HEAD : " & .Year & .Month & .Day)
            End With ' KaisaiDate
        End With

        mRS1.Update()

        For i = 0 To 242
            If mBuf.WFPayInfo(i).Kumiban <> "          " Then
                mRS2.AddNew()
                With mBuf
                    With .head.MakeDate
                        mRS2.Fields("MakeDate").Value = .Year & .Month & .Day '' 年月日
                    End With ' MakeDate
                    With .KaisaiDate
                        mRS2.Fields("Year").Value = .Year                     '' 開催年
                        mRS2.Fields("MonthDay").Value = .Month & .Day         '' 開催月日
                    End With ' KaisaiDate
                    With .WFPayInfo(i)
                        mRS2.Fields("Kumi").Value = .Kumiban                  '' 組番
                        mRS2.Fields("PayJyushosiki").Value = .Pay             '' 重勝式払戻金
                        mRS2.Fields("TekichuHyo").Value = .Tekichu_Hyo        '' 的中票数
                    End With ' WFPayInfo()

                    With .KaisaiDate
                        System.Diagnostics.Debug.WriteLine("Insert JYUSYOSIKI : " & .Year & .Month & .Day & mBuf.WFPayInfo(i).Kumiban)
                    End With ' KaisaiDate

                    mRS2.Update()
                End With
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
        Dim strSql As String '' SQL文
        Dim i As Short '' ループカウンタ

        gCon.BeginTrans()
        System.Diagnostics.Debug.WriteLine("BeginTrans")

        strSql = "UPDATE JYUSYOSIKI_HEAD SET "
        With mBuf
            With .head
                strSql = strSql & SS & "RecordSpec" & SE & "='" & Replace(.RecordSpec, "'", "''") & "',"   '' レコード種別
                strSql = strSql & SS & "DataKubun" & SE & "='" & Replace(.DataKubun, "'", "''") & "',"     '' データ区分
                strSql = strSql & SS & "MakeDate" & SE & "='" & Replace(strMakeDate, "'", "''") & "',"     '' 年月日
            End With ' head
            With .KaisaiDate
                strSql = strSql & SS & "Year" & SE & "='" & Replace(.Year, "'", "''") & "',"               '' 開催年
                strSql = strSql & SS & "MonthDay" & SE & "='" & Replace(.Month & .Day, "'", "''") & "',"   '' 開催月日
            End With ' KaisaiDate
            strSql = strSql & SS & "reserved1" & SE & "='" & Replace(.reserved1, "'", "''") & "',"         '' 予備1
            For i = 0 To 4
                With .WFRaceInfo(i)
                    strSql = strSql & SS & "JyoCD" & i + 1 & SE & "='" & Replace(.JyoCD, "'", "''") & "',"      '' 競馬場コード
                    strSql = strSql & SS & "Kaiji" & i + 1 & SE & "='" & Replace(.Kaiji, "'", "''") & "',"      '' 開催回
                    strSql = strSql & SS & "Nichiji" & i + 1 & SE & "='" & Replace(.Nichiji, "'", "''") & "',"  '' 開催日目
                    strSql = strSql & SS & "RaceNum" & i + 1 & SE & "='" & Replace(.RaceNum, "'", "''") & "',"  '' レース番号
                End With ' WFRaceInfo()
            Next i
            strSql = strSql & SS & "reserved2" & SE & "='" & Replace(.reserved2, "'", "''") & "',"  '' 予備2
            For i = 0 To 4
                With .WFYukoHyoInfo(i)
                    strSql = strSql & SS & "YukoHyosu" & i + 1 & SE & "='" & Replace(.Yuko_Hyo, "'", "''") & "',"  '' 有効票数
                End With ' WFYukoHyoInfo()
            Next i
            strSql = strSql & SS & "HenkanFlag" & SE & "='" & Replace(.HenkanFlag, "'", "''") & "',"              '' 返還フラグ
            strSql = strSql & SS & "FuseirituFlag" & SE & "='" & Replace(.FuseiritsuFlag, "'", "''") & "',"       '' 不成立フラグ
            strSql = strSql & SS & "TekichunashiFlag" & SE & "='" & Replace(.TekichunashiFlag, "'", "''") & "',"  '' 的中無フラグ
            strSql = strSql & SS & "CarryoverSyoki" & SE & "='" & Replace(.COShoki, "'", "''") & "',"             '' キャリーオーバー金額初期
            strSql = strSql & SS & "CarryoverZandaka" & SE & "='" & Replace(.COZanDaka, "'", "''") & "',"         '' キャリーオーバー金額残高

            strSql = Left(strSql, Len(strSql) - 1) ' Delete Last Comma

            With .KaisaiDate
                strSql = strSql & " WHERE " & SS & "Year" & SE & "='" & Replace(.Year, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "MonthDay" & SE & "='" & Replace(.Month & .Day, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "MakeDate" & SE & " <= '" & Replace(strMakeDate, "'", "''") & "'"
                System.Diagnostics.Debug.WriteLine("UPDATE JYUSYOSIKI_HEAD : " & .Year & .Month & .Day)
            End With ' KaisaiDate
        End With

        gCon.Execute(strSql)

        For i = 0 To 242
            If mBuf.WFPayInfo(i).Kumiban <> "          " Then
                strSql = "UPDATE JYUSYOSIKI SET "
                With mBuf
                    strSql = strSql & SS & "MakeDate" & SE & "='" & Replace(strMakeDate, "'", "''") & "',"          '' 年月日
                    With .KaisaiDate
                        strSql = strSql & SS & "Year" & SE & "='" & Replace(.Year, "'", "''") & "',"                '' 開催年
                        strSql = strSql & SS & "MonthDay" & SE & "='" & Replace(.Month & .Day, "'", "''") & "',"    '' 開催月日
                    End With ' KaisaiDate
                    With .WFPayInfo(i)
                        strSql = strSql & SS & "Kumi" & SE & "='" & Replace(.Kumiban, "'", "''") & "',"             '' 組番
                        strSql = strSql & SS & "PayJyushosiki" & SE & "='" & Replace(.Pay, "'", "''") & "',"        '' 重勝式払戻金
                        strSql = strSql & SS & "TekichuHyo" & SE & "='" & Replace(.Tekichu_Hyo, "'", "''") & "',"   '' 的中票数
                    End With ' WFPayInfo

                    strSql = Left(strSql, Len(strSql) - 1) ' Delete Last Comma

                    With .KaisaiDate
                        strSql = strSql & " WHERE " & SS & "Year" & SE & "='" & Replace(.Year, "'", "''") & "'"
                        strSql = strSql & " AND " & SS & "MonthDay" & SE & "='" & Replace(.Month & .Day, "'", "''") & "'"
                        strSql = strSql & " AND " & SS & "Kumi" & SE & "='" & Replace(mBuf.WFPayInfo(i).Kumiban, "'", "''") & "'"
                        strSql = strSql & " AND " & SS & "MakeDate" & SE & " <= '" & Replace(strMakeDate, "'", "''") & "'"
                        System.Diagnostics.Debug.WriteLine("UPDATE JYUSYOSIKI : " & .Year & .Month & .Day & mBuf.WFPayInfo(i).Kumiban)
                    End With ' KaisaiDate
                End With
                gCon.Execute(strSql)
            End If
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