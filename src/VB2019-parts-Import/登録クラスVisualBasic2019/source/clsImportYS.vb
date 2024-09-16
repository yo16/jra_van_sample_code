Option Strict Off
Option Explicit On
Option Compare Binary
Friend Class clsImportYS
	' @(h) clsReadYS.cls
	' @(s)
	' JVData "YS" データベース登録クラス
	'
	
	Private mBuf As JV_YS_SCHEDULE ''年間スケジュール構造体
	Private mRS As ADODB.Recordset
	
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
        strSql = "SELECT * FROM SCHEDULE"
        mRS = New ADODB.Recordset()
        mRS.Open(strSql, gCon, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockPessimistic)

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
        mRS.Close()

        mRS = Nothing


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
    ' 引き数    : lBuf - JVData 識別子"YS" の１行
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

        mRS.AddNew()

        With mBuf
            With .head
                mRS.Fields("RecordSpec").Value = .RecordSpec '' レコード種別
                mRS.Fields("DataKubun").Value = .DataKubun '' データ区分
                With .MakeDate
                    mRS.Fields("MakeDate").Value = .Year & .Month & .Day '' 年月日
                End With ' MakeDate
            End With ' head
            With .id
                mRS.Fields("Year").Value = .Year '' 開催年
                mRS.Fields("MonthDay").Value = .MonthDay '' 開催月日
                mRS.Fields("JyoCD").Value = .JyoCD '' 競馬場コード
                mRS.Fields("Kaiji").Value = .Kaiji '' 開催回第N回
                mRS.Fields("Nichiji").Value = .Nichiji '' 開催日目N日目
            End With ' id
            mRS.Fields("YoubiCD").Value = .YoubiCD '' 曜日コード
            For i = 0 To 2
                With .JyusyoInfo(i)
                    mRS.Fields("Jyusyo" & i + 1 & "TokuNum").Value = .TokuNum '' 特別競走番号
                    mRS.Fields("Jyusyo" & i + 1 & "Hondai").Value = .Hondai '' 競走名本題
                    mRS.Fields("Jyusyo" & i + 1 & "Ryakusyo10").Value = .Ryakusyo10 '' 競走名略称10字
                    mRS.Fields("Jyusyo" & i + 1 & "Ryakusyo6").Value = .Ryakusyo6 '' 競走名略称6字
                    mRS.Fields("Jyusyo" & i + 1 & "Ryakusyo3").Value = .Ryakusyo3 '' 競走名略称3字
                    mRS.Fields("Jyusyo" & i + 1 & "Nkai").Value = .Nkai '' 重賞回次第N回
                    mRS.Fields("Jyusyo" & i + 1 & "GradeCD").Value = .GradeCD '' グレードコード
                    mRS.Fields("Jyusyo" & i + 1 & "SyubetuCD").Value = .SyubetuCD '' 競走種別コード
                    mRS.Fields("Jyusyo" & i + 1 & "KigoCD").Value = .KigoCD '' 競走記号コード
                    mRS.Fields("Jyusyo" & i + 1 & "JyuryoCD").Value = .JyuryoCD '' 重量種別コード
                    mRS.Fields("Jyusyo" & i + 1 & "Kyori").Value = .Kyori '' 距離
                    mRS.Fields("Jyusyo" & i + 1 & "TrackCD").Value = .TrackCD '' トラックコード
                End With ' JyusyoInfo
            Next i
        End With

        With mBuf.id
            System.Diagnostics.Debug.WriteLine("Insert SCHEDULE : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji)
        End With ' id

        mRS.Update()

        gCon.CommitTrans()
        System.Diagnostics.Debug.WriteLine("CommitTrans")

        InsertDB = True

ExitHandler:
        Exit Function
ErrorHandler:
        InsertDB = False
        gCon.RollbackTrans()
        mRS.CancelUpdate()
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

        System.Diagnostics.Debug.WriteLine("BeginTrans")
        gCon.BeginTrans()

        strSql = "UPDATE SCHEDULE SET "
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
            End With ' id
            strSql = strSql & SS & "YoubiCD" & SE & "='" & Replace(.YoubiCD, "'", "''") & "'," '' 曜日コード
            With .JyusyoInfo(0)
                strSql = strSql & SS & "Jyusyo1TokuNum" & SE & "='" & Replace(.TokuNum, "'", "''") & "'," '' 特別競走番号
                strSql = strSql & SS & "Jyusyo1Hondai" & SE & "='" & Replace(.Hondai, "'", "''") & "'," '' 競走名本題
                strSql = strSql & SS & "Jyusyo1Ryakusyo10" & SE & "='" & Replace(.Ryakusyo10, "'", "''") & "'," '' 競走名略称10字
                strSql = strSql & SS & "Jyusyo1Ryakusyo6" & SE & "='" & Replace(.Ryakusyo6, "'", "''") & "'," '' 競走名略称6字
                strSql = strSql & SS & "Jyusyo1Ryakusyo3" & SE & "='" & Replace(.Ryakusyo3, "'", "''") & "'," '' 競走名略称3字
                strSql = strSql & SS & "Jyusyo1Nkai" & SE & "='" & Replace(.Nkai, "'", "''") & "'," '' 重賞回次第N回
                strSql = strSql & SS & "Jyusyo1GradeCD" & SE & "='" & Replace(.GradeCD, "'", "''") & "'," '' グレードコード
                strSql = strSql & SS & "Jyusyo1SyubetuCD" & SE & "='" & Replace(.SyubetuCD, "'", "''") & "'," '' 競走種別コード
                strSql = strSql & SS & "Jyusyo1KigoCD" & SE & "='" & Replace(.KigoCD, "'", "''") & "'," '' 競走記号コード
                strSql = strSql & SS & "Jyusyo1JyuryoCD" & SE & "='" & Replace(.JyuryoCD, "'", "''") & "'," '' 重量種別コード
                strSql = strSql & SS & "Jyusyo1Kyori" & SE & "='" & Replace(.Kyori, "'", "''") & "'," '' 距離
                strSql = strSql & SS & "Jyusyo1TrackCD" & SE & "='" & Replace(.TrackCD, "'", "''") & "'," '' トラックコード
            End With ' JyusyoInfo(0)
            With .JyusyoInfo(1)
                strSql = strSql & SS & "Jyusyo2TokuNum" & SE & "='" & Replace(.TokuNum, "'", "''") & "'," '' 特別競走番号
                strSql = strSql & SS & "Jyusyo2Hondai" & SE & "='" & Replace(.Hondai, "'", "''") & "'," '' 競走名本題
                strSql = strSql & SS & "Jyusyo2Ryakusyo10" & SE & "='" & Replace(.Ryakusyo10, "'", "''") & "'," '' 競走名略称10字
                strSql = strSql & SS & "Jyusyo2Ryakusyo6" & SE & "='" & Replace(.Ryakusyo6, "'", "''") & "'," '' 競走名略称6字
                strSql = strSql & SS & "Jyusyo2Ryakusyo3" & SE & "='" & Replace(.Ryakusyo3, "'", "''") & "'," '' 競走名略称3字
                strSql = strSql & SS & "Jyusyo2Nkai" & SE & "='" & Replace(.Nkai, "'", "''") & "'," '' 重賞回次第N回
                strSql = strSql & SS & "Jyusyo2GradeCD" & SE & "='" & Replace(.GradeCD, "'", "''") & "'," '' グレードコード
                strSql = strSql & SS & "Jyusyo2SyubetuCD" & SE & "='" & Replace(.SyubetuCD, "'", "''") & "'," '' 競走種別コード
                strSql = strSql & SS & "Jyusyo2KigoCD" & SE & "='" & Replace(.KigoCD, "'", "''") & "'," '' 競走記号コード
                strSql = strSql & SS & "Jyusyo2JyuryoCD" & SE & "='" & Replace(.JyuryoCD, "'", "''") & "'," '' 重量種別コード
                strSql = strSql & SS & "Jyusyo2Kyori" & SE & "='" & Replace(.Kyori, "'", "''") & "'," '' 距離
                strSql = strSql & SS & "Jyusyo2TrackCD" & SE & "='" & Replace(.TrackCD, "'", "''") & "'," '' トラックコード
            End With ' JyusyoInfo(1)
            With .JyusyoInfo(2)
                strSql = strSql & SS & "Jyusyo3TokuNum" & SE & "='" & Replace(.TokuNum, "'", "''") & "'," '' 特別競走番号
                strSql = strSql & SS & "Jyusyo3Hondai" & SE & "='" & Replace(.Hondai, "'", "''") & "'," '' 競走名本題
                strSql = strSql & SS & "Jyusyo3Ryakusyo10" & SE & "='" & Replace(.Ryakusyo10, "'", "''") & "'," '' 競走名略称10字
                strSql = strSql & SS & "Jyusyo3Ryakusyo6" & SE & "='" & Replace(.Ryakusyo6, "'", "''") & "'," '' 競走名略称6字
                strSql = strSql & SS & "Jyusyo3Ryakusyo3" & SE & "='" & Replace(.Ryakusyo3, "'", "''") & "'," '' 競走名略称3字
                strSql = strSql & SS & "Jyusyo3Nkai" & SE & "='" & Replace(.Nkai, "'", "''") & "'," '' 重賞回次第N回
                strSql = strSql & SS & "Jyusyo3GradeCD" & SE & "='" & Replace(.GradeCD, "'", "''") & "'," '' グレードコード
                strSql = strSql & SS & "Jyusyo3SyubetuCD" & SE & "='" & Replace(.SyubetuCD, "'", "''") & "'," '' 競走種別コード
                strSql = strSql & SS & "Jyusyo3KigoCD" & SE & "='" & Replace(.KigoCD, "'", "''") & "'," '' 競走記号コード
                strSql = strSql & SS & "Jyusyo3JyuryoCD" & SE & "='" & Replace(.JyuryoCD, "'", "''") & "'," '' 重量種別コード
                strSql = strSql & SS & "Jyusyo3Kyori" & SE & "='" & Replace(.Kyori, "'", "''") & "'," '' 距離
                strSql = strSql & SS & "Jyusyo3TrackCD" & SE & "='" & Replace(.TrackCD, "'", "''") & "'" '' トラックコード
            End With ' JyusyoInfo(2)
            With .id
                strSql = strSql & " WHERE " & SS & "Year" & SE & "='" & Replace(.Year, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "MonthDay" & SE & "='" & Replace(.MonthDay, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "JyoCD" & SE & "='" & Replace(.JyoCD, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "Kaiji" & SE & "='" & Replace(.Kaiji, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "Nichiji" & SE & "='" & Replace(.Nichiji, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "MakeDate" & SE & " <= '" & Replace(strMakeDate, "'", "''") & "'"
            End With
        End With
        gCon.Execute(strSql)
        With mBuf.id
            System.Diagnostics.Debug.WriteLine("UPDATE SCHEDULE : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji)
        End With ' id
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