Option Strict Off
Option Explicit On
Option Compare Binary
Friend Class clsImportHC
	' @(h) clsReadHC.cls
	' @(s)
	' JVData "HC" データベース登録クラス
	'
	
	Private mBuf As JV_HC_HANRO '' 坂道調教構造体
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
        strSql = "SELECT * FROM HANRO"
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
        System.Diagnostics.Debug.WriteLine("mRS.Close")
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
    ' 引き数    : lBuf - JVData 識別子"HC" の１行
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
            mRS.Fields("TresenKubun").Value = .TresenKubun '' トレセン区分
            With .ChokyoDate
                mRS.Fields("ChokyoDate").Value = .Year & .Month & .Day '' 年月日
            End With ' ChokyoDate
            mRS.Fields("ChokyoTime").Value = .ChokyoTime '' 調教時刻
            mRS.Fields("KettoNum").Value = .KettoNum '' 血統登録番号
            mRS.Fields("HaronTime4").Value = .HaronTime4 '' 4ハロンタイム合計(800M-0M)
            mRS.Fields("LapTime4").Value = .LapTime4 '' ラップタイム(800M-600M)
            mRS.Fields("HaronTime3").Value = .HaronTime3 '' 3ハロンタイム合計(600M-0M)
            mRS.Fields("LapTime3").Value = .LapTime3 '' ラップタイム(600M-400M)
            mRS.Fields("HaronTime2").Value = .HaronTime2 '' 2ハロンタイム合計(400M-0M)
            mRS.Fields("LapTime2").Value = .LapTime2 '' ラップタイム(400M-200M)
            mRS.Fields("LapTime1").Value = .LapTime1 '' ラップタイム(200M-0M)
        End With

        With mBuf
            System.Diagnostics.Debug.WriteLine("INSERT HANRO : " & .TresenKubun & .ChokyoDate.Year & .ChokyoDate.Month & .ChokyoDate.Day & .ChokyoTime & .KettoNum)
        End With ' mBuf

        mRS.Update()

        gCon.CommitTrans()
        System.Diagnostics.Debug.WriteLine("CommitTrans")

        InsertDB = True

ExitHandler:
        Exit Function
ErrorHandler:
        InsertDB = False
        mRS.CancelUpdate()
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

        strSql = "UPDATE HANRO SET "
        With mBuf
            With .head
                strSql = strSql & SS & "RecordSpec" & SE & "='" & Replace(.RecordSpec, "'", "''") & "'," '' レコード種別
                strSql = strSql & SS & "DataKubun" & SE & "='" & Replace(.DataKubun, "'", "''") & "'," '' データ区分
                strSql = strSql & SS & "MakeDate" & SE & "= '" & Replace(strMakeDate, "'", "''") & "'," '' 年月日
            End With ' head
            strSql = strSql & SS & "TresenKubun" & SE & "='" & Replace(.TresenKubun, "'", "''") & "'," '' トレセン区分
            With .ChokyoDate
                strSql = strSql & SS & "ChokyoDate" & SE & "='" & Replace(.Year & .Month & .Day, "'", "''") & "'," '' 年月日
            End With ' ChokyoDate
            strSql = strSql & SS & "ChokyoTime" & SE & "='" & Replace(.ChokyoTime, "'", "''") & "'," '' 調教時刻
            strSql = strSql & SS & "KettoNum" & SE & "='" & Replace(.KettoNum, "'", "''") & "'," '' 血統登録番号
            strSql = strSql & SS & "HaronTime4" & SE & "='" & Replace(.HaronTime4, "'", "''") & "'," '' 4ハロンタイム合計(800M-0M)
            strSql = strSql & SS & "LapTime4" & SE & "='" & Replace(.LapTime4, "'", "''") & "'," '' ラップタイム(800M-600M)
            strSql = strSql & SS & "HaronTime3" & SE & "='" & Replace(.HaronTime3, "'", "''") & "'," '' 3ハロンタイム合計(600M-0M)
            strSql = strSql & SS & "LapTime3" & SE & "='" & Replace(.LapTime3, "'", "''") & "'," '' ラップタイム(600M-400M)
            strSql = strSql & SS & "HaronTime2" & SE & "='" & Replace(.HaronTime2, "'", "''") & "'," '' 2ハロンタイム合計(400M-0M)
            strSql = strSql & SS & "LapTime2" & SE & "='" & Replace(.LapTime2, "'", "''") & "'," '' ラップタイム(400M-200M)
            strSql = strSql & SS & "LapTime1" & SE & "='" & Replace(.LapTime1, "'", "''") & "'," '' ラップタイム(200M-0M)

            strSql = Left(strSql, Len(strSql) - 1) ' Delete Last Comma

            strSql = strSql & " WHERE " & SS & "TresenKubun" & SE & "='" & Replace(.TresenKubun, "'", "''") & "'"
            With .ChokyoDate
                strSql = strSql & " AND " & SS & "ChokyoDate" & SE & "='" & Replace(.Year & .Month & .Day, "'", "''") & "'"
            End With ' ChokyoDate
            strSql = strSql & " AND " & SS & "ChokyoTime" & SE & "='" & Replace(.ChokyoTime, "'", "''") & "'"
            strSql = strSql & " AND " & SS & "KettoNum" & SE & "='" & Replace(.KettoNum, "'", "''") & "'"
            strSql = strSql & " AND " & SS & "MakeDate" & SE & " <= '" & Replace(strMakeDate, "'", "''") & "'"
        End With ' mBuf

        With mBuf
            System.Diagnostics.Debug.WriteLine("UPDATE HANRO : " & .TresenKubun & .ChokyoDate.Year & .ChokyoDate.Month & .ChokyoDate.Day & .ChokyoTime & .KettoNum)
        End With ' mBuf

        gCon.Execute(strSql)

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