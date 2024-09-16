Option Strict Off
Option Explicit On
Option Compare Binary
Friend Class clsImportWE
	' @(h) clsReadWE.cls
	' @(s)
	' JVData "WE" データベース登録クラス
	'
	
	Private mBuf As JV_WE_WEATHER ''天候馬場状態
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
        strSql = "SELECT * FROM TENKO_BABA"
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
    ' 引き数    : lBuf - JVData 識別子"WE" の１行
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
            With .HappyoTime
                mRS.Fields("HappyoTime").Value = .Month & .Day & .Hour & .Minute '' 月日時分
            End With ' HappyoTime
            mRS.Fields("HenkoID").Value = .HenkoID '' 変更識別
            With .TenkoBaba
                mRS.Fields("AtoTenkoCD").Value = .TenkoCD '' 天候コード
                mRS.Fields("AtoSibaBabaCD").Value = .SibaBabaCD '' 芝馬場状態コード
                mRS.Fields("AtoDirtBabaCD").Value = .DirtBabaCD '' ダート馬場状態コード
            End With ' TenkoBaba
            With .TenkoBabaBefore
                mRS.Fields("MaeTenkoCD").Value = .TenkoCD '' 天候コード
                mRS.Fields("MaeSibaBabaCD").Value = .SibaBabaCD '' 芝馬場状態コード
                mRS.Fields("MaeDirtBabaCD").Value = .DirtBabaCD '' ダート馬場状態コード
            End With ' TenkoBabaBefore
        End With

        With mBuf.id
            System.Diagnostics.Debug.WriteLine("Insert JV_WE_WEATHER : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & mBuf.HappyoTime.Month & mBuf.HappyoTime.Day & mBuf.HappyoTime.Hour & mBuf.HappyoTime.Minute & mBuf.HenkoID)
        End With ' id

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

        strSql = "UPDATE TENKO_BABA SET "
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
            With .HappyoTime
                strSql = strSql & SS & "HappyoTime" & SE & "='" & Replace(.Month & .Day & .Hour & .Minute, "'", "''") & "',"
            End With ' HappyoTime
            strSql = strSql & SS & "HenkoID" & SE & "='" & Replace(.HenkoID, "'", "''") & "'," '' 変更識別
            With .TenkoBaba
                strSql = strSql & SS & "AtoTenkoCD" & SE & "='" & Replace(.TenkoCD, "'", "''") & "'," '' 天候コード
                strSql = strSql & SS & "AtoSibaBabaCD" & SE & "='" & Replace(.SibaBabaCD, "'", "''") & "'," '' 芝馬場状態コード
                strSql = strSql & SS & "AtoDirtBabaCD" & SE & "='" & Replace(.DirtBabaCD, "'", "''") & "'," '' ダート馬場状態コード
            End With ' TenkoBaba
            With .TenkoBabaBefore
                strSql = strSql & SS & "MaeTenkoCD" & SE & "='" & Replace(.TenkoCD, "'", "''") & "'," '' 天候コード
                strSql = strSql & SS & "MaeSibaBabaCD" & SE & "='" & Replace(.SibaBabaCD, "'", "''") & "'," '' 芝馬場状態コード
                strSql = strSql & SS & "MaeDirtBabaCD" & SE & "='" & Replace(.DirtBabaCD, "'", "''") & "'" '' ダート馬場状態コード
            End With ' TenkoBabaBefore
            With .id
                strSql = strSql & " WHERE " & SS & "Year" & SE & "='" & Replace(.Year, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "MonthDay" & SE & "='" & Replace(.MonthDay, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "JyoCD" & SE & "='" & Replace(.JyoCD, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "Kaiji" & SE & "='" & Replace(.Kaiji, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "Nichiji" & SE & "='" & Replace(.Nichiji, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "HappyoTime" & SE & "='" & Replace(mBuf.HappyoTime.Month & mBuf.HappyoTime.Day & mBuf.HappyoTime.Hour & mBuf.HappyoTime.Minute, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "HenkoID" & SE & "='" & Replace(mBuf.HenkoID, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "MakeDate" & SE & " <= '" & Replace(strMakeDate, "'", "''") & "'"
            End With
        End With

        gCon.Execute(strSql)

        With mBuf.id
            System.Diagnostics.Debug.WriteLine("UPDATE JV_WE_WEATHER : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & mBuf.HappyoTime.Month & mBuf.HappyoTime.Day & mBuf.HappyoTime.Hour & mBuf.HappyoTime.Minute & mBuf.HenkoID)
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