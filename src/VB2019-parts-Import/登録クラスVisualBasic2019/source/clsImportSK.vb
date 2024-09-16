Option Strict Off
Option Explicit On
Option Compare Binary
Friend Class clsImportSK
	' @(h) clsReadSK.cls
	' @(s)
	' JVData "SK" データベース登録クラス
	'
	
	Private mBuf As JV_SK_SANKU '' 産駒マスタ構造体
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
        strSql = "SELECT * FROM SANKU"
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
    ' 引き数    : lBuf - JVData 識別子"SK" の１行
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
            mRS.Fields("KettoNum").Value = .KettoNum '' 血統登録番号
            With .BirthDate
                mRS.Fields("BirthDate").Value = .Year & .Month & .Day '' 年月日
            End With ' BirthDate
            mRS.Fields("SexCD").Value = .SexCD '' 性別コード
            mRS.Fields("HinsyuCD").Value = .HinsyuCD '' 品種コード
            mRS.Fields("KeiroCD").Value = .KeiroCD '' 毛色コード
            mRS.Fields("SankuMochiKubun").Value = .SankuMochiKubun '' 産駒持込区分
            mRS.Fields("ImportYear").Value = .ImportYear '' 輸入年
            mRS.Fields("BreederCode").Value = .BreederCode '' 生産者コード
            mRS.Fields("SanchiName").Value = .SanchiName '' 産地名
            mRS.Fields("FNum").Value = .HansyokuNum(0)
            mRS.Fields("MNum").Value = .HansyokuNum(1)
            mRS.Fields("FFNum").Value = .HansyokuNum(2)
            mRS.Fields("FMNum").Value = .HansyokuNum(3)
            mRS.Fields("MFNum").Value = .HansyokuNum(4)
            mRS.Fields("MMNum").Value = .HansyokuNum(5)
            mRS.Fields("FFFNum").Value = .HansyokuNum(6)
            mRS.Fields("FFMNum").Value = .HansyokuNum(7)
            mRS.Fields("FMFNum").Value = .HansyokuNum(8)
            mRS.Fields("FMMNum").Value = .HansyokuNum(9)
            mRS.Fields("MFFNum").Value = .HansyokuNum(10)
            mRS.Fields("MFMNum").Value = .HansyokuNum(11)
            mRS.Fields("MMFNum").Value = .HansyokuNum(12)
            mRS.Fields("MMMNum").Value = .HansyokuNum(13)
        End With

        With mBuf
            System.Diagnostics.Debug.WriteLine("Insert SANKU : " & .KettoNum)
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
    Public Function UpdateDB(ByRef strMakeDate As String) As Boolean
        On Error GoTo ErrorHandler
        Dim i As Short '' ループカウンタ
        Dim j As Short '' ループカウンタ
        Dim k As Short '' ループカウンタ
        Dim strSql As String '' SQL文

        gCon.BeginTrans()
        System.Diagnostics.Debug.WriteLine("BeginTrans")

        strSql = "UPDATE SANKU SET "
        With mBuf
            With .head
                strSql = strSql & SS & "RecordSpec" & SE & "='" & Replace(.RecordSpec, "'", "''") & "'," '' レコード種別
                strSql = strSql & SS & "DataKubun" & SE & "='" & Replace(.DataKubun, "'", "''") & "'," '' データ区分
                strSql = strSql & SS & "MakeDate" & SE & "='" & Replace(strMakeDate, "'", "''") & "'," '' 年月日
            End With ' head
            strSql = strSql & SS & "KettoNum" & SE & "='" & Replace(.KettoNum, "'", "''") & "'," '' 血統登録番号
            With .BirthDate
                strSql = strSql & SS & "BirthDate" & SE & "='" & Replace(.Year & .Month & .Day, "'", "''") & "'," '' 年月日
            End With ' BirthDate
            strSql = strSql & SS & "SexCD" & SE & "='" & Replace(.SexCD, "'", "''") & "'," '' 性別コード
            strSql = strSql & SS & "HinsyuCD" & SE & "='" & Replace(.HinsyuCD, "'", "''") & "'," '' 品種コード
            strSql = strSql & SS & "KeiroCD" & SE & "='" & Replace(.KeiroCD, "'", "''") & "'," '' 毛色コード
            strSql = strSql & SS & "SankuMochiKubun" & SE & "='" & Replace(.SankuMochiKubun, "'", "''") & "'," '' 産駒持込区分
            strSql = strSql & SS & "ImportYear" & SE & "='" & Replace(.ImportYear, "'", "''") & "'," '' 輸入年
            strSql = strSql & SS & "BreederCode" & SE & "='" & Replace(.BreederCode, "'", "''") & "'," '' 生産者コード
            strSql = strSql & SS & "SanchiName" & SE & "='" & Replace(.SanchiName, "'", "''") & "'," '' 産地名
            strSql = strSql & SS & "FNum" & SE & "='" & Replace(.HansyokuNum(0), "'", "''") & "'," '' 父繁殖登録番号
            strSql = strSql & SS & "MNum" & SE & "='" & Replace(.HansyokuNum(1), "'", "''") & "'," '' 母繁殖登録番号
            strSql = strSql & SS & "FFNum" & SE & "='" & Replace(.HansyokuNum(2), "'", "''") & "'," '' 父父繁殖登録番号
            strSql = strSql & SS & "FMNum" & SE & "='" & Replace(.HansyokuNum(3), "'", "''") & "'," '' 父母繁殖登録番号
            strSql = strSql & SS & "MFNum" & SE & "='" & Replace(.HansyokuNum(4), "'", "''") & "'," '' 母父繁殖登録番号
            strSql = strSql & SS & "MMNum" & SE & "='" & Replace(.HansyokuNum(5), "'", "''") & "'," '' 母母繁殖登録番号
            strSql = strSql & SS & "FFFNum" & SE & "='" & Replace(.HansyokuNum(6), "'", "''") & "'," '' 父父父繁殖登録番号
            strSql = strSql & SS & "FFMNum" & SE & "='" & Replace(.HansyokuNum(7), "'", "''") & "'," '' 父父母繁殖登録番号
            strSql = strSql & SS & "FMFNum" & SE & "='" & Replace(.HansyokuNum(8), "'", "''") & "'," '' 父母父繁殖登録番号
            strSql = strSql & SS & "FMMNum" & SE & "='" & Replace(.HansyokuNum(9), "'", "''") & "'," '' 父母母繁殖登録番号
            strSql = strSql & SS & "MFFNum" & SE & "='" & Replace(.HansyokuNum(10), "'", "''") & "'," '' 母父父繁殖登録番号
            strSql = strSql & SS & "MFMNum" & SE & "='" & Replace(.HansyokuNum(11), "'", "''") & "'," '' 母父母繁殖登録番号
            strSql = strSql & SS & "MMFNum" & SE & "='" & Replace(.HansyokuNum(12), "'", "''") & "'," '' 母母父繁殖登録番号
            strSql = strSql & SS & "MMMNum" & SE & "='" & Replace(.HansyokuNum(13), "'", "''") & "'," '' 母母母繁殖登録番号

            strSql = Left(strSql, Len(strSql) - 1) ' Delete Last Comma

            strSql = strSql & " WHERE " & SS & "KettoNum" & SE & "='" & Replace(.KettoNum, "'", "''") & "'"
            strSql = strSql & " AND " & SS & "MakeDate" & SE & " <= '" & Replace(strMakeDate, "'", "''") & "'"
        End With
        With mBuf
            System.Diagnostics.Debug.WriteLine("UPDATE SANKU : " & .KettoNum)
        End With ' id

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