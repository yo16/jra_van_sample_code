Option Strict Off
Option Explicit On
Option Compare Binary
Friend Class clsImportHN
	' @(h) clsReadHN.cls
	' @(s)
	' JVData "HN" データベース登録クラス
	'
	
	Private mBuf As JV_HN_HANSYOKU '' 繁殖馬マスタ構造体
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
        strSql = "SELECT * FROM HANSYOKU"
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
    ' 引き数    : lBuf - JVData 識別子"HN" の１行
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
            mRS.Fields("HansyokuNum").Value = .HansyokuNum '' 繁殖登録番号
            mRS.Fields("reserved").Value = .reserved '' 予備
            mRS.Fields("KettoNum").Value = .KettoNum '' 血統登録番号
            mRS.Fields("DelKubun").Value = .DelKubun '' 繁殖馬抹消区分(現在は予備として使用)
            mRS.Fields("Bamei").Value = .Bamei '' 馬名
            mRS.Fields("BameiKana").Value = .BameiKana '' 馬名半角カナ
            mRS.Fields("BameiEng").Value = .BameiEng '' 馬名欧字
            mRS.Fields("BirthYear").Value = .BirthYear '' 生年
            mRS.Fields("SexCD").Value = .SexCD '' 性別コード
            mRS.Fields("HinsyuCD").Value = .HinsyuCD '' 品種コード
            mRS.Fields("KeiroCD").Value = .KeiroCD '' 毛色コード
            mRS.Fields("HansyokuMochiKubun").Value = .HansyokuMochiKubun '' 繁殖馬持込区分
            mRS.Fields("ImportYear").Value = .ImportYear '' 輸入年
            mRS.Fields("SanchiName").Value = .SanchiName '' 産地名
            mRS.Fields("HansyokuFNum").Value = .HansyokuFNum '' 父馬繁殖登録番号
            mRS.Fields("HansyokuMNum").Value = .HansyokuMNum '' 母馬繁殖登録番号
        End With

        With mBuf
            System.Diagnostics.Debug.WriteLine("Insert HANSYOKU : " & .HansyokuNum)
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

        strSql = "UPDATE HANSYOKU SET "
        With mBuf
            With .head
                strSql = strSql & SS & "RecordSpec" & SE & "='" & Replace(.RecordSpec, "'", "''") & "'," '' レコード種別
                strSql = strSql & SS & "DataKubun" & SE & "='" & Replace(.DataKubun, "'", "''") & "'," '' データ区分
                strSql = strSql & SS & "MakeDate" & SE & "= '" & Replace(strMakeDate, "'", "''") & "'," '' 年月日
            End With ' head
            strSql = strSql & SS & "HansyokuNum" & SE & "='" & Replace(.HansyokuNum, "'", "''") & "'," '' 繁殖登録番号
            strSql = strSql & SS & "reserved" & SE & "='" & Replace(.reserved, "'", "''") & "'," '' 予備
            strSql = strSql & SS & "KettoNum" & SE & "='" & Replace(.KettoNum, "'", "''") & "'," '' 血統登録番号
            strSql = strSql & SS & "DelKubun" & SE & "='" & Replace(.DelKubun, "'", "''") & "'," '' 繁殖馬抹消区分(現在は予備として使用)
            strSql = strSql & SS & "Bamei" & SE & "='" & Replace(.Bamei, "'", "''") & "'," '' 馬名
            strSql = strSql & SS & "BameiKana" & SE & "='" & Replace(.BameiKana, "'", "''") & "'," '' 馬名半角カナ
            strSql = strSql & SS & "BameiEng" & SE & "='" & Replace(.BameiEng, "'", "''") & "'," '' 馬名欧字
            strSql = strSql & SS & "BirthYear" & SE & "='" & Replace(.BirthYear, "'", "''") & "'," '' 生年
            strSql = strSql & SS & "SexCD" & SE & "='" & Replace(.SexCD, "'", "''") & "'," '' 性別コード
            strSql = strSql & SS & "HinsyuCD" & SE & "='" & Replace(.HinsyuCD, "'", "''") & "'," '' 品種コード
            strSql = strSql & SS & "KeiroCD" & SE & "='" & Replace(.KeiroCD, "'", "''") & "'," '' 毛色コード
            strSql = strSql & SS & "HansyokuMochiKubun" & SE & "='" & Replace(.HansyokuMochiKubun, "'", "''") & "'," '' 繁殖馬持込区分
            strSql = strSql & SS & "ImportYear" & SE & "='" & Replace(.ImportYear, "'", "''") & "'," '' 輸入年
            strSql = strSql & SS & "SanchiName" & SE & "='" & Replace(.SanchiName, "'", "''") & "'," '' 産地名
            strSql = strSql & SS & "HansyokuFNum" & SE & "='" & Replace(.HansyokuFNum, "'", "''") & "'," '' 父馬繁殖登録番号
            strSql = strSql & SS & "HansyokuMNum" & SE & "='" & Replace(.HansyokuMNum, "'", "''") & "'," '' 母馬繁殖登録番号

            strSql = Left(strSql, Len(strSql) - 1) ' Delete Last Comma

            strSql = strSql & " WHERE " & SS & "HansyokuNum" & SE & "='" & Replace(.HansyokuNum, "'", "''") & "'"
            strSql = strSql & " AND " & SS & "MakeDate" & SE & " <= '" & Replace(strMakeDate, "'", "''") & "'"
        End With

        With mBuf
            System.Diagnostics.Debug.WriteLine("UPDATE HANSYOKU : " & .HansyokuNum)
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