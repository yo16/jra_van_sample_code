Option Strict Off
Option Explicit On
Option Compare Binary
Friend Class clsImportBN
	' @(h) clsReadBN.cls
	' @(s)
	' JVData "BN" データベース登録クラス
	'
	
	Private mBuf As JV_BN_BANUSI '' 馬主マスタ構造体
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
        strSql = "SELECT * FROM BANUSI"
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
    ' 引き数    : lBuf - JVData 識別子"BN" の１行
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
        Dim s1 As String = "" '' 先頭文字列

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
            mRS.Fields("BanusiCode").Value = .BanusiCode '' 馬主コード
            mRS.Fields("BanusiName_Co").Value = .BanusiName_Co '' 馬主名（法人格有）
            mRS.Fields("BanusiName").Value = .BanusiName '' 馬主名（法人格無）
            mRS.Fields("BanusiNameKana").Value = .BanusiNameKana '' 馬主名半角カナ
            mRS.Fields("BanusiNameEng").Value = .BanusiNameEng '' 馬主名欧字
            mRS.Fields("Fukusyoku").Value = .Fukusyoku '' 服色標示
            For i = 0 To 1
                With .HonRuikei(i)
                    If i = 0 Then s1 = "H"
                    If i = 1 Then s1 = "R"

                    mRS.Fields(s1 & "_SetYear").Value = .SetYear '' 設定年

                    mRS.Fields(s1 & "_HonSyokinTotal").Value = .HonSyokinTotal '' 本賞金合計

                    mRS.Fields(s1 & "_FukaSyokin").Value = .FukaSyokin '' 付加賞金合計
                    For j = 0 To 5

                        mRS.Fields(s1 & "_Chakukaisu" & j + 1).Value = .ChakuKaisu(j) '' 着回数
                    Next j
                End With ' HonRuikei
            Next i
        End With

        mRS.Update()

        With mBuf
            System.Diagnostics.Debug.WriteLine("INSERT BANUSI : " & .BanusiCode)
        End With ' mBuf

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

        strSql = "UPDATE BANUSI SET "
        With mBuf
            With .head

                strSql = strSql & SS & "RecordSpec" & SE & "='" & Replace(.RecordSpec, "'", "''") & "'," '' レコード種別

                strSql = strSql & SS & "DataKubun" & SE & "='" & Replace(.DataKubun, "'", "''") & "'," '' データ区分
                strSql = strSql & SS & "MakeDate" & SE & "= '" & Replace(strMakeDate, "'", "''") & "'," '' 年月日
            End With ' head
            strSql = strSql & SS & "BanusiCode" & SE & "='" & Replace(.BanusiCode, "'", "''") & "'," '' 馬主コード
            strSql = strSql & SS & "BanusiName_Co" & SE & "='" & Replace(.BanusiName_Co, "'", "''") & "'," '' 馬主名(法人格有)
            strSql = strSql & SS & "BanusiName" & SE & "='" & Replace(.BanusiName, "'", "''") & "'," '' 馬主名(法人格無)
            strSql = strSql & SS & "BanusiNameKana" & SE & "='" & Replace(.BanusiNameKana, "'", "''") & "'," '' 馬主名半角カナ
            strSql = strSql & SS & "BanusiNameEng" & SE & "='" & Replace(.BanusiNameEng, "'", "''") & "'," '' 馬主名欧字
            strSql = strSql & SS & "Fukusyoku" & SE & "='" & Replace(.Fukusyoku, "'", "''") & "'," '' 服色標示
            With .HonRuikei(0)
                strSql = strSql & SS & "H_SetYear" & SE & "='" & Replace(.SetYear, "'", "''") & "'," '' 設定年
                strSql = strSql & SS & "H_HonSyokinTotal" & SE & "='" & Replace(.HonSyokinTotal, "'", "''") & "'," '' 本賞金合計
                strSql = strSql & SS & "H_Fukasyokin" & SE & "='" & Replace(.FukaSyokin, "'", "''") & "'," '' 付加賞金合計
                For j = 0 To 5
                    strSql = strSql & SS & "H_Chakukaisu" & j + 1 & SE & "='" & Replace(.ChakuKaisu(j), "'", "''") & "'," '' 着回数
                Next j
            End With ' HonRuikei(0)
            With .HonRuikei(1)
                strSql = strSql & SS & "R_SetYear" & SE & "='" & Replace(.SetYear, "'", "''") & "'," '' 設定年
                strSql = strSql & SS & "R_HonSyokinTotal" & SE & "='" & Replace(.HonSyokinTotal, "'", "''") & "'," '' 本賞金合計
                strSql = strSql & SS & "R_Fukasyokin" & SE & "='" & Replace(.FukaSyokin, "'", "''") & "'," '' 付加賞金合計
                For j = 0 To 5
                    strSql = strSql & SS & "R_Chakukaisu" & j + 1 & SE & "='" & Replace(.ChakuKaisu(j), "'", "''") & "'," '' 着回数
                Next j
            End With ' HonRuikei(1)
            strSql = Left(strSql, Len(strSql) - 1) ' Delete Last Comma
            strSql = strSql & " WHERE " & SS & "BanusiCode" & SE & "='" & Replace(.BanusiCode, "'", "''") & "'"
            strSql = strSql & " AND " & SS & "MakeDate" & SE & " <= '" & Replace(strMakeDate, "'", "''") & "'"
        End With ' mBuf

        With mBuf
            System.Diagnostics.Debug.WriteLine("UPDATE BANUSI : " & .BanusiCode)
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