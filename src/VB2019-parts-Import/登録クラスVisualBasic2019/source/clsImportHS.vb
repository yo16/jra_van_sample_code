Option Strict Off
Option Explicit On
Option Compare Binary
Friend Class clsImportHS
	' @(h) clsReadHS.cls
	' @(s)
	' JVData "HS" データベース登録クラス
	'
	
	Private mBuf As JV_HS_SALE '' 競走馬市場取引価格構造体
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
        strSql = "SELECT * FROM SALE"
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
    ' 引き数    : lBuf - JVData 識別子"HS" の１行
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

        gCon.BeginTrans()
        System.Diagnostics.Debug.WriteLine("BeginTrans")

        mRS.AddNew()

        With mBuf
            With .head
                mRS.Fields("RecordSpec").Value = .RecordSpec             '' レコード種別
                mRS.Fields("DataKubun").Value = .DataKubun               '' データ区分
                With .MakeDate
                    mRS.Fields("MakeDate").Value = .Year & .Month & .Day '' 年月日
                End With ' MakeDate
            End With ' head
            mRS.Fields("KettoNum").Value = .KettoNum                     ''血統登録番号
            mRS.Fields("HansyokuFNum").Value = .HansyokuFNum             ''父馬繁殖登録番号
            mRS.Fields("HansyokuMNum").Value = .HansyokuMNum             ''母馬繁殖登録番号
            mRS.Fields("BirthYear").Value = .BirthYear                   ''生年
            mRS.Fields("SaleCode").Value = .SaleCode                     ''主催者・市場コード
            mRS.Fields("SaleHostName").Value = .SaleHostName             ''主催者名称
            mRS.Fields("SaleName").Value = .SaleName                     ''市場の名称
            With .FromDate
                mRS.Fields("FromDate").Value = .Year & .Month & .Day     ''市場の開催期間(開始日)
            End With ' FromDate
            With .ToDate
                mRS.Fields("ToDate").Value = .Year & .Month & .Day       ''市場の開催期間(終了日)
            End With ' ToDate
            mRS.Fields("Barei").Value = .Barei                           ''取引時の競走馬の年齢
            mRS.Fields("Price").Value = .Price                           ''取引価格
        End With

        With mBuf
            System.Diagnostics.Debug.WriteLine("Insert SALE : " & .KettoNum & .SaleCode & .FromDate.Year &  .FromDate.Month &  .FromDate.Day)
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
        Dim strSql As String '' SQL文

        gCon.BeginTrans()
        System.Diagnostics.Debug.WriteLine("BeginTrans")

        strSql = "UPDATE SALE SET "
        With mBuf
            With .head
                strSql = strSql & SS & "RecordSpec" & SE & "='" & Replace(.RecordSpec, "'", "''") & "',"          '' レコード種別
                strSql = strSql & SS & "DataKubun" & SE & "='" & Replace(.DataKubun, "'", "''") & "',"            '' データ区分
                strSql = strSql & SS & "MakeDate" & SE & "='" & Replace(strMakeDate, "'", "''") & "',"            '' 年月日
            End With ' head
            strSql = strSql & SS & "KettoNum" & SE & "='" & Replace(.KettoNum, "'", "''") & "',"                  ''血統登録番号
            strSql = strSql & SS & "HansyokuFNum" & SE & "='" & Replace(.HansyokuFNum, "'", "''") & "',"          ''父馬繁殖登録番号
            strSql = strSql & SS & "HansyokuMNum" & SE & "='" & Replace(.HansyokuMNum, "'", "''") & "',"          ''母馬繁殖登録番号
            strSql = strSql & SS & "BirthYear" & SE & "='" & Replace(.BirthYear, "'", "''") & "',"                ''生年
            strSql = strSql & SS & "SaleCode" & SE & "='" & Replace(.SaleCode, "'", "''") & "',"                  ''主催者・市場コード
            strSql = strSql & SS & "SaleHostName" & SE & "='" & Replace(.SaleHostName, "'", "''") & "',"          ''主催者名称
            strSql = strSql & SS & "SaleName" & SE & "='" & Replace(.SaleName, "'", "''") & "',"                  ''市場の名称
            With .FromDate
                strSql = strSql & SS & "FromDate" & SE & "='" & Replace(.Year & .Month & .Day, "'", "''") & "',"  ''市場の開催期間(開始日)
            End With ' FromDate
            With .ToDate
                strSql = strSql & SS & "ToDate" & SE & "='" & Replace(.Year & .Month & .Day, "'", "''") & "',"    ''市場の開催期間(終了日)
            End With ' ToDate
            strSql = strSql & SS & "Barei" & SE & "='" & Replace(.Barei, "'", "''") & "',"                        '' 取引時の競走馬の年齢
            strSql = strSql & SS & "Price" & SE & "='" & Replace(.Price, "'", "''") & "',"                        '' 取引価格

            strSql = Left(strSql, Len(strSql) - 1) ' Delete Last Comma

            strSql = strSql & " WHERE " & SS & "KettoNum" & SE & "='" & Replace(.KettoNum, "'", "''") & "'"                 ''血統登録番号
            strSql = strSql & " AND " & SS & "SaleCode" & SE & "='" & Replace(.SaleCode, "'", "''") & "'"                   ''主催者・市場コード
            With .FromDate
                strSql = strSql & "AND " & SS & "FromDate" & SE & "='" & Replace(.Year & .Month & .Day, "'", "''") & "'"   ''市場の開催期間(開始日)
            End With ' FromDate
            strSql = strSql & " AND " & SS & "MakeDate" & SE & " <= '" & Replace(strMakeDate, "'", "''") & "'"              '' 年月日
        End With
        With mBuf
            System.Diagnostics.Debug.WriteLine("UPDATE SALE : " & .KettoNum & .SaleCode & .FromDate.Year &  .FromDate.Month &  .FromDate.Day)
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