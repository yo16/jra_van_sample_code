Option Strict Off
Option Explicit On
Option Compare Binary
Friend Class clsImportRC
	' @(h) clsReadRC.cls
	' @(s)
	' JVData "RC" データベース登録クラス
	'
	
	Private mBuf As JV_RC_RECORD ''レコードマスタ構造体
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
		strSql = "SELECT * FROM RECORD"
		mRS = New ADODB.Recordset
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
    ' 引き数    : lBuf - JVData 識別子"RC" の１行
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
            mRS.Fields("RecInfoKubun").Value = .RecInfoKubun '' レコード識別区分
            With .id
                mRS.Fields("Year").Value = .Year '' 開催年
                mRS.Fields("MonthDay").Value = .MonthDay '' 開催月日
                mRS.Fields("JyoCD").Value = .JyoCD '' 競馬場コード
                mRS.Fields("Kaiji").Value = .Kaiji '' 開催回第N回
                mRS.Fields("Nichiji").Value = .Nichiji '' 開催日目N日目
                mRS.Fields("RaceNum").Value = .RaceNum '' レース番号
            End With ' id
            mRS.Fields("TokuNum").Value = .TokuNum '' 特別競走番号
            mRS.Fields("Hondai").Value = .Hondai '' 競走名本題
            mRS.Fields("GradeCD").Value = .GradeCD '' グレードコード
            mRS.Fields("SyubetuCD_TrackCD").Value = .SyubetuCD & .TrackCD '' 競走種別コード
            mRS.Fields("Kyori").Value = .Kyori '' 距離
            mRS.Fields("RecKubun").Value = .RecKubun '' レコード区分
            mRS.Fields("RecTime").Value = .RecTime '' レコードタイム
            With .TenkoBaba
                mRS.Fields("TenkoCD").Value = .TenkoCD '' 天候コード
                mRS.Fields("SibaBabaCD").Value = .SibaBabaCD '' 芝馬場状態コード
                mRS.Fields("DirtBabaCD").Value = .DirtBabaCD '' ダート馬場状態コード
            End With ' TenkoBaba
            For i = 0 To 2
                With .RecUmaInfo(i)
                    mRS.Fields("RecUmaKettoNum" & i + 1).Value = .KettoNum '' 血統登録番号
                    mRS.Fields("RecUmaBamei" & i + 1).Value = .Bamei '' 馬名
                    mRS.Fields("RecUmaUmaKigoCD" & i + 1).Value = .UmaKigoCD '' 馬記号コード
                    mRS.Fields("RecUmaSexCD" & i + 1).Value = .SexCD '' 性別コード
                    mRS.Fields("RecUmaChokyosiCode" & i + 1).Value = .ChokyosiCode '' 調教師コード
                    mRS.Fields("RecUmaChokyosiName" & i + 1).Value = .ChokyosiName '' 調教師名
                    mRS.Fields("RecUmaFutan" & i + 1).Value = .Futan '' 負担重量
                    mRS.Fields("RecUmaKisyuCode" & i + 1).Value = .KisyuCode '' 騎手コード
                    mRS.Fields("RecUmaKisyuName" & i + 1).Value = .KisyuName '' 騎手名
                End With ' RecUmaInfo
            Next i
        End With

        With mBuf.id
            System.Diagnostics.Debug.WriteLine("Insert RECORD : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum & mBuf.SyubetuCD & mBuf.TrackCD & mBuf.Kyori)
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

        strSql = "UPDATE RECORD SET "

        With mBuf
            With .head
                strSql = strSql & SS & "RecordSpec" & SE & "='" & Replace(.RecordSpec, "'", "''") & "'," '' レコード種別
                strSql = strSql & SS & "DataKubun" & SE & "='" & Replace(.DataKubun, "'", "''") & "'," '' データ区分
                strSql = strSql & SS & "MakeDate" & SE & "='" & Replace(strMakeDate, "'", "''") & "'," '' 年月日
            End With ' head
            strSql = strSql & SS & "RecInfoKubun" & SE & "='" & Replace(.RecInfoKubun, "'", "''") & "'," '' レコード識別区分
            With .id
                strSql = strSql & SS & "Year" & SE & "='" & Replace(.Year, "'", "''") & "'," '' 開催年
                strSql = strSql & SS & "MonthDay" & SE & "='" & Replace(.MonthDay, "'", "''") & "'," '' 開催月日
                strSql = strSql & SS & "JyoCD" & SE & "='" & Replace(.JyoCD, "'", "''") & "'," '' 競馬場コード
                strSql = strSql & SS & "Kaiji" & SE & "='" & Replace(.Kaiji, "'", "''") & "'," '' 開催回第N回
                strSql = strSql & SS & "Nichiji" & SE & "='" & Replace(.Nichiji, "'", "''") & "'," '' 開催日目N日目
                strSql = strSql & SS & "RaceNum" & SE & "='" & Replace(.RaceNum, "'", "''") & "'," '' レース番号
            End With ' id
            strSql = strSql & SS & "TokuNum" & SE & "='" & Replace(.TokuNum, "'", "''") & "'," '' 特別競走番号
            strSql = strSql & SS & "Hondai" & SE & "='" & Replace(.Hondai, "'", "''") & "'," '' 競走名本題
            strSql = strSql & SS & "GradeCD" & SE & "='" & Replace(.GradeCD, "'", "''") & "'," '' グレードコード
            strSql = strSql & SS & "SyubetuCD_TrackCD" & SE & "='" & Replace(.SyubetuCD & .TrackCD, "'", "''") & "'," '' 競走種別コード
            strSql = strSql & SS & "Kyori" & SE & "='" & Replace(.Kyori, "'", "''") & "'," '' 距離
            strSql = strSql & SS & "RecKubun" & SE & "='" & Replace(.RecKubun, "'", "''") & "'," '' レコード区分
            strSql = strSql & SS & "RecTime" & SE & "='" & Replace(.RecTime, "'", "''") & "'," '' レコードタイム
            With .TenkoBaba
                strSql = strSql & SS & "TenkoCD" & SE & "='" & Replace(.TenkoCD, "'", "''") & "'," '' 天候コード
                strSql = strSql & SS & "SibaBabaCD" & SE & "='" & Replace(.SibaBabaCD, "'", "''") & "'," '' 芝馬場状態コード
                strSql = strSql & SS & "DirtBabaCD" & SE & "='" & Replace(.DirtBabaCD, "'", "''") & "'," '' ダート馬場状態コード
            End With ' TenkoBaba
            With .RecUmaInfo(0)
                strSql = strSql & SS & "RecUmaKettoNum1" & SE & "='" & Replace(.KettoNum, "'", "''") & "'," '' 血統登録番号
                strSql = strSql & SS & "RecUmaBamei1" & SE & "='" & Replace(.Bamei, "'", "''") & "'," '' 馬名
                strSql = strSql & SS & "RecUmaUmaKigoCD1" & SE & "='" & Replace(.UmaKigoCD, "'", "''") & "'," '' 馬記号コード
                strSql = strSql & SS & "RecUmaSexCD1" & SE & "='" & Replace(.SexCD, "'", "''") & "'," '' 性別コード
                strSql = strSql & SS & "RecUmaChokyosiCode1" & SE & "='" & Replace(.ChokyosiCode, "'", "''") & "'," '' 調教師コード
                strSql = strSql & SS & "RecUmaChokyosiName1" & SE & "='" & Replace(.ChokyosiName, "'", "''") & "'," '' 調教師名
                strSql = strSql & SS & "RecUmaFutan1" & SE & "='" & Replace(.Futan, "'", "''") & "'," '' 負担重量
                strSql = strSql & SS & "RecUmaKisyuCode1" & SE & "='" & Replace(.KisyuCode, "'", "''") & "'," '' 騎手コード
                strSql = strSql & SS & "RecUmaKisyuName1" & SE & "='" & Replace(.KisyuName, "'", "''") & "'," '' 騎手名
            End With ' RecUmaInfo
            With .RecUmaInfo(1)
                strSql = strSql & SS & "RecUmaKettoNum2" & SE & "='" & Replace(.KettoNum, "'", "''") & "'," '' 血統登録番号
                strSql = strSql & SS & "RecUmaBamei2" & SE & "='" & Replace(.Bamei, "'", "''") & "'," '' 馬名
                strSql = strSql & SS & "RecUmaUmaKigoCD2" & SE & "='" & Replace(.UmaKigoCD, "'", "''") & "'," '' 馬記号コード
                strSql = strSql & SS & "RecUmaSexCD2" & SE & "='" & Replace(.SexCD, "'", "''") & "'," '' 性別コード
                strSql = strSql & SS & "RecUmaChokyosiCode2" & SE & "='" & Replace(.ChokyosiCode, "'", "''") & "'," '' 調教師コード
                strSql = strSql & SS & "RecUmaChokyosiName2" & SE & "='" & Replace(.ChokyosiName, "'", "''") & "'," '' 調教師名
                strSql = strSql & SS & "RecUmaFutan2" & SE & "='" & Replace(.Futan, "'", "''") & "'," '' 負担重量
                strSql = strSql & SS & "RecUmaKisyuCode2" & SE & "='" & Replace(.KisyuCode, "'", "''") & "'," '' 騎手コード
                strSql = strSql & SS & "RecUmaKisyuName2" & SE & "='" & Replace(.KisyuName, "'", "''") & "'," '' 騎手名
            End With ' RecUmaInfo
            With .RecUmaInfo(2)
                strSql = strSql & SS & "RecUmaKettoNum3" & SE & "='" & Replace(.KettoNum, "'", "''") & "'," '' 血統登録番号
                strSql = strSql & SS & "RecUmaBamei3" & SE & "='" & Replace(.Bamei, "'", "''") & "'," '' 馬名
                strSql = strSql & SS & "RecUmaUmaKigoCD3" & SE & "='" & Replace(.UmaKigoCD, "'", "''") & "'," '' 馬記号コード
                strSql = strSql & SS & "RecUmaSexCD3" & SE & "='" & Replace(.SexCD, "'", "''") & "'," '' 性別コード
                strSql = strSql & SS & "RecUmaChokyosiCode3" & SE & "='" & Replace(.ChokyosiCode, "'", "''") & "'," '' 調教師コード
                strSql = strSql & SS & "RecUmaChokyosiName3" & SE & "='" & Replace(.ChokyosiName, "'", "''") & "'," '' 調教師名
                strSql = strSql & SS & "RecUmaFutan3" & SE & "='" & Replace(.Futan, "'", "''") & "'," '' 負担重量
                strSql = strSql & SS & "RecUmaKisyuCode3" & SE & "='" & Replace(.KisyuCode, "'", "''") & "'," '' 騎手コード
                strSql = strSql & SS & "RecUmaKisyuName3" & SE & "='" & Replace(.KisyuName, "'", "''") & "'," '' 騎手名
            End With ' RecUmaInfo
            strSql = Left(strSql, Len(strSql) - 1) ' Delete Last Comma

            strSql = strSql & " WHERE " & SS & "RecInfoKubun" & SE & "='" & Replace(.RecInfoKubun, "'", "''") & "'"
            With .id
                strSql = strSql & " AND " & SS & "Year" & SE & "='" & Replace(.Year, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "MonthDay" & SE & "='" & Replace(.MonthDay, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "JyoCD" & SE & "='" & Replace(.JyoCD, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "Kaiji" & SE & "='" & Replace(.Kaiji, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "Nichiji" & SE & "='" & Replace(.Nichiji, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "RaceNum" & SE & "='" & Replace(.RaceNum, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "TokuNum" & SE & "='" & Replace(mBuf.TokuNum, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "SyubetuCD_TrackCD" & SE & "='" & Replace(mBuf.SyubetuCD & mBuf.TrackCD, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "Kyori" & SE & "='" & Replace(mBuf.Kyori, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "MakeDate" & SE & " <= '" & Replace(strMakeDate, "'", "''") & "'"
            End With
        End With
        With mBuf.id
            System.Diagnostics.Debug.WriteLine("UPDATE RECORD : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum & mBuf.SyubetuCD & mBuf.TrackCD & mBuf.Kyori)
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