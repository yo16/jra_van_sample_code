' @(h) clsImportRA.vb
'
' @(s)
' JVData "RA" データベースアクセスクラス

Option Strict Off
Option Explicit On
Option Compare Binary

Friend Class clsImportRA

    'レース詳細構造体
    Private mBuf As JV_RA_RACE
    Private mRS As ADODB.Recordset


    ' @(f)
    '
    ' 機能      : 初期化処理
    '
    ' 引き数    :
    '
    ' 返り値    :
    '
    ' 機能説明  :
    '
    Public Sub New()

        MyBase.New()
        Class_Initialize_Renamed()

    End Sub


    ' @(f)
    '
    ' 機能      : 終了処理
    '
    ' 引き数    :
    '
    ' 返り値    :
    '
    ' 機能説明  :
    '
    Protected Overrides Sub Finalize()

        Class_Terminate_Renamed()
        MyBase.Finalize()

    End Sub


    ' @(f)
    '
    ' 機能      : 初期化、コネクション、レコードセットオブジェクトのインスタンス生成
    '
    ' 引き数    :
    '
    ' 返り値    :
    '
    ' 機能説明  :
    '
    '
    Private Sub Class_Initialize_Renamed()
        On Error GoTo ErrorHandler

        ' SQL文
        Dim strSql As String
        strSql = "SELECT * FROM RACE"

        ' レコードセットオープン
        mRS = New ADODB.Recordset()
        mRS.Open(strSql, gCon, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockPessimistic)

ExitHandler:
        Exit Sub

ErrorHandler:
        System.Diagnostics.Debug.WriteLine(Err.Description)
        Resume ExitHandler

    End Sub


    ' @(f)
    '
    ' 機能      : 終了処理
    '
    ' 引き数    :
    '
    ' 返り値    :
    '
    ' 機能説明  :
    '
    Private Sub Class_Terminate_Renamed()
        On Error GoTo ErrorHandler

ExitHandler:
        Exit Sub

ErrorHandler:
        System.Diagnostics.Debug.WriteLine(Err.Description)
        Resume ExitHandler

    End Sub


    ' @(f)
    '
    ' 機能      : メンバー変数のレコードセットのクローズ処理
    '
    ' 引き数    :
    '
    ' 返り値    :
    '
    ' 機能説明  : ガーベッジコレクションにCloseを呼ばせると、何処で呼ばれるか
    '            分からない為、明示的に呼び出す必要があります。
    '
    Public Sub Close()

        'レコードセットクローズ
        mRS.Close()
        mRS = Nothing

    End Sub


    ' @(f)
    '
    ' 機能      : レコードの抽出(SELECT)処理
    '
    ' 引き数    : SQL文字列
    '
    ' 返り値    : レース詳細構造体配列
    '
    ' 機能説明  :
    '
    Public Function SelectDB(ByVal strSQL As String) As JV_RA_RACE()
        On Error GoTo ErrorHandler

        Dim IsDBOpen As Boolean = False

        ' ADODB.Recordset
        Dim dbRS As ADODB.Recordset

        ' ADODB.Fields
        Dim dbFld As ADODB.Fields

        ' レース詳細構造体
        Dim structRA(0) As JV_RA_RACE

        ' ループカウンタ
        Dim iLoopCnt As Integer

        ' レコード件数
        Dim lRecCount As Long
        lRecCount = 0

        ' レコード文字列
        Dim strBuff As String

        ' レコードセットの生成
        dbRS = New ADODB.Recordset()
        ' レコードセットのオープン
        dbRS.Open(strSQL, gCon, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockPessimistic)
        IsDBOpen = True

        While Not dbRS.EOF
            ' フィールドの取得
            dbFld = dbRS.Fields

            ReDim Preserve structRA(lRecCount)

            ' 構造体設定用パラメータ作成
            strBuff = dbFld("RecordSpec").Value().PadRight(2)
            strBuff = strBuff + dbFld("DataKubun").Value().PadRight(1)
            strBuff = strBuff + dbFld("MakeDate").Value().PadRight(8)
            strBuff = strBuff + dbFld("Year").Value().PadRight(4)
            strBuff = strBuff + dbFld("MonthDay").Value().PadRight(4)
            strBuff = strBuff + dbFld("JyoCD").Value().PadRight(2)
            strBuff = strBuff + dbFld("Kaiji").Value().PadRight(2)
            strBuff = strBuff + dbFld("Nichiji").Value().PadRight(2)
            strBuff = strBuff + dbFld("RaceNum").Value().PadRight(2)
            strBuff = strBuff + dbFld("YoubiCD").Value().PadRight(1)
            strBuff = strBuff + dbFld("TokuNum").Value().PadRight(4)
            strBuff = strBuff + bPadR(dbFld("Hondai").Value(), 60)
            strBuff = strBuff + bPadR(dbFld("Fukudai").Value(), 60)
            strBuff = strBuff + bPadR(dbFld("Kakko").Value(), 60)
            strBuff = strBuff + dbFld("HondaiEng").Value().PadRight(120)
            strBuff = strBuff + dbFld("FukudaiEng").Value().PadRight(120)
            strBuff = strBuff + dbFld("KakkoEng").Value().PadRight(120)
            strBuff = strBuff + bPadR(dbFld("Ryakusyo10").Value(), 20)
            strBuff = strBuff + bPadR(dbFld("Ryakusyo6").Value(), 12)
            strBuff = strBuff + bPadR(dbFld("Ryakusyo3").Value(), 6)
            strBuff = strBuff + dbFld("Kubun").Value().PadRight(1)
            strBuff = strBuff + dbFld("Nkai").Value().PadRight(3)
            strBuff = strBuff + dbFld("GradeCD").Value().PadRight(1)
            strBuff = strBuff + dbFld("GradeCDBefore").Value().PadRight(1)
            strBuff = strBuff + dbFld("SyubetuCD").Value().PadRight(2)
            strBuff = strBuff + dbFld("KigoCD").Value().PadRight(3)
            strBuff = strBuff + dbFld("JyuryoCD").Value().PadRight(1)
            For iLoopCnt = 0 To 4
                strBuff = strBuff + dbFld("JyokenCD" & iLoopCnt + 1).Value().PadRight(3)
            Next iLoopCnt
            strBuff = strBuff + bPadR(dbFld("JyokenName").Value(), 60)
            strBuff = strBuff + dbFld("Kyori").Value().PadRight(4)
            strBuff = strBuff + dbFld("KyoriBefore").Value().PadRight(4)
            strBuff = strBuff + dbFld("TrackCD").Value().PadRight(2)
            strBuff = strBuff + dbFld("TrackCDBefore").Value().PadRight(2)
            strBuff = strBuff + dbFld("CourseKubunCD").Value().PadRight(2)
            strBuff = strBuff + dbFld("CourseKubunCDBefore").Value().PadRight(2)
            For iLoopCnt = 0 To 6
                strBuff = strBuff + dbFld("Honsyokin" & iLoopCnt + 1).Value().PadRight(8)
            Next iLoopCnt
            For iLoopCnt = 0 To 4
                strBuff = strBuff + dbFld("HonsyokinBefore" & iLoopCnt + 1).Value().PadRight(8)
            Next iLoopCnt
            For iLoopCnt = 0 To 4
                strBuff = strBuff + dbFld("Fukasyokin" & iLoopCnt + 1).Value().PadRight(8)
            Next iLoopCnt
            For iLoopCnt = 0 To 2
                strBuff = strBuff + dbFld("FukasyokinBefore" & iLoopCnt + 1).Value().PadRight(8)
            Next iLoopCnt
            strBuff = strBuff + dbFld("HassoTime").Value().PadRight(4)
            strBuff = strBuff + dbFld("HassoTimeBefore").Value().PadRight(4)
            strBuff = strBuff + dbFld("TorokuTosu").Value().PadRight(2)
            strBuff = strBuff + dbFld("SyussoTosu").Value().PadRight(2)
            strBuff = strBuff + dbFld("NyusenTosu").Value().PadRight(2)
            strBuff = strBuff + dbFld("TenkoCD").Value().PadRight(1)
            strBuff = strBuff + dbFld("SibaBabaCD").Value().PadRight(1)
            strBuff = strBuff + dbFld("DirtBabaCD").Value().PadRight(1)
            For iLoopCnt = 0 To 24
                strBuff = strBuff + dbFld("LapTime" & iLoopCnt + 1).Value().PadRight(3)
            Next iLoopCnt
            strBuff = strBuff + dbFld("SyogaiMileTime").Value().PadRight(4)
            strBuff = strBuff + dbFld("HaronTimeS3").Value().PadRight(3)
            strBuff = strBuff + dbFld("HaronTimeS4").Value().PadRight(3)
            strBuff = strBuff + dbFld("HaronTimeL3").Value().PadRight(3)
            strBuff = strBuff + dbFld("HaronTimeL4").Value().PadRight(3)
            For iLoopCnt = 0 To 3
                strBuff = strBuff + dbFld("Corner" & iLoopCnt + 1).Value().PadRight(1)
                strBuff = strBuff + dbFld("Syukaisu" & iLoopCnt + 1).Value().PadRight(1)
                strBuff = strBuff + dbFld("Jyuni" & iLoopCnt + 1).Value().PadRight(70)
            Next iLoopCnt
            strBuff = strBuff + dbFld("RecordUpKubun").Value().PadRight(1) + vbCrLf

            ' 構造体へ格納
            structRA(lRecCount).SetData(strBuff)

            ' レコード件数カウント
            lRecCount = lRecCount + 1

            ' 次レコードへ
            dbRS.MoveNext()

        End While

ExitHandler:
        ' レコードセットのクローズ
        If dbRS Is Nothing = False And IsDBOpen = True Then
            dbRS.Close()
        End If
        dbRS = Nothing

        ' 取得した構造体配列をリターン
        SelectDB = structRA

        Exit Function

ErrorHandler:
        System.Diagnostics.Debug.WriteLine(Err.Description)
        Resume ExitHandler

    End Function


    ' @(f)
    '
    ' 機能      : レコードの削除(DELETE)処理
    '
    ' 引き数    : SQL文字列
    '
    ' 返り値    : True - 成功, False - 失敗
    '
    ' 機能説明  :
    '
    Public Function DeleteDB(ByVal strSQL As String) As Boolean
        On Error GoTo ErrorHandler

        Dim bRetStatus As Boolean
        bRetStatus = True

        ' トランザクション開始
        gCon.BeginTrans()

        'テーブルのレコードをパラメータのSQLで削除する
        gCon.Execute(strSQL)

        ' トランザクション終了(コミット)
        gCon.CommitTrans()
        System.Diagnostics.Debug.WriteLine("CommitTrans")

ExitHandler:
        DeleteDB = bRetStatus
        Exit Function

ErrorHandler:
        System.Diagnostics.Debug.WriteLine(Err.Description)
        bRetStatus = False

        ' トランザクション終了(ロールバック)
        gCon.RollbackTrans()
        System.Diagnostics.Debug.WriteLine("RollbackTrans")
        Resume ExitHandler

    End Function


    ' @(f)
    '
    ' 機能      : JVReadの返す１行をデータベースに登録する
    '
    ' 引き数    : strBuf - JVData 識別子"RA" の１行
    '             lngBufSize - 未使用
    '
    ' 返り値    : True - 成功, False - 失敗
    '
    ' 機能説明  : clsIReadインターフェイスAddプロシージャの実装
    '
    Public Function Add(ByRef strBuf As String, ByVal lngBufSize As Integer) As Boolean
        On Error GoTo ErrorHandler

        ' 登録するデータの作成年月日
        Dim strMakeDate As String

        '構造体にデータセット
        mBuf.SetData(strBuf)

        With mBuf.head.MakeDate
            strMakeDate = .Year & .Month & .Day
        End With

        ' INSERT処理
        If Not InsertDB() Then
            'UPDATE処理（INSERTが失敗した場合）
            If Not UpdateDB(strMakeDate) Then System.Diagnostics.Debug.WriteLine("更新に失敗しました。" & Left(strBuf, 2))
        End If

        Add = True

ExitHandler:
        Exit Function

ErrorHandler:
        System.Diagnostics.Debug.WriteLine(Err.Description)
        Add = False

        Resume ExitHandler

    End Function


    ' @(f)
    '
    ' 機能      : レコードの挿入(INSERT)処理
    '
    ' 引き数    : 
    '
    ' 返り値    : True - 成功, False - 失敗
    '
    ' 機能説明  :
    '
    Public Function InsertDB() As Boolean
        On Error GoTo ErrorHandler

        ' ループカウンタ
        Dim iLoopCnt As Integer

        ' トランザクション開始
        gCon.BeginTrans()
        System.Diagnostics.Debug.WriteLine("BeginTrans")

        mRS.AddNew()

        With mBuf
            With .head
                ' レコード種別
                mRS.Fields("RecordSpec").Value = .RecordSpec
                ' データ区分
                mRS.Fields("DataKubun").Value = .DataKubun
                With .MakeDate
                    ' 年月日
                    mRS.Fields("MakeDate").Value = .Year & .Month & .Day
                End With ' MakeDate
            End With ' head
            With .id
                ' 開催年
                mRS.Fields("Year").Value = .Year
                ' 開催月日
                mRS.Fields("MonthDay").Value = .MonthDay
                ' 競馬場コード
                mRS.Fields("JyoCD").Value = .JyoCD
                ' 開催回第N回
                mRS.Fields("Kaiji").Value = .Kaiji
                ' 開催日目N日目
                mRS.Fields("Nichiji").Value = .Nichiji
                ' レース番号
                mRS.Fields("RaceNum").Value = .RaceNum
            End With ' id
            With .RaceInfo
                ' 曜日コード
                mRS.Fields("YoubiCD").Value = .YoubiCD
                ' 特別競走番号
                mRS.Fields("TokuNum").Value = .TokuNum
                ' 競走名本題
                mRS.Fields("Hondai").Value = .Hondai
                ' 競走名副題
                mRS.Fields("Fukudai").Value = .Fukudai
                ' 競走名カッコ内
                mRS.Fields("Kakko").Value = .Kakko
                ' 競走名本題欧字
                mRS.Fields("HondaiEng").Value = .HondaiEng
                ' 競走名副題欧字
                mRS.Fields("FukudaiEng").Value = .FukudaiEng
                ' 競走名カッコ内欧字
                mRS.Fields("KakkoEng").Value = .KakkoEng
                ' 競走名略称１０字
                mRS.Fields("Ryakusyo10").Value = .Ryakusyo10
                ' 競走名略称６字
                mRS.Fields("Ryakusyo6").Value = .Ryakusyo6
                ' 競走名略称３字
                mRS.Fields("Ryakusyo3").Value = .Ryakusyo3
                ' 競走名区分
                mRS.Fields("Kubun").Value = .Kubun
                ' 重賞回次第N回
                mRS.Fields("Nkai").Value = .Nkai
            End With ' RaceInfo
            ' グレードコード
            mRS.Fields("GradeCD").Value = .GradeCD
            ' 変更前グレードコード
            mRS.Fields("GradeCDBefore").Value = .GradeCDBefore
            With .JyokenInfo
                ' 競走種別コード
                mRS.Fields("SyubetuCD").Value = .SyubetuCD
                ' 競走記号コード
                mRS.Fields("KigoCD").Value = .KigoCD
                ' 重量種別コード
                mRS.Fields("JyuryoCD").Value = .JyuryoCD
                For iLoopCnt = 0 To 4
                    ' 競走条件コード
                    mRS.Fields("JyokenCD" & iLoopCnt + 1).Value = .JyokenCD(iLoopCnt)
                Next iLoopCnt
            End With ' JyokenInfo
            ' 競走条件名称
            mRS.Fields("JyokenName").Value = .JyokenName
            ' 距離
            mRS.Fields("Kyori").Value = .Kyori
            ' 変更前距離
            mRS.Fields("KyoriBefore").Value = .KyoriBefore
            ' トラックコード
            mRS.Fields("TrackCD").Value = .TrackCD
            ' 変更前トラックコード
            mRS.Fields("TrackCDBefore").Value = .TrackCDBefore
            ' コース区分
            mRS.Fields("CourseKubunCD").Value = .CourseKubunCD
            ' 変更前コース区分
            mRS.Fields("CourseKubunCDBefore").Value = .CourseKubunCDBefore
            For iLoopCnt = 0 To 6
                ' 本賞金
                mRS.Fields("Honsyokin" & iLoopCnt + 1).Value = .Honsyokin(iLoopCnt)
            Next iLoopCnt
            For iLoopCnt = 0 To 4
                ' 変更前本賞金
                mRS.Fields("HonsyokinBefore" & iLoopCnt + 1).Value = .HonsyokinBefore(iLoopCnt)
            Next iLoopCnt
            For iLoopCnt = 0 To 4
                ' 付加賞金
                mRS.Fields("Fukasyokin" & iLoopCnt + 1).Value = .Fukasyokin(iLoopCnt)
            Next iLoopCnt
            For iLoopCnt = 0 To 2
                ' 変更前付加賞金
                mRS.Fields("FukasyokinBefore" & iLoopCnt + 1).Value = .FukasyokinBefore(iLoopCnt)
            Next iLoopCnt
            ' 発走時刻
            mRS.Fields("HassoTime").Value = .HassoTime
            ' 変更前発走時刻
            mRS.Fields("HassoTimeBefore").Value = .HassoTimeBefore
            ' 登録頭数
            mRS.Fields("TorokuTosu").Value = .TorokuTosu
            ' 出走頭数
            mRS.Fields("SyussoTosu").Value = .SyussoTosu
            ' 入線頭数
            mRS.Fields("NyusenTosu").Value = .NyusenTosu
            With .TenkoBaba
                ' 天候コード
                mRS.Fields("TenkoCD").Value = .TenkoCD
                ' 芝馬場状態コード
                mRS.Fields("SibaBabaCD").Value = .SibaBabaCD
                ' ダート馬場状態コード
                mRS.Fields("DirtBabaCD").Value = .DirtBabaCD
            End With ' TenkoBaba
            For iLoopCnt = 0 To 24
                ' ラップタイム
                mRS.Fields("LapTime" & iLoopCnt + 1).Value = .LapTime(iLoopCnt)
            Next iLoopCnt
            ' 障害マイルタイム
            mRS.Fields("SyogaiMileTime").Value = .SyogaiMileTime
            ' 前３ハロンタイム
            mRS.Fields("HaronTimeS3").Value = .HaronTimeS3
            ' 前４ハロンタイム
            mRS.Fields("HaronTimeS4").Value = .HaronTimeS4
            ' 後３ハロンタイム
            mRS.Fields("HaronTimeL3").Value = .HaronTimeL3
            ' 後４ハロンタイム
            mRS.Fields("HaronTimeL4").Value = .HaronTimeL4
            For iLoopCnt = 0 To 3
                With .CornerInfo(iLoopCnt)
                    ' コーナー
                    mRS.Fields("Corner" & iLoopCnt + 1).Value = .Corner
                    ' 周回数
                    mRS.Fields("Syukaisu" & iLoopCnt + 1).Value = .Syukaisu
                    ' 各通過順位
                    mRS.Fields("Jyuni" & iLoopCnt + 1).Value = .Jyuni
                End With ' CornerInfo
            Next iLoopCnt
            ' レコード更新区分
            mRS.Fields("RecordUpKubun").Value = .RecordUpKubun
        End With

        With mBuf.id
            System.Diagnostics.Debug.WriteLine("Insert RACE : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum)
        End With ' id

        mRS.Update()

        ' トランザクション終了(コミット)
        gCon.CommitTrans()
        System.Diagnostics.Debug.WriteLine("CommitTrans")

        InsertDB = True

ExitHandler:
        Exit Function

ErrorHandler:
        System.Diagnostics.Debug.WriteLine(Err.Description)
        InsertDB = False

        mRS.CancelUpdate()

        ' トランザクション終了(ロールバック)
        gCon.RollbackTrans()
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

        ' ループカウンタ
        Dim iLoopCnt As Short

        ' SQL文
        Dim strSql As String

        ' トランザクション開始
        gCon.BeginTrans()
        System.Diagnostics.Debug.WriteLine("BeginTrans")

        strSql = "UPDATE RACE SET "
        With mBuf
            With .head
                ' レコード種別
                strSql = strSql & SS & "RecordSpec" & SE & "='" & Replace(.RecordSpec, "'", "''") & "',"
                ' データ区分
                strSql = strSql & SS & "DataKubun" & SE & "='" & Replace(.DataKubun, "'", "''") & "',"
                ' 年月日
                strSql = strSql & SS & "MakeDate" & SE & "='" & Replace(strMakeDate, "'", "''") & "',"
            End With ' head
            With .id
                ' 開催年
                strSql = strSql & SS & "Year" & SE & "='" & Replace(.Year, "'", "''") & "',"
                ' 開催月日
                strSql = strSql & SS & "MonthDay" & SE & "='" & Replace(.MonthDay, "'", "''") & "',"
                ' 競馬場コード
                strSql = strSql & SS & "JyoCD" & SE & "='" & Replace(.JyoCD, "'", "''") & "',"
                ' 開催回第N回
                strSql = strSql & SS & "Kaiji" & SE & "='" & Replace(.Kaiji, "'", "''") & "',"
                ' 開催日目N日目
                strSql = strSql & SS & "Nichiji" & SE & "='" & Replace(.Nichiji, "'", "''") & "',"
                ' レース番号
                strSql = strSql & SS & "RaceNum" & SE & "='" & Replace(.RaceNum, "'", "''") & "',"
            End With ' id
            With .RaceInfo
                ' 曜日コード
                strSql = strSql & SS & "YoubiCD" & SE & "='" & Replace(.YoubiCD, "'", "''") & "',"
                ' 特別競走番号
                strSql = strSql & SS & "TokuNum" & SE & "='" & Replace(.TokuNum, "'", "''") & "',"
                ' 競走名本題
                strSql = strSql & SS & "Hondai" & SE & "='" & Replace(.Hondai, "'", "''") & "',"
                ' 競走名副題
                strSql = strSql & SS & "Fukudai" & SE & "='" & Replace(.Fukudai, "'", "''") & "',"
                ' 競走名カッコ内
                strSql = strSql & SS & "Kakko" & SE & "='" & Replace(.Kakko, "'", "''") & "',"
                ' 競走名本題欧字
                strSql = strSql & SS & "HondaiEng" & SE & "='" & Replace(.HondaiEng, "'", "''") & "',"
                ' 競走名副題欧字
                strSql = strSql & SS & "FukudaiEng" & SE & "='" & Replace(.FukudaiEng, "'", "''") & "',"
                ' 競走名カッコ内欧字
                strSql = strSql & SS & "KakkoEng" & SE & "='" & Replace(.KakkoEng, "'", "''") & "',"
                ' 競走名略称１０字
                strSql = strSql & SS & "Ryakusyo10" & SE & "='" & Replace(.Ryakusyo10, "'", "''") & "',"
                ' 競走名略称６字
                strSql = strSql & SS & "Ryakusyo6" & SE & "='" & Replace(.Ryakusyo6, "'", "''") & "',"
                ' 競走名略称３字
                strSql = strSql & SS & "Ryakusyo3" & SE & "='" & Replace(.Ryakusyo3, "'", "''") & "',"
                ' 競走名区分
                strSql = strSql & SS & "Kubun" & SE & "='" & Replace(.Kubun, "'", "''") & "',"
                ' 重賞回次第N回
                strSql = strSql & SS & "Nkai" & SE & "='" & Replace(.Nkai, "'", "''") & "',"
            End With ' RaceInfo
            ' グレードコード
            strSql = strSql & SS & "GradeCD" & SE & "='" & Replace(.GradeCD, "'", "''") & "',"
            ' 変更前グレードコード
            strSql = strSql & SS & "GradeCDBefore" & SE & "='" & Replace(.GradeCDBefore, "'", "''") & "',"
            With .JyokenInfo
                ' 競走種別コード
                strSql = strSql & SS & "SyubetuCD" & SE & "='" & Replace(.SyubetuCD, "'", "''") & "',"
                ' 競走記号コード
                strSql = strSql & SS & "KigoCD" & SE & "='" & Replace(.KigoCD, "'", "''") & "',"
                ' 重量種別コード
                strSql = strSql & SS & "JyuryoCD" & SE & "='" & Replace(.JyuryoCD, "'", "''") & "',"
                For iLoopCnt = 0 To 4
                    ' 競走条件コード
                    strSql = strSql & SS & "JyokenCD" & iLoopCnt + 1 & "" & SE & "='" & Replace(.JyokenCD(iLoopCnt), "'", "''") & "',"
                Next iLoopCnt
            End With ' JyokenInfo
            ' 競走条件名称
            strSql = strSql & SS & "JyokenName" & SE & "='" & Replace(.JyokenName, "'", "''") & "',"
            ' 距離
            strSql = strSql & SS & "Kyori" & SE & "='" & Replace(.Kyori, "'", "''") & "',"
            ' 変更前距離
            strSql = strSql & SS & "KyoriBefore" & SE & "='" & Replace(.KyoriBefore, "'", "''") & "',"
            ' トラックコード
            strSql = strSql & SS & "TrackCD" & SE & "='" & Replace(.TrackCD, "'", "''") & "',"
            ' 変更前トラックコード
            strSql = strSql & SS & "TrackCDBefore" & SE & "='" & Replace(.TrackCDBefore, "'", "''") & "',"
            ' コース区分
            strSql = strSql & SS & "CourseKubunCD" & SE & "='" & Replace(.CourseKubunCD, "'", "''") & "',"
            ' 変更前コース区分
            strSql = strSql & SS & "CourseKubunCDBefore" & SE & "='" & Replace(.CourseKubunCDBefore, "'", "''") & "',"
            For iLoopCnt = 0 To 6
                ' 本賞金
                strSql = strSql & SS & "Honsyokin" & iLoopCnt + 1 & "" & SE & "='" & Replace(.Honsyokin(iLoopCnt), "'", "''") & "',"
            Next iLoopCnt
            For iLoopCnt = 0 To 4
                ' 変更前本賞金
                strSql = strSql & SS & "HonsyokinBefore" & iLoopCnt + 1 & "" & SE & "='" & Replace(.HonsyokinBefore(iLoopCnt), "'", "''") & "',"
            Next iLoopCnt
            For iLoopCnt = 0 To 4
                ' 付加賞金
                strSql = strSql & SS & "Fukasyokin" & iLoopCnt + 1 & "" & SE & "='" & Replace(.Fukasyokin(iLoopCnt), "'", "''") & "',"
            Next iLoopCnt
            For iLoopCnt = 0 To 2
                ' 変更前付加賞金
                strSql = strSql & SS & "FukasyokinBefore" & iLoopCnt + 1 & "" & SE & "='" & Replace(.FukasyokinBefore(iLoopCnt), "'", "''") & "',"
            Next iLoopCnt
            ' 発走時刻
            strSql = strSql & SS & "HassoTime" & SE & "='" & Replace(.HassoTime, "'", "''") & "',"
            ' 変更前発走時刻
            strSql = strSql & SS & "HassoTimeBefore" & SE & "='" & Replace(.HassoTimeBefore, "'", "''") & "',"
            ' 登録頭数
            strSql = strSql & SS & "TorokuTosu" & SE & "='" & Replace(.TorokuTosu, "'", "''") & "',"
            ' 出走頭数
            strSql = strSql & SS & "SyussoTosu" & SE & "='" & Replace(.SyussoTosu, "'", "''") & "',"
            ' 入線頭数
            strSql = strSql & SS & "NyusenTosu" & SE & "='" & Replace(.NyusenTosu, "'", "''") & "',"
            With .TenkoBaba
                ' 天候コード
                strSql = strSql & SS & "TenkoCD" & SE & "='" & Replace(.TenkoCD, "'", "''") & "',"
                ' 芝馬場状態コード
                strSql = strSql & SS & "SibaBabaCD" & SE & "='" & Replace(.SibaBabaCD, "'", "''") & "',"
                ' ダート馬場状態コード
                strSql = strSql & SS & "DirtBabaCD" & SE & "='" & Replace(.DirtBabaCD, "'", "''") & "',"
            End With ' TenkoBaba
            For iLoopCnt = 0 To 24
                ' ラップタイム
                strSql = strSql & SS & "LapTime" & iLoopCnt + 1 & "" & SE & "='" & Replace(.LapTime(iLoopCnt), "'", "''") & "',"
            Next iLoopCnt
            ' 障害マイルタイム
            strSql = strSql & SS & "SyogaiMileTime" & SE & "='" & Replace(.SyogaiMileTime, "'", "''") & "',"
            ' 前３ハロンタイム
            strSql = strSql & SS & "HaronTimeS3" & SE & "='" & Replace(.HaronTimeS3, "'", "''") & "',"
            ' 前４ハロンタイム
            strSql = strSql & SS & "HaronTimeS4" & SE & "='" & Replace(.HaronTimeS4, "'", "''") & "',"
            ' 後３ハロンタイム
            strSql = strSql & SS & "HaronTimeL3" & SE & "='" & Replace(.HaronTimeL3, "'", "''") & "',"
            ' 後４ハロンタイム
            strSql = strSql & SS & "HaronTimeL4" & SE & "='" & Replace(.HaronTimeL4, "'", "''") & "',"
            For iLoopCnt = 0 To 3
                With .CornerInfo(iLoopCnt)
                    ' コーナー
                    strSql = strSql & SS & "Corner" & iLoopCnt + 1 & "" & SE & "='" & Replace(.Corner, "'", "''") & "',"
                    ' 周回数
                    strSql = strSql & SS & "Syukaisu" & iLoopCnt + 1 & "" & SE & "='" & Replace(.Syukaisu, "'", "''") & "',"
                    ' 各通過順位
                    strSql = strSql & SS & "Jyuni" & iLoopCnt + 1 & "" & SE & "='" & Replace(.Jyuni, "'", "''") & "',"
                End With ' CornerInfo
            Next iLoopCnt
            ' レコード更新区分
            strSql = strSql & SS & "RecordUpKubun" & SE & "='" & Replace(.RecordUpKubun, "'", "''") & "',"

            strSql = Left(strSql, Len(strSql) - 1) ' Delete Last Comma
            With .id
                strSql = strSql & " WHERE " & SS & "Year" & SE & "='" & Replace(.Year, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "MonthDay" & SE & "='" & Replace(.MonthDay, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "JyoCD" & SE & "='" & Replace(.JyoCD, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "Kaiji" & SE & "='" & Replace(.Kaiji, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "Nichiji" & SE & "='" & Replace(.Nichiji, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "RaceNum" & SE & "='" & Replace(.RaceNum, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "MakeDate" & SE & "<= '" & Replace(strMakeDate, "'", "''") & "'"
            End With
        End With

        With mBuf.id
            System.Diagnostics.Debug.WriteLine("UPDATE RACE : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum)
        End With ' id

        ' SQL実行
        gCon.Execute(strSql)

        ' トランザクション終了(コミット)
        gCon.CommitTrans()
        System.Diagnostics.Debug.WriteLine("CommitTrans")

        UpdateDB = True

ExitHandler:
        Exit Function

ErrorHandler:
        System.Diagnostics.Debug.WriteLine(Err.Description)
        UpdateDB = False

        ' トランザクション終了(ロールバック)
        gCon.RollbackTrans()
        System.Diagnostics.Debug.WriteLine("RollbackTrans")
        Resume ExitHandler

    End Function

End Class