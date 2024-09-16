Option Strict Off
Option Explicit On
Option Compare Binary
Friend Class clsImportRA
	' @(h) clsReadRA.cls
	'
	' @(s)
	' JVData "RA" データベース登録クラス
	'
	
	Private mBuf As JV_RA_RACE ''レース詳細構造体
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
        strSql = "SELECT * FROM RACE"
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
    ' 引き数    : lBuf - JVData 識別子"RA" の１行
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
                mRS.Fields("RaceNum").Value = .RaceNum '' レース番号
            End With ' id
            With .RaceInfo
                mRS.Fields("YoubiCD").Value = .YoubiCD '' 曜日コード
                mRS.Fields("TokuNum").Value = .TokuNum '' 特別競走番号
                mRS.Fields("Hondai").Value = .Hondai '' 競走名本題
                mRS.Fields("Fukudai").Value = .Fukudai '' 競走名副題
                mRS.Fields("Kakko").Value = .Kakko '' 競走名カッコ内
                mRS.Fields("HondaiEng").Value = .HondaiEng '' 競走名本題欧字
                mRS.Fields("FukudaiEng").Value = .FukudaiEng '' 競走名副題欧字
                mRS.Fields("KakkoEng").Value = .KakkoEng '' 競走名カッコ内欧字
                mRS.Fields("Ryakusyo10").Value = .Ryakusyo10 '' 競走名略称１０字
                mRS.Fields("Ryakusyo6").Value = .Ryakusyo6 '' 競走名略称６字
                mRS.Fields("Ryakusyo3").Value = .Ryakusyo3 '' 競走名略称３字
                mRS.Fields("Kubun").Value = .Kubun '' 競走名区分
                mRS.Fields("Nkai").Value = .Nkai '' 重賞回次第N回
            End With ' RaceInfo
            mRS.Fields("GradeCD").Value = .GradeCD '' グレードコード
            mRS.Fields("GradeCDBefore").Value = .GradeCDBefore '' 変更前グレードコード
            With .JyokenInfo
                mRS.Fields("SyubetuCD").Value = .SyubetuCD '' 競走種別コード
                mRS.Fields("KigoCD").Value = .KigoCD '' 競走記号コード
                mRS.Fields("JyuryoCD").Value = .JyuryoCD '' 重量種別コード
                For j = 0 To 4
                    mRS.Fields("JyokenCD" & j + 1).Value = .JyokenCD(j) '' 競走条件コード
                Next j
            End With ' JyokenInfo
            mRS.Fields("JyokenName").Value = .JyokenName '' 競走条件名称
            mRS.Fields("Kyori").Value = .Kyori '' 距離
            mRS.Fields("KyoriBefore").Value = .KyoriBefore '' 変更前距離
            mRS.Fields("TrackCD").Value = .TrackCD '' トラックコード
            mRS.Fields("TrackCDBefore").Value = .TrackCDBefore '' 変更前トラックコード
            mRS.Fields("CourseKubunCD").Value = .CourseKubunCD '' コース区分
            mRS.Fields("CourseKubunCDBefore").Value = .CourseKubunCDBefore '' 変更前コース区分
            For i = 0 To 6
                mRS.Fields("Honsyokin" & i + 1).Value = .Honsyokin(i) '' 本賞金
            Next i
            For i = 0 To 4
                mRS.Fields("HonsyokinBefore" & i + 1).Value = .HonsyokinBefore(i) '' 変更前本賞金
            Next i
            For i = 0 To 4
                mRS.Fields("Fukasyokin" & i + 1).Value = .Fukasyokin(i) '' 付加賞金
            Next i
            For i = 0 To 2
                mRS.Fields("FukasyokinBefore" & i + 1).Value = .FukasyokinBefore(i) '' 変更前付加賞金
            Next i
            mRS.Fields("HassoTime").Value = .HassoTime '' 発走時刻
            mRS.Fields("HassoTimeBefore").Value = .HassoTimeBefore '' 変更前発走時刻
            mRS.Fields("TorokuTosu").Value = .TorokuTosu '' 登録頭数
            mRS.Fields("SyussoTosu").Value = .SyussoTosu '' 出走頭数
            mRS.Fields("NyusenTosu").Value = .NyusenTosu '' 入線頭数
            With .TenkoBaba
                mRS.Fields("TenkoCD").Value = .TenkoCD '' 天候コード
                mRS.Fields("SibaBabaCD").Value = .SibaBabaCD '' 芝馬場状態コード
                mRS.Fields("DirtBabaCD").Value = .DirtBabaCD '' ダート馬場状態コード
            End With ' TenkoBaba
            For i = 0 To 24
                mRS.Fields("LapTime" & i + 1).Value = .LapTime(i) '' ラップタイム
            Next i
            mRS.Fields("SyogaiMileTime").Value = .SyogaiMileTime '' 障害マイルタイム
            mRS.Fields("HaronTimeS3").Value = .HaronTimeS3 '' 前３ハロンタイム
            mRS.Fields("HaronTimeS4").Value = .HaronTimeS4 '' 前４ハロンタイム
            mRS.Fields("HaronTimeL3").Value = .HaronTimeL3 '' 後３ハロンタイム
            mRS.Fields("HaronTimeL4").Value = .HaronTimeL4 '' 後４ハロンタイム
            For i = 0 To 3
                With .CornerInfo(i)
                    mRS.Fields("Corner" & i + 1).Value = .Corner '' コーナー
                    mRS.Fields("Syukaisu" & i + 1).Value = .Syukaisu '' 周回数
                    mRS.Fields("Jyuni" & i + 1).Value = .Jyuni '' 各通過順位
                End With ' CornerInfo
            Next i
            mRS.Fields("RecordUpKubun").Value = .RecordUpKubun '' レコード更新区分
        End With

        With mBuf.id
            System.Diagnostics.Debug.WriteLine("Insert RACE : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum)
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

        strSql = "UPDATE RACE SET "
        With mBuf
            With .head
                strSql = strSql & SS & "RecordSpec" & SE & "='" & Replace(.RecordSpec, "'", "''") & "'," '' レコード種別
                strSql = strSql & SS & "DataKubun" & SE & "='" & Replace(.DataKubun, "'", "''") & "'," '' データ区分
                strSql = strSql & SS & "MakeDate" & SE & "= '" & Replace(strMakeDate, "'", "''") & "'," '' 年月日
            End With ' head
            With .id
                strSql = strSql & SS & "Year" & SE & "='" & Replace(.Year, "'", "''") & "'," '' 開催年
                strSql = strSql & SS & "MonthDay" & SE & "='" & Replace(.MonthDay, "'", "''") & "'," '' 開催月日
                strSql = strSql & SS & "JyoCD" & SE & "='" & Replace(.JyoCD, "'", "''") & "'," '' 競馬場コード
                strSql = strSql & SS & "Kaiji" & SE & "='" & Replace(.Kaiji, "'", "''") & "'," '' 開催回第N回
                strSql = strSql & SS & "Nichiji" & SE & "='" & Replace(.Nichiji, "'", "''") & "'," '' 開催日目N日目
                strSql = strSql & SS & "RaceNum" & SE & "='" & Replace(.RaceNum, "'", "''") & "'," '' レース番号
            End With ' id
            With .RaceInfo
                strSql = strSql & SS & "YoubiCD" & SE & "='" & Replace(.YoubiCD, "'", "''") & "'," '' 曜日コード
                strSql = strSql & SS & "TokuNum" & SE & "='" & Replace(.TokuNum, "'", "''") & "'," '' 特別競走番号
                strSql = strSql & SS & "Hondai" & SE & "='" & Replace(.Hondai, "'", "''") & "'," '' 競走名本題
                strSql = strSql & SS & "Fukudai" & SE & "='" & Replace(.Fukudai, "'", "''") & "'," '' 競走名副題
                strSql = strSql & SS & "Kakko" & SE & "='" & Replace(.Kakko, "'", "''") & "'," '' 競走名カッコ内
                strSql = strSql & SS & "HondaiEng" & SE & "='" & Replace(.HondaiEng, "'", "''") & "'," '' 競走名本題欧字
                strSql = strSql & SS & "FukudaiEng" & SE & "='" & Replace(.FukudaiEng, "'", "''") & "'," '' 競走名副題欧字
                strSql = strSql & SS & "KakkoEng" & SE & "='" & Replace(.KakkoEng, "'", "''") & "'," '' 競走名カッコ内欧字
                strSql = strSql & SS & "Ryakusyo10" & SE & "='" & Replace(.Ryakusyo10, "'", "''") & "'," '' 競走名略称１０字
                strSql = strSql & SS & "Ryakusyo6" & SE & "='" & Replace(.Ryakusyo6, "'", "''") & "'," '' 競走名略称６字
                strSql = strSql & SS & "Ryakusyo3" & SE & "='" & Replace(.Ryakusyo3, "'", "''") & "'," '' 競走名略称３字
                strSql = strSql & SS & "Kubun" & SE & "='" & Replace(.Kubun, "'", "''") & "'," '' 競走名区分
                strSql = strSql & SS & "Nkai" & SE & "='" & Replace(.Nkai, "'", "''") & "'," '' 重賞回次第N回
            End With ' RaceInfo
            strSql = strSql & SS & "GradeCD" & SE & "='" & Replace(.GradeCD, "'", "''") & "'," '' グレードコード
            strSql = strSql & SS & "GradeCDBefore" & SE & "='" & Replace(.GradeCDBefore, "'", "''") & "'," '' 変更前グレードコード
            With .JyokenInfo
                strSql = strSql & SS & "SyubetuCD" & SE & "='" & Replace(.SyubetuCD, "'", "''") & "'," '' 競走種別コード
                strSql = strSql & SS & "KigoCD" & SE & "='" & Replace(.KigoCD, "'", "''") & "'," '' 競走記号コード
                strSql = strSql & SS & "JyuryoCD" & SE & "='" & Replace(.JyuryoCD, "'", "''") & "'," '' 重量種別コード
                For j = 0 To 4
                    strSql = strSql & SS & "JyokenCD" & j + 1 & SE & "='" & Replace(.JyokenCD(j), "'", "''") & "'," '' 競走条件コード
                Next j
            End With ' JyokenInfo
            strSql = strSql & SS & "JyokenName" & SE & "='" & Replace(.JyokenName, "'", "''") & "'," '' 競走条件名称
            strSql = strSql & SS & "Kyori" & SE & "='" & Replace(.Kyori, "'", "''") & "'," '' 距離
            strSql = strSql & SS & "KyoriBefore" & SE & "='" & Replace(.KyoriBefore, "'", "''") & "'," '' 変更前距離
            strSql = strSql & SS & "TrackCD" & SE & "='" & Replace(.TrackCD, "'", "''") & "'," '' トラックコード
            strSql = strSql & SS & "TrackCDBefore" & SE & "='" & Replace(.TrackCDBefore, "'", "''") & "'," '' 変更前トラックコード
            strSql = strSql & SS & "CourseKubunCD" & SE & "='" & Replace(.CourseKubunCD, "'", "''") & "'," '' コース区分
            strSql = strSql & SS & "CourseKubunCDBefore" & SE & "='" & Replace(.CourseKubunCDBefore, "'", "''") & "'," '' 変更前コース区分
            For i = 0 To 6
                strSql = strSql & SS & "Honsyokin" & i + 1 & SE & "='" & Replace(.Honsyokin(i), "'", "''") & "'," '' 本賞金
            Next i
            For i = 0 To 4
                strSql = strSql & SS & "HonsyokinBefore" & i + 1 & SE & "='" & Replace(.HonsyokinBefore(i), "'", "''") & "'," '' 変更前本賞金
            Next i
            For i = 0 To 4
                strSql = strSql & SS & "Fukasyokin" & i + 1 & SE & "='" & Replace(.Fukasyokin(i), "'", "''") & "'," '' 付加賞金
            Next i
            For i = 0 To 2
                strSql = strSql & SS & "FukasyokinBefore" & i + 1 & SE & "='" & Replace(.FukasyokinBefore(i), "'", "''") & "'," '' 変更前付加賞金
            Next i
            strSql = strSql & SS & "HassoTime" & SE & "='" & Replace(.HassoTime, "'", "''") & "'," '' 発走時刻
            strSql = strSql & SS & "HassoTimeBefore" & SE & "='" & Replace(.HassoTimeBefore, "'", "''") & "'," '' 変更前発走時刻
            strSql = strSql & SS & "TorokuTosu" & SE & "='" & Replace(.TorokuTosu, "'", "''") & "'," '' 登録頭数
            strSql = strSql & SS & "SyussoTosu" & SE & "='" & Replace(.SyussoTosu, "'", "''") & "'," '' 出走頭数
            strSql = strSql & SS & "NyusenTosu" & SE & "='" & Replace(.NyusenTosu, "'", "''") & "'," '' 入線頭数
            With .TenkoBaba
                strSql = strSql & SS & "TenkoCD" & SE & "='" & Replace(.TenkoCD, "'", "''") & "'," '' 天候コード
                strSql = strSql & SS & "SibaBabaCD" & SE & "='" & Replace(.SibaBabaCD, "'", "''") & "'," '' 芝馬場状態コード
                strSql = strSql & SS & "DirtBabaCD" & SE & "='" & Replace(.DirtBabaCD, "'", "''") & "'," '' ダート馬場状態コード
            End With ' TenkoBaba
            For i = 0 To 24
                strSql = strSql & SS & "LapTime" & i + 1 & SE & "='" & Replace(.LapTime(i), "'", "''") & "'," '' ラップタイム
            Next i
            strSql = strSql & SS & "SyogaiMileTime" & SE & "='" & Replace(.SyogaiMileTime, "'", "''") & "'," '' 障害マイルタイム
            strSql = strSql & SS & "HaronTimeS3" & SE & "='" & Replace(.HaronTimeS3, "'", "''") & "'," '' 前３ハロンタイム
            strSql = strSql & SS & "HaronTimeS4" & SE & "='" & Replace(.HaronTimeS4, "'", "''") & "'," '' 前４ハロンタイム
            strSql = strSql & SS & "HaronTimeL3" & SE & "='" & Replace(.HaronTimeL3, "'", "''") & "'," '' 後３ハロンタイム
            strSql = strSql & SS & "HaronTimeL4" & SE & "='" & Replace(.HaronTimeL4, "'", "''") & "'," '' 後４ハロンタイム
            For i = 0 To 3
                With .CornerInfo(i)
                    strSql = strSql & SS & "Corner" & i + 1 & SE & "='" & Replace(.Corner, "'", "''") & "'," '' コーナー
                    strSql = strSql & SS & "Syukaisu" & i + 1 & SE & "='" & Replace(.Syukaisu, "'", "''") & "'," '' 周回数
                    strSql = strSql & SS & "Jyuni" & i + 1 & SE & "='" & Replace(.Jyuni, "'", "''") & "'," '' 各通過順位
                End With ' CornerInfo
            Next i
            strSql = strSql & SS & "RecordUpKubun" & SE & "='" & Replace(.RecordUpKubun, "'", "''") & "'," '' レコード更新区分
            strSql = Left(strSql, Len(strSql) - 1) ' Delete Last Comma
            With .id
                strSql = strSql & " WHERE " & SS & "Year" & SE & "='" & Replace(.Year, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "MonthDay" & SE & "='" & Replace(.MonthDay, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "JyoCD" & SE & "='" & Replace(.JyoCD, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "Kaiji" & SE & "='" & Replace(.Kaiji, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "Nichiji" & SE & "='" & Replace(.Nichiji, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "RaceNum" & SE & "='" & Replace(.RaceNum, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "MakeDate" & SE & " <= '" & Replace(strMakeDate, "'", "''") & "'"
            End With
        End With

        With mBuf.id
            System.Diagnostics.Debug.WriteLine("UPDATE RACE : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum)
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