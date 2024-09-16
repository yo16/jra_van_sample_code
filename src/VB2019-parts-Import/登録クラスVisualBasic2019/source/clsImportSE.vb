Option Strict Off
Option Explicit On
Option Compare Binary
Friend Class clsImportSE
	' @(h) clsReadSE.cls
	' @(s)
	' JVData "SE" データベース登録クラス
	'
	
	Private mBuf As JV_SE_RACE_UMA ''馬毎レース情報構造体
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
        strSql = "SELECT * FROM UMA_RACE"
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
    ' 引き数    : lBuf - JVData 識別子"SE" の１行
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
            mRS.Fields("Wakuban").Value = .Wakuban '' 枠番
            mRS.Fields("Umaban").Value = .Umaban '' 馬番
            mRS.Fields("KettoNum").Value = .KettoNum '' 血統登録番号
            mRS.Fields("Bamei").Value = .Bamei '' 馬名
            mRS.Fields("UmaKigoCD").Value = .UmaKigoCD '' 馬記号コード
            mRS.Fields("SexCD").Value = .SexCD '' 性別コード
            mRS.Fields("HinsyuCD").Value = .HinsyuCD '' 品種コード
            mRS.Fields("KeiroCD").Value = .KeiroCD '' 毛色コード
            mRS.Fields("Barei").Value = .Barei '' 馬齢
            mRS.Fields("TozaiCD").Value = .TozaiCD '' 東西所属コード
            mRS.Fields("ChokyosiCode").Value = .ChokyosiCode '' 調教師コード
            mRS.Fields("ChokyosiRyakusyo").Value = .ChokyosiRyakusyo '' 調教師名略称
            mRS.Fields("BanusiCode").Value = .BanusiCode '' 馬主コード
            mRS.Fields("BanusiName").Value = .BanusiName '' 馬主名
            mRS.Fields("Fukusyoku").Value = .Fukusyoku '' 服色標示
            mRS.Fields("reserved1").Value = .reserved1 '' 予備
            mRS.Fields("Futan").Value = .Futan '' 負担重量
            mRS.Fields("FutanBefore").Value = .FutanBefore '' 変更前負担重量
            mRS.Fields("Blinker").Value = .Blinker '' ブリンカー使用区分
            mRS.Fields("reserved2").Value = .reserved2 '' 予備
            mRS.Fields("KisyuCode").Value = .KisyuCode '' 騎手コード
            mRS.Fields("KisyuCodeBefore").Value = .KisyuCodeBefore '' 変更前騎手コード
            mRS.Fields("KisyuRyakusyo").Value = .KisyuRyakusyo '' 騎手名略称
            mRS.Fields("KisyuRyakusyoBefore").Value = .KisyuRyakusyoBefore '' 変更前騎手名略称
            mRS.Fields("MinaraiCD").Value = .MinaraiCD '' 騎手見習コード
            mRS.Fields("MinaraiCDBefore").Value = .MinaraiCDBefore '' 変更前騎手見習コード
            mRS.Fields("BaTaijyu").Value = .BaTaijyu '' 馬体重
            mRS.Fields("ZogenFugo").Value = .ZogenFugo '' 増減符号
            mRS.Fields("ZogenSa").Value = .ZogenSa '' 増減差
            mRS.Fields("IJyoCD").Value = .IJyoCD '' 異常区分コード
            mRS.Fields("NyusenJyuni").Value = .NyusenJyuni '' 入線順位
            mRS.Fields("KakuteiJyuni").Value = .KakuteiJyuni '' 確定着順
            mRS.Fields("DochakuKubun").Value = .DochakuKubun '' 同着区分
            mRS.Fields("DochakuTosu").Value = .DochakuTosu '' 同着頭数
            mRS.Fields("Time").Value = .Time '' 走破タイム
            mRS.Fields("ChakusaCD").Value = .ChakusaCD '' 着差コード
            mRS.Fields("ChakusaCDP").Value = .ChakusaCDP '' +着差コード
            mRS.Fields("ChakusaCDPP").Value = .ChakusaCDPP '' ++着差コード
            mRS.Fields("Jyuni1c").Value = .Jyuni1c '' 1コーナーでの順位
            mRS.Fields("Jyuni2c").Value = .Jyuni2c '' 2コーナーでの順位
            mRS.Fields("Jyuni3c").Value = .Jyuni3c '' 3コーナーでの順位
            mRS.Fields("Jyuni4c").Value = .Jyuni4c '' 4コーナーでの順位
            mRS.Fields("Odds").Value = .Odds '' 単勝オッズ
            mRS.Fields("Ninki").Value = .Ninki '' 単勝人気順
            mRS.Fields("Honsyokin").Value = .Honsyokin '' 獲得本賞金
            mRS.Fields("Fukasyokin").Value = .Fukasyokin '' 獲得付加賞金
            mRS.Fields("reserved3").Value = .reserved3 '' 予備
            mRS.Fields("reserved4").Value = .reserved4 '' 予備
            mRS.Fields("HaronTimeL4").Value = .HaronTimeL4 '' 後４ハロンタイム
            mRS.Fields("HaronTimeL3").Value = .HaronTimeL3 '' 後３ハロンタイム
            For i = 0 To 2
                With .ChakuUmaInfo(i)
                    mRS.Fields("KettoNum" & i + 1).Value = .KettoNum '' 血統登録番号
                    mRS.Fields("Bamei" & i + 1).Value = .Bamei '' 馬名
                End With ' ChakuUmaInfo
            Next i
            mRS.Fields("TimeDiff").Value = .TimeDiff '' タイム差
            mRS.Fields("RecordUpKubun").Value = .RecordUpKubun '' レコード更新区分
            mRS.Fields("DMKubun").Value = .DMKubun '' マイニング区分
            mRS.Fields("DMTime").Value = .DMTime '' マイニング予想走破タイム
            mRS.Fields("DMGosaP").Value = .DMGosaP '' 予測誤差(信頼度)＋
            mRS.Fields("DMGosaM").Value = .DMGosaM '' 予測誤差(信頼度)－
            mRS.Fields("DMJyuni").Value = .DMJyuni '' マイニング予想順位
            mRS.Fields("KyakusituKubun").Value = .KyakusituKubun '' 今回レース脚質判定
        End With

        With mBuf.id
            System.Diagnostics.Debug.WriteLine("Insert UMA_RACE : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum)
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

        strSql = "UPDATE UMA_RACE SET "

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
                strSql = strSql & SS & "RaceNum" & SE & "='" & Replace(.RaceNum, "'", "''") & "'," '' レース番号
            End With ' id
            strSql = strSql & SS & "Wakuban" & SE & "='" & Replace(.Wakuban, "'", "''") & "'," '' 枠番
            strSql = strSql & SS & "Umaban" & SE & "='" & Replace(.Umaban, "'", "''") & "'," '' 馬番
            strSql = strSql & SS & "KettoNum" & SE & "='" & Replace(.KettoNum, "'", "''") & "'," '' 血統登録番号
            strSql = strSql & SS & "Bamei" & SE & "='" & Replace(.Bamei, "'", "''") & "'," '' 馬名
            strSql = strSql & SS & "UmaKigoCD" & SE & "='" & Replace(.UmaKigoCD, "'", "''") & "'," '' 馬記号コード
            strSql = strSql & SS & "SexCD" & SE & "='" & Replace(.SexCD, "'", "''") & "'," '' 性別コード
            strSql = strSql & SS & "HinsyuCD" & SE & "='" & Replace(.HinsyuCD, "'", "''") & "'," '' 品種コード
            strSql = strSql & SS & "KeiroCD" & SE & "='" & Replace(.KeiroCD, "'", "''") & "'," '' 毛色コード
            strSql = strSql & SS & "Barei" & SE & "='" & Replace(.Barei, "'", "''") & "'," '' 馬齢
            strSql = strSql & SS & "TozaiCD" & SE & "='" & Replace(.TozaiCD, "'", "''") & "'," '' 東西所属コード
            strSql = strSql & SS & "ChokyosiCode" & SE & "='" & Replace(.ChokyosiCode, "'", "''") & "'," '' 調教師コード
            strSql = strSql & SS & "ChokyosiRyakusyo" & SE & "='" & Replace(.ChokyosiRyakusyo, "'", "''") & "'," '' 調教師名略称
            strSql = strSql & SS & "BanusiCode" & SE & "='" & Replace(.BanusiCode, "'", "''") & "'," '' 馬主コード
            strSql = strSql & SS & "BanusiName" & SE & "='" & Replace(.BanusiName, "'", "''") & "'," '' 馬主名
            strSql = strSql & SS & "Fukusyoku" & SE & "='" & Replace(.Fukusyoku, "'", "''") & "'," '' 服色標示
            strSql = strSql & SS & "reserved1" & SE & "='" & Replace(.reserved1, "'", "''") & "'," '' 予備
            strSql = strSql & SS & "Futan" & SE & "='" & Replace(.Futan, "'", "''") & "'," '' 負担重量
            strSql = strSql & SS & "FutanBefore" & SE & "='" & Replace(.FutanBefore, "'", "''") & "'," '' 変更前負担重量
            strSql = strSql & SS & "Blinker" & SE & "='" & Replace(.Blinker, "'", "''") & "'," '' ブリンカー使用区分
            strSql = strSql & SS & "reserved2" & SE & "='" & Replace(.reserved2, "'", "''") & "'," '' 予備
            strSql = strSql & SS & "KisyuCode" & SE & "='" & Replace(.KisyuCode, "'", "''") & "'," '' 騎手コード
            strSql = strSql & SS & "KisyuCodeBefore" & SE & "='" & Replace(.KisyuCodeBefore, "'", "''") & "'," '' 変更前騎手コード
            strSql = strSql & SS & "KisyuRyakusyo" & SE & "='" & Replace(.KisyuRyakusyo, "'", "''") & "'," '' 騎手名略称
            strSql = strSql & SS & "KisyuRyakusyoBefore" & SE & "='" & Replace(.KisyuRyakusyoBefore, "'", "''") & "'," '' 変更前騎手名略称
            strSql = strSql & SS & "MinaraiCD" & SE & "='" & Replace(.MinaraiCD, "'", "''") & "'," '' 騎手見習コード
            strSql = strSql & SS & "MinaraiCDBefore" & SE & "='" & Replace(.MinaraiCDBefore, "'", "''") & "'," '' 変更前騎手見習コード
            strSql = strSql & SS & "BaTaijyu" & SE & "='" & Replace(.BaTaijyu, "'", "''") & "'," '' 馬体重
            strSql = strSql & SS & "ZogenFugo" & SE & "='" & Replace(.ZogenFugo, "'", "''") & "'," '' 増減符号
            strSql = strSql & SS & "ZogenSa" & SE & "='" & Replace(.ZogenSa, "'", "''") & "'," '' 増減差
            strSql = strSql & SS & "IJyoCD" & SE & "='" & Replace(.IJyoCD, "'", "''") & "'," '' 異常区分コード
            strSql = strSql & SS & "NyusenJyuni" & SE & "='" & Replace(.NyusenJyuni, "'", "''") & "'," '' 入線順位
            strSql = strSql & SS & "KakuteiJyuni" & SE & "='" & Replace(.KakuteiJyuni, "'", "''") & "'," '' 確定着順
            strSql = strSql & SS & "DochakuKubun" & SE & "='" & Replace(.DochakuKubun, "'", "''") & "'," '' 同着区分
            strSql = strSql & SS & "DochakuTosu" & SE & "='" & Replace(.DochakuTosu, "'", "''") & "'," '' 同着頭数
            strSql = strSql & SS & "Time" & SE & "='" & Replace(.Time, "'", "''") & "'," '' 走破タイム
            strSql = strSql & SS & "ChakusaCD" & SE & "='" & Replace(.ChakusaCD, "'", "''") & "'," '' 着差コード
            strSql = strSql & SS & "ChakusaCDP" & SE & "='" & Replace(.ChakusaCDP, "'", "''") & "'," '' +着差コード
            strSql = strSql & SS & "ChakusaCDPP" & SE & "='" & Replace(.ChakusaCDPP, "'", "''") & "'," '' ++着差コード
            strSql = strSql & SS & "Jyuni1c" & SE & "='" & Replace(.Jyuni1c, "'", "''") & "'," '' 1コーナーでの順位
            strSql = strSql & SS & "Jyuni2c" & SE & "='" & Replace(.Jyuni2c, "'", "''") & "'," '' 2コーナーでの順位
            strSql = strSql & SS & "Jyuni3c" & SE & "='" & Replace(.Jyuni3c, "'", "''") & "'," '' 3コーナーでの順位
            strSql = strSql & SS & "Jyuni4c" & SE & "='" & Replace(.Jyuni4c, "'", "''") & "'," '' 4コーナーでの順位
            strSql = strSql & SS & "Odds" & SE & "='" & Replace(.Odds, "'", "''") & "'," '' 単勝オッズ
            strSql = strSql & SS & "Ninki" & SE & "='" & Replace(.Ninki, "'", "''") & "'," '' 単勝人気順
            strSql = strSql & SS & "Honsyokin" & SE & "='" & Replace(.Honsyokin, "'", "''") & "'," '' 獲得本賞金
            strSql = strSql & SS & "Fukasyokin" & SE & "='" & Replace(.Fukasyokin, "'", "''") & "'," '' 獲得付加賞金
            strSql = strSql & SS & "reserved3" & SE & "='" & Replace(.reserved3, "'", "''") & "'," '' 予備
            strSql = strSql & SS & "reserved4" & SE & "='" & Replace(.reserved4, "'", "''") & "'," '' 予備
            strSql = strSql & SS & "HaronTimeL4" & SE & "='" & Replace(.HaronTimeL4, "'", "''") & "'," '' 後４ハロンタイム
            strSql = strSql & SS & "HaronTimeL3" & SE & "='" & Replace(.HaronTimeL3, "'", "''") & "'," '' 後３ハロンタイム
            For i = 0 To 2
                With .ChakuUmaInfo(i)
                    strSql = strSql & SS & "KettoNum" & i + 1 & SE & "='" & Replace(.KettoNum, "'", "''") & "'," '' 血統登録番号（相手馬1～3）
                    strSql = strSql & SS & "Bamei" & i + 1 & SE & "='" & Replace(.Bamei, "'", "''") & "'," '' 馬名（相手馬1～3）
                End With ' ChakuUmaInfo
            Next i
            strSql = strSql & SS & "TimeDiff" & SE & "='" & Replace(.TimeDiff, "'", "''") & "'," '' タイム差
            strSql = strSql & SS & "RecordUpKubun" & SE & "='" & Replace(.RecordUpKubun, "'", "''") & "'," '' レコード更新区分
            strSql = strSql & SS & "DMKubun" & SE & "='" & Replace(.DMKubun, "'", "''") & "'," '' マイニング区分
            strSql = strSql & SS & "DMTime" & SE & "='" & Replace(.DMTime, "'", "''") & "'," '' マイニング予想走破タイム
            strSql = strSql & SS & "DMGosaP" & SE & "='" & Replace(.DMGosaP, "'", "''") & "'," '' 予測誤差(信頼度)＋
            strSql = strSql & SS & "DMGosaM" & SE & "='" & Replace(.DMGosaM, "'", "''") & "'," '' 予測誤差(信頼度)－
            strSql = strSql & SS & "DMJyuni" & SE & "='" & Replace(.DMJyuni, "'", "''") & "'," '' マイニング予想順位
            strSql = strSql & SS & "KyakusituKubun" & SE & "='" & Replace(.KyakusituKubun, "'", "''") & "'," '' 今回レース脚質判定

            strSql = Left(strSql, Len(strSql) - 1) ' Delete Last Comma

            With .id
                strSql = strSql & " WHERE " & SS & "Year" & SE & "='" & Replace(.Year, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "MonthDay" & SE & "='" & Replace(.MonthDay, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "JyoCD" & SE & "='" & Replace(.JyoCD, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "Kaiji" & SE & "='" & Replace(.Kaiji, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "Nichiji" & SE & "='" & Replace(.Nichiji, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "RaceNum" & SE & "='" & Replace(.RaceNum, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "Umaban" & SE & "='" & Replace(mBuf.Umaban, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "KettoNum" & SE & "='" & Replace(mBuf.KettoNum, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "MakeDate" & SE & " <= '" & Replace(strMakeDate, "'", "''") & "'"
            End With
        End With

        With mBuf.id
            System.Diagnostics.Debug.WriteLine("UPDATE UMA_RACE : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum & mBuf.Umaban & mBuf.KettoNum)
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