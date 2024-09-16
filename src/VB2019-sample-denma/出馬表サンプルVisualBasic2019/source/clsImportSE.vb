' @(h) clsImportSE.vb
'
' @(s)
' JVData "SE" データベースアクセスクラス

Option Strict Off
Option Explicit On
Option Compare Binary

Friend Class clsImportSE

    '馬毎レース情報構造体
    Private mBuf As JV_SE_RACE_UMA
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
        strSql = "SELECT * FROM UMA_RACE"

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
    Public Function SelectDB(ByVal strSQL As String) As JV_SE_RACE_UMA()
        On Error GoTo ErrorHandler

        Dim IsDBOpen As Boolean = False

        ' ADODB.Recordset
        Dim dbRS As ADODB.Recordset

        ' ADODB.Fields
        Dim dbFld As ADODB.Fields

        ' レース詳細構造体
        Dim structSE(0) As JV_SE_RACE_UMA

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

            ReDim Preserve structSE(lRecCount)

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
            strBuff = strBuff + dbFld("Wakuban").Value().PadRight(1)
            strBuff = strBuff + dbFld("Umaban").Value().PadRight(2)
            strBuff = strBuff + dbFld("KettoNum").Value().PadRight(10)
            strBuff = strBuff + bPadR(dbFld("Bamei").Value(), 36)
            strBuff = strBuff + dbFld("UmaKigoCD").Value().PadRight(2)
            strBuff = strBuff + dbFld("SexCD").Value().PadRight(1)
            strBuff = strBuff + dbFld("HinsyuCD").Value().PadRight(1)
            strBuff = strBuff + dbFld("KeiroCD").Value().PadRight(2)
            strBuff = strBuff + dbFld("Barei").Value().PadRight(2)
            strBuff = strBuff + dbFld("TozaiCD").Value().PadRight(1)
            strBuff = strBuff + dbFld("ChokyosiCode").Value().PadRight(5)
            strBuff = strBuff + bPadR(dbFld("ChokyosiRyakusyo").Value(), 8)
            strBuff = strBuff + dbFld("BanusiCode").Value().PadRight(6)
            strBuff = strBuff + bPadR(dbFld("BanusiName").Value(), 64)
            strBuff = strBuff + bPadR(dbFld("Fukusyoku").Value(), 60)
            strBuff = strBuff + bPadR(dbFld("reserved1").Value(), 60)
            strBuff = strBuff + dbFld("Futan").Value().PadRight(3)
            strBuff = strBuff + dbFld("FutanBefore").Value().PadRight(3)
            strBuff = strBuff + dbFld("Blinker").Value().PadRight(1)
            strBuff = strBuff + dbFld("reserved2").Value().PadRight(1)
            strBuff = strBuff + dbFld("KisyuCode").Value().PadRight(5)
            strBuff = strBuff + dbFld("KisyuCodeBefore").Value().PadRight(5)
            strBuff = strBuff + bPadR(dbFld("KisyuRyakusyo").Value(), 8)
            strBuff = strBuff + bPadR(dbFld("KisyuRyakusyoBefore").Value(), 8)
            strBuff = strBuff + dbFld("MinaraiCD").Value().PadRight(1)
            strBuff = strBuff + dbFld("MinaraiCDBefore").Value().PadRight(1)
            strBuff = strBuff + dbFld("BaTaijyu").Value().PadRight(3)
            strBuff = strBuff + dbFld("ZogenFugo").Value().PadRight(1)
            strBuff = strBuff + dbFld("ZogenSa").Value().PadRight(3)
            strBuff = strBuff + dbFld("IJyoCD").Value().PadRight(1)
            strBuff = strBuff + dbFld("NyusenJyuni").Value().PadRight(2)
            strBuff = strBuff + dbFld("KakuteiJyuni").Value().PadRight(2)
            strBuff = strBuff + dbFld("DochakuKubun").Value().PadRight(1)
            strBuff = strBuff + dbFld("DochakuTosu").Value().PadRight(1)
            strBuff = strBuff + dbFld("Time").Value().PadRight(4)
            strBuff = strBuff + dbFld("ChakusaCD").Value().PadRight(3)
            strBuff = strBuff + dbFld("ChakusaCDP").Value().PadRight(3)
            strBuff = strBuff + dbFld("ChakusaCDPP").Value().PadRight(3)
            strBuff = strBuff + dbFld("Jyuni1c").Value().PadRight(2)
            strBuff = strBuff + dbFld("Jyuni2c").Value().PadRight(2)
            strBuff = strBuff + dbFld("Jyuni3c").Value().PadRight(2)
            strBuff = strBuff + dbFld("Jyuni4c").Value().PadRight(2)
            strBuff = strBuff + dbFld("Odds").Value().PadRight(4)
            strBuff = strBuff + dbFld("Ninki").Value().PadRight(2)
            strBuff = strBuff + dbFld("Honsyokin").Value().PadRight(8)
            strBuff = strBuff + dbFld("Fukasyokin").Value().PadRight(8)
            strBuff = strBuff + dbFld("reserved3").Value().PadRight(3)
            strBuff = strBuff + dbFld("reserved4").Value().PadRight(3)
            strBuff = strBuff + dbFld("HaronTimeL4").Value().PadRight(3)
            strBuff = strBuff + dbFld("HaronTimeL3").Value().PadRight(3)
            For iLoopCnt = 0 To 2
                strBuff = strBuff + dbFld("KettoNum" & iLoopCnt + 1).Value().PadRight(10)
                strBuff = strBuff + bPadR(dbFld("Bamei" & iLoopCnt + 1).Value(), 36)
            Next iLoopCnt
            strBuff = strBuff + dbFld("TimeDiff").Value().PadRight(4)
            strBuff = strBuff + dbFld("RecordUpKubun").Value().PadRight(1)
            strBuff = strBuff + dbFld("DMKubun").Value().PadRight(1)
            strBuff = strBuff + dbFld("DMTime").Value().PadRight(5)
            strBuff = strBuff + dbFld("DMGosaP").Value().PadRight(4)
            strBuff = strBuff + dbFld("DMGosaM").Value().PadRight(4)
            strBuff = strBuff + dbFld("DMJyuni").Value().PadRight(2)
            strBuff = strBuff + dbFld("KyakusituKubun").Value().PadRight(1) + vbCrLf

            ' 構造体へ格納
            structSE(lRecCount).SetData(strBuff)

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
        SelectDB = structSE

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
    ' 引き数    : strBuf - JVData 識別子"SE" の１行
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
            ' 枠番
            mRS.Fields("Wakuban").Value = .Wakuban
            ' 馬番
            mRS.Fields("Umaban").Value = .Umaban
            ' 血統登録番号
            mRS.Fields("KettoNum").Value = .KettoNum
            ' 馬名
            mRS.Fields("Bamei").Value = .Bamei
            ' 馬記号コード
            mRS.Fields("UmaKigoCD").Value = .UmaKigoCD
            ' 性別コード
            mRS.Fields("SexCD").Value = .SexCD
            ' 品種コード
            mRS.Fields("HinsyuCD").Value = .HinsyuCD
            ' 毛色コード
            mRS.Fields("KeiroCD").Value = .KeiroCD
            ' 馬齢
            mRS.Fields("Barei").Value = .Barei
            ' 東西所属コード
            mRS.Fields("TozaiCD").Value = .TozaiCD
            ' 調教師コード
            mRS.Fields("ChokyosiCode").Value = .ChokyosiCode
            ' 調教師名略称
            mRS.Fields("ChokyosiRyakusyo").Value = .ChokyosiRyakusyo
            ' 馬主コード
            mRS.Fields("BanusiCode").Value = .BanusiCode
            ' 馬主名
            mRS.Fields("BanusiName").Value = .BanusiName
            ' 服色標示
            mRS.Fields("Fukusyoku").Value = .Fukusyoku
            ' 予備
            mRS.Fields("reserved1").Value = .reserved1
            ' 負担重量
            mRS.Fields("Futan").Value = .Futan
            ' 変更前負担重量
            mRS.Fields("FutanBefore").Value = .FutanBefore
            ' ブリンカー使用区分
            mRS.Fields("Blinker").Value = .Blinker
            ' 予備
            mRS.Fields("reserved2").Value = .reserved2
            ' 騎手コード
            mRS.Fields("KisyuCode").Value = .KisyuCode
            ' 変更前騎手コード
            mRS.Fields("KisyuCodeBefore").Value = .KisyuCodeBefore
            ' 騎手名略称
            mRS.Fields("KisyuRyakusyo").Value = .KisyuRyakusyo
            ' 変更前騎手名略称
            mRS.Fields("KisyuRyakusyoBefore").Value = .KisyuRyakusyoBefore
            ' 騎手見習コード
            mRS.Fields("MinaraiCD").Value = .MinaraiCD
            ' 変更前騎手見習コード
            mRS.Fields("MinaraiCDBefore").Value = .MinaraiCDBefore
            ' 馬体重
            mRS.Fields("BaTaijyu").Value = .BaTaijyu
            ' 増減符号
            mRS.Fields("ZogenFugo").Value = .ZogenFugo
            ' 増減差
            mRS.Fields("ZogenSa").Value = .ZogenSa
            ' 異常区分コード
            mRS.Fields("IJyoCD").Value = .IJyoCD
            ' 入線順位
            mRS.Fields("NyusenJyuni").Value = .NyusenJyuni
            ' 確定着順
            mRS.Fields("KakuteiJyuni").Value = .KakuteiJyuni
            ' 同着区分
            mRS.Fields("DochakuKubun").Value = .DochakuKubun
            ' 同着頭数
            mRS.Fields("DochakuTosu").Value = .DochakuTosu
            ' 走破タイム
            mRS.Fields("Time").Value = .Time
            ' 着差コード
            mRS.Fields("ChakusaCD").Value = .ChakusaCD
            ' +着差コード
            mRS.Fields("ChakusaCDP").Value = .ChakusaCDP
            ' ++着差コード
            mRS.Fields("ChakusaCDPP").Value = .ChakusaCDPP
            ' 1コーナーでの順位
            mRS.Fields("Jyuni1c").Value = .Jyuni1c
            ' 2コーナーでの順位
            mRS.Fields("Jyuni2c").Value = .Jyuni2c
            ' 3コーナーでの順位
            mRS.Fields("Jyuni3c").Value = .Jyuni3c
            ' 4コーナーでの順位
            mRS.Fields("Jyuni4c").Value = .Jyuni4c
            ' 単勝オッズ
            mRS.Fields("Odds").Value = .Odds
            ' 単勝人気順
            mRS.Fields("Ninki").Value = .Ninki
            ' 獲得本賞金
            mRS.Fields("Honsyokin").Value = .Honsyokin
            ' 獲得付加賞金
            mRS.Fields("Fukasyokin").Value = .Fukasyokin
            ' 予備
            mRS.Fields("reserved3").Value = .reserved3
            ' 予備
            mRS.Fields("reserved4").Value = .reserved4
            ' 後４ハロンタイム
            mRS.Fields("HaronTimeL4").Value = .HaronTimeL4
            ' 後３ハロンタイム
            mRS.Fields("HaronTimeL3").Value = .HaronTimeL3
            For iLoopCnt = 0 To 2
                With .ChakuUmaInfo(iLoopCnt)
                    ' 血統登録番号
                    mRS.Fields("KettoNum" & iLoopCnt + 1).Value = .KettoNum
                    ' 馬名
                    mRS.Fields("Bamei" & iLoopCnt + 1).Value = .Bamei
                End With ' ChakuUmaInfo
            Next iLoopCnt
            ' タイム差
            mRS.Fields("TimeDiff").Value = .TimeDiff
            ' レコード更新区分
            mRS.Fields("RecordUpKubun").Value = .RecordUpKubun
            ' マイニング区分
            mRS.Fields("DMKubun").Value = .DMKubun
            ' マイニング予想走破タイム
            mRS.Fields("DMTime").Value = .DMTime
            ' 予測誤差(信頼度)＋
            mRS.Fields("DMGosaP").Value = .DMGosaP
            ' 予測誤差(信頼度)－
            mRS.Fields("DMGosaM").Value = .DMGosaM
            ' マイニング予想順位
            mRS.Fields("DMJyuni").Value = .DMJyuni
            ' 今回レース脚質判定
            mRS.Fields("KyakusituKubun").Value = .KyakusituKubun
        End With

        With mBuf.id
            System.Diagnostics.Debug.WriteLine("Insert UMA_RACE : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum)
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

        strSql = "UPDATE UMA_RACE SET "
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
            ' 枠番
            strSql = strSql & SS & "Wakuban" & SE & "='" & Replace(.Wakuban, "'", "''") & "',"
            ' 馬番
            strSql = strSql & SS & "Umaban" & SE & "='" & Replace(.Umaban, "'", "''") & "',"
            ' 血統登録番号
            strSql = strSql & SS & "KettoNum" & SE & "='" & Replace(.KettoNum, "'", "''") & "',"
            ' 馬名
            strSql = strSql & SS & "Bamei" & SE & "='" & Replace(.Bamei, "'", "''") & "',"
            ' 馬記号コード
            strSql = strSql & SS & "UmaKigoCD" & SE & "='" & Replace(.UmaKigoCD, "'", "''") & "',"
            ' 性別コード
            strSql = strSql & SS & "SexCD" & SE & "='" & Replace(.SexCD, "'", "''") & "',"
            ' 品種コード
            strSql = strSql & SS & "HinsyuCD" & SE & "='" & Replace(.HinsyuCD, "'", "''") & "',"
            ' 毛色コード
            strSql = strSql & SS & "KeiroCD" & SE & "='" & Replace(.KeiroCD, "'", "''") & "',"
            ' 馬齢
            strSql = strSql & SS & "Barei" & SE & "='" & Replace(.Barei, "'", "''") & "',"
            ' 東西所属コード
            strSql = strSql & SS & "TozaiCD" & SE & "='" & Replace(.TozaiCD, "'", "''") & "',"
            ' 調教師コード
            strSql = strSql & SS & "ChokyosiCode" & SE & "='" & Replace(.ChokyosiCode, "'", "''") & "',"
            ' 調教師名略称
            strSql = strSql & SS & "ChokyosiRyakusyo" & SE & "='" & Replace(.ChokyosiRyakusyo, "'", "''") & "',"
            ' 馬主コード
            strSql = strSql & SS & "BanusiCode" & SE & "='" & Replace(.BanusiCode, "'", "''") & "',"
            ' 馬主名
            strSql = strSql & SS & "BanusiName" & SE & "='" & Replace(.BanusiName, "'", "''") & "',"
            ' 服色標示
            strSql = strSql & SS & "Fukusyoku" & SE & "='" & Replace(.Fukusyoku, "'", "''") & "',"
            ' 予備
            strSql = strSql & SS & "reserved1" & SE & "='" & Replace(.reserved1, "'", "''") & "',"
            ' 負担重量
            strSql = strSql & SS & "Futan" & SE & "='" & Replace(.Futan, "'", "''") & "',"
            ' 変更前負担重量
            strSql = strSql & SS & "FutanBefore" & SE & "='" & Replace(.FutanBefore, "'", "''") & "',"
            ' ブリンカー使用区分
            strSql = strSql & SS & "Blinker" & SE & "='" & Replace(.Blinker, "'", "''") & "',"
            ' 予備
            strSql = strSql & SS & "reserved2" & SE & "='" & Replace(.reserved2, "'", "''") & "',"
            ' 騎手コード
            strSql = strSql & SS & "KisyuCode" & SE & "='" & Replace(.KisyuCode, "'", "''") & "',"
            ' 変更前騎手コード
            strSql = strSql & SS & "KisyuCodeBefore" & SE & "='" & Replace(.KisyuCodeBefore, "'", "''") & "',"
            ' 騎手名略称
            strSql = strSql & SS & "KisyuRyakusyo" & SE & "='" & Replace(.KisyuRyakusyo, "'", "''") & "',"
            ' 変更前騎手名略称
            strSql = strSql & SS & "KisyuRyakusyoBefore" & SE & "='" & Replace(.KisyuRyakusyoBefore, "'", "''") & "',"
            ' 騎手見習コード
            strSql = strSql & SS & "MinaraiCD" & SE & "='" & Replace(.MinaraiCD, "'", "''") & "',"
            ' 変更前騎手見習コード
            strSql = strSql & SS & "MinaraiCDBefore" & SE & "='" & Replace(.MinaraiCDBefore, "'", "''") & "',"
            ' 馬体重
            strSql = strSql & SS & "BaTaijyu" & SE & "='" & Replace(.BaTaijyu, "'", "''") & "',"
            ' 増減符号
            strSql = strSql & SS & "ZogenFugo" & SE & "='" & Replace(.ZogenFugo, "'", "''") & "',"
            ' 増減差
            strSql = strSql & SS & "ZogenSa" & SE & "='" & Replace(.ZogenSa, "'", "''") & "',"
            ' 異常区分コード
            strSql = strSql & SS & "IJyoCD" & SE & "='" & Replace(.IJyoCD, "'", "''") & "',"
            ' 入線順位
            strSql = strSql & SS & "NyusenJyuni" & SE & "='" & Replace(.NyusenJyuni, "'", "''") & "',"
            ' 確定着順
            strSql = strSql & SS & "KakuteiJyuni" & SE & "='" & Replace(.KakuteiJyuni, "'", "''") & "',"
            ' 同着区分
            strSql = strSql & SS & "DochakuKubun" & SE & "='" & Replace(.DochakuKubun, "'", "''") & "',"
            ' 同着頭数
            strSql = strSql & SS & "DochakuTosu" & SE & "='" & Replace(.DochakuTosu, "'", "''") & "',"
            ' 走破タイム
            strSql = strSql & SS & "Time" & SE & "='" & Replace(.Time, "'", "''") & "',"
            ' 着差コード
            strSql = strSql & SS & "ChakusaCD" & SE & "='" & Replace(.ChakusaCD, "'", "''") & "',"
            ' +着差コード
            strSql = strSql & SS & "ChakusaCDP" & SE & "='" & Replace(.ChakusaCDP, "'", "''") & "',"
            ' ++着差コード
            strSql = strSql & SS & "ChakusaCDPP" & SE & "='" & Replace(.ChakusaCDPP, "'", "''") & "',"
            ' 1コーナーでの順位
            strSql = strSql & SS & "Jyuni1c" & SE & "='" & Replace(.Jyuni1c, "'", "''") & "',"
            ' 2コーナーでの順位
            strSql = strSql & SS & "Jyuni2c" & SE & "='" & Replace(.Jyuni2c, "'", "''") & "',"
            ' 3コーナーでの順位
            strSql = strSql & SS & "Jyuni3c" & SE & "='" & Replace(.Jyuni3c, "'", "''") & "',"
            ' 4コーナーでの順位
            strSql = strSql & SS & "Jyuni4c" & SE & "='" & Replace(.Jyuni4c, "'", "''") & "',"
            ' 単勝オッズ
            strSql = strSql & SS & "Odds" & SE & "='" & Replace(.Odds, "'", "''") & "',"
            ' 単勝人気順
            strSql = strSql & SS & "Ninki" & SE & "='" & Replace(.Ninki, "'", "''") & "',"
            ' 獲得本賞金
            strSql = strSql & SS & "Honsyokin" & SE & "='" & Replace(.Honsyokin, "'", "''") & "',"
            ' 獲得付加賞金
            strSql = strSql & SS & "Fukasyokin" & SE & "='" & Replace(.Fukasyokin, "'", "''") & "',"
            ' 予備
            strSql = strSql & SS & "reserved3" & SE & "='" & Replace(.reserved3, "'", "''") & "',"
            ' 予備
            strSql = strSql & SS & "reserved4" & SE & "='" & Replace(.reserved4, "'", "''") & "',"
            ' 後４ハロンタイム
            strSql = strSql & SS & "HaronTimeL4" & SE & "='" & Replace(.HaronTimeL4, "'", "''") & "',"
            ' 後３ハロンタイム
            strSql = strSql & SS & "HaronTimeL3" & SE & "='" & Replace(.HaronTimeL3, "'", "''") & "',"
            For iLoopCnt = 0 To 2
                With .ChakuUmaInfo(iLoopCnt)
                    ' 血統登録番号
                    strSql = strSql & SS & "KettoNum" & iLoopCnt + 1 & "" & SE & "='" & Replace(.KettoNum, "'", "''") & "',"
                    ' 馬名
                    strSql = strSql & SS & "Bamei" & iLoopCnt + 1 & "" & SE & "='" & Replace(.Bamei, "'", "''") & "',"
                End With ' ChakuUmaInfo
            Next iLoopCnt

            ' タイム差
            strSql = strSql & SS & "TimeDiff" & SE & "='" & Replace(.TimeDiff, "'", "''") & "',"
            ' レコード更新区分
            strSql = strSql & SS & "RecordUpKubun" & SE & "='" & Replace(.RecordUpKubun, "'", "''") & "',"
            ' マイニング区分
            strSql = strSql & SS & "DMKubun" & SE & "='" & Replace(.DMKubun, "'", "''") & "',"
            ' マイニング予想走破タイム
            strSql = strSql & SS & "DMTime" & SE & "='" & Replace(.DMTime, "'", "''") & "',"
            ' 予測誤差(信頼度)＋
            strSql = strSql & SS & "DMGosaP" & SE & "='" & Replace(.DMGosaP, "'", "''") & "',"
            ' 予測誤差(信頼度)－
            strSql = strSql & SS & "DMGosaM" & SE & "='" & Replace(.DMGosaM, "'", "''") & "',"
            ' マイニング予想順位
            strSql = strSql & SS & "DMJyuni" & SE & "='" & Replace(.DMJyuni, "'", "''") & "',"
            ' 今回レース脚質判定
            strSql = strSql & SS & "KyakusituKubun" & SE & "='" & Replace(.KyakusituKubun, "'", "''") & "',"

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
                strSql = strSql & " AND " & SS & "MakeDate" & SE & "<= '" & Replace(strMakeDate, "'", "''") & "'"
            End With
        End With

        With mBuf.id
            System.Diagnostics.Debug.WriteLine("UPDATE UMA_RACE : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum & mBuf.Umaban & mBuf.KettoNum)
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