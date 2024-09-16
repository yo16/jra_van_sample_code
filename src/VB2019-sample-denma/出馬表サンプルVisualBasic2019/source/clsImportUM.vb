' @(h) clsImportUM.vb
'
' @(s)
' JVData "UM" データベースアクセスクラス

Option Strict Off
Option Explicit On
Option Compare Binary

Friend Class clsImportUM

    '競走馬マスタ構造体
    Private mBuf As JV_UM_UMA
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
        strSql = "SELECT * FROM UMA"

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
    Public Function SelectDB(ByVal strSQL As String) As JV_UM_UMA()
        On Error GoTo ErrorHandler

        Dim IsDBOpen As Boolean = False

        ' ADODB.Recordset
        Dim dbRS As ADODB.Recordset

        ' ADODB.Fields
        Dim dbFld As ADODB.Fields

        ' レース詳細構造体
        Dim structUM(0) As JV_UM_UMA

        ' ループカウンタ
        Dim iLoopCnt1 As Integer
        Dim iLoopCnt2 As Integer

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

            ReDim Preserve structUM(lRecCount)

            ' 構造体設定用パラメータ作成
            strBuff = dbFld("RecordSpec").Value().PadRight(2)
            strBuff = strBuff + dbFld("DataKubun").Value().PadRight(1)
            strBuff = strBuff + dbFld("MakeDate").Value().PadRight(8)
            strBuff = strBuff + dbFld("KettoNum").Value().PadRight(10)
            strBuff = strBuff + dbFld("DelKubun").Value().PadRight(1)
            strBuff = strBuff + dbFld("RegDate").Value().PadRight(8)
            strBuff = strBuff + dbFld("DelDate").Value().PadRight(8)
            strBuff = strBuff + dbFld("BirthDate").Value().PadRight(8)
            strBuff = strBuff + bPadR(dbFld("Bamei").Value(), 36)
            strBuff = strBuff + dbFld("BameiKana").Value().PadRight(36)
            strBuff = strBuff + dbFld("BameiEng").Value().PadRight(60)
            strBuff = strBuff + dbFld("ZaikyuFlag").Value().PadRight(1)
            strBuff = strBuff + dbFld("Reserved").Value().PadRight(19)
            strBuff = strBuff + dbFld("UmaKigoCD").Value().PadRight(2)
            strBuff = strBuff + dbFld("SexCD").Value().PadRight(1)
            strBuff = strBuff + dbFld("HinsyuCD").Value().PadRight(1)
            strBuff = strBuff + dbFld("KeiroCD").Value().PadRight(2)
            For iLoopCnt1 = 0 To 13
                strBuff = strBuff + dbFld("Ketto3InfoHansyokuNum" & iLoopCnt1 + 1).Value().PadRight(10)
                strBuff = strBuff + bPadR(dbFld("Ketto3InfoBamei" & iLoopCnt1 + 1).Value(), 36)
            Next iLoopCnt1
            strBuff = strBuff + dbFld("TozaiCD").Value().PadRight(1)
            strBuff = strBuff + dbFld("ChokyosiCode").Value().PadRight(5)
            strBuff = strBuff + bPadR(dbFld("ChokyosiRyakusyo").Value(), 8)
            strBuff = strBuff + bPadR(dbFld("Syotai").Value(), 20)
            strBuff = strBuff + dbFld("BreederCode").Value().PadRight(8)
            strBuff = strBuff + bPadR(dbFld("BreederName").Value(), 72)
            strBuff = strBuff + bPadR(dbFld("SanchiName").Value(), 20)
            strBuff = strBuff + dbFld("BanusiCode").Value().PadRight(6)
            strBuff = strBuff + bPadR(dbFld("BanusiName").Value(), 64)
            strBuff = strBuff + dbFld("RuikeiHonsyoHeiti").Value().PadRight(9)
            strBuff = strBuff + dbFld("RuikeiHonsyoSyogai").Value().PadRight(9)
            strBuff = strBuff + dbFld("RuikeiFukaHeichi").Value().PadRight(9)
            strBuff = strBuff + dbFld("RuikeiFukaSyogai").Value().PadRight(9)
            strBuff = strBuff + dbFld("RuikeiSyutokuHeichi").Value().PadRight(9)
            strBuff = strBuff + dbFld("RuikeiSyutokuSyogai").Value().PadRight(9)
            For iLoopCnt1 = 0 To 5
                strBuff = strBuff + dbFld("SogoChakukaisu" & iLoopCnt1 + 1).Value().PadRight(3)
            Next iLoopCnt1
            For iLoopCnt1 = 0 To 5
                strBuff = strBuff + dbFld("ChuoChakukaisu" & iLoopCnt1 + 1).Value().PadRight(3)
            Next iLoopCnt1
            For iLoopCnt1 = 0 To 6
                For iLoopCnt2 = 0 To 5
                    strBuff = strBuff + dbFld("Ba" & iLoopCnt1 + 1 & "Chakukaisu" & iLoopCnt2 + 1).Value().PadRight(3)
                Next iLoopCnt2
            Next iLoopCnt1
            For iLoopCnt1 = 0 To 11
                For iLoopCnt2 = 0 To 5
                    strBuff = strBuff + dbFld("Jyotai" & iLoopCnt1 + 1 & "Chakukaisu" & iLoopCnt2 + 1).Value().PadRight(3)
                Next iLoopCnt2
            Next iLoopCnt1
            For iLoopCnt1 = 0 To 5
                For iLoopCnt2 = 0 To 5
                    strBuff = strBuff + dbFld("Kyori" & iLoopCnt1 + 1 & "Chakukaisu" & iLoopCnt2 + 1).Value().PadRight(3)
                Next iLoopCnt2
            Next iLoopCnt1
            For iLoopCnt1 = 0 To 3
                strBuff = strBuff + dbFld("Kyakusitu" & iLoopCnt1 + 1).Value().PadRight(3)
            Next iLoopCnt1
            strBuff = strBuff + dbFld("RaceCount").Value().PadRight(3) & vbCrLf

            ' 構造体へ格納
            structUM(lRecCount).SetData(strBuff)

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
        SelectDB = structUM

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
    ' 引き数    : strBuf - JVData 識別子"UM" の１行
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
        Dim iLoopCnt1 As Integer
        Dim iLoopCnt2 As Integer

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
            ' 血統登録番号
            mRS.Fields("KettoNum").Value = .KettoNum
            ' 競走馬抹消区分
            mRS.Fields("DelKubun").Value = .DelKubun
            With .RegDate
                ' 年月日
                mRS.Fields("RegDate").Value = .Year & .Month & .Day
            End With ' RegDate
            With .DelDate
                ' 年月日
                mRS.Fields("DelDate").Value = .Year & .Month & .Day
            End With ' DelDate
            With .BirthDate
                ' 年月日
                mRS.Fields("BirthDate").Value = .Year & .Month & .Day
            End With ' BirthDate
            ' 馬名
            mRS.Fields("Bamei").Value = .Bamei
            ' 馬名半角カナ
            mRS.Fields("BameiKana").Value = .BameiKana
            ' 馬名欧字
            mRS.Fields("BameiEng").Value = .BameiEng
            ' JRA施設在きゅうフラグ
            mRS.Fields("ZaikyuFlag").Value = .ZaikyuFlag
            ' 予備
            mRS.Fields("Reserved").Value = .Reserved
            ' 馬記号コード
            mRS.Fields("UmaKigoCD").Value = .UmaKigoCD
            ' 性別コード
            mRS.Fields("SexCD").Value = .SexCD
            ' 品種コード
            mRS.Fields("HinsyuCD").Value = .HinsyuCD
            ' 毛色コード
            mRS.Fields("KeiroCD").Value = .KeiroCD
            For iLoopCnt1 = 0 To 13
                With .Ketto3Info(iLoopCnt1)
                    ' 繁殖登録番号
                    mRS.Fields("Ketto3InfoHansyokuNum" & iLoopCnt1 + 1).Value = .HansyokuNum
                    ' 馬名
                    mRS.Fields("Ketto3InfoBamei" & iLoopCnt1 + 1).Value = .Bamei
                End With ' Ketto3Info
            Next iLoopCnt1
            ' 東西所属コード
            mRS.Fields("TozaiCD").Value = .TozaiCD
            ' 調教師コード
            mRS.Fields("ChokyosiCode").Value = .ChokyosiCode
            ' 調教師名略称
            mRS.Fields("ChokyosiRyakusyo").Value = .ChokyosiRyakusyo
            ' 招待地域名
            mRS.Fields("Syotai").Value = .Syotai
            ' 生産者コード
            mRS.Fields("BreederCode").Value = .BreederCode
            ' 生産者名
            mRS.Fields("BreederName").Value = .BreederName
            ' 産地名
            mRS.Fields("SanchiName").Value = .SanchiName
            ' 馬主コード
            mRS.Fields("BanusiCode").Value = .BanusiCode
            ' 馬主名
            mRS.Fields("BanusiName").Value = .BanusiName
            ' 平地本賞金累計
            mRS.Fields("RuikeiHonsyoHeiti").Value = .RuikeiHonsyoHeiti
            ' 障害本賞金累計
            mRS.Fields("RuikeiHonsyoSyogai").Value = .RuikeiHonsyoSyogai
            ' 平地付加賞金累計
            mRS.Fields("RuikeiFukaHeichi").Value = .RuikeiFukaHeichi
            ' 障害付加賞金累計
            mRS.Fields("RuikeiFukaSyogai").Value = .RuikeiFukaSyogai
            ' 平地収得賞金累計
            mRS.Fields("RuikeiSyutokuHeichi").Value = .RuikeiSyutokuHeichi
            ' 障害収得賞金累計
            mRS.Fields("RuikeiSyutokuSyogai").Value = .RuikeiSyutokuSyogai
            With .ChakuSogo
                For iLoopCnt1 = 0 To 5
                    mRS.Fields("SogoChakukaisu" & iLoopCnt1 + 1).Value = .Chakukaisu(iLoopCnt1)
                Next iLoopCnt1
            End With ' ChakuSogo
            With .ChakuChuo
                For iLoopCnt1 = 0 To 5
                    mRS.Fields("ChuoChakukaisu" & iLoopCnt1 + 1).Value = .Chakukaisu(iLoopCnt1)
                Next iLoopCnt1
            End With ' ChakuChuo
            For iLoopCnt1 = 0 To 6
                With .ChakuKaisuBa(iLoopCnt1)
                    For iLoopCnt2 = 0 To 5
                        mRS.Fields("Ba" & iLoopCnt1 + 1 & "Chakukaisu" & iLoopCnt2 + 1).Value = .Chakukaisu(iLoopCnt2)
                    Next iLoopCnt2
                End With ' ChakuKaisuBa
            Next iLoopCnt1
            For iLoopCnt1 = 0 To 11
                With .ChakuKaisuJyotai(iLoopCnt1)
                    For iLoopCnt2 = 0 To 5
                        mRS.Fields("Jyotai" & iLoopCnt1 + 1 & "Chakukaisu" & iLoopCnt2 + 1).Value = .Chakukaisu(iLoopCnt2)
                    Next iLoopCnt2
                End With ' ChakuKaisuJyotai
            Next iLoopCnt1
            For iLoopCnt1 = 0 To 5
                With .ChakuKaisuKyori(iLoopCnt1)
                    For iLoopCnt2 = 0 To 5
                        mRS.Fields("Kyori" & iLoopCnt1 + 1 & "Chakukaisu" & iLoopCnt2 + 1).Value = .Chakukaisu(iLoopCnt2)
                    Next iLoopCnt2
                End With ' ChakuKaisuKyori
            Next iLoopCnt1
            For iLoopCnt1 = 0 To 3
                ' 脚質傾向
                mRS.Fields("Kyakusitu" & iLoopCnt1 + 1).Value = .Kyakusitu(iLoopCnt1)
            Next iLoopCnt1
            ' 登録レース数
            mRS.Fields("RaceCount").Value = .RaceCount
        End With

        With mBuf
            System.Diagnostics.Debug.WriteLine("Insert UMA : " & .KettoNum)
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
        Dim iLoopCnt1 As Short
        Dim iLoopCnt2 As Short

        ' SQL文
        Dim strSql As String

        ' トランザクション開始
        gCon.BeginTrans()
        System.Diagnostics.Debug.WriteLine("BeginTrans")

        strSql = "UPDATE UMA SET "
        With mBuf
            ' 血統登録番号
            strSql = strSql & SS & "KettoNum" & SE & "='" & Replace(.KettoNum, "'", "''") & "',"
            ' 競走馬抹消区分
            strSql = strSql & SS & "DelKubun" & SE & "='" & Replace(.DelKubun, "'", "''") & "',"
            With .RegDate
                ' 年月日
                strSql = strSql & SS & "RegDate" & SE & "='" & Replace(.Year & .Month & .Day, "'", "''") & "',"
            End With ' RegDate
            With .DelDate
                ' 年月日
                strSql = strSql & SS & "DelDate" & SE & "='" & Replace(.Year & .Month & .Day, "'", "''") & "',"
            End With ' DelDate
            With .BirthDate
                ' 年月日
                strSql = strSql & SS & "BirthDate" & SE & "='" & Replace(.Year & .Month & .Day, "'", "''") & "',"
            End With ' BirthDate
            ' 馬名
            strSql = strSql & SS & "Bamei" & SE & "='" & Replace(.Bamei, "'", "''") & "',"
            ' 馬名半角カナ
            strSql = strSql & SS & "BameiKana" & SE & "='" & Replace(.BameiKana, "'", "''") & "',"
            ' 馬名欧字
            strSql = strSql & SS & "BameiEng" & SE & "='" & Replace(.BameiEng, "'", "''") & "',"
            ' JRA施設在きゅうフラグ
            strSql = strSql & SS & "ZaikyuFlag" & SE & "='" & Replace(.ZaikyuFlag, "'", "''") & "',"
            ' 予備
            strSql = strSql & SS & "Reserved" & SE & "='" & Replace(.Reserved, "'", "''") & "',"
            ' 馬記号コード
            strSql = strSql & SS & "UmaKigoCD" & SE & "='" & Replace(.UmaKigoCD, "'", "''") & "',"
            ' 性別コード
            strSql = strSql & SS & "SexCD" & SE & "='" & Replace(.SexCD, "'", "''") & "',"
            ' 品種コード
            strSql = strSql & SS & "HinsyuCD" & SE & "='" & Replace(.HinsyuCD, "'", "''") & "',"
            ' 毛色コード
            strSql = strSql & SS & "KeiroCD" & SE & "='" & Replace(.KeiroCD, "'", "''") & "',"
            For iLoopCnt1 = 0 To 13
                With .Ketto3Info(iLoopCnt1)
                    ' 繁殖登録番号
                    strSql = strSql & SS & "Ketto3InfoHansyokuNum" & iLoopCnt1 + 1 & "" & SE & "='" & Replace(.HansyokuNum, "'", "''") & "',"
                    ' 馬名
                    strSql = strSql & SS & "Ketto3InfoBamei" & iLoopCnt1 + 1 & "" & SE & "='" & Replace(.Bamei, "'", "''") & "',"
                End With ' Ketto3Info
            Next iLoopCnt1
            ' 東西所属コード
            strSql = strSql & SS & "TozaiCD" & SE & "='" & Replace(.TozaiCD, "'", "''") & "',"
            ' 調教師コード
            strSql = strSql & SS & "ChokyosiCode" & SE & "='" & Replace(.ChokyosiCode, "'", "''") & "',"
            ' 調教師名略称
            strSql = strSql & SS & "ChokyosiRyakusyo" & SE & "='" & Replace(.ChokyosiRyakusyo, "'", "''") & "',"
            ' 招待地域名
            strSql = strSql & SS & "Syotai" & SE & "='" & Replace(.Syotai, "'", "''") & "',"
            ' 生産者コード
            strSql = strSql & SS & "BreederCode" & SE & "='" & Replace(.BreederCode, "'", "''") & "',"
            ' 生産者名
            strSql = strSql & SS & "BreederName" & SE & "='" & Replace(.BreederName, "'", "''") & "',"
            ' 産地名
            strSql = strSql & SS & "SanchiName" & SE & "='" & Replace(.SanchiName, "'", "''") & "',"
            ' 馬主コード
            strSql = strSql & SS & "BanusiCode" & SE & "='" & Replace(.BanusiCode, "'", "''") & "',"
            ' 馬主名
            strSql = strSql & SS & "BanusiName" & SE & "='" & Replace(.BanusiName, "'", "''") & "',"
            ' 平地本賞金累計
            strSql = strSql & SS & "RuikeiHonsyoHeiti" & SE & "='" & Replace(.RuikeiHonsyoHeiti, "'", "''") & "',"
            ' 障害本賞金累計
            strSql = strSql & SS & "RuikeiHonsyoSyogai" & SE & "='" & Replace(.RuikeiHonsyoSyogai, "'", "''") & "',"
            ' 平地付加賞金累計
            strSql = strSql & SS & "RuikeiFukaHeichi" & SE & "='" & Replace(.RuikeiFukaHeichi, "'", "''") & "',"
            ' 障害付加賞金累計
            strSql = strSql & SS & "RuikeiFukaSyogai" & SE & "='" & Replace(.RuikeiFukaSyogai, "'", "''") & "',"
            ' 平地収得賞金累計
            strSql = strSql & SS & "RuikeiSyutokuHeichi" & SE & "='" & Replace(.RuikeiSyutokuHeichi, "'", "''") & "',"
            ' 障害収得賞金累計
            strSql = strSql & SS & "RuikeiSyutokuSyogai" & SE & "='" & Replace(.RuikeiSyutokuSyogai, "'", "''") & "',"
            With .ChakuSogo
                For iLoopCnt1 = 0 To 5
                    strSql = strSql & SS & "SogoChakukaisu" & iLoopCnt1 + 1 & "" & SE & "='" & Replace(.Chakukaisu(iLoopCnt1), "'", "''") & "',"
                Next iLoopCnt1
            End With ' ChakuSogo
            With .ChakuChuo
                For iLoopCnt1 = 0 To 5
                    strSql = strSql & SS & "ChuoChakukaisu" & iLoopCnt1 + 1 & "" & SE & "='" & Replace(.Chakukaisu(iLoopCnt1), "'", "''") & "',"
                Next iLoopCnt1
            End With ' ChakuChuo
            For iLoopCnt1 = 0 To 6
                With .ChakuKaisuBa(iLoopCnt1)
                    For iLoopCnt2 = 0 To 5
                        strSql = strSql & SS & "Ba" & iLoopCnt1 + 1 & "Chakukaisu" & iLoopCnt2 + 1 & "" & SE & "='" & Replace(.Chakukaisu(iLoopCnt2), "'", "''") & "'"
                        If iLoopCnt1 <> 6 Or iLoopCnt2 <> 5 Then
                            strSql = strSql & ","
                        End If
                    Next iLoopCnt2
                End With ' ChakuKaisuBa
            Next iLoopCnt1

            'strSql = Left(strSql, Len(strSql) - 1) ' Delete Last Comma

            strSql = strSql & " WHERE " & SS & "KettoNum" & SE & "='" & Replace(.KettoNum, "'", "''") & "'"
            strSql = strSql & " AND " & SS & "MakeDate" & SE & "<='" & Replace(strMakeDate, "'", "''") & "'"


            gCon.Execute(strSql)

            '一度に更新できるフィールド数が約127までの為 分割更新（JET仕様） 

            strSql = "UPDATE UMA SET "
            'ヘッダの更新は後半の更新で行う
            With .head
                ' レコード種別
                strSql = strSql & SS & "RecordSpec" & SE & "='" & Replace(.RecordSpec, "'", "''") & "',"
                ' データ区分
                strSql = strSql & SS & "DataKubun" & SE & "='" & Replace(.DataKubun, "'", "''") & "',"
                ' 年月日
                strSql = strSql & SS & "MakeDate" & SE & "='" & Replace(strMakeDate, "'", "''") & "',"
            End With ' head
            For iLoopCnt1 = 0 To 11
                With .ChakuKaisuJyotai(iLoopCnt1)
                    For iLoopCnt2 = 0 To 5
                        strSql = strSql & SS & "Jyotai" & iLoopCnt1 + 1 & "Chakukaisu" & iLoopCnt2 + 1 & "" & SE & "='" & Replace(.Chakukaisu(iLoopCnt2), "'", "''") & "',"
                    Next iLoopCnt2
                End With ' ChakuKaisuJyotai
            Next iLoopCnt1
            For iLoopCnt1 = 0 To 5
                With .ChakuKaisuKyori(iLoopCnt1)
                    For iLoopCnt2 = 0 To 5
                        strSql = strSql & SS & "Kyori" & iLoopCnt1 + 1 & "Chakukaisu" & iLoopCnt2 + 1 & "" & SE & "='" & Replace(.Chakukaisu(iLoopCnt2), "'", "''") & "',"
                    Next iLoopCnt2
                End With ' ChakuKaisuKyori
            Next iLoopCnt1
            For iLoopCnt1 = 0 To 3
                ' 脚質傾向
                strSql = strSql & SS & "Kyakusitu" & iLoopCnt1 + 1 & "" & SE & "='" & Replace(.Kyakusitu(iLoopCnt1), "'", "''") & "',"
            Next iLoopCnt1
            ' 登録レース数
            strSql = strSql & SS & "RaceCount" & SE & "='" & Replace(.RaceCount, "'", "''") & "'"
            strSql = strSql & " WHERE " & SS & "KettoNum" & SE & "='" & Replace(.KettoNum, "'", "''") & "'"
            strSql = strSql & " AND " & SS & "MakeDate" & SE & "<= '" & Replace(strMakeDate, "'", "''") & "'"
        End With

        With mBuf
            System.Diagnostics.Debug.WriteLine("UPDATE UMA : " & .KettoNum)
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