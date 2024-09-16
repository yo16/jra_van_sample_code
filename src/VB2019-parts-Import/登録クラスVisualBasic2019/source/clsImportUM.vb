Option Strict Off
Option Explicit On
Option Compare Binary
Friend Class clsImportUM
	' @(h) clsReadUM.cls
	' @(s)
	' JVData "UM" データベース登録クラス
	'
	
	Private mBuf As JV_UM_UMA ''競走馬マスタ構造体
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
        strSql = "SELECT * FROM UMA"
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
    ' 引き数    : lBuf - JVData 識別子"UM" の１行
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


    '
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
    Private Function InsertDB() As Boolean
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
            mRS.Fields("DelKubun").Value = .DelKubun '' 競走馬抹消区分
            With .RegDate
                mRS.Fields("RegDate").Value = .Year & .Month & .Day '' 年月日
            End With ' RegDate
            With .DelDate
                mRS.Fields("DelDate").Value = .Year & .Month & .Day '' 年月日
            End With ' DelDate
            With .BirthDate
                mRS.Fields("BirthDate").Value = .Year & .Month & .Day '' 年月日
            End With ' BirthDate
            mRS.Fields("Bamei").Value = .Bamei '' 馬名
            mRS.Fields("BameiKana").Value = .BameiKana '' 馬名半角カナ
            mRS.Fields("BameiEng").Value = .BameiEng '' 馬名欧字
            mRS.Fields("ZaikyuFlag").Value = .ZaikyuFlag '' JRA施設在きゅうフラグ
            mRS.Fields("Reserved").Value = .Reserved '' 予備
            mRS.Fields("UmaKigoCD").Value = .UmaKigoCD '' 馬記号コード
            mRS.Fields("SexCD").Value = .SexCD '' 性別コード
            mRS.Fields("HinsyuCD").Value = .HinsyuCD '' 品種コード
            mRS.Fields("KeiroCD").Value = .KeiroCD '' 毛色コード
            For i = 0 To 13
                With .Ketto3Info(i)
                    mRS.Fields("Ketto3InfoHansyokuNum" & i + 1).Value = .HansyokuNum '' 繁殖登録番号
                    mRS.Fields("Ketto3InfoBamei" & i + 1).Value = .Bamei '' 馬名
                End With ' Ketto3Info
            Next i
            mRS.Fields("TozaiCD").Value = .TozaiCD '' 東西所属コード
            mRS.Fields("ChokyosiCode").Value = .ChokyosiCode '' 調教師コード
            mRS.Fields("ChokyosiRyakusyo").Value = .ChokyosiRyakusyo '' 調教師名略称
            mRS.Fields("Syotai").Value = .Syotai '' 招待地域名
            mRS.Fields("BreederCode").Value = .BreederCode '' 生産者コード
            mRS.Fields("BreederName").Value = .BreederName '' 生産者名
            mRS.Fields("SanchiName").Value = .SanchiName '' 産地名
            mRS.Fields("BanusiCode").Value = .BanusiCode '' 馬主コード
            mRS.Fields("BanusiName").Value = .BanusiName '' 馬主名
            mRS.Fields("RuikeiHonsyoHeiti").Value = .RuikeiHonsyoHeiti '' 平地本賞金累計
            mRS.Fields("RuikeiHonsyoSyogai").Value = .RuikeiHonsyoSyogai '' 障害本賞金累計
            mRS.Fields("RuikeiFukaHeichi").Value = .RuikeiFukaHeichi '' 平地付加賞金累計
            mRS.Fields("RuikeiFukaSyogai").Value = .RuikeiFukaSyogai '' 障害付加賞金累計
            mRS.Fields("RuikeiSyutokuHeichi").Value = .RuikeiSyutokuHeichi '' 平地収得賞金累計
            mRS.Fields("RuikeiSyutokuSyogai").Value = .RuikeiSyutokuSyogai '' 障害収得賞金累計
            With .ChakuSogo
                For j = 0 To 5
                    mRS.Fields("SogoChakukaisu" & j + 1).Value = .Chakukaisu(j)
                Next j
            End With ' ChakuSogo
            With .ChakuChuo
                For j = 0 To 5
                    mRS.Fields("ChuoChakukaisu" & j + 1).Value = .Chakukaisu(j)
                Next j
            End With ' ChakuChuo
            For i = 0 To 6
                With .ChakuKaisuBa(i)
                    For j = 0 To 5
                        mRS.Fields("Ba" & i + 1 & "Chakukaisu" & j + 1).Value = .Chakukaisu(j)
                    Next j
                End With ' ChakuKaisuBa
            Next i
            For i = 0 To 11
                With .ChakuKaisuJyotai(i)
                    For j = 0 To 5
                        mRS.Fields("Jyotai" & i + 1 & "Chakukaisu" & j + 1).Value = .Chakukaisu(j)
                    Next j
                End With ' ChakuKaisuJyotai
            Next i
            For i = 0 To 5
                With .ChakuKaisuKyori(i)
                    For j = 0 To 5
                        mRS.Fields("Kyori" & i + 1 & "Chakukaisu" & j + 1).Value = .Chakukaisu(j)
                    Next j
                End With ' ChakuKaisuKyori
            Next i
            For i = 0 To 3
                mRS.Fields("Kyakusitu" & i + 1).Value = .Kyakusitu(i) '' 脚質傾向
            Next i
            mRS.Fields("RaceCount").Value = .RaceCount '' 登録レース数
        End With

        With mBuf
            System.Diagnostics.Debug.WriteLine("Insert UMA : " & .KettoNum)
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

        With mBuf
            strSql = "UPDATE UMA SET "
            strSql = strSql & SS & "KettoNum" & SE & "='" & Replace(.KettoNum, "'", "''") & "'," '' 血統登録番号
            strSql = strSql & SS & "DelKubun" & SE & "='" & Replace(.DelKubun, "'", "''") & "'," '' 競走馬抹消区分
            With .RegDate
                strSql = strSql & SS & "RegDate" & SE & "='" & Replace(.Year & .Month & .Day, "'", "''") & "'," '' 年月日
            End With ' RegDate
            With .DelDate
                strSql = strSql & SS & "DelDate" & SE & "='" & Replace(.Year & .Month & .Day, "'", "''") & "'," '' 年月日
            End With ' DelDate
            With .BirthDate
                strSql = strSql & SS & "BirthDate" & SE & "='" & Replace(.Year & .Month & .Day, "'", "''") & "'," '' 年月日
            End With ' BirthDate
            strSql = strSql & SS & "Bamei" & SE & "='" & Replace(.Bamei, "'", "''") & "'," '' 馬名
            strSql = strSql & SS & "BameiKana" & SE & "='" & Replace(.BameiKana, "'", "''") & "'," '' 馬名半角カナ
            strSql = strSql & SS & "BameiEng" & SE & "='" & Replace(.BameiEng, "'", "''") & "'," '' 馬名欧字
            strSql = strSql & SS & "ZaikyuFlag" & SE & "='" & Replace(.ZaikyuFlag, "'", "''") & "'," '' JRA施設在きゅうフラグ
            strSql = strSql & SS & "Reserved" & SE & "='" & Replace(.Reserved, "'", "''") & "'," '' 予備
            strSql = strSql & SS & "UmaKigoCD" & SE & "='" & Replace(.UmaKigoCD, "'", "''") & "'," '' 馬記号コード
            strSql = strSql & SS & "SexCD" & SE & "='" & Replace(.SexCD, "'", "''") & "'," '' 性別コード
            strSql = strSql & SS & "HinsyuCD" & SE & "='" & Replace(.HinsyuCD, "'", "''") & "'," '' 品種コード
            strSql = strSql & SS & "KeiroCD" & SE & "='" & Replace(.KeiroCD, "'", "''") & "'," '' 毛色コード
            For i = 0 To 13
                With .Ketto3Info(i)
                    strSql = strSql & SS & "Ketto3InfoHansyokuNum" & i + 1 & SE & "='" & Replace(.HansyokuNum, "'", "''") & "'," '' 繁殖登録番号
                    strSql = strSql & SS & "Ketto3InfoBamei" & i + 1 & SE & "='" & Replace(.Bamei, "'", "''") & "'," '' 馬名
                End With ' Ketto3Info
            Next i
            strSql = strSql & SS & "TozaiCD" & SE & "='" & Replace(.TozaiCD, "'", "''") & "'," '' 東西所属コード
            strSql = strSql & SS & "ChokyosiCode" & SE & "='" & Replace(.ChokyosiCode, "'", "''") & "'," '' 調教師コード
            strSql = strSql & SS & "ChokyosiRyakusyo" & SE & "='" & Replace(.ChokyosiRyakusyo, "'", "''") & "'," '' 調教師名略称
            strSql = strSql & SS & "Syotai" & SE & "='" & Replace(.Syotai, "'", "''") & "'," '' 招待地域名
            strSql = strSql & SS & "BreederCode" & SE & "='" & Replace(.BreederCode, "'", "''") & "'," '' 生産者コード
            strSql = strSql & SS & "BreederName" & SE & "='" & Replace(.BreederName, "'", "''") & "'," '' 生産者名
            strSql = strSql & SS & "SanchiName" & SE & "='" & Replace(.SanchiName, "'", "''") & "'," '' 産地名
            strSql = strSql & SS & "BanusiCode" & SE & "='" & Replace(.BanusiCode, "'", "''") & "'," '' 馬主コード
            strSql = strSql & SS & "BanusiName" & SE & "='" & Replace(.BanusiName, "'", "''") & "'," '' 馬主名
            strSql = strSql & SS & "RuikeiHonsyoHeiti" & SE & "='" & Replace(.RuikeiHonsyoHeiti, "'", "''") & "'," '' 平地本賞金累計
            strSql = strSql & SS & "RuikeiHonsyoSyogai" & SE & "='" & Replace(.RuikeiHonsyoSyogai, "'", "''") & "'," '' 障害本賞金累計
            strSql = strSql & SS & "RuikeiFukaHeichi" & SE & "='" & Replace(.RuikeiFukaHeichi, "'", "''") & "'," '' 平地付加賞金累計
            strSql = strSql & SS & "RuikeiFukaSyogai" & SE & "='" & Replace(.RuikeiFukaSyogai, "'", "''") & "'," '' 障害付加賞金累計
            strSql = strSql & SS & "RuikeiSyutokuHeichi" & SE & "='" & Replace(.RuikeiSyutokuHeichi, "'", "''") & "'," '' 平地収得賞金累計
            strSql = strSql & SS & "RuikeiSyutokuSyogai" & SE & "='" & Replace(.RuikeiSyutokuSyogai, "'", "''") & "'," '' 障害収得賞金累計
            With .ChakuSogo
                For j = 0 To 5
                    strSql = strSql & SS & "SogoChakukaisu" & j + 1 & SE & "='" & Replace(.Chakukaisu(j), "'", "''") & "',"
                Next j
            End With ' ChakuSogo
            With .ChakuChuo
                For j = 0 To 5
                    strSql = strSql & SS & "ChuoChakukaisu" & j + 1 & SE & "='" & Replace(.Chakukaisu(j), "'", "''") & "',"
                Next j
            End With ' ChakuChuo
            For i = 0 To 6
                With .ChakuKaisuBa(i)
                    For j = 0 To 5
                        strSql = strSql & SS & "Ba" & i + 1 & "Chakukaisu" & j + 1 & SE & "='" & Replace(.Chakukaisu(j), "'", "''") & "'"
                        If i <> 6 Or j <> 5 Then
                            strSql = strSql & ","
                        End If
                    Next j
                End With ' ChakuKaisuBa
            Next i

            'strSql = Left(strSql, Len(strSql) - 1) ' Delete Last Comma

            strSql = strSql & " WHERE " & SS & "KettoNum" & SE & " ='" & Replace(.KettoNum, "'", "''") & "'"
            strSql = strSql & " AND " & SS & "MakeDate" & SE & " <='" & Replace(strMakeDate, "'", "''") & "'"


            gCon.Execute(strSql)

            ''一度に更新できるフィールド数が約127までの為 分割更新（JET仕様）
            strSql = "UPDATE UMA SET "
            'ヘッダの更新は後半の更新で行う
            With .head
                strSql = strSql & SS & "RecordSpec" & SE & "='" & Replace(.RecordSpec, "'", "''") & "'," '' レコード種別
                strSql = strSql & SS & "DataKubun" & SE & "='" & Replace(.DataKubun, "'", "''") & "'," '' データ区分
                strSql = strSql & SS & "MakeDate" & SE & "='" & Replace(strMakeDate, "'", "''") & "'," '' 年月日
            End With ' head
            For i = 0 To 11
                With .ChakuKaisuJyotai(i)
                    For j = 0 To 5
                        strSql = strSql & SS & "Jyotai" & i + 1 & "Chakukaisu" & j + 1 & SE & "='" & Replace(.Chakukaisu(j), "'", "''") & "',"
                    Next j
                End With ' ChakuKaisuJyotai
            Next i
            For i = 0 To 5
                With .ChakuKaisuKyori(i)
                    For j = 0 To 5
                        strSql = strSql & SS & "Kyori" & i + 1 & "Chakukaisu" & j + 1 & SE & "='" & Replace(.Chakukaisu(j), "'", "''") & "',"
                    Next j
                End With ' ChakuKaisuKyori
            Next i
            For i = 0 To 3
                strSql = strSql & SS & "Kyakusitu" & i + 1 & SE & "='" & Replace(.Kyakusitu(i), "'", "''") & "'," '' 脚質傾向
            Next i
            strSql = strSql & SS & "RaceCount" & SE & "='" & Replace(.RaceCount, "'", "''") & "'" '' 登録レース数
            strSql = strSql & " WHERE " & SS & "KettoNum" & SE & "='" & Replace(.KettoNum, "'", "''") & "'"
            strSql = strSql & " AND " & SS & "MakeDate" & SE & " <= '" & Replace(strMakeDate, "'", "''") & "'"


            gCon.Execute(strSql)

        End With

        With mBuf
            System.Diagnostics.Debug.WriteLine("UPDATE UMA : " & .KettoNum)
        End With ' id

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