Option Strict Off
Option Explicit On
Option Compare Binary
Friend Class clsImportKS
	' @(h) clsReadKS.cls
	' @(s)
	' JVData "KS" データベース登録クラス
	'
	
	Private mBuf As JV_KS_KISYU '' 騎手マスタ構造体
	Private mRS1 As ADODB.Recordset
	Private mRS2 As ADODB.Recordset
	
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
        strSql = "SELECT * FROM KISYU"
        mRS1 = New ADODB.Recordset()
        mRS1.Open(strSql, gCon, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockPessimistic)

        strSql = "SELECT * FROM KISYU_SEISEKI"
        mRS2 = New ADODB.Recordset()
        mRS2.Open(strSql, gCon, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockPessimistic)


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
        mRS1.Close()

        mRS1 = Nothing
        mRS2.Close()

        mRS2 = Nothing
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
    ' 引き数    : lBuf - JVData 識別子"KS" の１行
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

        mRS1.AddNew()

        With mBuf
            With .head
                mRS1.Fields("RecordSpec").Value = .RecordSpec '' レコード種別
                mRS1.Fields("DataKubun").Value = .DataKubun '' データ区分
                With .MakeDate
                    mRS1.Fields("MakeDate").Value = .Year & .Month & .Day '' 年月日
                End With ' MakeDate
            End With ' head
            mRS1.Fields("KisyuCode").Value = .KisyuCode '' 騎手コード
            mRS1.Fields("DelKubun").Value = .DelKubun '' 騎手抹消区分
            With .IssueDate
                mRS1.Fields("IssueDate").Value = .Year & .Month & .Day '' 年月日
            End With ' IssueDate
            With .DelDate
                mRS1.Fields("DelDate").Value = .Year & .Month & .Day '' 年月日
            End With ' DelDate
            With .BirthDate
                mRS1.Fields("BirthDate").Value = .Year & .Month & .Day '' 年月日
            End With ' BirthDate
            mRS1.Fields("KisyuName").Value = .KisyuName '' 騎手名漢字
            mRS1.Fields("reserved").Value = .reserved '' 予備
            mRS1.Fields("KisyuNameKana").Value = .KisyuNameKana '' 騎手名半角カナ
            mRS1.Fields("KisyuRyakusyo").Value = .KisyuRyakusyo '' 騎手名略称
            mRS1.Fields("KisyuNameEng").Value = .KisyuNameEng '' 騎手名欧字
            mRS1.Fields("SexCD").Value = .SexCD '' 性別区分
            mRS1.Fields("SikakuCD").Value = .SikakuCD '' 騎乗資格コード
            mRS1.Fields("MinaraiCD").Value = .MinaraiCD '' 騎手見習コード
            mRS1.Fields("TozaiCD").Value = .TozaiCD '' 騎手東西所属コード
            mRS1.Fields("Syotai").Value = .Syotai '' 招待地域名
            mRS1.Fields("ChokyosiCode").Value = .ChokyosiCode '' 所属調教師コード
            mRS1.Fields("ChokyosiRyakusyo").Value = .ChokyosiRyakusyo '' 所属調教師名略称
            For i = 0 To 1
                With .HatuKiJyo(i)
                    With .Hatukijyoid
                        mRS1.Fields("HatuKiJyo" & i + 1 & "Hatukijyoid").Value = .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum
                        '' 開催年 開催月日 競馬場コード 開催回[第N回] 開催日目[N日目] レース番号
                    End With ' Hatukijyoid
                    mRS1.Fields("HatuKiJyo" & i + 1 & "SyussoTosu").Value = .SyussoTosu '' 出走頭数
                    mRS1.Fields("HatuKiJyo" & i + 1 & "KettoNum").Value = .KettoNum '' 血統登録番号
                    mRS1.Fields("HatuKiJyo" & i + 1 & "Bamei").Value = .Bamei '' 馬名
                    mRS1.Fields("HatuKiJyo" & i + 1 & "KakuteiJyuni").Value = .KakuteiJyuni '' 確定着順
                    mRS1.Fields("HatuKiJyo" & i + 1 & "IJyoCD").Value = .IJyoCD '' 異常区分コード
                End With ' HatuKiJyo
            Next i
            For i = 0 To 1
                With .HatuSyori(i)
                    With .Hatusyoriid
                        mRS1.Fields("HatuSyori" & i + 1 & "Hatusyoriid").Value = .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum
                        '' 開催年 開催月日 競馬場コード 開催回[第N回] 開催日目[N日目] レース番号
                    End With ' Hatusyoriid
                    mRS1.Fields("HatuSyori" & i + 1 & "SyussoTosu").Value = .SyussoTosu '' 出走頭数
                    mRS1.Fields("HatuSyori" & i + 1 & "KettoNum").Value = .KettoNum '' 血統登録番号
                    mRS1.Fields("HatuSyori" & i + 1 & "Bamei").Value = .Bamei '' 馬名
                End With ' HatuSyori
            Next i
            For i = 0 To 2
                With .SaikinJyusyo(i)
                    With .SaikinJyusyoid
                        mRS1.Fields("SaikinJyusyo" & i + 1 & "SaikinJyusyoid").Value = .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum
                        '' 開催年 開催月日 競馬場コード 開催回[第N回] 開催日目[N日目] レース番号
                    End With ' SaikinJyusyoid
                    mRS1.Fields("SaikinJyusyo" & i + 1 & "Hondai").Value = .Hondai '' 競走名本題
                    mRS1.Fields("SaikinJyusyo" & i + 1 & "Ryakusyo10").Value = .Ryakusyo10 '' 競走名略称10字
                    mRS1.Fields("SaikinJyusyo" & i + 1 & "Ryakusyo6").Value = .Ryakusyo6 '' 競走名略称6字
                    mRS1.Fields("SaikinJyusyo" & i + 1 & "Ryakusyo3").Value = .Ryakusyo3 '' 競走名略称3字
                    mRS1.Fields("SaikinJyusyo" & i + 1 & "GradeCD").Value = .GradeCD '' グレードコード
                    mRS1.Fields("SaikinJyusyo" & i + 1 & "SyussoTosu").Value = .SyussoTosu '' 出走頭数
                    mRS1.Fields("SaikinJyusyo" & i + 1 & "KettoNum").Value = .KettoNum '' 血統登録番号
                    mRS1.Fields("SaikinJyusyo" & i + 1 & "Bamei").Value = .Bamei '' 馬名
                End With ' SaikinJyusyo
            Next i
        End With ' mBuf

        With mBuf
            System.Diagnostics.Debug.WriteLine("Insert KISYU : " & .KisyuCode)
        End With

        mRS1.Update()

        '' KISYU_SEISEKI (騎手マスタ_成績)
        For i = 0 To 2
            With mBuf
                mRS2.AddNew()
                With .head
                    With .MakeDate
                        mRS2.Fields("MakeDate").Value = .Year & .Month & .Day '' 年月日
                    End With ' MakeDate
                End With ' head
                mRS2.Fields("KisyuCode").Value = .KisyuCode '' 騎手コード
                mRS2.Fields("Num").Value = i '' 連番
                With .HonZenRuikei(i)
                    mRS2.Fields("SetYear").Value = .SetYear '' 設定年
                    mRS2.Fields("HonSyokinHeichi").Value = .HonSyokinHeichi '' 平地本賞金合計
                    mRS2.Fields("HonSyokinSyogai").Value = .HonSyokinSyogai '' 障害本賞金合計
                    mRS2.Fields("FukaSyokinHeichi").Value = .FukaSyokinHeichi '' 平地付加賞金合計
                    mRS2.Fields("FukaSyokinSyogai").Value = .FukaSyokinSyogai '' 障害付加賞金合計
                    With .ChakuKaisuHeichi
                        For k = 0 To 5
                            mRS2.Fields("HeichiChakukaisu" & k + 1).Value = .Chakukaisu(k)
                        Next k
                    End With ' ChakuKaisuHeichi
                    With .ChakuKaisuSyogai
                        For k = 0 To 5
                            mRS2.Fields("SyogaiChakukaisu" & k + 1).Value = .Chakukaisu(k)
                        Next k
                    End With ' ChakuKaisuSyogai
                    For j = 0 To 19
                        With .ChakuKaisuJyo(j)
                            For k = 0 To 5
                                mRS2.Fields("Jyo" & j + 1 & "Chakukaisu" & k + 1).Value = .Chakukaisu(k)
                            Next k
                        End With ' ChakuKaisuJyo
                    Next j
                    For j = 0 To 5
                        With .ChakuKaisuKyori(j)
                            For k = 0 To 5
                                mRS2.Fields("Kyori" & j + 1 & "Chakukaisu" & k + 1).Value = .Chakukaisu(k)
                            Next k
                        End With ' ChakuKaisuKyori
                    Next j
                End With ' HonZenRuikei
            End With

            With mBuf
                System.Diagnostics.Debug.WriteLine("Insert KISYU_SEISEKI : " & .KisyuCode & CStr(i))
            End With

            mRS2.Update()

        Next i

        gCon.CommitTrans()
        System.Diagnostics.Debug.WriteLine("CommitTrans")

        InsertDB = True

ExitHandler:
        Exit Function
ErrorHandler:
        InsertDB = False
        gCon.RollbackTrans()
        mRS1.CancelUpdate()
        mRS2.CancelUpdate()
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

        '' KISYU(騎手マスタ)
        strSql = "UPDATE KISYU SET "
        With mBuf
            With .head
                strSql = strSql & SS & "RecordSpec" & SE & "='" & Replace(.RecordSpec, "'", "''") & "'," '' レコード種別
                strSql = strSql & SS & "DataKubun" & SE & "='" & Replace(.DataKubun, "'", "''") & "'," '' データ区分
                strSql = strSql & SS & "MakeDate" & SE & "= '" & Replace(strMakeDate, "'", "''") & "'," '' 年月日
            End With ' head
            strSql = strSql & SS & "KisyuCode" & SE & "='" & Replace(.KisyuCode, "'", "''") & "'," '' 騎手コード
            strSql = strSql & SS & "DelKubun" & SE & "='" & Replace(.DelKubun, "'", "''") & "'," '' 騎手抹消区分
            With .IssueDate
                strSql = strSql & SS & "IssueDate" & SE & "= '" & Replace(.Year & .Month & .Day, "'", "''") & "'," '' 年月日
            End With ' IssueDate
            With .DelDate
                strSql = strSql & SS & "DelDate" & SE & "= '" & Replace(.Year & .Month & .Day, "'", "''") & "'," '' 年月日
            End With ' DelDate
            With .BirthDate
                strSql = strSql & SS & "BirthDate" & SE & "= '" & Replace(.Year & .Month & .Day, "'", "''") & "'," '' 年月日
            End With ' BirthDate
            strSql = strSql & SS & "KisyuName" & SE & "='" & Replace(.KisyuName, "'", "''") & "'," '' 騎手名漢字
            strSql = strSql & SS & "reserved" & SE & "='" & Replace(.reserved, "'", "''") & "'," '' 予備
            strSql = strSql & SS & "KisyuNameKana" & SE & "='" & Replace(.KisyuNameKana, "'", "''") & "'," '' 騎手名半角カナ
            strSql = strSql & SS & "KisyuRyakusyo" & SE & "='" & Replace(.KisyuRyakusyo, "'", "''") & "'," '' 騎手名略称
            strSql = strSql & SS & "KisyuNameEng" & SE & "='" & Replace(.KisyuNameEng, "'", "''") & "'," '' 騎手名欧字
            strSql = strSql & SS & "SexCD" & SE & "='" & Replace(.SexCD, "'", "''") & "'," '' 性別区分
            strSql = strSql & SS & "SikakuCD" & SE & "='" & Replace(.SikakuCD, "'", "''") & "'," '' 騎乗資格コード
            strSql = strSql & SS & "MinaraiCD" & SE & "='" & Replace(.MinaraiCD, "'", "''") & "'," '' 騎手見習コード
            strSql = strSql & SS & "TozaiCD" & SE & "='" & Replace(.TozaiCD, "'", "''") & "'," '' 騎手東西所属コード
            strSql = strSql & SS & "Syotai" & SE & "='" & Replace(.Syotai, "'", "''") & "'," '' 招待地域名
            strSql = strSql & SS & "ChokyosiCode" & SE & "='" & Replace(.ChokyosiCode, "'", "''") & "'," '' 所属調教師コード
            strSql = strSql & SS & "ChokyosiRyakusyo" & SE & "='" & Replace(.ChokyosiRyakusyo, "'", "''") & "'," '' 所属調教師名略称
            For i = 0 To 1
                With .HatuKiJyo(i)
                    With .Hatukijyoid
                        strSql = strSql & SS & "HatuKiJyo" & i + 1 & "Hatukijyoid" & SE & "='" & Replace(.Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum, "'", "''") & "',"
                        '' 開催年 開催月日 競馬場コード 開催回[第N回] 開催日目[N日目] レース番号
                    End With ' Hatukijyoid
                    strSql = strSql & SS & "HatuKiJyo" & i + 1 & "SyussoTosu" & SE & "='" & Replace(.SyussoTosu, "'", "''") & "'," '' 出走頭数
                    strSql = strSql & SS & "HatuKiJyo" & i + 1 & "KettoNum" & SE & "='" & Replace(.KettoNum, "'", "''") & "'," '' 血統登録番号
                    strSql = strSql & SS & "HatuKiJyo" & i + 1 & "Bamei" & SE & "='" & Replace(.Bamei, "'", "''") & "'," '' 馬名
                    strSql = strSql & SS & "HatuKiJyo" & i + 1 & "KakuteiJyuni" & SE & "='" & Replace(.KakuteiJyuni, "'", "''") & "'," '' 確定着順
                    strSql = strSql & SS & "HatuKiJyo" & i + 1 & "IJyoCD" & SE & "='" & Replace(.IJyoCD, "'", "''") & "'," '' 異常区分コード
                End With ' HatuKiJyo
            Next i
            For i = 0 To 1
                With .HatuSyori(i)
                    With .Hatusyoriid
                        strSql = strSql & SS & "HatuSyori" & i + 1 & "Hatusyoriid" & SE & "='" & Replace(.Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum, "'", "''") & "',"
                        '' 開催年 開催月日 競馬場コード 開催回[第N回] 開催日目[N日目] レース番号
                    End With ' Hatusyoriid
                    strSql = strSql & SS & "HatuSyori" & i + 1 & "SyussoTosu" & SE & "='" & Replace(.SyussoTosu, "'", "''") & "'," '' 出走頭数
                    strSql = strSql & SS & "HatuSyori" & i + 1 & "KettoNum" & SE & "='" & Replace(.KettoNum, "'", "''") & "'," '' 血統登録番号
                    strSql = strSql & SS & "HatuSyori" & i + 1 & "Bamei" & SE & "='" & Replace(.Bamei, "'", "''") & "'," '' 馬名
                End With ' HatuSyori
            Next i
            For i = 0 To 2
                With .SaikinJyusyo(i)
                    With .SaikinJyusyoid
                        strSql = strSql & SS & "SaikinJyusyo" & i + 1 & "SaikinJyusyoid" & SE & "='" & Replace(.Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum, "'", "''") & "',"
                        '' 開催年 開催月日 競馬場コード 開催回[第N回] 開催日目[N日目] レース番号
                    End With ' SaikinJyusyoid
                    strSql = strSql & SS & "SaikinJyusyo" & i + 1 & "Hondai" & SE & "='" & Replace(.Hondai, "'", "''") & "'," '' 競走名本題
                    strSql = strSql & SS & "SaikinJyusyo" & i + 1 & "Ryakusyo10" & SE & "='" & Replace(.Ryakusyo10, "'", "''") & "'," '' 競走名略称10字
                    strSql = strSql & SS & "SaikinJyusyo" & i + 1 & "Ryakusyo6" & SE & "='" & Replace(.Ryakusyo6, "'", "''") & "'," '' 競走名略称6字
                    strSql = strSql & SS & "SaikinJyusyo" & i + 1 & "Ryakusyo3" & SE & "='" & Replace(.Ryakusyo3, "'", "''") & "'," '' 競走名略称3字
                    strSql = strSql & SS & "SaikinJyusyo" & i + 1 & "GradeCD" & SE & "='" & Replace(.GradeCD, "'", "''") & "'," '' グレードコード
                    strSql = strSql & SS & "SaikinJyusyo" & i + 1 & "SyussoTosu" & SE & "='" & Replace(.SyussoTosu, "'", "''") & "'," '' 出走頭数
                    strSql = strSql & SS & "SaikinJyusyo" & i + 1 & "KettoNum" & SE & "='" & Replace(.KettoNum, "'", "''") & "'," '' 血統登録番号
                    strSql = strSql & SS & "SaikinJyusyo" & i + 1 & "Bamei" & SE & "='" & Replace(.Bamei, "'", "''") & "'," '' 馬名
                End With ' SaikinJyusyo
            Next i

            strSql = Left(strSql, Len(strSql) - 1) ' Delete Last Comma

            strSql = strSql & " WHERE " & SS & "KisyuCode" & SE & "='" & Replace(.KisyuCode, "'", "''") & "'"
            strSql = strSql & " AND " & SS & "MakeDate" & SE & " <= '" & Replace(strMakeDate, "'", "''") & "'"
        End With ' mBuf

        With mBuf
            System.Diagnostics.Debug.WriteLine("UPDATE KISYU : " & .KisyuCode)
        End With ' mBuf

        gCon.Execute(strSql)

        '' KISYU_SEISEKI (騎手マスタ_成績)
        For i = 0 To 2
            With mBuf
                strSql = "UPDATE KISYU_SEISEKI SET "
                strSql = strSql & SS & "KisyuCode" & SE & "='" & Replace(.KisyuCode, "'", "''") & "'," '' 調教師コード
                strSql = strSql & SS & "Num" & SE & "='" & Replace(CStr(i), "'", "''") & "'," '' 連番
                With .HonZenRuikei(i)
                    strSql = strSql & SS & "SetYear" & SE & "='" & Replace(.SetYear, "'", "''") & "'," '' 設定年
                    strSql = strSql & SS & "HonSyokinHeichi" & SE & "='" & Replace(.HonSyokinHeichi, "'", "''") & "'," '' 平地本賞金合計
                    strSql = strSql & SS & "HonSyokinSyogai" & SE & "='" & Replace(.HonSyokinSyogai, "'", "''") & "'," '' 障害本賞金合計
                    strSql = strSql & SS & "FukaSyokinHeichi" & SE & "='" & Replace(.FukaSyokinHeichi, "'", "''") & "'," '' 平地付加賞金合計
                    strSql = strSql & SS & "FukaSyokinSyogai" & SE & "='" & Replace(.FukaSyokinSyogai, "'", "''") & "'," '' 障害付加賞金合計
                    With .ChakuKaisuHeichi
                        For k = 0 To 5
                            strSql = strSql & SS & "HeichiChakukaisu" & k + 1 & SE & "='" & Replace(.Chakukaisu(k), "'", "''") & "',"
                        Next k
                    End With ' ChakuKaisuHeichi

                    With .ChakuKaisuSyogai
                        For k = 0 To 5
                            strSql = strSql & SS & "SyogaiChakukaisu" & k + 1 & SE & "='" & Replace(.Chakukaisu(k), "'", "''") & "',"
                        Next k
                    End With ' ChakuKaisuSyogai
                    For j = 0 To 5
                        With .ChakuKaisuKyori(j)
                            For k = 0 To 5
                                strSql = strSql & SS & "Kyori" & j + 1 & "Chakukaisu" & k + 1 & SE & "='" & Replace(.Chakukaisu(k), "'", "''") & "',"
                            Next k
                        End With ' ChakuKaisuKyori
                    Next j
                End With ' HonZenRuikei

                strSql = Left(strSql, Len(strSql) - 1) ' Delete Last Comma

                strSql = strSql & " WHERE " & SS & "KisyuCode" & SE & "='" & Replace(.KisyuCode, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "Num" & SE & "='" & Replace(CStr(i), "'", "''") & "'"
                strSql = strSql & " AND " & SS & "MakeDate" & SE & "<= '" & Replace(strMakeDate, "'", "''") & "'"

                gCon.Execute(strSql)

                ''一度に更新できるフィールド数が約127までの為 分割更新（JET仕様）
                strSql = "UPDATE KISYU_SEISEKI SET "
                With .head
                    With .MakeDate
                        strSql = strSql & SS & "MakeDate" & SE & "='" & Replace(strMakeDate, "'", "''") & "'," '' 年月日
                    End With ' MakeDate
                End With ' head
                With .HonZenRuikei(i)
                    For j = 0 To 19
                        With .ChakuKaisuJyo(j)
                            For k = 0 To 5
                                strSql = strSql & SS & "Jyo" & j + 1 & "Chakukaisu" & k + 1 & SE & "='" & Replace(.Chakukaisu(k), "'", "''") & "',"
                            Next k
                        End With ' ChakuKaisuJyo
                    Next j
                End With ' HonZenRuikei

                strSql = Left(strSql, Len(strSql) - 1) ' Delete Last Comma

                strSql = strSql & " WHERE " & SS & "KisyuCode" & SE & "='" & Replace(.KisyuCode, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "Num" & SE & "='" & Replace(CStr(i), "'", "''") & "'"
                strSql = strSql & " AND " & SS & "MakeDate" & SE & "<= '" & Replace(strMakeDate, "'", "''") & "'"
            End With ' mBuf

            With mBuf
                System.Diagnostics.Debug.WriteLine("UPDATE KISYU_SEISEKI : " & .KisyuCode & CStr(i))
            End With ' mBuf

            gCon.Execute(strSql)

        Next i

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