Option Strict Off
Option Explicit On
Option Compare Binary
Friend Class clsImportTK
	' @(h) clsReadTK
	' @(s)
	' JVData "TK" データベース登録クラス
	'
	
	Private mBuf As JV_TK_TOKUUMA '' 特別登録馬構造体
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
        strSql = "SELECT * FROM TOKU_RACE"
        mRS1 = New ADODB.Recordset()
        mRS1.Open(strSql, gCon, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockPessimistic)

        'レコードセットオープン
        strSql = "SELECT * FROM TOKU"
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
        'レコードセットクローズ
        mRS1.Close()
        mRS2.Close()

        mRS1 = Nothing
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
    ' 引き数    : lBuf - JVData 識別子"TK" の１行
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

        ' ヘッダ部分
        With mBuf
            With .head
                mRS1.Fields("RecordSpec").Value = .RecordSpec '' レコード種別
                mRS1.Fields("DataKubun").Value = .DataKubun '' データ区分
                With .MakeDate
                    mRS1.Fields("MakeDate").Value = .Year & .Month & .Day '' 年月日
                End With ' MakeDate
            End With ' head
            With .id
                mRS1.Fields("Year").Value = .Year '' 開催年
                mRS1.Fields("MonthDay").Value = .MonthDay '' 開催月日
                mRS1.Fields("JyoCD").Value = .JyoCD '' 競馬場コード
                mRS1.Fields("Kaiji").Value = .Kaiji '' 開催回第N回
                mRS1.Fields("Nichiji").Value = .Nichiji '' 開催日目N日目
                mRS1.Fields("RaceNum").Value = .RaceNum '' レース番号
            End With ' id
            With .RaceInfo
                mRS1.Fields("YoubiCD").Value = .YoubiCD '' 曜日コード
                mRS1.Fields("TokuNum").Value = .TokuNum '' 特別競走番号
                mRS1.Fields("Hondai").Value = .Hondai '' 競走名本題
                mRS1.Fields("Fukudai").Value = .Fukudai '' 競走名副題
                mRS1.Fields("Kakko").Value = .Kakko '' 競走名カッコ内
                mRS1.Fields("HondaiEng").Value = .HondaiEng '' 競走名本題欧字
                mRS1.Fields("FukudaiEng").Value = .FukudaiEng '' 競走名副題欧字
                mRS1.Fields("KakkoEng").Value = .KakkoEng '' 競走名カッコ内欧字
                mRS1.Fields("Ryakusyo10").Value = .Ryakusyo10 '' 競走名略称１０字
                mRS1.Fields("Ryakusyo6").Value = .Ryakusyo6 '' 競走名略称６字
                mRS1.Fields("Ryakusyo3").Value = .Ryakusyo3 '' 競走名略称３字
                mRS1.Fields("Kubun").Value = .Kubun '' 競走名区分
                mRS1.Fields("Nkai").Value = .Nkai '' 重賞回次第N回
            End With ' RaceInfo
            mRS1.Fields("GradeCD").Value = .GradeCD '' グレードコード
            With .JyokenInfo
                mRS1.Fields("SyubetuCD").Value = .SyubetuCD '' 競走種別コード
                mRS1.Fields("KigoCD").Value = .KigoCD '' 競走記号コード
                mRS1.Fields("JyuryoCD").Value = .JyuryoCD '' 重量種別コード
                For j = 0 To 4
                    mRS1.Fields("JyokenCD" & j + 1).Value = .JyokenCD(j) '' 競走条件コード
                Next j
            End With ' JyokenInfo
            mRS1.Fields("Kyori").Value = .Kyori '' 距離
            mRS1.Fields("TrackCD").Value = .TrackCD '' トラックコード
            mRS1.Fields("CourseKubunCD").Value = .CourseKubunCD '' コース区分
            With .HandiDate
                mRS1.Fields("HandiDate").Value = .Year & .Month & .Day '' 年月日
            End With ' HandiDate
            mRS1.Fields("TorokuTosu").Value = .TorokuTosu '' 登録頭数
        End With

        With mBuf.id
            System.Diagnostics.Debug.WriteLine("Insert TOKU_RACE : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum)
        End With ' id

        mRS1.Update()


        ' 馬毎部分
        For i = 0 To CDbl(mBuf.TorokuTosu) - 1
            mRS2.AddNew()
            With mBuf
                With .head
                    With .MakeDate
                        mRS2.Fields("MakeDate").Value = .Year & .Month & .Day '' 年月日
                    End With ' MakeDate
                End With
                With .id
                    mRS2.Fields("Year").Value = .Year '' 開催年
                    mRS2.Fields("MonthDay").Value = .MonthDay '' 開催月日
                    mRS2.Fields("JyoCD").Value = .JyoCD '' 競馬場コード
                    mRS2.Fields("Kaiji").Value = .Kaiji '' 開催回第N回
                    mRS2.Fields("Nichiji").Value = .Nichiji '' 開催日目N日目
                    mRS2.Fields("RaceNum").Value = .RaceNum '' レース番号
                End With ' id
                With .TokuUmaInfo(i)
                    mRS2.Fields("Num").Value = .Num '' 連番
                    mRS2.Fields("KettoNum").Value = .KettoNum '' 血統登録番号
                    mRS2.Fields("Bamei").Value = .Bamei '' 馬名
                    mRS2.Fields("UmaKigoCD").Value = .UmaKigoCD '' 馬記号コード
                    mRS2.Fields("SexCD").Value = .SexCD '' 性別コード
                    mRS2.Fields("TozaiCD").Value = .TozaiCD '' 調教師東西所属コード
                    mRS2.Fields("ChokyosiCode").Value = .ChokyosiCode '' 調教師コード
                    mRS2.Fields("ChokyosiRyakusyo").Value = .ChokyosiRyakusyo '' 調教師名略称
                    mRS2.Fields("Futan").Value = .Futan '' 負担重量
                    mRS2.Fields("Koryu").Value = .Koryu '' 交流区分
                End With ' TokuUmaInfo
            End With

            With mBuf.id
                System.Diagnostics.Debug.WriteLine("Insert TOKU : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum & mBuf.TokuUmaInfo(i).Num)
            End With ' id
            mRS2.Update()
        Next i

        gCon.CommitTrans()
        System.Diagnostics.Debug.WriteLine("CommitTrans")

        InsertDB = True

ExitHandler:
        Exit Function
ErrorHandler:
        InsertDB = False
        mRS1.CancelUpdate()
        mRS2.CancelUpdate()
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

        ' ヘッダ部分
        strSql = "UPDATE TOKU_RACE SET "
        With mBuf
            With .head
                strSql = strSql & SS & "RecordSpec" & SE & " = '" & Replace(.RecordSpec, "'", "''") & "'," '' レコード種別
                strSql = strSql & SS & "DataKubun" & SE & " = '" & Replace(.DataKubun, "'", "''") & "'," '' データ区分
                With .MakeDate
                    strSql = strSql & SS & "MakeDate" & SE & " = '" & Replace(.Year & .Month & .Day, "'", "''") & "'," '' 年
                End With ' MakeDate
            End With ' head
            With .RaceInfo
                strSql = strSql & SS & "YoubiCD" & SE & " = '" & Replace(.YoubiCD, "'", "''") & "'," '' 曜日コード
                strSql = strSql & SS & "TokuNum" & SE & " = '" & Replace(.TokuNum, "'", "''") & "'," '' 特別競走番号
                strSql = strSql & SS & "Hondai" & SE & " = '" & Replace(.Hondai, "'", "''") & "'," '' 競走名本題
                strSql = strSql & SS & "Fukudai" & SE & " = '" & Replace(.Fukudai, "'", "''") & "'," '' 競走名副題
                strSql = strSql & SS & "Kakko" & SE & " = '" & Replace(.Kakko, "'", "''") & "'," '' 競走名カッコ内
                strSql = strSql & SS & "HondaiEng" & SE & " = '" & Replace(.HondaiEng, "'", "''") & "'," '' 競走名本題欧字
                strSql = strSql & SS & "FukudaiEng" & SE & " = '" & Replace(.FukudaiEng, "'", "''") & "'," '' 競走名副題欧字
                strSql = strSql & SS & "KakkoEng" & SE & " = '" & Replace(.KakkoEng, "'", "''") & "'," '' 競走名カッコ内欧字
                strSql = strSql & SS & "Ryakusyo10" & SE & " = '" & Replace(.Ryakusyo10, "'", "''") & "'," '' 競走名略称１０字
                strSql = strSql & SS & "Ryakusyo6" & SE & " = '" & Replace(.Ryakusyo6, "'", "''") & "'," '' 競走名略称６字
                strSql = strSql & SS & "Ryakusyo3" & SE & " = '" & Replace(.Ryakusyo3, "'", "''") & "'," '' 競走名略称３字
                strSql = strSql & SS & "Kubun" & SE & " = '" & Replace(.Kubun, "'", "''") & "'," '' 競走名区分
                strSql = strSql & SS & "Nkai" & SE & " = '" & Replace(.Nkai, "'", "''") & "'," '' 重賞回次第N回
            End With ' RaceInfo
            strSql = strSql & SS & "GradeCD" & SE & " = '" & Replace(.GradeCD, "'", "''") & "'," '' グレードコード
            With .JyokenInfo
                strSql = strSql & SS & "SyubetuCD" & SE & " = '" & Replace(.SyubetuCD, "'", "''") & "'," '' 競走種別コード
                strSql = strSql & SS & "KigoCD" & SE & " = '" & Replace(.KigoCD, "'", "''") & "'," '' 競走記号コード
                strSql = strSql & SS & "JyuryoCD" & SE & " = '" & Replace(.JyuryoCD, "'", "''") & "'," '' 重量種別コード
                For j = 0 To 4
                    strSql = strSql & SS & "JyokenCD" & j + 1 & SE & " = '" & Replace(.JyokenCD(j), "'", "''") & "'," '' 競走条件コード
                Next j
            End With ' JyokenInfo
            strSql = strSql & SS & "Kyori" & SE & " = '" & Replace(.Kyori, "'", "''") & "'," '' 距離
            strSql = strSql & SS & "TrackCD" & SE & " = '" & Replace(.TrackCD, "'", "''") & "'," '' トラックコード
            strSql = strSql & SS & "CourseKubunCD" & SE & " = '" & Replace(.CourseKubunCD, "'", "''") & "'," '' コース区分
            With .HandiDate
                strSql = strSql & SS & "HandiDate" & SE & " = '" & Replace(.Year & .Month & .Day, "'", "''") & "'," '' HandiDate
            End With ' HandiDate
            strSql = strSql & SS & "TorokuTosu" & SE & " = '" & Replace(.TorokuTosu, "'", "''") & "'," '' 登録頭数
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
            System.Diagnostics.Debug.WriteLine("UPDATE TOKU_RACE : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum)
        End With ' id

        Dim tmpRS As ADODB.Recordset
        tmpRS = gCon.Execute(strSql)
        If tmpRS Is Nothing Then
        Else

            ' 馬毎部分
            ' 馬毎情報はレコード数が可変するため、既存のデータを全て削除してから登録しなおします。

            With mBuf.id
                strSql = "DELETE FROM TOKU"
                strSql = strSql & " WHERE " & SS & "Year" & SE & "='" & Replace(.Year, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "MonthDay" & SE & "='" & Replace(.MonthDay, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "JyoCD" & SE & "='" & Replace(.JyoCD, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "Kaiji" & SE & "='" & Replace(.Kaiji, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "Nichiji" & SE & "='" & Replace(.Nichiji, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "RaceNum" & SE & "='" & Replace(.RaceNum, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "MakeDate" & SE & " <= '" & Replace(strMakeDate, "'", "''") & "'"

                With mBuf.id
                    System.Diagnostics.Debug.WriteLine("DELETE TOKU : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum)
                End With ' id

                gCon.Execute(strSql)
            End With

            ' 全部再登録する
            For i = 0 To CDbl(mBuf.TorokuTosu) - 1
                strSql = "INSERT INTO TOKU VALUES( "
                With mBuf
                    With .head
                        strSql = strSql & "'" & Replace(strMakeDate, "'", "''") & "'," ''提供年月日
                    End With
                    With .id
                        strSql = strSql & "'" & Replace(.Year, "'", "''") & "'," '' 開催年
                        strSql = strSql & "'" & Replace(.MonthDay, "'", "''") & "'," '' 開催月日
                        strSql = strSql & "'" & Replace(.JyoCD, "'", "''") & "'," '' 競馬場コード
                        strSql = strSql & "'" & Replace(.Kaiji, "'", "''") & "'," '' 開催回第N回
                        strSql = strSql & "'" & Replace(.Nichiji, "'", "''") & "'," '' 開催日目N日目
                        strSql = strSql & "'" & Replace(.RaceNum, "'", "''") & "'," '' レース番号
                    End With ' id
                    With .TokuUmaInfo(i)
                        strSql = strSql & "'" & Replace(.Num, "'", "''") & "'," '' 連番
                        strSql = strSql & "'" & Replace(.KettoNum, "'", "''") & "'," '' 血統登録番号
                        strSql = strSql & "'" & Replace(.Bamei, "'", "''") & "'," '' 馬名
                        strSql = strSql & "'" & Replace(.UmaKigoCD, "'", "''") & "'," '' 馬記号コード
                        strSql = strSql & "'" & Replace(.SexCD, "'", "''") & "'," '' 性別コード
                        strSql = strSql & "'" & Replace(.TozaiCD, "'", "''") & "'," '' 調教師東西所属コード
                        strSql = strSql & "'" & Replace(.ChokyosiCode, "'", "''") & "'," '' 調教師コード
                        strSql = strSql & "'" & Replace(.ChokyosiRyakusyo, "'", "''") & "'," '' 調教師名略称
                        strSql = strSql & "'" & Replace(.Futan, "'", "''") & "'," '' 負担重量
                        strSql = strSql & "'" & Replace(.Koryu, "'", "''") & "'," '' 交流区分
                    End With ' TokuUmaInfo
                End With

                strSql = Left(strSql, Len(strSql) - 1) ' Delete Last Comma
                strSql = strSql & ")"

                With mBuf.id
                    System.Diagnostics.Debug.WriteLine("Insert TOKU : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum & mBuf.TokuUmaInfo(i).Num)
                End With ' id
                gCon.Execute(strSql)
            Next i
            tmpRS = Nothing
        End If

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