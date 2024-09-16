Option Strict Off
Option Explicit On
Option Compare Binary
Friend Class clsImportHR
	' @(h) clsReadHR.cls
	' @(s)
	' JVData "HR" データベース登録クラス
	'
	
	Private mBuf As JV_HR_PAY '' 払戻構造体
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
        strSql = "SELECT * FROM HARAI"
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
    ' 機能      : Closeのコーディング
    '
    ' 機能説明  : ガーベッジコレクションにCloseを呼ばせるとどこで呼ばれるか分からない為、
    '           　明示的に呼び出す必要がある。
    '
    Public Sub Close()
        'レコードセットクローズ
        mRS.Close()
        System.Diagnostics.Debug.WriteLine("mRS.Close")
        mRS = Nothing

    End Sub


    ' @(f)
    '
    ' 機能      : Addプロシージャを呼ぶ
    '
    ' 引き数    : lBuf - JVData 識別子"HR" の１行
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
        System.Diagnostics.Debug.WriteLine("mRs.AddNew")
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
            mRS.Fields("TorokuTosu").Value = .TorokuTosu '' 登録頭数
            mRS.Fields("SyussoTosu").Value = .SyussoTosu '' 出走頭数
            For i = 0 To 8
                mRS.Fields("FuseirituFlag" & i + 1).Value = .FuseirituFlag(i) '' 不成立フラグ
            Next i
            For i = 0 To 8
                mRS.Fields("TokubaraiFlag" & i + 1).Value = .TokubaraiFlag(i) '' 特払フラグ
            Next i
            For i = 0 To 8
                mRS.Fields("HenkanFlag" & i + 1).Value = .HenkanFlag(i) '' 返還フラグ
            Next i
            For i = 0 To 27
                mRS.Fields("HenkanUma" & i + 1).Value = .HenkanUma(i) '' 返還馬番情報(馬番01～28)
            Next i
            For i = 0 To 7
                mRS.Fields("HenkanWaku" & i + 1).Value = .HenkanWaku(i) '' 返還枠番情報(枠番1～8)
            Next i
            For i = 0 To 7
                mRS.Fields("HenkanDoWaku" & i + 1).Value = .HenkanDoWaku(i) '' 返還同枠情報(枠番1～8)
            Next i
            For i = 0 To 2
                With .PayTansyo(i)
                    mRS.Fields("PayTansyoUmaban" & i + 1).Value = .Umaban '' 馬番
                    mRS.Fields("PayTansyoPay" & i + 1).Value = .Pay '' 払戻金
                    mRS.Fields("PayTansyoNinki" & i + 1).Value = .Ninki '' 人気順
                End With ' PayTansyo
            Next i
            For i = 0 To 4
                With .PayFukusyo(i)
                    mRS.Fields("PayFukusyoUmaban" & i + 1).Value = .Umaban '' 馬番
                    mRS.Fields("PayFukusyoPay" & i + 1).Value = .Pay '' 払戻金
                    mRS.Fields("PayFukusyoNinki" & i + 1).Value = .Ninki '' 人気順
                End With ' PayFukusyo
            Next i
            For i = 0 To 2
                With .PayWakuren(i)
                    mRS.Fields("PayWakurenKumi" & i + 1).Value = .Umaban '' 組番
                    mRS.Fields("PayWakurenPay" & i + 1).Value = .Pay '' 払戻金
                    mRS.Fields("PayWakurenNinki" & i + 1).Value = .Ninki '' 人気順
                End With ' PayWakuren
            Next i
            For i = 0 To 2
                With .PayUmaren(i)
                    mRS.Fields("PayUmarenKumi" & i + 1).Value = .Kumi '' 組番
                    mRS.Fields("PayUmarenPay" & i + 1).Value = .Pay '' 払戻金
                    mRS.Fields("PayUmarenNinki" & i + 1).Value = .Ninki '' 人気順
                End With ' PayUmaren
            Next i
            For i = 0 To 6
                With .PayWide(i)
                    mRS.Fields("PayWideKumi" & i + 1).Value = .Kumi '' 組番
                    mRS.Fields("PayWidePay" & i + 1).Value = .Pay '' 払戻金
                    mRS.Fields("PayWideNinki" & i + 1).Value = .Ninki '' 人気順
                End With ' PayWide
            Next i
            For i = 0 To 2
                With .PayReserved1(i)
                    mRS.Fields("PayReserved1Kumi" & i + 1).Value = .Kumi '' 組番
                    mRS.Fields("PayReserved1Pay" & i + 1).Value = .Pay '' 払戻金
                    mRS.Fields("PayReserved1Ninki" & i + 1).Value = .Ninki '' 人気順
                End With ' PayReserved1
            Next i
            For i = 0 To 5
                With .PayUmatan(i)
                    mRS.Fields("PayUmatanKumi" & i + 1).Value = .Kumi '' 組番
                    mRS.Fields("PayUmatanPay" & i + 1).Value = .Pay '' 払戻金
                    mRS.Fields("PayUmatanNinki" & i + 1).Value = .Ninki '' 人気順
                End With ' PayUmatan
            Next i
            For i = 0 To 2
                With .PaySanrenpuku(i)
                    mRS.Fields("PaySanrenpukuKumi" & i + 1).Value = .Kumi '' 組番
                    mRS.Fields("PaySanrenpukuPay" & i + 1).Value = .Pay '' 払戻金
                    mRS.Fields("PaySanrenpukuNinki" & i + 1).Value = .Ninki '' 人気順
                End With ' PaySanrenpuku
            Next i
            For i = 0 To 5
                With .PaySanrentan(i)
                    mRS.Fields("PaySanrentanKumi" & i + 1).Value = .Kumi '' 組番
                    mRS.Fields("PaySanrentanPay" & i + 1).Value = .Pay '' 払戻金
                    mRS.Fields("PaySanrentanNinki" & i + 1).Value = .Ninki '' 人気順
                End With ' PayReserved2
            Next i
        End With

        With mBuf.id
            System.Diagnostics.Debug.WriteLine("Insert HARAI : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum)
        End With ' id

        mRS.Update()
        System.Diagnostics.Debug.WriteLine("mRS.update")

        gCon.CommitTrans()
        System.Diagnostics.Debug.WriteLine("CommitTrans")

        InsertDB = True

ExitHandler:
        Exit Function
ErrorHandler:
        InsertDB = False
        mRS.CancelUpdate()
        System.Diagnostics.Debug.WriteLine("mRS.CancelUpdate")
        gCon.RollbackTrans()
        System.Diagnostics.Debug.WriteLine(Err.Description)
        System.Diagnostics.Debug.WriteLine("Insert RollbackTrans")
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

        strSql = "UPDATE HARAI SET "
        With mBuf

            With .id
                strSql = strSql & SS & "Year" & SE & "='" & Replace(.Year, "'", "''") & "'," '' 開催年
                strSql = strSql & SS & "MonthDay" & SE & "='" & Replace(.MonthDay, "'", "''") & "'," '' 開催月日
                strSql = strSql & SS & "JyoCD" & SE & "='" & Replace(.JyoCD, "'", "''") & "'," '' 競馬場コード
                strSql = strSql & SS & "Kaiji" & SE & "='" & Replace(.Kaiji, "'", "''") & "'," '' 開催回第N回
                strSql = strSql & SS & "Nichiji" & SE & "='" & Replace(.Nichiji, "'", "''") & "'," '' 開催日目N日目
                strSql = strSql & SS & "RaceNum" & SE & "='" & Replace(.RaceNum, "'", "''") & "'," '' レース番号
            End With ' id
            strSql = strSql & SS & "TorokuTosu" & SE & "='" & Replace(.TorokuTosu, "'", "''") & "'," '' 登録頭数
            strSql = strSql & SS & "SyussoTosu" & SE & "='" & Replace(.SyussoTosu, "'", "''") & "'," '' 出走頭数
            For i = 0 To 8
                strSql = strSql & SS & "FuseirituFlag" & i + 1 & SE & "='" & Replace(.FuseirituFlag(i), "'", "''") & "'," '' 不成立フラグ
            Next i
            For i = 0 To 8
                strSql = strSql & SS & "TokubaraiFlag" & i + 1 & SE & "='" & Replace(.TokubaraiFlag(i), "'", "''") & "'," '' 特払フラグ
            Next i
            For i = 0 To 8
                strSql = strSql & SS & "HenkanFlag" & i + 1 & SE & "='" & Replace(.HenkanFlag(i), "'", "''") & "'," '' 返還フラグ
            Next i
            For i = 0 To 27
                strSql = strSql & SS & "HenkanUma" & i + 1 & SE & "='" & Replace(.HenkanUma(i), "'", "''") & "'," '' 返還馬番情報(馬番01～28)
            Next i
            For i = 0 To 7
                strSql = strSql & SS & "HenkanWaku" & i + 1 & SE & "='" & Replace(.HenkanWaku(i), "'", "''") & "'," '' 返還枠番情報(枠番1～8)
            Next i
            For i = 0 To 7
                strSql = strSql & SS & "HenkanDoWaku" & i + 1 & SE & "='" & Replace(.HenkanDoWaku(i), "'", "''") & "'," '' 返還同枠情報(枠番1～8)
            Next i
            For i = 0 To 2
                With .PayTansyo(i)
                    strSql = strSql & SS & "PayTansyoUmaban" & i + 1 & SE & "='" & Replace(.Umaban, "'", "''") & "'," '' 馬番
                    strSql = strSql & SS & "PayTansyoPay" & i + 1 & SE & "='" & Replace(.Pay, "'", "''") & "'," '' 払戻金
                    strSql = strSql & SS & "PayTansyoNinki" & i + 1 & SE & "='" & Replace(.Ninki, "'", "''") & "'," '' 人気順
                End With ' PayTansyo
            Next i
            For i = 0 To 4
                With .PayFukusyo(i)
                    strSql = strSql & SS & "PayFukusyoUmaban" & i + 1 & SE & "='" & Replace(.Umaban, "'", "''") & "'," '' 馬番
                    strSql = strSql & SS & "PayFukusyoPay" & i + 1 & SE & "='" & Replace(.Pay, "'", "''") & "'," '' 払戻金
                    strSql = strSql & SS & "PayFukusyoNinki" & i + 1 & SE & "='" & Replace(.Ninki, "'", "''") & "'," '' 人気順
                End With ' PayFukusyo
            Next i

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

            gCon.Execute(strSql)

            ''一度に更新できるフィールド数が約127までの為 分割更新（JET仕様）
            strSql = "UPDATE HARAI SET "
            With .head

                strSql = strSql & SS & "RecordSpec" & SE & "='" & Replace(.RecordSpec, "'", "''") & "'," '' レコード種別

                strSql = strSql & SS & "DataKubun" & SE & "='" & Replace(.DataKubun, "'", "''") & "'," '' データ区分
                strSql = strSql & SS & "MakeDate" & SE & "= '" & Replace(strMakeDate, "'", "''") & "'," '' 年月日
            End With ' head
            For i = 0 To 2
                With .PayWakuren(i)
                    strSql = strSql & SS & "PayWakurenKumi" & i + 1 & SE & "='" & Replace(.Umaban, "'", "''") & "'," '' 馬番
                    strSql = strSql & SS & "PayWakurenPay" & i + 1 & SE & "='" & Replace(.Pay, "'", "''") & "'," '' 払戻金
                    strSql = strSql & SS & "PayWakurenNinki" & i + 1 & SE & "='" & Replace(.Ninki, "'", "''") & "'," '' 人気順
                End With ' PayWakuren
            Next i
            For i = 0 To 2
                With .PayUmaren(i)
                    strSql = strSql & SS & "PayUmarenKumi" & i + 1 & SE & "='" & Replace(.Kumi, "'", "''") & "'," '' 組番
                    strSql = strSql & SS & "PayUmarenPay" & i + 1 & SE & "='" & Replace(.Pay, "'", "''") & "'," '' 払戻金
                    strSql = strSql & SS & "PayUmarenNinki" & i + 1 & SE & "='" & Replace(.Ninki, "'", "''") & "'," '' 人気順
                End With ' PayUmaren
            Next i
            For i = 0 To 6
                With .PayWide(i)
                    strSql = strSql & SS & "PayWideKumi" & i + 1 & SE & "='" & Replace(.Kumi, "'", "''") & "'," '' 組番
                    strSql = strSql & SS & "PayWidePay" & i + 1 & SE & "='" & Replace(.Pay, "'", "''") & "'," '' 払戻金
                    strSql = strSql & SS & "PayWideNinki" & i + 1 & SE & "='" & Replace(.Ninki, "'", "''") & "'," '' 人気順
                End With ' PayWide
            Next i
            For i = 0 To 2
                With .PayReserved1(i)
                    strSql = strSql & SS & "PayReserved1Kumi" & i + 1 & SE & "='" & Replace(.Kumi, "'", "''") & "'," '' 組番
                    strSql = strSql & SS & "PayReserved1Pay" & i + 1 & SE & "='" & Replace(.Pay, "'", "''") & "'," '' 払戻金
                    strSql = strSql & SS & "PayReserved1Ninki" & i + 1 & SE & "='" & Replace(.Ninki, "'", "''") & "'," '' 人気順
                End With ' PayReserved1
            Next i
            For i = 0 To 5
                With .PayUmatan(i)
                    strSql = strSql & SS & "PayUmatanKumi" & i + 1 & SE & "='" & Replace(.Kumi, "'", "''") & "'," '' 組番
                    strSql = strSql & SS & "PayUmatanPay" & i + 1 & SE & "='" & Replace(.Pay, "'", "''") & "'," '' 払戻金
                    strSql = strSql & SS & "PayUmatanNinki" & i + 1 & SE & "='" & Replace(.Ninki, "'", "''") & "'," '' 人気順
                End With ' PayUmatan
            Next i
            For i = 0 To 2
                With .PaySanrenpuku(i)
                    strSql = strSql & SS & "PaySanrenpukuKumi" & i + 1 & SE & "='" & Replace(.Kumi, "'", "''") & "'," '' 組番
                    strSql = strSql & SS & "PaySanrenpukuPay" & i + 1 & SE & "='" & Replace(.Pay, "'", "''") & "'," '' 払戻金
                    strSql = strSql & SS & "PaySanrenpukuNinki" & i + 1 & SE & "='" & Replace(.Ninki, "'", "''") & "'," '' 人気順
                End With ' PaySanrenpuku
            Next i
            For i = 0 To 5
                With .PaySanrentan(i)
                    strSql = strSql & SS & "PaySanrentanKumi" & i + 1 & SE & "='" & Replace(.Kumi, "'", "''") & "'," '' 組番
                    strSql = strSql & SS & "PaySanrentanPay" & i + 1 & SE & "='" & Replace(.Pay, "'", "''") & "'," '' 払戻金
                    strSql = strSql & SS & "PaySanrentanNinki" & i + 1 & SE & "='" & Replace(.Ninki, "'", "''") & "'," '' 人気順
                End With ' PayReserved2
            Next i

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
            System.Diagnostics.Debug.WriteLine("UPDATE HARAI : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum)
        End With ' id

        gCon.Execute(strSql)

        System.Diagnostics.Debug.WriteLine("CommitTrans")
        gCon.CommitTrans()

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