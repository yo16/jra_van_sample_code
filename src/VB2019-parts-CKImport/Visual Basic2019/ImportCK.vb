Imports System.Collections.Generic
Imports System.Diagnostics

Public Class ImportCK
    Inherits ImportBase

    Public Sub New()
        MyBase.New()

        Dim SQL As New List(Of String)

        mBuf = CType(mBuf, JV_CK_CHAKU)

        '' 実行するSQLをListに格納
        SQL.Add("SELECT * FROM CHAKU")

        '' レコードセットOPEN処理実行
        Class_Initialize_Renamed(SQL)
    End Sub

    Protected Overrides Function InsertDB() As Boolean
        Dim i As Short '' ループカウンタ
        Dim j As Short '' ループカウンタ
        Dim k As Short '' ループカウンタ
        Dim temp As String
        Dim s As String

        Try
            gCon.BeginTrans()
            Debug.WriteLine("BeginTrans")

            mRS(0).AddNew()

            With mBuf
                With .head
                    mRS(0).Fields("RecordSpec").Value = .RecordSpec '' レコード種別
                    mRS(0).Fields("DataKubun").Value = .DataKubun '' データ区分
                    With .MakeDate
                        mRS(0).Fields("MakeDate").Value = .Year & .Month & .Day '' 年月日
                    End With ' MakeDate
                End With ' head

                With .id
                    mRS(0).Fields("Year").Value = .Year '' 開催年
                    mRS(0).Fields("MonthDay").Value = .MonthDay '' 開催月日
                    mRS(0).Fields("JyoCD").Value = .JyoCD '' 競馬場コード
                    mRS(0).Fields("Kaiji").Value = .Kaiji '' 開催回第N回
                    mRS(0).Fields("Nichiji").Value = .Nichiji '' 開催日目N日目
                    mRS(0).Fields("RaceNum").Value = .RaceNum '' レース番号
                End With ' id

                With .UmaChaku
                    mRS(0).Fields("KettoNum").Value = .KettoNum '' 血統登録番号
                    mRS(0).Fields("Bamei").Value = .Bamei '' 馬名
                    mRS(0).Fields("RuikeiHonsyoHeiti").Value = .RuikeiHonsyoHeiti '' 平地本賞金累計
                    mRS(0).Fields("RuikeiHonsyoSyogai").Value = .RuikeiHonsyoSyogai '' 障害本賞金累計
                    mRS(0).Fields("RuikeiFukaHeichi").Value = .RuikeiFukaHeichi '' 平地付加賞金累計
                    mRS(0).Fields("RuikeiFukaSyogai").Value = .RuikeiFukaSyogai '' 障害付加賞金累計
                    mRS(0).Fields("RuikeiSyutokuHeichi").Value = .RuikeiSyutokuHeichi '' 平地収得賞金累計
                    mRS(0).Fields("RuikeiSyutokuSyogai").Value = .RuikeiSyutokuSyogai '' 障害収得賞金累計

                    temp = ""
                    '' 総合着回数
                    With .ChakuSogo
                        For j = 0 To 5
                            temp &= .Chakukaisu(j)
                        Next j
                    End With
                    '' 中央合計着回数
                    With .ChakuChuo
                        For j = 0 To 5
                            temp &= .Chakukaisu(j)
                        Next j
                    End With
                    '' 馬場別着回数
                    For j = 0 To 6
                        With .ChakuKaisuBa(j)
                            For k = 0 To 5
                                temp &= .Chakukaisu(k)
                            Next k
                        End With
                    Next j
                    '' 馬場状態別着回数
                    For j = 0 To 11
                        With .ChakuKaisuJyotai(j)
                            For k = 0 To 5
                                temp &= .Chakukaisu(k)
                            Next k
                        End With
                    Next j
                    '' 距離別着回数(芝)
                    For j = 0 To 8
                        With .ChakuKaisuSibaKyori(j)
                            For k = 0 To 5
                                temp &= .Chakukaisu(k)
                            Next k
                        End With
                    Next j
                    '' 距離別着回数(ダート)
                    For j = 0 To 8
                        With .ChakuKaisuDirtKyori(j)
                            For k = 0 To 5
                                temp &= .Chakukaisu(k)
                            Next k
                        End With
                    Next j
                    '' 競馬場別着回数(芝)
                    For j = 0 To 9
                        With .ChakuKaisuJyoSiba(j)
                            For k = 0 To 5
                                temp &= .Chakukaisu(k)
                            Next k
                        End With
                    Next j
                    '' 競馬場別着回数(ダート)
                    For j = 0 To 9
                        With .ChakuKaisuJyoDirt(j)
                            For k = 0 To 5
                                temp &= .Chakukaisu(k)
                            Next k
                        End With
                    Next j
                    '' 競馬場別着回数(障害)
                    For j = 0 To 9
                        With .ChakuKaisuJyoSyogai(j)
                            For k = 0 To 5
                                temp &= .Chakukaisu(k)
                            Next k
                        End With
                    Next j
                    mRS(0).Fields("Chakukaisu").Value = temp

                    '' 脚質傾向
                    temp = ""
                    For j = 0 To 3
                        temp &= .Kyakusitu(j)
                    Next j
                    mRS(0).Fields("Kyakusitu").Value = temp
                    mRS(0).Fields("RaceCount").Value = .RaceCount '' 登録レース数
                End With

                With .KisyuChaku
                    mRS(0).Fields("KisyuCode").Value = .KisyuCode '' 騎手コード
                    mRS(0).Fields("KisyuName").Value = .KisyuName '' 騎手名

                    '' 騎手本年･累計成績情報
                    For i = 0 To 1
                        With .HonRuikei(i)
                            If i = 0 Then
                                s = "H"
                            Else
                                s = "R"
                            End If

                            mRS(0).Fields("K_" & s & "_SetYear").Value = .SetYear '' 設定年
                            mRS(0).Fields("K_" & s & "_HonSyokinHeichi").Value = .HonSyokinHeichi '' 平地本賞金合計
                            mRS(0).Fields("K_" & s & "_HonSyokinSyogai").Value = .HonSyokinSyogai '' 障害本賞金合計
                            mRS(0).Fields("K_" & s & "_FukaSyokinHeichi").Value = .FukaSyokinHeichi '' 平地付加賞金合計
                            mRS(0).Fields("K_" & s & "_FukaSyokinSyogai").Value = .FukaSyokinSyogai '' 障害付加賞金合計
                            temp = ""
                            '' 芝着回数
                            With .ChakuKaisuSiba
                                For k = 0 To 5
                                    temp &= .Chakukaisu(k)
                                Next
                            End With
                            '' ダート着回数
                            With .ChakuKaisuDirt
                                For k = 0 To 5
                                    temp &= .Chakukaisu(k)
                                Next
                            End With
                            '' 障害着回数
                            With .ChakuKaisuSyogai
                                For k = 0 To 5
                                    temp &= .Chakukaisu(k)
                                Next
                            End With
                            '' 距離別着回数(芝)
                            For j = 0 To 8
                                With .ChakuKaisuSibaKyori(j)
                                    For k = 0 To 5
                                        temp &= .Chakukaisu(k)
                                    Next
                                End With
                            Next
                            '' 距離別着回数(ダート)
                            For j = 0 To 8
                                With .ChakuKaisuDirtKyori(j)
                                    For k = 0 To 5
                                        temp &= .Chakukaisu(k)
                                    Next
                                End With
                            Next
                            '' 競馬場別着回数(芝)
                            For j = 0 To 9
                                With .ChakuKaisuJyoSiba(j)
                                    For k = 0 To 5
                                        temp &= .Chakukaisu(k)
                                    Next
                                End With
                            Next
                            '' 競馬場別着回数(ダート)
                            For j = 0 To 9
                                With .ChakuKaisuJyoDirt(j)
                                    For k = 0 To 5
                                        temp &= .Chakukaisu(k)
                                    Next
                                End With
                            Next
                            '' 競馬場別着回数(障害)
                            For j = 0 To 9
                                With .ChakuKaisuJyoSyogai(j)
                                    For k = 0 To 5
                                        temp &= .Chakukaisu(k)
                                    Next
                                End With
                            Next

                            mRS(0).Fields("K_" & s & "_Chakukaisu").Value = temp
                        End With
                    Next
                End With

                With .ChokyoChaku
                    mRS(0).Fields("ChokyosiCode").Value = .ChokyosiCode '' 調教師コード
                    mRS(0).Fields("ChokyosiName").Value = .ChokyosiName '' 調教師名

                    '' 調教師本年･累計成績情報
                    For i = 0 To 1
                        With .HonRuikei(i)
                            If i = 0 Then
                                s = "H"
                            Else
                                s = "R"
                            End If

                            mRS(0).Fields("C_" & s & "_SetYear").Value = .SetYear '' 設定年
                            mRS(0).Fields("C_" & s & "_HonSyokinHeichi").Value = .HonSyokinHeichi '' 平地本賞金合計
                            mRS(0).Fields("C_" & s & "_HonSyokinSyogai").Value = .HonSyokinSyogai '' 障害本賞金合計
                            mRS(0).Fields("C_" & s & "_FukaSyokinHeichi").Value = .FukaSyokinHeichi '' 平地付加賞金合計
                            mRS(0).Fields("C_" & s & "_FukaSyokinSyogai").Value = .FukaSyokinSyogai '' 障害付加賞金合計
                            temp = ""
                            '' 芝着回数
                            With .ChakuKaisuSiba
                                For k = 0 To 5
                                    temp &= .Chakukaisu(k)
                                Next
                            End With
                            '' ダート着回数
                            With .ChakuKaisuDirt
                                For k = 0 To 5
                                    temp &= .Chakukaisu(k)
                                Next
                            End With
                            '' 障害着回数
                            With .ChakuKaisuSyogai
                                For k = 0 To 5
                                    temp &= .Chakukaisu(k)
                                Next
                            End With
                            '' 距離別着回数(芝)
                            For j = 0 To 8
                                With .ChakuKaisuSibaKyori(j)
                                    For k = 0 To 5
                                        temp &= .Chakukaisu(k)
                                    Next
                                End With
                            Next
                            '' 距離別着回数(ダート)
                            For j = 0 To 8
                                With .ChakuKaisuDirtKyori(j)
                                    For k = 0 To 5
                                        temp &= .Chakukaisu(k)
                                    Next
                                End With
                            Next
                            '' 競馬場別着回数(芝)
                            For j = 0 To 9
                                With .ChakuKaisuJyoSiba(j)
                                    For k = 0 To 5
                                        temp &= .Chakukaisu(k)
                                    Next
                                End With
                            Next
                            '' 競馬場別着回数(ダート)
                            For j = 0 To 9
                                With .ChakuKaisuJyoDirt(j)
                                    For k = 0 To 5
                                        temp &= .Chakukaisu(k)
                                    Next
                                End With
                            Next
                            '' 競馬場別着回数(障害)
                            For j = 0 To 9
                                With .ChakuKaisuJyoSyogai(j)
                                    For k = 0 To 5
                                        temp &= .Chakukaisu(k)
                                    Next
                                End With
                            Next
                            mRS(0).Fields("C_" & s & "_Chakukaisu").Value = temp
                        End With
                    Next
                End With
                
                With .BanusiChaku
                    mRS(0).Fields("BanusiCode").Value = .BanusiCode '' 馬主コード
                    mRS(0).Fields("BanusiName_Co").Value = .BanusiName_Co '' 馬主名(法人格有)
                    mRS(0).Fields("BanusiName").Value = .BanusiName '' 馬主名(法人格無)

                    '' 馬主本年･累計成績情報
                    For i = 0 To 1
                        With .HonRuikei(i)
                            If i = 0 Then
                                s = "H"
                            Else
                                s = "R"
                            End If

                            mRS(0).Fields("Ba_" & s & "_SetYear").Value = .SetYear '' 設定年
                            mRS(0).Fields("Ba_" & s & "_HonSyokin").Value = .HonSyokinTotal '' 本賞金合計
                            mRS(0).Fields("Ba_" & s & "_FukaSyokin").Value = .FukaSyokin '' 付加賞金合計
                            '' 着回数
                            temp = ""
                            For j = 0 To 5
                                temp &= .ChakuKaisu(j)
                            Next j
                            mRS(0).Fields("Ba_" & s & "_Chakukaisu").Value = temp
                        End With
                    Next i
                End With

                With .BreederChaku
                    mRS(0).Fields("BreederCode").Value = .BreederCode '' 生産者コード
                    mRS(0).Fields("BreederName_Co").Value = .BreederName '' 生産者名(法人格有)
                    mRS(0).Fields("BreederName").Value = .BreederName '' 生産者名(法人格無)

                    '' 生産者本年･累計成績情報
                    For i = 0 To 1
                        With .HonRuikei(i)
                            If i = 0 Then
                                s = "H"
                            Else
                                s = "R"
                            End If

                            mRS(0).Fields("Br_" & s & "_SetYear").Value = .SetYear '' 設定年
                            mRS(0).Fields("Br_" & s & "_HonSyokin").Value = .HonSyokinTotal '' 本賞金合計
                            mRS(0).Fields("Br_" & s & "_FukaSyokin").Value = .FukaSyokin '' 付加賞金合計
                            '' 着回数
                            temp = ""
                            For j = 0 To 5
                                temp &= .ChakuKaisu(j)
                            Next j
                            mRS(0).Fields("Br_" & s & "_Chakukaisu").Value = temp
                        End With
                    Next i
                End With

            End With

            mRS(0).Update()

            gCon.CommitTrans()
            Debug.WriteLine("CommitTrans")

            Return True
        Catch ex As Exception
            mRS(0).CancelUpdate()
            gCon.RollbackTrans()
            Debug.WriteLine("RollbackTrans")
            Return False
        End Try
    End Function

    Protected Overrides Function UpdateDB(ByRef strMakeDate As String) As Boolean
        Dim i As Short '' ループカウンタ
        Dim j As Short '' ループカウンタ
        Dim k As Short '' ループカウンタ
        Dim SQL As String '' SQL文
        Dim temp As String
        Dim s As String

        Try
            gCon.BeginTrans()
            System.Diagnostics.Debug.WriteLine("BeginTrans")

            SQL = "UPDATE CHAKU SET "
            With mBuf
                With .head
                    SQL = SQL & "[RecordSpec]='" & Replace(.RecordSpec, "'", "''") & "'," '' レコード種別
                    SQL = SQL & "[DataKubun]='" & Replace(.DataKubun, "'", "''") & "'," '' データ区分
                    SQL = SQL & "[MakeDate]= '" & Replace(strMakeDate, "'", "''") & "'," '' 年月日
                End With ' head
                With .id
                    SQL = SQL & "[Year]='" & Replace(.Year, "'", "''") & "'," '' 開催年
                    SQL = SQL & "[MonthDay]='" & Replace(.MonthDay, "'", "''") & "'," '' 開催月日
                    SQL = SQL & "[JyoCD]='" & Replace(.JyoCD, "'", "''") & "'," '' 競馬場コード
                    SQL = SQL & "[Kaiji]='" & Replace(.Kaiji, "'", "''") & "'," '' 開催回第N回
                    SQL = SQL & "[Nichiji]='" & Replace(.Nichiji, "'", "''") & "'," '' 開催日目N日目
                    SQL = SQL & "[RaceNum]='" & Replace(.RaceNum, "'", "''") & "'," '' レース番号
                End With ' id

                With .UmaChaku
                    SQL = SQL & "[KettoNum]='" & Replace(.KettoNum, "'", "''") & "'," '' 血統登録番号
                    SQL = SQL & "[Bamei]='" & Replace(.Bamei, "'", "''") & "'," '' 馬名
                    SQL = SQL & "[RuikeiHonsyoHeiti]='" & Replace(.RuikeiHonsyoHeiti, "'", "''") & "'," '' 平地本賞金累計
                    SQL = SQL & "[RuikeiHonsyoSyogai]='" & Replace(.RuikeiHonsyoSyogai, "'", "''") & "'," '' 障害本賞金累計
                    SQL = SQL & "[RuikeiFukaHeichi]='" & Replace(.RuikeiFukaHeichi, "'", "''") & "'," '' 平地付加賞金累計
                    SQL = SQL & "[RuikeiFukaSyogai]='" & Replace(.RuikeiFukaSyogai, "'", "''") & "'," '' 障害付加賞金累計
                    SQL = SQL & "[RuikeiSyutokuHeichi]='" & Replace(.RuikeiSyutokuHeichi, "'", "''") & "'," '' 平地収得賞金累計
                    SQL = SQL & "[RuikeiSyutokuSyogai]='" & Replace(.RuikeiSyutokuSyogai, "'", "''") & "'," '' 障害収得賞金累計

                    temp = ""
                    '' 総合着回数
                    With .ChakuSogo
                        For j = 0 To 5
                            temp &= .Chakukaisu(j)
                        Next j
                    End With
                    '' 中央合計着回数
                    With .ChakuChuo
                        For j = 0 To 5
                            temp &= .Chakukaisu(j)
                        Next j
                    End With
                    '' 馬場別着回数
                    For j = 0 To 6
                        With .ChakuKaisuBa(j)
                            For k = 0 To 5
                                temp &= .Chakukaisu(k)
                            Next k
                        End With
                    Next j
                    '' 馬場状態別着回数
                    For j = 0 To 11
                        With .ChakuKaisuJyotai(j)
                            For k = 0 To 5
                                temp &= .Chakukaisu(k)
                            Next k
                        End With
                    Next j
                    '' 距離別着回数(芝)
                    For j = 0 To 8
                        With .ChakuKaisuSibaKyori(j)
                            For k = 0 To 5
                                temp &= .Chakukaisu(k)
                            Next k
                        End With
                    Next j
                    '' 距離別着回数(ダート)
                    For j = 0 To 8
                        With .ChakuKaisuDirtKyori(j)
                            For k = 0 To 5
                                temp &= .Chakukaisu(k)
                            Next k
                        End With
                    Next j
                    '' 競馬場別着回数(芝)
                    For j = 0 To 9
                        With .ChakuKaisuJyoSiba(j)
                            For k = 0 To 5
                                temp &= .Chakukaisu(k)
                            Next k
                        End With
                    Next j
                    '' 競馬場別着回数(ダート)
                    For j = 0 To 9
                        With .ChakuKaisuJyoDirt(j)
                            For k = 0 To 5
                                temp &= .Chakukaisu(k)
                            Next k
                        End With
                    Next j
                    '' 競馬場別着回数(障害)
                    For j = 0 To 9
                        With .ChakuKaisuJyoSyogai(j)
                            For k = 0 To 5
                                temp &= .Chakukaisu(k)
                            Next k
                        End With
                    Next j
                    SQL = SQL & "[Chakukaisu]='" & Replace(temp, "'", "''") & "',"
                    '' 脚質傾向
                    temp = ""
                    For j = 0 To 3
                        temp &= .Kyakusitu(j)
                    Next j
                    SQL = SQL & "[Kyakusitu]='" & Replace(temp, "'", "''") & "',"
                    SQL = SQL & "[RaceCount]='" & Replace(.RaceCount, "'", "''") & "'," '' 登録レース数
                End With
                
                With .KisyuChaku
                    SQL = SQL & "[KisyuCode]='" & Replace(.KisyuCode, "'", "''") & "'," '' 騎手コード
                    SQL = SQL & "[KisyuName]='" & Replace(.KisyuName, "'", "''") & "'," '' 騎手名

                    '' 騎手本年･累計成績情報
                    For i = 0 To 1
                        With .HonRuikei(i)
                            If i = 0 Then
                                s = "H"
                            Else
                                s = "R"
                            End If

                            SQL = SQL & "[K_" & s & "_SetYear]='" & Replace(.SetYear, "'", "''") & "'," '' 設定年
                            SQL = SQL & "[K_" & s & "_HonSyokinHeichi]='" & Replace(.HonSyokinHeichi, "'", "''") & "'," '' 平地本賞金合計
                            SQL = SQL & "[K_" & s & "_HonSyokinSyogai]='" & Replace(.HonSyokinSyogai, "'", "''") & "'," '' 障害本賞金合計
                            SQL = SQL & "[K_" & s & "_FukaSyokinHeichi]='" & Replace(.FukaSyokinHeichi, "'", "''") & "'," '' 平地付加賞金合計
                            SQL = SQL & "[K_" & s & "_FukaSyokinSyogai]='" & Replace(.FukaSyokinSyogai, "'", "''") & "'," '' 障害付加賞金合計
                            temp = ""
                            '' 芝着回数
                            With .ChakuKaisuSiba
                                For k = 0 To 5
                                    temp &= .Chakukaisu(k)
                                Next
                            End With
                            '' ダート着回数
                            With .ChakuKaisuDirt
                                For k = 0 To 5
                                    temp &= .Chakukaisu(k)
                                Next
                            End With
                            '' 障害着回数
                            With .ChakuKaisuSyogai
                                For k = 0 To 5
                                    temp &= .Chakukaisu(k)
                                Next
                            End With
                            '' 距離別着回数(芝)
                            For j = 0 To 8
                                With .ChakuKaisuSibaKyori(j)
                                    For k = 0 To 5
                                        temp &= .Chakukaisu(k)
                                    Next
                                End With
                            Next
                            '' 距離別着回数(ダート)
                            For j = 0 To 8
                                With .ChakuKaisuDirtKyori(j)
                                    For k = 0 To 5
                                        temp &= .Chakukaisu(k)
                                    Next
                                End With
                            Next
                            '' 競馬場別着回数(芝)
                            For j = 0 To 9
                                With .ChakuKaisuJyoSiba(j)
                                    For k = 0 To 5
                                        temp &= .Chakukaisu(k)
                                    Next
                                End With
                            Next
                            '' 競馬場別着回数(ダート)
                            For j = 0 To 9
                                With .ChakuKaisuJyoDirt(j)
                                    For k = 0 To 5
                                        temp &= .Chakukaisu(k)
                                    Next
                                End With
                            Next
                            '' 競馬場別着回数(障害)
                            For j = 0 To 9
                                With .ChakuKaisuJyoSyogai(j)
                                    For k = 0 To 5
                                        temp &= .Chakukaisu(k)
                                    Next
                                End With
                            Next
                            SQL = SQL & "[K_" & s & "_Chakukaisu]='" & Replace(temp, "'", "''") & "',"
                        End With
                    Next
                End With

                With .ChokyoChaku
                    SQL = SQL & "[ChokyosiCode]='" & Replace(.ChokyosiCode, "'", "''") & "'," '' 調教師コード
                    SQL = SQL & "[ChokyosiName]='" & Replace(.ChokyosiName, "'", "''") & "'," '' 調教師名

                    '' 調教師本年･累計成績情報
                    For i = 0 To 1
                        With .HonRuikei(i)
                            If i = 0 Then
                                s = "H"
                            Else
                                s = "R"
                            End If

                            SQL = SQL & "[C_" & s & "_SetYear]='" & Replace(.SetYear, "'", "''") & "'," '' 設定年
                            SQL = SQL & "[C_" & s & "_HonSyokinHeichi]='" & Replace(.HonSyokinHeichi, "'", "''") & "'," '' 平地本賞金合計
                            SQL = SQL & "[C_" & s & "_HonSyokinSyogai]='" & Replace(.HonSyokinSyogai, "'", "''") & "'," '' 障害本賞金合計
                            SQL = SQL & "[C_" & s & "_FukaSyokinHeichi]='" & Replace(.FukaSyokinHeichi, "'", "''") & "'," '' 平地付加賞金合計
                            SQL = SQL & "[C_" & s & "_FukaSyokinSyogai]='" & Replace(.FukaSyokinSyogai, "'", "''") & "'," '' 障害付加賞金合計
                            temp = ""
                            '' 芝着回数
                            With .ChakuKaisuSiba
                                For k = 0 To 5
                                    temp &= .Chakukaisu(k)
                                Next
                            End With
                            '' ダート着回数
                            With .ChakuKaisuDirt
                                For k = 0 To 5
                                    temp &= .Chakukaisu(k)
                                Next
                            End With
                            '' 障害着回数
                            With .ChakuKaisuSyogai
                                For k = 0 To 5
                                    temp &= .Chakukaisu(k)
                                Next
                            End With
                            '' 距離別着回数(芝)
                            For j = 0 To 8
                                With .ChakuKaisuSibaKyori(j)
                                    For k = 0 To 5
                                        temp &= .Chakukaisu(k)
                                    Next
                                End With
                            Next
                            '' 距離別着回数(ダート)
                            For j = 0 To 8
                                With .ChakuKaisuDirtKyori(j)
                                    For k = 0 To 5
                                        temp &= .Chakukaisu(k)
                                    Next
                                End With
                            Next
                            '' 競馬場別着回数(芝)
                            For j = 0 To 9
                                With .ChakuKaisuJyoSiba(j)
                                    For k = 0 To 5
                                        temp &= .Chakukaisu(k)
                                    Next
                                End With
                            Next
                            '' 競馬場別着回数(ダート)
                            For j = 0 To 9
                                With .ChakuKaisuJyoDirt(j)
                                    For k = 0 To 5
                                        temp &= .Chakukaisu(k)
                                    Next
                                End With
                            Next
                            '' 競馬場別着回数(障害)
                            For j = 0 To 9
                                With .ChakuKaisuJyoSyogai(j)
                                    For k = 0 To 5
                                        temp &= .Chakukaisu(k)
                                    Next
                                End With
                            Next
                            SQL = SQL & "[C_" & s & "_Chakukaisu]='" & Replace(temp, "'", "''") & "',"
                        End With
                    Next
                End With

                With .BanusiChaku
                    SQL = SQL & "[BanusiCode]='" & Replace(.BanusiCode, "'", "''") & "'," '' 馬主コード
                    SQL = SQL & "[BanusiName_Co]='" & Replace(.BanusiName_Co, "'", "''") & "'," '' 馬主名(法人格有)
                    SQL = SQL & "[BanusiName]='" & Replace(.BanusiName, "'", "''") & "'," '' 馬主名(法人格無)

                    '' 馬主本年･累計成績情報
                    For i = 0 To 1
                        With .HonRuikei(i)
                            If i = 0 Then
                                s = "H"
                            Else
                                s = "R"
                            End If

                            SQL = SQL & "[Ba_" & s & "_SetYear]='" & Replace(.SetYear, "'", "''") & "',"
                            SQL = SQL & "[Ba_" & s & "_HonSyokin]='" & Replace(.HonSyokinTotal, "'", "''") & "',"
                            SQL = SQL & "[Ba_" & s & "_FukaSyokin]='" & Replace(.FukaSyokin, "'", "''") & "',"
                            '' 着回数
                            temp = ""
                            For j = 0 To 5
                                temp &= .ChakuKaisu(j)
                            Next j
                            SQL = SQL & "[Ba_" & s & "_Chakukaisu]='" & Replace(temp, "'", "''") & "',"
                        End With
                    Next
                End With

                With .BreederChaku
                    SQL = SQL & "[BreederCode]='" & Replace(.BreederCode, "'", "''") & "'," '' 生産者コード
                    SQL = SQL & "[BreederName_Co]='" & Replace(.BreederName_Co, "'", "''") & "'," '' 生産者名(法人格有)
                    SQL = SQL & "[BreederName]='" & Replace(.BreederName, "'", "''") & "'," '' 生産者名(法人格無)
                    '' 生産者本年･累計成績情報
                    For i = 0 To 1
                        With .HonRuikei(i)
                            If i = 0 Then
                                s = "H"
                            Else
                                s = "R"
                            End If

                            SQL = SQL & "[Br_" & s & "_SetYear]='" & Replace(.SetYear, "'", "''") & "',"
                            SQL = SQL & "[Br_" & s & "_HonSyokin]='" & Replace(.HonSyokinTotal, "'", "''") & "',"
                            SQL = SQL & "[Br_" & s & "_FukaSyokin]='" & Replace(.FukaSyokin, "'", "''") & "',"
                            '' 着回数
                            temp = ""
                            For j = 0 To 5
                                temp &= .ChakuKaisu(j)
                            Next j
                            SQL = SQL & "[Br_" & s & "_Chakukaisu]='" & Replace(temp, "'", "''") & "'"
                            If i = 0 Then
                                SQL = SQL & ","
                            End If
                        End With
                    Next
                End With

                SQL = SQL & " WHERE [Year]='" & Replace(.id.Year, "'", "''") & "'"
                SQL = SQL & " AND [MonthDay] = '" & Replace(.id.MonthDay, "'", "''") & "'"
                SQL = SQL & " AND [JyoCD] = '" & Replace(.id.JyoCD, "'", "''") & "'"
                SQL = SQL & " AND [Kaiji] = '" & Replace(.id.Kaiji, "'", "''") & "'"
                SQL = SQL & " AND [Nichiji] = '" & Replace(.id.Nichiji, "'", "''") & "'"
                SQL = SQL & " AND [RaceNum] = '" & Replace(.id.RaceNum, "'", "''") & "'"
                SQL = SQL & " AND [KettoNum] = '" & Replace(.UmaChaku.KettoNum, "'", "''") & "'"
            End With
            gCon.Execute(SQL)

            With mBuf
                Debug.WriteLine("UPDATE RACE : " & .id.Year & .id.MonthDay & .id.JyoCD & .id.Kaiji & .id.Nichiji & .id.RaceNum & .UmaChaku.KettoNum)
            End With ' id

            gCon.CommitTrans()
            Debug.WriteLine("CommitTrans")

            Return True
        Catch ex As Exception
            Debug.WriteLine("RollbackTrans")
            gCon.RollbackTrans()

            Throw
        End Try

        Return False
    End Function
End Class
