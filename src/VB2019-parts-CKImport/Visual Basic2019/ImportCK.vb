Imports System.Collections.Generic
Imports System.Diagnostics

Public Class ImportCK
    Inherits ImportBase

    Public Sub New()
        MyBase.New()

        Dim SQL As New List(Of String)

        mBuf = CType(mBuf, JV_CK_CHAKU)

        '' ���s����SQL��List�Ɋi�[
        SQL.Add("SELECT * FROM CHAKU")

        '' ���R�[�h�Z�b�gOPEN�������s
        Class_Initialize_Renamed(SQL)
    End Sub

    Protected Overrides Function InsertDB() As Boolean
        Dim i As Short '' ���[�v�J�E���^
        Dim j As Short '' ���[�v�J�E���^
        Dim k As Short '' ���[�v�J�E���^
        Dim temp As String
        Dim s As String

        Try
            gCon.BeginTrans()
            Debug.WriteLine("BeginTrans")

            mRS(0).AddNew()

            With mBuf
                With .head
                    mRS(0).Fields("RecordSpec").Value = .RecordSpec '' ���R�[�h���
                    mRS(0).Fields("DataKubun").Value = .DataKubun '' �f�[�^�敪
                    With .MakeDate
                        mRS(0).Fields("MakeDate").Value = .Year & .Month & .Day '' �N����
                    End With ' MakeDate
                End With ' head

                With .id
                    mRS(0).Fields("Year").Value = .Year '' �J�ÔN
                    mRS(0).Fields("MonthDay").Value = .MonthDay '' �J�Ì���
                    mRS(0).Fields("JyoCD").Value = .JyoCD '' ���n��R�[�h
                    mRS(0).Fields("Kaiji").Value = .Kaiji '' �J�É��N��
                    mRS(0).Fields("Nichiji").Value = .Nichiji '' �J�Ó���N����
                    mRS(0).Fields("RaceNum").Value = .RaceNum '' ���[�X�ԍ�
                End With ' id

                With .UmaChaku
                    mRS(0).Fields("KettoNum").Value = .KettoNum '' �����o�^�ԍ�
                    mRS(0).Fields("Bamei").Value = .Bamei '' �n��
                    mRS(0).Fields("RuikeiHonsyoHeiti").Value = .RuikeiHonsyoHeiti '' ���n�{�܋��݌v
                    mRS(0).Fields("RuikeiHonsyoSyogai").Value = .RuikeiHonsyoSyogai '' ��Q�{�܋��݌v
                    mRS(0).Fields("RuikeiFukaHeichi").Value = .RuikeiFukaHeichi '' ���n�t���܋��݌v
                    mRS(0).Fields("RuikeiFukaSyogai").Value = .RuikeiFukaSyogai '' ��Q�t���܋��݌v
                    mRS(0).Fields("RuikeiSyutokuHeichi").Value = .RuikeiSyutokuHeichi '' ���n�����܋��݌v
                    mRS(0).Fields("RuikeiSyutokuSyogai").Value = .RuikeiSyutokuSyogai '' ��Q�����܋��݌v

                    temp = ""
                    '' ��������
                    With .ChakuSogo
                        For j = 0 To 5
                            temp &= .Chakukaisu(j)
                        Next j
                    End With
                    '' �������v����
                    With .ChakuChuo
                        For j = 0 To 5
                            temp &= .Chakukaisu(j)
                        Next j
                    End With
                    '' �n��ʒ���
                    For j = 0 To 6
                        With .ChakuKaisuBa(j)
                            For k = 0 To 5
                                temp &= .Chakukaisu(k)
                            Next k
                        End With
                    Next j
                    '' �n���ԕʒ���
                    For j = 0 To 11
                        With .ChakuKaisuJyotai(j)
                            For k = 0 To 5
                                temp &= .Chakukaisu(k)
                            Next k
                        End With
                    Next j
                    '' �����ʒ���(��)
                    For j = 0 To 8
                        With .ChakuKaisuSibaKyori(j)
                            For k = 0 To 5
                                temp &= .Chakukaisu(k)
                            Next k
                        End With
                    Next j
                    '' �����ʒ���(�_�[�g)
                    For j = 0 To 8
                        With .ChakuKaisuDirtKyori(j)
                            For k = 0 To 5
                                temp &= .Chakukaisu(k)
                            Next k
                        End With
                    Next j
                    '' ���n��ʒ���(��)
                    For j = 0 To 9
                        With .ChakuKaisuJyoSiba(j)
                            For k = 0 To 5
                                temp &= .Chakukaisu(k)
                            Next k
                        End With
                    Next j
                    '' ���n��ʒ���(�_�[�g)
                    For j = 0 To 9
                        With .ChakuKaisuJyoDirt(j)
                            For k = 0 To 5
                                temp &= .Chakukaisu(k)
                            Next k
                        End With
                    Next j
                    '' ���n��ʒ���(��Q)
                    For j = 0 To 9
                        With .ChakuKaisuJyoSyogai(j)
                            For k = 0 To 5
                                temp &= .Chakukaisu(k)
                            Next k
                        End With
                    Next j
                    mRS(0).Fields("Chakukaisu").Value = temp

                    '' �r���X��
                    temp = ""
                    For j = 0 To 3
                        temp &= .Kyakusitu(j)
                    Next j
                    mRS(0).Fields("Kyakusitu").Value = temp
                    mRS(0).Fields("RaceCount").Value = .RaceCount '' �o�^���[�X��
                End With

                With .KisyuChaku
                    mRS(0).Fields("KisyuCode").Value = .KisyuCode '' �R��R�[�h
                    mRS(0).Fields("KisyuName").Value = .KisyuName '' �R�薼

                    '' �R��{�N��݌v���я��
                    For i = 0 To 1
                        With .HonRuikei(i)
                            If i = 0 Then
                                s = "H"
                            Else
                                s = "R"
                            End If

                            mRS(0).Fields("K_" & s & "_SetYear").Value = .SetYear '' �ݒ�N
                            mRS(0).Fields("K_" & s & "_HonSyokinHeichi").Value = .HonSyokinHeichi '' ���n�{�܋����v
                            mRS(0).Fields("K_" & s & "_HonSyokinSyogai").Value = .HonSyokinSyogai '' ��Q�{�܋����v
                            mRS(0).Fields("K_" & s & "_FukaSyokinHeichi").Value = .FukaSyokinHeichi '' ���n�t���܋����v
                            mRS(0).Fields("K_" & s & "_FukaSyokinSyogai").Value = .FukaSyokinSyogai '' ��Q�t���܋����v
                            temp = ""
                            '' �Œ���
                            With .ChakuKaisuSiba
                                For k = 0 To 5
                                    temp &= .Chakukaisu(k)
                                Next
                            End With
                            '' �_�[�g����
                            With .ChakuKaisuDirt
                                For k = 0 To 5
                                    temp &= .Chakukaisu(k)
                                Next
                            End With
                            '' ��Q����
                            With .ChakuKaisuSyogai
                                For k = 0 To 5
                                    temp &= .Chakukaisu(k)
                                Next
                            End With
                            '' �����ʒ���(��)
                            For j = 0 To 8
                                With .ChakuKaisuSibaKyori(j)
                                    For k = 0 To 5
                                        temp &= .Chakukaisu(k)
                                    Next
                                End With
                            Next
                            '' �����ʒ���(�_�[�g)
                            For j = 0 To 8
                                With .ChakuKaisuDirtKyori(j)
                                    For k = 0 To 5
                                        temp &= .Chakukaisu(k)
                                    Next
                                End With
                            Next
                            '' ���n��ʒ���(��)
                            For j = 0 To 9
                                With .ChakuKaisuJyoSiba(j)
                                    For k = 0 To 5
                                        temp &= .Chakukaisu(k)
                                    Next
                                End With
                            Next
                            '' ���n��ʒ���(�_�[�g)
                            For j = 0 To 9
                                With .ChakuKaisuJyoDirt(j)
                                    For k = 0 To 5
                                        temp &= .Chakukaisu(k)
                                    Next
                                End With
                            Next
                            '' ���n��ʒ���(��Q)
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
                    mRS(0).Fields("ChokyosiCode").Value = .ChokyosiCode '' �����t�R�[�h
                    mRS(0).Fields("ChokyosiName").Value = .ChokyosiName '' �����t��

                    '' �����t�{�N��݌v���я��
                    For i = 0 To 1
                        With .HonRuikei(i)
                            If i = 0 Then
                                s = "H"
                            Else
                                s = "R"
                            End If

                            mRS(0).Fields("C_" & s & "_SetYear").Value = .SetYear '' �ݒ�N
                            mRS(0).Fields("C_" & s & "_HonSyokinHeichi").Value = .HonSyokinHeichi '' ���n�{�܋����v
                            mRS(0).Fields("C_" & s & "_HonSyokinSyogai").Value = .HonSyokinSyogai '' ��Q�{�܋����v
                            mRS(0).Fields("C_" & s & "_FukaSyokinHeichi").Value = .FukaSyokinHeichi '' ���n�t���܋����v
                            mRS(0).Fields("C_" & s & "_FukaSyokinSyogai").Value = .FukaSyokinSyogai '' ��Q�t���܋����v
                            temp = ""
                            '' �Œ���
                            With .ChakuKaisuSiba
                                For k = 0 To 5
                                    temp &= .Chakukaisu(k)
                                Next
                            End With
                            '' �_�[�g����
                            With .ChakuKaisuDirt
                                For k = 0 To 5
                                    temp &= .Chakukaisu(k)
                                Next
                            End With
                            '' ��Q����
                            With .ChakuKaisuSyogai
                                For k = 0 To 5
                                    temp &= .Chakukaisu(k)
                                Next
                            End With
                            '' �����ʒ���(��)
                            For j = 0 To 8
                                With .ChakuKaisuSibaKyori(j)
                                    For k = 0 To 5
                                        temp &= .Chakukaisu(k)
                                    Next
                                End With
                            Next
                            '' �����ʒ���(�_�[�g)
                            For j = 0 To 8
                                With .ChakuKaisuDirtKyori(j)
                                    For k = 0 To 5
                                        temp &= .Chakukaisu(k)
                                    Next
                                End With
                            Next
                            '' ���n��ʒ���(��)
                            For j = 0 To 9
                                With .ChakuKaisuJyoSiba(j)
                                    For k = 0 To 5
                                        temp &= .Chakukaisu(k)
                                    Next
                                End With
                            Next
                            '' ���n��ʒ���(�_�[�g)
                            For j = 0 To 9
                                With .ChakuKaisuJyoDirt(j)
                                    For k = 0 To 5
                                        temp &= .Chakukaisu(k)
                                    Next
                                End With
                            Next
                            '' ���n��ʒ���(��Q)
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
                    mRS(0).Fields("BanusiCode").Value = .BanusiCode '' �n��R�[�h
                    mRS(0).Fields("BanusiName_Co").Value = .BanusiName_Co '' �n�喼(�@�l�i�L)
                    mRS(0).Fields("BanusiName").Value = .BanusiName '' �n�喼(�@�l�i��)

                    '' �n��{�N��݌v���я��
                    For i = 0 To 1
                        With .HonRuikei(i)
                            If i = 0 Then
                                s = "H"
                            Else
                                s = "R"
                            End If

                            mRS(0).Fields("Ba_" & s & "_SetYear").Value = .SetYear '' �ݒ�N
                            mRS(0).Fields("Ba_" & s & "_HonSyokin").Value = .HonSyokinTotal '' �{�܋����v
                            mRS(0).Fields("Ba_" & s & "_FukaSyokin").Value = .FukaSyokin '' �t���܋����v
                            '' ����
                            temp = ""
                            For j = 0 To 5
                                temp &= .ChakuKaisu(j)
                            Next j
                            mRS(0).Fields("Ba_" & s & "_Chakukaisu").Value = temp
                        End With
                    Next i
                End With

                With .BreederChaku
                    mRS(0).Fields("BreederCode").Value = .BreederCode '' ���Y�҃R�[�h
                    mRS(0).Fields("BreederName_Co").Value = .BreederName '' ���Y�Җ�(�@�l�i�L)
                    mRS(0).Fields("BreederName").Value = .BreederName '' ���Y�Җ�(�@�l�i��)

                    '' ���Y�Җ{�N��݌v���я��
                    For i = 0 To 1
                        With .HonRuikei(i)
                            If i = 0 Then
                                s = "H"
                            Else
                                s = "R"
                            End If

                            mRS(0).Fields("Br_" & s & "_SetYear").Value = .SetYear '' �ݒ�N
                            mRS(0).Fields("Br_" & s & "_HonSyokin").Value = .HonSyokinTotal '' �{�܋����v
                            mRS(0).Fields("Br_" & s & "_FukaSyokin").Value = .FukaSyokin '' �t���܋����v
                            '' ����
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
        Dim i As Short '' ���[�v�J�E���^
        Dim j As Short '' ���[�v�J�E���^
        Dim k As Short '' ���[�v�J�E���^
        Dim SQL As String '' SQL��
        Dim temp As String
        Dim s As String

        Try
            gCon.BeginTrans()
            System.Diagnostics.Debug.WriteLine("BeginTrans")

            SQL = "UPDATE CHAKU SET "
            With mBuf
                With .head
                    SQL = SQL & "[RecordSpec]='" & Replace(.RecordSpec, "'", "''") & "'," '' ���R�[�h���
                    SQL = SQL & "[DataKubun]='" & Replace(.DataKubun, "'", "''") & "'," '' �f�[�^�敪
                    SQL = SQL & "[MakeDate]= '" & Replace(strMakeDate, "'", "''") & "'," '' �N����
                End With ' head
                With .id
                    SQL = SQL & "[Year]='" & Replace(.Year, "'", "''") & "'," '' �J�ÔN
                    SQL = SQL & "[MonthDay]='" & Replace(.MonthDay, "'", "''") & "'," '' �J�Ì���
                    SQL = SQL & "[JyoCD]='" & Replace(.JyoCD, "'", "''") & "'," '' ���n��R�[�h
                    SQL = SQL & "[Kaiji]='" & Replace(.Kaiji, "'", "''") & "'," '' �J�É��N��
                    SQL = SQL & "[Nichiji]='" & Replace(.Nichiji, "'", "''") & "'," '' �J�Ó���N����
                    SQL = SQL & "[RaceNum]='" & Replace(.RaceNum, "'", "''") & "'," '' ���[�X�ԍ�
                End With ' id

                With .UmaChaku
                    SQL = SQL & "[KettoNum]='" & Replace(.KettoNum, "'", "''") & "'," '' �����o�^�ԍ�
                    SQL = SQL & "[Bamei]='" & Replace(.Bamei, "'", "''") & "'," '' �n��
                    SQL = SQL & "[RuikeiHonsyoHeiti]='" & Replace(.RuikeiHonsyoHeiti, "'", "''") & "'," '' ���n�{�܋��݌v
                    SQL = SQL & "[RuikeiHonsyoSyogai]='" & Replace(.RuikeiHonsyoSyogai, "'", "''") & "'," '' ��Q�{�܋��݌v
                    SQL = SQL & "[RuikeiFukaHeichi]='" & Replace(.RuikeiFukaHeichi, "'", "''") & "'," '' ���n�t���܋��݌v
                    SQL = SQL & "[RuikeiFukaSyogai]='" & Replace(.RuikeiFukaSyogai, "'", "''") & "'," '' ��Q�t���܋��݌v
                    SQL = SQL & "[RuikeiSyutokuHeichi]='" & Replace(.RuikeiSyutokuHeichi, "'", "''") & "'," '' ���n�����܋��݌v
                    SQL = SQL & "[RuikeiSyutokuSyogai]='" & Replace(.RuikeiSyutokuSyogai, "'", "''") & "'," '' ��Q�����܋��݌v

                    temp = ""
                    '' ��������
                    With .ChakuSogo
                        For j = 0 To 5
                            temp &= .Chakukaisu(j)
                        Next j
                    End With
                    '' �������v����
                    With .ChakuChuo
                        For j = 0 To 5
                            temp &= .Chakukaisu(j)
                        Next j
                    End With
                    '' �n��ʒ���
                    For j = 0 To 6
                        With .ChakuKaisuBa(j)
                            For k = 0 To 5
                                temp &= .Chakukaisu(k)
                            Next k
                        End With
                    Next j
                    '' �n���ԕʒ���
                    For j = 0 To 11
                        With .ChakuKaisuJyotai(j)
                            For k = 0 To 5
                                temp &= .Chakukaisu(k)
                            Next k
                        End With
                    Next j
                    '' �����ʒ���(��)
                    For j = 0 To 8
                        With .ChakuKaisuSibaKyori(j)
                            For k = 0 To 5
                                temp &= .Chakukaisu(k)
                            Next k
                        End With
                    Next j
                    '' �����ʒ���(�_�[�g)
                    For j = 0 To 8
                        With .ChakuKaisuDirtKyori(j)
                            For k = 0 To 5
                                temp &= .Chakukaisu(k)
                            Next k
                        End With
                    Next j
                    '' ���n��ʒ���(��)
                    For j = 0 To 9
                        With .ChakuKaisuJyoSiba(j)
                            For k = 0 To 5
                                temp &= .Chakukaisu(k)
                            Next k
                        End With
                    Next j
                    '' ���n��ʒ���(�_�[�g)
                    For j = 0 To 9
                        With .ChakuKaisuJyoDirt(j)
                            For k = 0 To 5
                                temp &= .Chakukaisu(k)
                            Next k
                        End With
                    Next j
                    '' ���n��ʒ���(��Q)
                    For j = 0 To 9
                        With .ChakuKaisuJyoSyogai(j)
                            For k = 0 To 5
                                temp &= .Chakukaisu(k)
                            Next k
                        End With
                    Next j
                    SQL = SQL & "[Chakukaisu]='" & Replace(temp, "'", "''") & "',"
                    '' �r���X��
                    temp = ""
                    For j = 0 To 3
                        temp &= .Kyakusitu(j)
                    Next j
                    SQL = SQL & "[Kyakusitu]='" & Replace(temp, "'", "''") & "',"
                    SQL = SQL & "[RaceCount]='" & Replace(.RaceCount, "'", "''") & "'," '' �o�^���[�X��
                End With
                
                With .KisyuChaku
                    SQL = SQL & "[KisyuCode]='" & Replace(.KisyuCode, "'", "''") & "'," '' �R��R�[�h
                    SQL = SQL & "[KisyuName]='" & Replace(.KisyuName, "'", "''") & "'," '' �R�薼

                    '' �R��{�N��݌v���я��
                    For i = 0 To 1
                        With .HonRuikei(i)
                            If i = 0 Then
                                s = "H"
                            Else
                                s = "R"
                            End If

                            SQL = SQL & "[K_" & s & "_SetYear]='" & Replace(.SetYear, "'", "''") & "'," '' �ݒ�N
                            SQL = SQL & "[K_" & s & "_HonSyokinHeichi]='" & Replace(.HonSyokinHeichi, "'", "''") & "'," '' ���n�{�܋����v
                            SQL = SQL & "[K_" & s & "_HonSyokinSyogai]='" & Replace(.HonSyokinSyogai, "'", "''") & "'," '' ��Q�{�܋����v
                            SQL = SQL & "[K_" & s & "_FukaSyokinHeichi]='" & Replace(.FukaSyokinHeichi, "'", "''") & "'," '' ���n�t���܋����v
                            SQL = SQL & "[K_" & s & "_FukaSyokinSyogai]='" & Replace(.FukaSyokinSyogai, "'", "''") & "'," '' ��Q�t���܋����v
                            temp = ""
                            '' �Œ���
                            With .ChakuKaisuSiba
                                For k = 0 To 5
                                    temp &= .Chakukaisu(k)
                                Next
                            End With
                            '' �_�[�g����
                            With .ChakuKaisuDirt
                                For k = 0 To 5
                                    temp &= .Chakukaisu(k)
                                Next
                            End With
                            '' ��Q����
                            With .ChakuKaisuSyogai
                                For k = 0 To 5
                                    temp &= .Chakukaisu(k)
                                Next
                            End With
                            '' �����ʒ���(��)
                            For j = 0 To 8
                                With .ChakuKaisuSibaKyori(j)
                                    For k = 0 To 5
                                        temp &= .Chakukaisu(k)
                                    Next
                                End With
                            Next
                            '' �����ʒ���(�_�[�g)
                            For j = 0 To 8
                                With .ChakuKaisuDirtKyori(j)
                                    For k = 0 To 5
                                        temp &= .Chakukaisu(k)
                                    Next
                                End With
                            Next
                            '' ���n��ʒ���(��)
                            For j = 0 To 9
                                With .ChakuKaisuJyoSiba(j)
                                    For k = 0 To 5
                                        temp &= .Chakukaisu(k)
                                    Next
                                End With
                            Next
                            '' ���n��ʒ���(�_�[�g)
                            For j = 0 To 9
                                With .ChakuKaisuJyoDirt(j)
                                    For k = 0 To 5
                                        temp &= .Chakukaisu(k)
                                    Next
                                End With
                            Next
                            '' ���n��ʒ���(��Q)
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
                    SQL = SQL & "[ChokyosiCode]='" & Replace(.ChokyosiCode, "'", "''") & "'," '' �����t�R�[�h
                    SQL = SQL & "[ChokyosiName]='" & Replace(.ChokyosiName, "'", "''") & "'," '' �����t��

                    '' �����t�{�N��݌v���я��
                    For i = 0 To 1
                        With .HonRuikei(i)
                            If i = 0 Then
                                s = "H"
                            Else
                                s = "R"
                            End If

                            SQL = SQL & "[C_" & s & "_SetYear]='" & Replace(.SetYear, "'", "''") & "'," '' �ݒ�N
                            SQL = SQL & "[C_" & s & "_HonSyokinHeichi]='" & Replace(.HonSyokinHeichi, "'", "''") & "'," '' ���n�{�܋����v
                            SQL = SQL & "[C_" & s & "_HonSyokinSyogai]='" & Replace(.HonSyokinSyogai, "'", "''") & "'," '' ��Q�{�܋����v
                            SQL = SQL & "[C_" & s & "_FukaSyokinHeichi]='" & Replace(.FukaSyokinHeichi, "'", "''") & "'," '' ���n�t���܋����v
                            SQL = SQL & "[C_" & s & "_FukaSyokinSyogai]='" & Replace(.FukaSyokinSyogai, "'", "''") & "'," '' ��Q�t���܋����v
                            temp = ""
                            '' �Œ���
                            With .ChakuKaisuSiba
                                For k = 0 To 5
                                    temp &= .Chakukaisu(k)
                                Next
                            End With
                            '' �_�[�g����
                            With .ChakuKaisuDirt
                                For k = 0 To 5
                                    temp &= .Chakukaisu(k)
                                Next
                            End With
                            '' ��Q����
                            With .ChakuKaisuSyogai
                                For k = 0 To 5
                                    temp &= .Chakukaisu(k)
                                Next
                            End With
                            '' �����ʒ���(��)
                            For j = 0 To 8
                                With .ChakuKaisuSibaKyori(j)
                                    For k = 0 To 5
                                        temp &= .Chakukaisu(k)
                                    Next
                                End With
                            Next
                            '' �����ʒ���(�_�[�g)
                            For j = 0 To 8
                                With .ChakuKaisuDirtKyori(j)
                                    For k = 0 To 5
                                        temp &= .Chakukaisu(k)
                                    Next
                                End With
                            Next
                            '' ���n��ʒ���(��)
                            For j = 0 To 9
                                With .ChakuKaisuJyoSiba(j)
                                    For k = 0 To 5
                                        temp &= .Chakukaisu(k)
                                    Next
                                End With
                            Next
                            '' ���n��ʒ���(�_�[�g)
                            For j = 0 To 9
                                With .ChakuKaisuJyoDirt(j)
                                    For k = 0 To 5
                                        temp &= .Chakukaisu(k)
                                    Next
                                End With
                            Next
                            '' ���n��ʒ���(��Q)
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
                    SQL = SQL & "[BanusiCode]='" & Replace(.BanusiCode, "'", "''") & "'," '' �n��R�[�h
                    SQL = SQL & "[BanusiName_Co]='" & Replace(.BanusiName_Co, "'", "''") & "'," '' �n�喼(�@�l�i�L)
                    SQL = SQL & "[BanusiName]='" & Replace(.BanusiName, "'", "''") & "'," '' �n�喼(�@�l�i��)

                    '' �n��{�N��݌v���я��
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
                            '' ����
                            temp = ""
                            For j = 0 To 5
                                temp &= .ChakuKaisu(j)
                            Next j
                            SQL = SQL & "[Ba_" & s & "_Chakukaisu]='" & Replace(temp, "'", "''") & "',"
                        End With
                    Next
                End With

                With .BreederChaku
                    SQL = SQL & "[BreederCode]='" & Replace(.BreederCode, "'", "''") & "'," '' ���Y�҃R�[�h
                    SQL = SQL & "[BreederName_Co]='" & Replace(.BreederName_Co, "'", "''") & "'," '' ���Y�Җ�(�@�l�i�L)
                    SQL = SQL & "[BreederName]='" & Replace(.BreederName, "'", "''") & "'," '' ���Y�Җ�(�@�l�i��)
                    '' ���Y�Җ{�N��݌v���я��
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
                            '' ����
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
