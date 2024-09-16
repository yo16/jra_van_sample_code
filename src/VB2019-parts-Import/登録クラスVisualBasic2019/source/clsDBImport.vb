Option Strict Off
Option Explicit On
Option Compare Binary
Friend Class clsDBImport
    '========================================================================
    '  JRA-VAN Data Lab.�v���O���~���O�p�[�c�uJV-Data�o�^�N���X�v
    '
    '
    '   �쐬: JRA-VAN �\�t�g�E�F�A�H�[  2003�N6�� 3��
    '	�X�V:                           2006�N11�� 7��
    '	�X�V:                           2007�N11�� 8��
    '   �X�V:                           2012�N1��17��
    '========================================================================
    '   (C) Copyright JRA SYSTEM SERVICE CO.,LTD. 2007 All rights reserved
    '========================================================================

    Private strRecIDCur As String ''���R�[�h��ʁi���݁j
    Private strRecIDOld As String ''���R�[�h��ʁi�P�O�̃o�b�t�@�j
    Private ImportObj As Object

    ' @(f)
    '
    ' �@�\      : ��������
    '
    ' ������    :
    '
    ' �Ԃ�l    :
    '
    ' �@�\����  : �R�l�N�V�����I�[�v��
    '

    Private Sub Class_Initialize_Renamed()
        If ConnectDB() Then
            System.Diagnostics.Debug.WriteLine("gCon.Open")
        End If
ExitHandler:
        Exit Sub
    End Sub
    Public Sub New()
        MyBase.New()
        Class_Initialize_Renamed()
    End Sub

    ' @(f)
    '
    ' �@�\      : �I������
    '
    ' ������    :
    '
    ' �Ԃ�l    :
    '
    ' �@�\����  : �R�l�N�V�����N���[�Y
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


    Public Sub Close()
        '���R�[�h���ID�ɑO��ǂ񂾂��̂��c���Ă����
        If strRecIDOld <> "" Then
            '�����N���X��j��
            ImportObj.Close()
            ImportObj = Nothing
        End If
        '�R�l�N�V�����N���[�Y
        gCon.Close()
        System.Diagnostics.Debug.WriteLine("gCon.Close")

    End Sub


    ' @(f)
    '
    ' �@�\      : �e�[�u���N���A
    '
    ' ������    :
    '
    ' �Ԃ�l    :
    '
    ' �@�\����  :
    '
    Public Sub ClearData(Optional ByVal strTBLName As String = "")
        On Error GoTo ErrorHandler
        Dim strDel As String ''SQL��

        '�J�n����
        gCon.BeginTrans()

        If strTBLName <> "" Then

            '�w�肵���e�[�u�����폜����
            strDel = "DELETE FROM " & strTBLName
            gCon.Execute(strDel)

        Else

            '�e�[�u���̓��e��S�č폜����
            strDel = "DELETE FROM BANUSI"
            gCon.Execute(strDel)

            strDel = "DELETE FROM BATAIJYU"
            gCon.Execute(strDel)

            strDel = "DELETE FROM CHOKYO"
            gCon.Execute(strDel)

            strDel = "DELETE FROM CHOKYO_SEISEKI"
            gCon.Execute(strDel)

            strDel = "DELETE FROM HANRO"
            gCon.Execute(strDel)

            strDel = "DELETE FROM HANSYOKU"
            gCon.Execute(strDel)

            strDel = "DELETE FROM HARAI"
            gCon.Execute(strDel)

            strDel = "DELETE FROM HYOSU"
            gCon.Execute(strDel)

            strDel = "DELETE FROM HYOSU2"
            gCon.Execute(strDel)

            strDel = "DELETE FROM HYOSU_SANRENTAN"
            gCon.Execute(strDel)

            strDel = "DELETE FROM HYOSU_SANREN"
            gCon.Execute(strDel)

            strDel = "DELETE FROM HYOSU_TANPUKU"
            gCon.Execute(strDel)

            strDel = "DELETE FROM HYOSU_UMARENWIDE"
            gCon.Execute(strDel)

            strDel = "DELETE FROM HYOSU_UMATAN"
            gCon.Execute(strDel)

            strDel = "DELETE FROM HYOSU_WAKU"
            gCon.Execute(strDel)

            strDel = "DELETE FROM KISYU"
            gCon.Execute(strDel)

            strDel = "DELETE FROM KISYU_CHANGE"
            gCon.Execute(strDel)

            strDel = "DELETE FROM KISYU_SEISEKI"
            gCon.Execute(strDel)

            strDel = "DELETE FROM MINING"
            gCon.Execute(strDel)

            strDel = "DELETE FROM ODDS_SANRENTAN"
            gCon.Execute(strDel)

            strDel = "DELETE FROM ODDS_SANRENTAN_HEAD"
            gCon.Execute(strDel)

            strDel = "DELETE FROM ODDS_SANREN"
            gCon.Execute(strDel)

            strDel = "DELETE FROM ODDS_SANREN_HEAD"
            gCon.Execute(strDel)

            strDel = "DELETE FROM ODDS_TANPUKU"
            gCon.Execute(strDel)

            strDel = "DELETE FROM ODDS_TANPUKUWAKU_HEAD"
            gCon.Execute(strDel)

            strDel = "DELETE FROM ODDS_UMAREN"
            gCon.Execute(strDel)

            strDel = "DELETE FROM ODDS_UMAREN_HEAD"
            gCon.Execute(strDel)

            strDel = "DELETE FROM ODDS_UMATAN"
            gCon.Execute(strDel)

            strDel = "DELETE FROM ODDS_UMATAN_HEAD"
            gCon.Execute(strDel)

            strDel = "DELETE FROM ODDS_WAKU"
            gCon.Execute(strDel)

            strDel = "DELETE FROM ODDS_WIDE"
            gCon.Execute(strDel)

            strDel = "DELETE FROM ODDS_WIDE_HEAD"
            gCon.Execute(strDel)

            strDel = "DELETE FROM RACE"
            gCon.Execute(strDel)

            strDel = "DELETE FROM RECORD"
            gCon.Execute(strDel)

            strDel = "DELETE FROM SANKU"
            gCon.Execute(strDel)

            strDel = "DELETE FROM SCHEDULE"
            gCon.Execute(strDel)

            strDel = "DELETE FROM SEISAN"
            gCon.Execute(strDel)

            strDel = "DELETE FROM TENKO_BABA"
            gCon.Execute(strDel)

            strDel = "DELETE FROM TOKU"
            gCon.Execute(strDel)

            strDel = "DELETE FROM TOKU_RACE"
            gCon.Execute(strDel)

            strDel = "DELETE FROM TORIKESI_JYOGAI"
            gCon.Execute(strDel)

            strDel = "DELETE FROM UMA"
            gCon.Execute(strDel)

            strDel = "DELETE FROM UMA_RACE"
            gCon.Execute(strDel)

            strDel = "DELETE FROM COURSE_CHANGE"
            gCon.Execute(strDel)

            strDel = "DELETE FROM HASSOU_JIKOKU_CHANGE"
            gCon.Execute(strDel)

            strDel = "DELETE FROM SALE"
            gCon.Execute(strDel)

            strDel = "DELETE FROM BAMEIORIGIN"
            gCon.Execute(strDel)

            strDel = "DELETE FROM KEITO"
            gCon.Execute(strDel)

            strDel = "DELETE FROM COURSE"
            gCon.Execute(strDel)

            strDel = "DELETE FROM TAISENGATA_MINING"
            gCon.Execute(strDel)

            strDel = "DELETE FROM JYUSYOSIKI_HEAD"
            gCon.Execute(strDel)

            strDel = "DELETE FROM JYUSYOSIKI"
            gCon.Execute(strDel)

            strDel = "DELETE FROM JOGAIBA"
            gCon.Execute(strDel)

        End If

        '�I������
        gCon.CommitTrans()
        System.Diagnostics.Debug.WriteLine("gCon.CommitTrans")

ExitHandler:
        Exit Sub
ErrorHandler:
        '���~����
        gCon.RollbackTrans()

        System.Diagnostics.Debug.WriteLine(Err.Description)
        Resume ExitHandler
    End Sub

    Public Sub SetData(ByRef strBuff As String, ByVal lngBuffSize As Integer)
        On Error GoTo ErrorHandler

        '���R�[�h���ID���擾
        strRecIDCur = Left(strBuff, 2)
        System.Diagnostics.Debug.WriteLine("SetData " & strRecIDCur)


        If (strRecIDOld <> strRecIDCur) Then

            '���R�[�h���ID�ɑO��ǂ񂾂��̂��c���Ă����
            If ImportObj IsNot Nothing Then
                '�����N���X��j��
                ImportObj.Close()
                ImportObj = Nothing
            End If


            '�����N���X�쐬
            Select Case strRecIDCur
                Case "TK"
                    ImportObj = New clsImportTK()
                Case "RA"
                    ImportObj = New clsImportRA()
                Case "SE"
                    ImportObj = New clsImportSE()
                Case "HR"
                    ImportObj = New clsImportHR()
                Case "H1"
                    ImportObj = New clsImportH1()
                Case "H6"
                    ImportObj = New clsImportH6()
                Case "O1"
                    ImportObj = New clsImportO1()
                Case "O2"
                    ImportObj = New clsImportO2()
                Case "O3"
                    ImportObj = New clsImportO3()
                Case "O4"
                    ImportObj = New clsImportO4()
                Case "O5"
                    ImportObj = New clsImportO5()
                Case "O6"
                    ImportObj = New clsImportO6()
                Case "UM"
                    ImportObj = New clsImportUM()
                Case "KS"
                    ImportObj = New clsImportKS()
                Case "CH"
                    ImportObj = New clsImportCH()
                Case "BR"
                    ImportObj = New clsImportBR()
                Case "BN"
                    ImportObj = New clsImportBN()
                Case "RC"
                    ImportObj = New clsImportRC()
                Case "HN"
                    ImportObj = New clsImportHN()
                Case "SK"
                    ImportObj = New clsImportSK()
                Case "HC"
                    ImportObj = New clsImportHC()
                Case "WH"
                    ImportObj = New clsImportWH()
                Case "WE"
                    ImportObj = New clsImportWE()
                Case "AV"
                    ImportObj = New clsImportAV()
                Case "JC"
                    ImportObj = New clsImportJC()
                Case "TC"
                    ImportObj = New clsImportTC()
                Case "CC"
                    ImportObj = New clsImportCC()
                Case "DM"
                    ImportObj = New clsImportDM()
                Case "YS"
                    ImportObj = New clsImportYS()
                Case "HS"
                    ImportObj = New clsImportHS()
                Case "HY"
                    ImportObj = New clsImportHY()
                Case "BT"
                    ImportObj = New clsImportBT()
                Case "CS"
                    ImportObj = New clsImportCS()
                Case "TM"
                    ImportObj = New clsImportTM()
                Case "WF"
                    ImportObj = New clsImportWF()
                Case "JG"
                    ImportObj = New clsImportJG()
                Case Else
                    System.Diagnostics.Debug.WriteLine("����`�̃��R�[�h���[" & strRecIDCur & "]")
                    Exit Sub
            End Select
        End If

        'DB�ǉ�����

        Call ImportObj.Add(strBuff, lngBuffSize)

        '���R�[�h���ID��ێ�
        strRecIDOld = strRecIDCur


ExitHandler:
        Exit Sub

ErrorHandler:
        System.Diagnostics.Debug.WriteLine(Err.Description)
        Resume ExitHandler
    End Sub
End Class