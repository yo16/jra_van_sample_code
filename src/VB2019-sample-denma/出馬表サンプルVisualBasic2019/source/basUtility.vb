Option Explicit On 

Imports System.Text

Module basUtility

    Declare Function GetPrivateProfileString Lib "KERNEL32.DLL" Alias "GetPrivateProfileStringA" ( _
     ByVal lpAppName As String, _
     ByVal lpKeyName As String, ByVal lpDefault As String, _
     ByVal lpReturnedString As StringBuilder, ByVal nSize As Integer, _
     ByVal lpFileName As String) As Integer

    Declare Function WritePrivateProfileString Lib "KERNEL32.DLL" Alias "WritePrivateProfileStringA" ( _
     ByVal lpAppName As String, _
     ByVal lpKeyName As String, _
     ByVal lpString As String, _
     ByVal lpFileName As String) As Integer

    Public ImportRA As clsImportRA
    Public ImportSE As clsImportSE
    Public ImportUM As clsImportUM

    Public objCDCv As clsCodeConv

    Public strConnectString As String

    ' @(f)
    '
    ' �@�\�@�@�@: �w��o�C�g���܂ŋ󔒂��u�E�Ɂv�t��������
    '
    ' �Ԃ�l�@�@: �w��o�C�g�� ������
    '
    ' �������@�@: strpad - �Ώە�����(byte�w����ɖ���)
    '          : totalBytes - �w�肷��o�C�g��
    ' 
    ' ���l     : �ʏ��PadRight�͑S�p����(2byte)��1�����Ɛ�����̂ŁA
    '            �o�C�g���Ŏw�肷�邱�ƂőS�p���p��������������������������킹�邱�Ƃ��e�ՂƂȂ�
    ' 
    Public Function bPadR(ByVal strPad As String, ByVal totalBytes As Integer) As String

        Dim strReturn As String
        Dim intTmp As Integer
        Dim bBuff As Byte()
        Dim bSize As Long

        If strPad Is Nothing Then
            strPad = ""
        End If
        bSize = Str2Byte(strPad).Length
        bBuff = New Byte(bSize) {}

        bBuff = Str2Byte(strPad)
        If bBuff.Length < totalBytes Then
            If bBuff.Length.Equals(strPad.Length) Then
                strReturn = strPad.PadRight(totalBytes)
            Else
                intTmp = totalBytes - (bBuff.Length - strPad.Length)
                strReturn = strPad.PadRight(intTmp)
            End If
        Else
            strReturn = strPad
        End If

        bPadR = strReturn
    End Function

    ' @(f)
    '
    ' �@�\�@�@�@: �w��o�C�g���܂ŋ󔒂��u���Ɂv�t��������
    '
    ' �Ԃ�l�@�@: �w��o�C�g�� ������
    '
    ' �������@�@: strpad - �Ώە�����(byte�w����ɖ���)
    '          : totalBytes - �w�肷��o�C�g��
    ' 
    ' ���l     : �ʏ��PadLeft�͑S�p����(2byte)��1�����Ɛ�����̂ŁA
    '            �o�C�g���Ŏw�肷�邱�ƂőS�p���p��������������������������킹�邱�Ƃ��e�ՂƂȂ�
    ' 
    Public Function bPadL(ByVal strPad As String, ByVal totalBytes As Integer) As String

        Dim strReturn As String
        Dim intTmp As Integer
        Dim bBuff As Byte()
        Dim bSize As Long

        If strPad Is Nothing Then
            strPad = ""
        End If
        bSize = Str2Byte(strPad).Length
        bBuff = New Byte(bSize) {}

        bBuff = Str2Byte(strPad)
        If bBuff.Length < totalBytes Then
            If bBuff.Length.Equals(strPad.Length) Then
                strReturn = strPad.PadLeft(totalBytes)
            Else
                intTmp = totalBytes - (bBuff.Length - strPad.Length)
                strReturn = strPad.PadLeft(intTmp)
            End If
        Else
            strReturn = strPad
        End If

        bPadL = strReturn
    End Function


    ' @(f)
    '
    ' �@�\�@�@�@: �����񂩂�󔒕����𔲂�
    '
    ' �Ԃ�l�@�@: �󔒂𔲂��Ώۂ̕�����
    '
    ' �������@�@: strTrim - ������(byte�w����ɖ���)
    '
    Public Function TrimSP(ByVal strTrim As String) As String

        Dim strReturn As String
        Dim i As Short ' ���[�v�J�E���^
        Dim strTmp As String
        strReturn = ""

        For i = 0 To strTrim.Length - 1
            ' �Ώە�����̐擪����1���������ׂ�
            strTmp = strTrim.Substring(i, 1)
            ' ���p�A�S�p�̋󔒂������A������Ɋi�[����
            If strTmp.Equals(" ") Or strTmp.Equals("�@") Then
            Else
                strReturn = strReturn & strTmp
            End If
        Next i

        TrimSP = strReturn
    End Function


    ' @(f)
    '
    ' �@�\�@�@�@: �O���[�h�R�[�h �� ���� ���Ǖ�����ɂ���
    '
    ' �Ԃ�l�@�@: ���� ������
    '
    ' �������@�@: strCD - �O���[�h�R�[�h������(1byte)
    '
    Public Function GRAD2(ByVal strCD As String) As String
        GRAD2 = String.Empty
        Select Case strCD
            Case "A"
                GRAD2 = "(�f�T)"
            Case "B"
                GRAD2 = "(�f�U)"
            Case "C"
                GRAD2 = "(�f�V)"
            Case "D"
                GRAD2 = ""
            Case "E"
                GRAD2 = ""
            Case "F"
                GRAD2 = "(J��f�T)"
            Case "G"
                GRAD2 = "(J��f�U)"
            Case "H"
                GRAD2 = "(J��f�V)"
            Case " "
                GRAD2 = ""
        End Select
    End Function

    ' @(f)
    '
    ' �@�\�@�@�@: �O���[�h�R�[�h �� ���� ���Ǖ�����ɂ���
    '
    ' �Ԃ�l�@�@: ���� ������
    '
    ' �������@�@: strCD - �O���[�h�R�[�h������(1byte)
    '
    Public Function GRAD3(ByVal strCD As String) As String
        GRAD3 = String.Empty
        Select Case strCD
            Case "A"
                GRAD3 = "(�f�T)"
            Case "B"
                GRAD3 = "(�f�U)"
            Case "C"
                GRAD3 = "(�f�V)"
            Case "D"
                GRAD3 = ""
            Case "E"
                GRAD3 = ""
            Case "F"
                GRAD3 = "(JG�T)"
            Case "G"
                GRAD3 = "(JG�U)"
            Case "H"
                GRAD3 = "(JG�V)"
            Case " "
                GRAD3 = ""
        End Select
    End Function

    ' @(f)
    '
    ' �@�\�@�@�@: ������ʃR�[�h �� ���� ���Ǖ�����ɂ���
    '
    ' �Ԃ�l�@�@: ���� ������
    '
    ' �������@�@: strCD - ������ʃR�[�h������(2byte)
    '
    Public Function KSSB6(ByVal strCD As String) As String
        KSSB6 = String.Empty
        Select Case strCD
            Case "0"
                KSSB6 = ""
            Case "11"
                KSSB6 = "�Q��" '"�T���u���b�h�n2��"
            Case "12"
                KSSB6 = "�R��" '"�T���u���b�h�n3��"
            Case "13"
                KSSB6 = "�R�Ώ�" '"�T���u���b�h�n3�Έȏ�"
            Case "14"
                KSSB6 = "�S�Ώ�" '"�T���u���b�h�n4�Έȏ�"
            Case "18"
                KSSB6 = "�R�Ώ�" '"�T���u���b�h�n��Q3�Έȏ�"
            Case "19"
                KSSB6 = "�S�Ώ�" '"�T���u���b�h�n��Q4�Έȏ�"
            Case "21"
                KSSB6 = "�Q��" '"�A���u�n2��"
            Case "22"
                KSSB6 = "�R��" '"�A���u�n3��"
            Case "23"
                KSSB6 = "�R�Ώ�" '"�A���u�n3�Έȏ�"
            Case "24"
                KSSB6 = "�S�Ώ�" '"�A���u�n4�Έȏ�"
        End Select
    End Function

    ' @(f)
    '
    ' �@�\�@�@�@: ������ʃR�[�h �� ���� ���Ǖ�����ɂ���
    '
    ' �Ԃ�l�@�@: ���� ������
    '
    ' �������@�@: strCD - ������ʃR�[�h������(2byte)
    '
    Public Function KSSB7(ByVal strCD As String) As String
        KSSB7 = String.Empty
        Select Case strCD
            Case "0"
                KSSB7 = ""
            Case "11"
                KSSB7 = "�Q��" '"�T���u���b�h�n2��"
            Case "12"
                KSSB7 = "�R��" '"�T���u���b�h�n3��"
            Case "13"
                KSSB7 = "�R�Ώ�" '"�T���u���b�h�n3�Έȏ�"
            Case "14"
                KSSB7 = "�S�Ώ�" '"�T���u���b�h�n4�Έȏ�"
            Case "18"
                KSSB7 = "��Q�R�Ώ�" '"�T���u���b�h�n��Q3�Έȏ�"
            Case "19"
                KSSB7 = "��Q�S�Ώ�" '"�T���u���b�h�n��Q4�Έȏ�"
            Case "21"
                KSSB7 = "�Q��" '"�A���u�n2��"
            Case "22"
                KSSB7 = "�R��" '"�A���u�n3��"
            Case "23"
                KSSB7 = "�R�Ώ�" '"�A���u�n3�Έȏ�"
            Case "24"
                KSSB7 = "�S�Ώ�" '"�A���u�n4�Έȏ�"
        End Select
    End Function

    ' @(f)
    '
    ' �@�\�@�@�@: ���������R�[�h �� ���� ���Ǖ�����ɂ���
    '
    ' �Ԃ�l�@�@: ���� ������
    '
    ' �������@�@: strCD - ���������R�[�h������(3byte)
    '
    Public Function KSJK4(ByVal strCD As String) As String
        KSJK4 = String.Empty
        Select Case Val(strCD)
            Case 0
                KSJK4 = ""
            Case 1 To 99
                KSJK4 = 100 * Val(strCD) & "����"
            Case 100
                KSJK4 = "�P��"
            Case "701"
                KSJK4 = "�V�n"
            Case "702"
                KSJK4 = "���o��"
            Case "703"
                KSJK4 = "������"
            Case "999"
                KSJK4 = "�����"
        End Select
    End Function

    ' @(f)
    '
    ' �@�\�@�@�@: ���ʃR�[�h �� ���� ���Ǖ�����ɂ���
    '
    ' �Ԃ�l�@�@: ���� ������
    '
    ' �������@�@: strCD - ���ʃR�[�h������(1byte)
    '
    Public Function SEIB4(ByVal strCD As String) As String
        SEIB4 = String.Empty
        Select Case strCD
            Case "0"
                SEIB4 = ""
            Case "1"
                SEIB4 = "��"
            Case "2"
                SEIB4 = "��"
            Case "3"
                SEIB4 = "�x"
        End Select
    End Function

    ' @(f)
    '
    ' �@�\�@�@�@: �g�w�i�F�̎w��
    '
    ' �Ԃ�l�@�@: �g�w�i�F RGB�l(&H��16�i���\�L)
    '
    ' �������@�@: strCD - �g�ԕ�����(1byte)
    '
    Public Function CELBK1(ByVal strCD As String) As String
        CELBK1 = String.Empty
        Select Case strCD
            Case "0"
                CELBK1 = "&HFFFFFF"
            Case "1"
                CELBK1 = "&HFFFFFF"
            Case "2"
                CELBK1 = "&H010000"
            Case "3"
                CELBK1 = "&HFF0000"
            Case "4"
                CELBK1 = "&H0000FF"
            Case "5"
                CELBK1 = "&HFFFF00"
            Case "6"
                CELBK1 = "&H00FF00"
            Case "7"
                CELBK1 = "&HFF8000"
            Case "8"
                CELBK1 = "&HFF8080"
        End Select
    End Function

    ' @(f)
    '
    ' �@�\�@�@�@: �����w�i�F�̎w��
    '
    ' �Ԃ�l�@�@: �����w�i�F RGB�l(&H��16�i���\�L)
    '
    ' �������@�@: strCD - ����������(1byte)
    '
    Public Function CELBK2(ByVal strCD As String) As String
        CELBK2 = String.Empty
        Select Case strCD
            Case "01"
                CELBK2 = "&HFFCCCC"
            Case "02"
                CELBK2 = "&HFFCC80"
            Case "03"
                CELBK2 = "&HCCFFFF"
        End Select
    End Function

    ' @(f)
    '
    ' �@�\�@�@�@: �g�����F�̎w��
    '
    ' �Ԃ�l�@�@: �g�����F ������
    '
    ' �������@�@: strCD - �g�ԕ�����(1byte)
    '
    Public Function CELFK(ByVal strCD As String) As String
        CELFK = String.Empty
        Select Case strCD
            Case "0"
                CELFK = ""
            Case "1"
                CELFK = ""
            Case "2"
                CELFK = "White"
            Case "3"
                CELFK = "White"
            Case "4"
                CELFK = "White"
            Case "5"
                CELFK = ""
            Case "6"
                CELFK = ""
            Case "7"
                CELFK = ""
            Case "8"
                CELFK = ""
        End Select
    End Function

    ' @(f)
    '
    ' �@�\�@�@�@: �g���b�N�R�[�h �� ���� ���Ǖ�����ɂ���
    '
    ' �Ԃ�l�@�@: ���� ������
    '
    ' �������@�@: strCD - �g���b�N�R�[�h������(2byte)
    '
    Public Function TRCK4(ByVal strCD As String) As String
        TRCK4 = String.Empty
        Select Case strCD
            Case "00"
                TRCK4 = ""
            Case "10" To "22"
                TRCK4 = "��"
            Case "23" To 26, "29"
                TRCK4 = "�_"
            Case "27", "28"
                TRCK4 = "��"
            Case "51" To "59"
                TRCK4 = "��"
        End Select
    End Function

    ' @(f)
    '
    ' �@�\�@�@�@: �g���b�N�R�[�h �� ���� ���Ǖ�����ɂ���
    '
    ' �Ԃ�l�@�@: ���� ������
    '
    ' �������@�@: strCD - �g���b�N�R�[�h������(2byte)
    '
    Public Function TRCK5(ByVal strCD As String) As String
        TRCK5 = String.Empty
        Select Case strCD
            Case "00"
                TRCK5 = ""
            Case "10"
                TRCK5 = "�Œ�"
            Case "11" To "16"
                TRCK5 = "�ō�"
            Case "17" To "22"
                TRCK5 = "�ŉE"
            Case "23", "25"
                TRCK5 = "�_��"
            Case "24", "26"
                TRCK5 = "�_�E"
            Case "27"
                TRCK5 = "����"
            Case "28"
                TRCK5 = "���E"
            Case "29"
                TRCK5 = "�_��"
            Case "51" To "59"
                TRCK5 = "��Q"
        End Select
    End Function

    ' @(f)
    '
    ' �@�\�@�@�@: �n���ԃR�[�h �� ���� ���Ǖ�����ɂ���
    '
    ' �Ԃ�l�@�@: ���� ������
    '
    ' �������@�@: strCD - �n���ԃR�[�h������(1byte)
    '
    Public Function BBJT4(ByVal strCD As String) As String
        BBJT4 = String.Empty
        Select Case strCD
            Case "0"
                BBJT4 = ""
            Case "1"
                BBJT4 = "��"
            Case "2"
                BBJT4 = "�c"
            Case "3"
                BBJT4 = "�d"
            Case "4"
                BBJT4 = "�s"
        End Select
    End Function

    ' @(f)
    '
    ' �@�\�@�@�@: �f�[�^�敪 ���Ǖ�����ɂ���
    '
    ' �Ԃ�l�@�@: �敪 ������
    '
    ' �������@�@: strCD - �f�[�^�敪������(1byte)
    '
    Public Function DTKB1(ByVal strCD As String) As String
        DTKB1 = String.Empty
        Select Case strCD
            Case "1"
                DTKB1 = "�o���n���\(�ؗj)"
            Case "2"
                DTKB1 = "�o�n�\(���E�y�j)"
            Case "3"
                DTKB1 = "���񐬐�(3���܂Ŋm��)"
            Case "4"
                DTKB1 = "���񐬐�(5���܂Ŋm��)"
            Case "5"
                DTKB1 = "���񐬐�(�S�n�����m��)"
            Case "6"
                DTKB1 = "���񐬐�(�S�n����+��Ű�ʉߏ�)"
            Case "7"
                DTKB1 = "����(���j)"
            Case "A"
                DTKB1 = "�n�����n"
            Case "B"
                DTKB1 = "�C�O���ۃ��[�X"
            Case "9"
                DTKB1 = "���[�X���~"
            Case "0"
                DTKB1 = "�Y�����R�[�h�폜"
        End Select
    End Function

    ' @(f)
    '
    ' �@�\�@�@�@: �v���t�@�C������̃f�[�^�擾
    '
    ' �Ԃ�l�@�@: �v���t�@�C���f�[�^
    '
    ' �������@�@: strAppName - �Z�N�V������
    '            strKeyName - �L�[��
    '            strFileName - �v���t�@�C����
    '
    Public Function GetProfileDataStr(ByVal strAppName As String, ByVal strKeyName As String, ByVal strFileName As String) As String

        Dim iReturnCode As String
        Dim sb As StringBuilder = New StringBuilder(1024)

        ' �������ǂݏo��
        iReturnCode = GetPrivateProfileString(strAppName, strKeyName, "default", sb, sb.Capacity, strFileName)

        GetProfileDataStr = sb.ToString

    End Function

    ' @(f)
    '
    ' �@�\�@�@�@: �v���t�@�C���ւ̃f�[�^�ݒ�
    '
    ' �������@�@: strAppName - �Z�N�V������
    '            strKeyName - �L�[��
    '            strValue - �ݒ�l
    '            strFileName - �v���t�@�C����
    '
    Public Sub WriteProfileDataStr(ByVal strAppName As String, ByVal strKeyName As String, ByVal strValue As String, ByVal strFileName As String)

        WritePrivateProfileString(strAppName, strKeyName, strValue, strFileName)

    End Sub


    ' @(f)
    '
    ' �@�\�@�@�@: �f�[�^�x�[�X�ڑ����s��
    '
    ' �Ԃ�l�@�@: ��������(True-����I��, False-�ُ�I��)
    '
    Public Function ConnectDB() As Boolean
        On Error GoTo ErrorHandler

        Dim bReturnCode As Boolean
        bReturnCode = False

        '�ڑ�������

        ' �f�[�^�x�[�X�Ƃ̐ڑ����s���B
        gCon = New ADODB.Connection()
        Dim strPath As String

        ' �f�[�^�x�[�X�̃I�[�v��
        gCon.Open(strConnectString)

        bReturnCode = True

ExitHandler:
        ConnectDB = bReturnCode
        Exit Function

ErrorHandler:
        System.Diagnostics.Debug.WriteLine(Err.Description)
        bReturnCode = False

        MsgBox(Err.Description)
        Resume ExitHandler

    End Function

End Module
