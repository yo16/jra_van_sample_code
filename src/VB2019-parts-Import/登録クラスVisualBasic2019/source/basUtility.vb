Option Explicit On 

Imports System.Text

Module basUtility

    Declare Function GetPrivateProfileString Lib "KERNEL32.DLL" Alias "GetPrivateProfileStringA" ( _
     ByVal lpAppName As String, _
     ByVal lpKeyName As String, ByVal lpDefault As String, _
     ByVal lpReturnedString As StringBuilder, ByVal nSize As Integer, _
     ByVal lpFileName As String) As Integer

    Public strConnectString As String

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
        iReturnCode = GetPrivateProfileString(strAppName, strKeyName, "", sb, sb.Capacity, strFileName)

        GetProfileDataStr = sb.ToString

    End Function

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
