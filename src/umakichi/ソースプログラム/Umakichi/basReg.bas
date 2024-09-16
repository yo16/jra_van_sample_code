Attribute VB_Name = "basReg"
'
'   ���W�X�g���Ɋւ��郂�W���[��
'

Option Explicit

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �O���萔
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Const HKEY_CLASSES_ROOT = &H80000000     '�t�@�C���̊֘A�t��
Public Const HKEY_CURRENT_CONFIG = &H80000005   '
Public Const HKEY_CURRENT_USER = &H80000001     '���ݎg�p���Ă��郆�[�U�[�̐ݒ�
Public Const HKEY_DYN_DATA = &H80000006         '
Public Const HKEY_LOCAL_MACHINE = &H80000002    '�����̃��[�U�[�ɋ��ʂ̐ݒ�
Public Const HKEY_PERFORMANCE_DATA = &H80000004 '
Public Const HKEY_USERS = &H80000003            '

Public Const REG_DWORD = 4  '������
Public Const REG_SZ = 1     '����

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   API�֐��錾
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   �@�\: ���W�X�g���̃L�[���J��
'
'   ���l: �Ȃ�
'
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" ( _
    ByVal hKey As Long, _
    ByVal lpSubKey As String, _
    phkResult As Long) As Long


'
'   �@�\: ���W�X�g���̃L�[���쐬
'
'   ���l: �Ȃ�
'
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" ( _
    ByVal hKey As Long, _
    ByVal lpSubKey As String, _
    phkResult As Long) As Long


'
'   �@�\: ���W�X�g���f�[�^�̎擾
'
'   ���l: �Ȃ�
'
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" ( _
    ByVal hKey As Long, _
    ByVal lpValueName As String, _
    ByVal lpReserved As Long, _
    lpType As Long, _
    lpData As Any, _
    lpcbData As Long) As Long


'
'   �@�\: ���W�X�g���f�[�^����������
'
'   ���l: �Ȃ�
'
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" ( _
    ByVal hKey As Long, _
    ByVal lpValueName As String, _
    ByVal Reserved As Long, _
    ByVal dwType As Long, _
    lpData As Any, _
    ByVal cbData As Long) As Long


'
'   �@�\: ���W�X�g���L�[�����
'
'   ���l: �Ȃ�
'
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �O���֐�
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   �@�\: ���W�X�g���f�[�^�̎擾
'
'   ���l: ������ RootKey - ���[�g�L�[
' �@�@�@         SubKey  - �T�u�L�[
' �@�@�@         ValueName   - �l�̖��O
'
Public Function GetRegData(RootKey As Long, SubKey As String, ValueName As String) As String
On Error GoTo ErrH

    Dim hnd         As Long             '�L�[�n���h��
    Dim rtype       As Long             '�^�C�v
    Dim rlen        As Long             '�o�b�t�@�̒���
    Dim str         As String * 1024    '�o�b�t�@
    Dim regret      As Long             '�߂�l
    
    GetRegData = ""
    '���W�X�g���L�[���J��
    If RegOpenKey(RootKey, SubKey, hnd) = 0 Then
        rlen = 1024
        '���W�X�g���f�[�^�擾 ( ByVal�̎g�p )
        If RegQueryValueEx(hnd, ValueName, 0, rtype, ByVal str, rlen) = 0 Then
            GetRegData = Left$(str, InStr(str, Chr(0)) - 1)
        End If
    End If
    '���W�X�g���L�[�����
    regret = RegCloseKey(hnd)
    
    Exit Function
    
ErrH:
    GetRegData = ""
End Function


'
'   �@�\: ���W�X�g���f�[�^�̏�������
'
'   ���l: ������ RootKey - ���[�g�L�[
' �@�@�@         SubKey  - �T�u�L�[
' �@�@�@         ValueName   - �l�̖��O
' �@�@�@         Value - �������ޒl�̃f�[�^
'         �Ԃ�l True:����I��, False:�ُ�I��
'
Public Function SetRegData(RootKey As Long, SubKey As String, ValueName As String, value As String) As Boolean
On Error GoTo ErrH

    Dim hnd         As Long
    Dim regret      As Long
    
    SetRegData = False
    '���W�X�g���L�[���J���A�L�[���Ȃ���΍쐬
    If RegCreateKey(RootKey, SubKey, hnd) = 0 Then
        '�f�[�^����������
        If RegSetValueEx(hnd, ValueName, 0, REG_SZ, ByVal value, LenB(value)) = 0 Then
            SetRegData = True
        End If
    End If
    '���W�X�g���L�[�����
    regret = RegCloseKey(hnd)
    
    Exit Function
    
ErrH:
    SetRegData = False
End Function

