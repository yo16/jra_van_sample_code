Attribute VB_Name = "basIni"
'
'   INI�t�@�C���Ɋւ��郂�W���[��
'

Option Explicit


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   API�֐��錾
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   �@�\: �w�肳�ꂽ�������t�@�C�� (.INI �t�@�C��) �́A�w�肳�ꂽ�Z�N�V�������ɂ���A
'         �w�肳�ꂽ�L�[�Ɋ֘A�t�����Ă��镶������擾���܂��B
'         �֐�����������ƁA�o�b�t�@�Ɋi�[���ꂽ���������Ԃ�܂� (�I�[�� NULL �����͊܂܂Ȃ�) �B
'
'   ���l: �Ȃ�
'
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Long, _
    ByVal lpFileName As String) As Long
    

'
'   �@�\: �w�肳�ꂽ�������t�@�C���i.INI�t�@�C���j�́A�w�肳�ꂽ�Z�N�V�������ɁA
'         �w�肳�ꂽ�L�[�Ƃ���Ɋ֘A�t����ꂽ������̃y�A�𕡐��i�[���܂��B
'
'   ���l: �Ȃ�
'
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, _
    ByVal lpString As Any, _
    ByVal lpFileName As String) As Long


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �O���֐�
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   �@�\: INI�t�@�C���f�[�^�̎擾
'
'   ���l: ������ AppName  - ���[�g�L�[
' �@�@�@         KeyName  - �T�u�L�[
' �@�@�@         Default  - �l�̖��O
'                FileName - INI�t�@�C����
'
Public Function GetIniData(AppName As String, KeyName As String, Default As String, filename As String) As String
On Error GoTo ErrorHandler
    Dim str         As String * 1024    '�o�b�t�@
    Dim retuenValue As Long

    GetIniData = ""
    
    'INI�t�@�C���f�[�^�擾 ( ByVal�̎g�p )
    retuenValue = GetPrivateProfileString(AppName, KeyName, Default, ByVal str, 1024, filename)
    If retuenValue > 0 Then
        GetIniData = Left$(str, retuenValue)
    End If
    
    Exit Function
ErrorHandler:
    GetIniData = ""
End Function

'
'   �@�\: INI�t�@�C���f�[�^�̏�������
'
'   ���l: ������ AppName  - �Z�N�V������
' �@�@�@         KeyName  - �L�[��
' �@�@�@         Value    - �l
' �@�@�@         FileName - INI�t�@�C����
'
Public Function SetIniData(AppName As String, KeyName As String, value As String, filename As String) As Boolean
On Error GoTo ErrorHandler
    SetIniData = False

    '�f�[�^����������
    If WritePrivateProfileString(AppName, KeyName, ByVal value, filename) <> 0 Then
        SetIniData = True
    End If
    
    gApp.Log "SetIniData : " & AppName & "." & KeyName & "=" & value & " : " & filename
    
    Exit Function
ErrorHandler:
    gApp.ErrLog
    SetIniData = False
End Function

