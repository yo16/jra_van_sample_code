VERSION 5.00
Object = "{2AB17740-0C41-11D7-916F-0003479BEB3F}#1.18#0"; "JVDTLab.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDBPrompt 
   BorderStyle     =   3  '�Œ��޲�۸�
   Caption         =   "����f�[�^�擾��"
   ClientHeight    =   1395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3825
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   3825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows �̊���l
   Begin MSComctlLib.StatusBar stbBar 
      Align           =   2  '������
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   1110
      Width           =   3825
      _ExtentX        =   6747
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6694
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmrRefresh 
      Interval        =   100
      Left            =   540
      Top             =   30
   End
   Begin VB.Timer tmrStartTrigger 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   30
      Top             =   30
   End
   Begin JVDTLabLibCtl.JVLink axJVLink 
      Left            =   2730
      OleObjectBlob   =   "frmDBPrompt.frx":0000
      Top             =   150
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "����f�[�^���擾���Ă��܂��B���΂炭���҂����������B"
      Height          =   405
      Left            =   780
      TabIndex        =   0
      Top             =   510
      Width           =   2400
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmDBPrompt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'   ����f�[�^�擾 �_�C�A���O
'
Option Explicit

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �����ϐ�
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private mblnFinishFlag As Boolean     '' �I���t���O ����I�����ɐ^
Private mblnCancelFlag As Boolean     '' ���f�t���O ���f���擾���[�v�ɓ`����

Private mMode    As ukPromptMode    '' ����擾���[�h
Private mstrKey  As String          '' JVData�L�[
Private mstrType As String

Private mlngReadByte As Long

Private objMP As clsPointer '' �}�E�X�|�C���^����N���X


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �v���p�e�B
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   �@�\: ����擾���[�h��ݒ肷��
'
'   ���l: �Ȃ�
'
Public Property Let Mode(RHS As ukPromptMode)
    mMode = RHS
End Property

'
'   �@�\: JVLink����擾����f�[�^�̃L�[��ݒ肷��
'
'   ���l: �Ȃ�
'
Public Property Let key(RHS As String)
    mstrKey = RHS
End Property


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �O���֐�
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8


'
'   �@�\: �擾���C�����[�v
'
'   ���l: �Ȃ�
'
Public Function FetchJVData(DataSpec As String) As Boolean
On Error GoTo ErrorHandler

    Const lngBuffSize         As Long = "102901"   '' JVRead�p�o�b�t�@
    Const lngFileNameSize     As Long = "256"     '' �t�@�C�����o�b�t�@�̃T�C�Y

    Dim lngReturnCode       As Long                  '' JVLink����̖߂�l
    Dim lngPrevReturnCode   As Long                  '' �O���JVLink����̖߂�l
    Dim strDataSpec         As String                '' JVOpen �f�[�^���
    Dim strFromTime         As String                '' �擾�J�n����
    Dim lngOptionFlag       As Long                  '' JVLink�擾�I�v�V����

    Dim strLastTime         As String                '' JVOpen���Ԃ��擾����
    Dim strFileName         As String                '' JVRead���Ԃ��t�@�C�����p�o�b�t�@
    Dim strBuff             As String                '' JVRead�p�o�b�t�@
    Dim ImportObj           As clsIImport            '' Import Object Interface
    Dim strRecordIDOld      As String                '' ���R�[�h���ID
    Dim strRecordIDNew      As String                '' ���R�[�h���ID��ۑ�����
    Dim lngFileCount        As Long                  '' �����t�@�C����
    Dim varStartTime        As Variant               '' �ǂݍ��݊J�n����
    Dim lngCount            As Long                  '' �S�̃J�E���^
    Dim sngTimerStart       As Single                '' �S�̃^�C�}�[
    Dim sngTimerEnd         As Single                '' �S�̃^�C�}�[�I��
    Dim lngSubCount         As Long                  '' ���R�[�h��ʖ��J�E���^
    Dim sngSubTimerStart    As Single                '' ���R�[�h��ʖ��^�C�}�[
    Dim sngSubTimerEnd      As Single                '' ���R�[�h��ʖ��^�C�}�[�I��

    Me.Icon = LoadResPicture(100, vbResIcon)
    Set objMP = New clsPointer
    Call objMP.SetBusyPointer(Me)

    '-------
    ' JVInit

    lngReturnCode = axJVLink.JVInit(gJVLinkSID)
    If lngReturnCode <> 0 Then
        gApp.Log "JVLink - JVInit�G���["
        MsgBox "JVLink - JVInit�G���[", vbExclamation, "�n�g�F�G���["
        ' �G���[�I��
        FetchJVData = False
        Exit Function
    End If


    '--------
    ' JVRTOpen

    lngReturnCode = axJVLink.JVRTOpen(DataSpec, mstrKey)

    Select Case lngReturnCode
        Case 0
            ' ���s
        Case -1
            FetchJVData = True  ' �����ŏI��
            axJVLink.JVClose
            Exit Function
        Case -504 ' �����e�i���X��
            ' �����e�i���X���̃��b�Z�[�W��JVLink���\������̂�
            ' �n�g���ł͉����\�����Ȃ��ŏI������B
            FetchJVData = False ' ���s�ŏI��
            axJVLink.JVClose
            Exit Function
        Case Else
            MsgBox "JV-Link�֐ڑ����s���܂����B" & vbCrLf _
                & "JV-Link����̃G���[���b�Z�[�W: " & vbCrLf _
                & ErrMsgJVOpen(lngReturnCode), vbInformation, "�n�g�FJV-Link�G���["
            FetchJVData = False ' ���s�ŏI��
            axJVLink.JVClose
            Exit Function
    End Select


    '-------------------
    ' JVRead�̃��[�v����

    lngFileCount = 0
    varStartTime = Now

    ' �S�̃J�E���^������
    lngCount = 0

    ' �������Ԍv���p�@�J�n���Ԑݒ�
    sngTimerStart = Timer
    
    mlngReadByte = 0

    Do
        DoEvents    ' �o�b�N�O���E���h����

        If mblnCancelFlag Then
            
            FetchJVData = False
            Exit Function ' �擾���~�{�^���������ꂽ��I��
        End If

        
        '�o�b�t�@�쐬
        strBuff = String$(lngBuffSize, vbNullChar)
        strFileName = String$(lngFileNameSize, vbNullChar)

        lngReturnCode = axJVLink.JVRead(strBuff, lngBuffSize, strFileName)
        
        '���^�[���R�[�h�ɂ�菈���𕪊�
        Select Case lngReturnCode
        Case 0      ' �S�t�@�C���ǂݍ��ݏI��
            ' ���[�v����E�o����
            Exit Do
        Case -1     ' �t�@�C���؂�ւ��
            If lngReturnCode <> lngPrevReturnCode Then
                gApp.Log "�t�@�C���؂�ւ��" & strFileName
            End If
            lngFileCount = lngFileCount + 1
        Case -3     ' �_�E�����[�h��
            If lngReturnCode <> lngPrevReturnCode Then gApp.Log "�_�E�����[�h��"
            mstrType = "�_�E�����[�h�ҋ@"
        Case -201   ' Init����ĂȂ�
            MsgBox "JVInit���s���Ă��܂���B", vbExclamation, "�n�g�F�G���["
            Exit Do
        Case -203   ' Open����ĂȂ�
            MsgBox "JVOpen���s���Ă��܂���B", vbExclamation, "�n�g�F�G���["
            Exit Do
        Case -502   ' �_�E�����[�h���s
            MsgBox "�_�E�����[�h���ɃG���[���������܂����B", vbExclamation, "�n�g�F�G���["
            Exit Do
        Case -503   ' �t�@�C�����Ȃ�
            MsgBox "�t�@�C��������܂���B", vbExclamation, "�n�g�F�G���["
            Exit Do
        Case Is > 0 ' ����ǂݍ���
            mlngReadByte = mlngReadByte + lngReturnCode
            
            '���R�[�h���ID���擾
            strRecordIDNew = Left$(strBuff, 2)
            mstrType = Left$(strBuff, 20)


            '���R�[�h���ID���ύX���ꂽ�ꍇ�i�܂��͏��������j
            If strRecordIDNew <> strRecordIDOld Then
                mstrType = strRecordIDNew

                ' ���茋�ʂ�\��
                If lngSubCount <> 0 Then
                    sngSubTimerEnd = Timer
                    gApp.Log "���R�[�h��ʂ� " & strRecordIDOld & " ���� " & strRecordIDNew & "�ɕς��܂���"
                    gApp.Log vbTab & strRecordIDOld & "�����s��:" & CStr(lngSubCount)
                    gApp.Log vbTab & strRecordIDOld & "�ǂݍ��ݎ���:" & CStr(sngSubTimerEnd - sngSubTimerStart) & "�b"
                    If (sngSubTimerEnd - sngSubTimerStart) > 0 Then
                        gApp.Log vbTab & strRecordIDOld & "�������x:" & CStr(lngSubCount / (sngSubTimerEnd - sngSubTimerStart)) & "�s/�b"
                    End If
                End If
                
                If Not ImportObj Is Nothing Then
                    '�C���|�[�g�I�u�W�F�N�g�ɏI��������������
                    Call ImportObj.CloseDB
                End If
                
                '�C���|�[�g�I�u�W�F�N�g��ύX����
                Set ImportObj = SelectImportObj(strRecordIDNew)
                Call ImportObj.OpenDB

                '���R�[�h���ID�̕ۑ�
                strRecordIDOld = strRecordIDNew

                '�J�n������ۑ�
                sngSubTimerStart = Timer

                '��ʖ��J�E���^������
                lngSubCount = 0

            End If

            If Not ImportObj Is Nothing Then

                'DB�ǉ�����
                If Not ImportObj.Add(StrConv(strBuff, vbFromUnicode)) Then
                    gApp.Log "���R�[�h�o�^�̎��s" & strBuff

                End If

                '�J�E���g�A�b�v
                lngSubCount = lngSubCount + 1
                lngCount = lngCount + 1
            End If


        Case Else
            gApp.Log "�s���ȃ��^�[���R�[�h" & lngReturnCode
        End Select

        '���^�[���R�[�h��ۑ�
        lngPrevReturnCode = lngReturnCode

    Loop

    sngSubTimerEnd = Timer
    sngTimerEnd = Timer
    gApp.Log "���R�[�h��� " & strRecordIDOld & "�̓ǂݍ��ݏ������I�����܂���"
    gApp.Log vbTab & strRecordIDOld & "�����s��:" & CStr(lngSubCount)
    gApp.Log vbTab & strRecordIDOld & "�ǂݍ��ݎ���:" & CStr(sngSubTimerEnd - sngSubTimerStart) & "�b"
    If (sngSubTimerEnd - sngSubTimerStart) > 0 Then
        gApp.Log vbTab & strRecordIDOld & "�������x:" & CStr(lngSubCount / (sngSubTimerEnd - sngSubTimerStart)) & "�s/�b"
    End If

    gApp.Log "���ׂẴf�[�^�ǂݍ��ݏ������I�����܂���"
    gApp.Log vbTab & "�S�̏����s��:" & CStr(lngCount)
    gApp.Log vbTab & "�S�̓ǂݍ��ݎ���:" & CStr(sngTimerEnd - sngTimerStart) & "�b"
    If (sngTimerEnd - sngTimerStart) > 0 Then
        gApp.Log "�S�̏������x:" & CStr(lngCount / (sngTimerEnd - sngTimerStart)) & "�s/�b"
    End If

    '---------
    ' JVClose
    
    axJVLink.JVClose

    Set objMP = Nothing

    FetchJVData = True
    Exit Function
ErrorHandler:
    gApp.ErrLog
    axJVLink.JVClose
    FetchJVData = False
End Function


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �����֐�
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8


'
'   �@�\: �L�����Z���{�^���C�x���g
'
'   ���l: �Ȃ�
'
Private Sub cmdCancel_Click()
On Error GoTo ErrorHandler
    Unload Me
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �t�H�[���������C�x���g
'
'   ���l: �Ȃ�
'
Private Sub Form_Initialize()
On Error GoTo ErrorHandler
    mMode = ukpRA
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �t�H�[�����[�h�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub Form_Load()
On Error GoTo ErrorHandler


    If mstrKey = "" Then
        gApp.Log TypeName(Me) & "�L�[���ݒ�"
        mblnFinishFlag = True
        Unload Me
    Else
        mblnFinishFlag = False
    
        tmrStartTrigger.Interval = 1
        tmrStartTrigger.Enabled = True
    End If

    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �t�H�[���A�����[�h�m�F�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrorHandler
    If UnloadMode <> vbFormCode Then
        If MsgBox("�f�[�^�̎擾�𒆎~���܂����H", vbYesNo + vbQuestion, "�n�g�F�f�[�^�擾���~�̊m�F") = vbYes Then
            gApp.Log "�f�[�^�擾�̃L�����Z��"
            axJVLink.JVClose
            Set objMP = Nothing
            Cancel = True
            mblnCancelFlag = True
        Else
            Cancel = True
        End If
    End If
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: ���t���b�V���^�C�}�[�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub tmrRefresh_Timer()
On Error GoTo ErrorHandler
    Dim max         As Long
    Dim persent     As Long
    
    max = axJVLink.m_TotalReadFilesize
    If max <> 0 Then
        persent = mlngReadByte / max
    Else
        persent = 100
    End If
    
    
    ' �v���O���X�o�[�ύX
    If max > 0 Then
        '
    End If
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �擾���C�����[�v
'
'   ���l: �Ȃ�
'
Private Sub tmrStartTrigger_Timer()
On Error GoTo ErrorHandler
    Dim returnFlag As Boolean
    Dim AllKeys()  As String
    Dim i As Long

    Log "�f�[�^�擾�̊J�n"
    tmrStartTrigger.Enabled = False
    Select Case mMode
    
    Case ukpRA ' ���񃌁[�X
        Log "���[�X���ʎ擾"
        returnFlag = returnFlag Or Not FetchJVData("0B12") ' ���[�X����
        If returnFlag = False Then
            Log "�n�̏d�擾"
            returnFlag = returnFlag Or Not FetchJVData("0B11") ' �n�̏d
        End If
        If returnFlag = False Then
            Log "�f�[�^�}�C�j���O�擾"
            returnFlag = returnFlag Or Not FetchJVData("0B13") ' �f�[�^�}�C�j���O
        End If
        If returnFlag = False Then
            Log "�J�Ï��擾"
            returnFlag = returnFlag Or Not FetchJVData("0B14") ' �J�Ï��
        End If
        

    Case ukpOD ' ����I�b�Y
        Log "�I�b�Y�擾"
        returnFlag = returnFlag Or Not FetchJVData("0B30") ' ����I�b�Y(�S�q��)
        If returnFlag = False Then
            Log "�[���擾"
            returnFlag = returnFlag Or Not FetchJVData("0B20") ' ����[��(�S�q��)
        End If
    Case ukpPALLET ' ���j���[�p���b�g
        AllKeys = Split(mstrKey, ",")
                        
        For i = 0 To UBound(AllKeys)
            mstrKey = AllKeys(i)
            Log "RT�擾:" & mstrKey
            If returnFlag = False Then
                Log "���[�X���ʎ擾"
                returnFlag = returnFlag Or Not FetchJVData("0B12") ' ���[�X����
            End If
            If returnFlag = False Then
                Log "�n�̏d�擾"
                returnFlag = returnFlag Or Not FetchJVData("0B11") ' �n�̏d
            End If
            If returnFlag = False Then
                Log "�f�[�^�}�C�j���O�擾"
                returnFlag = returnFlag Or Not FetchJVData("0B13") ' �f�[�^�}�C�j���O
            End If
            If returnFlag = False Then
                Log "�J�Ï��擾"
                returnFlag = returnFlag Or Not FetchJVData("0B14") ' �J�Ï��
            End If
        Next i
    End Select
    
    If returnFlag Then
        gApp.Log "�f�[�^�擾�̎��s"
    Else
        gApp.Log "�f�[�^�擾�̐���I��"
        MsgBox "�擾���I�����܂����", vbOKOnly + vbInformation, "�n�g�F�擾�I��"
    End If
    
    mblnFinishFlag = True
    Unload Me
    Exit Sub
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   �@�\: �C���|�[�g�I�u�W�F�N�g�I������
'
'   ���l: �Ȃ�
'
Private Function SelectImportObj(strRecrodID As String) As clsIImport
    Select Case strRecrodID
    Case "AV"
        Set SelectImportObj = New clsImportAV
    Case "BN"
        Set SelectImportObj = New clsImportBN
    Case "BR"
        Set SelectImportObj = New clsImportBR
    Case "CH"
        Set SelectImportObj = New clsImportCH
    Case "DM"
        Set SelectImportObj = New clsImportDM
    Case "HC"
        Set SelectImportObj = New clsImportHC
    Case "HN"
        Set SelectImportObj = New clsImportHN
    Case "HR"
        Set SelectImportObj = New clsImportHR
    Case "JC"
        Set SelectImportObj = New clsImportJC
    Case "KS"
        Set SelectImportObj = New clsImportKS
    Case "O1"
        Set SelectImportObj = New clsImportO1
    Case "O2"
        Set SelectImportObj = New clsImportO2
    Case "O3"
        Set SelectImportObj = New clsImportO3
    Case "O4"
        Set SelectImportObj = New clsImportO4
    Case "O5"
        Set SelectImportObj = New clsImportO5
    Case "RA"
        Set SelectImportObj = New clsImportRA
    Case "RC"
        Set SelectImportObj = New clsImportRC
    Case "SE"
        Set SelectImportObj = New clsImportSE
    Case "SK"
        Set SelectImportObj = New clsImportSK
    Case "TK"
        Set SelectImportObj = New clsImportTK
    Case "UM"
        Set SelectImportObj = New clsImportUM
    Case "WE"
        Set SelectImportObj = New clsImportWE
    Case "WH"
        Set SelectImportObj = New clsImportWH
    Case "YS"
        Set SelectImportObj = New clsImportYS
    Case "TC"
        Set SelectImportObj = New clsImportTC
    Case "CC"
        Set SelectImportObj = New clsImportCC
    Case "H1", "H6", "O6"
        Set SelectImportObj = New clsImportODDS
    Case Else
        Set SelectImportObj = Nothing
    End Select
End Function


'
'   �@�\: ���O�o�͏���
'
'   ���l: �Ȃ�
'
Private Sub Log(strText As String)
    stbBar.Panels(1).Text = strText
    gApp.Log strText
End Sub
