VERSION 5.00
Object = "{2AB17740-0C41-11D7-916F-0003479BEB3F}#1.18#0"; "JVDTLab.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDBUpdate 
   BorderStyle     =   3  '�Œ��޲�۸�
   Caption         =   "�f�[�^�X�V"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '��Ű ̫�т̒���
   Begin VB.Timer tmrRefresh 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   4440
      Top             =   1560
   End
   Begin VB.Timer tmrStartTrigger 
      Enabled         =   0   'False
      Left            =   4860
      Top             =   1560
   End
   Begin VB.Frame fraInfo 
      Height          =   1065
      Left            =   90
      TabIndex        =   7
      Top             =   1440
      Width           =   3735
      Begin VB.Label lblType 
         AutoSize        =   -1  'True
         Height          =   180
         Left            =   1470
         TabIndex        =   13
         Top             =   780
         Width           =   60
      End
      Begin VB.Label lblFix 
         AutoSize        =   -1  'True
         Caption         =   "�擾�����"
         Height          =   180
         Index           =   5
         Left            =   120
         TabIndex        =   12
         Top             =   780
         Width           =   900
      End
      Begin VB.Label lblFinish 
         AutoSize        =   -1  'True
         Height          =   180
         Left            =   1470
         TabIndex        =   11
         Top             =   510
         Width           =   60
      End
      Begin VB.Label lblPass 
         AutoSize        =   -1  'True
         Height          =   180
         Left            =   1470
         TabIndex        =   10
         Top             =   240
         Width           =   60
      End
      Begin VB.Label lblFix 
         AutoSize        =   -1  'True
         Caption         =   "�o�ߎ���"
         Height          =   180
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   720
      End
      Begin VB.Label lblFix 
         AutoSize        =   -1  'True
         Caption         =   "�\�z�c�莞��"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   510
         Width           =   1035
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "�L�����Z��"
      Height          =   285
      Left            =   3960
      TabIndex        =   1
      Top             =   2130
      Width           =   1365
   End
   Begin MSComctlLib.ProgressBar prgPart 
      Height          =   195
      Left            =   90
      TabIndex        =   0
      Top             =   1200
      Width           =   5205
      _ExtentX        =   9181
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.ProgressBar prgAll 
      Height          =   195
      Left            =   90
      TabIndex        =   2
      Top             =   690
      Width           =   5205
      _ExtentX        =   9181
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.ProgressBar prgDown 
      Height          =   195
      Left            =   90
      TabIndex        =   14
      Top             =   210
      Width           =   5205
      _ExtentX        =   9181
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Timer tmrOptimize 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   4020
      Top             =   1560
   End
   Begin VB.Label lblFix 
      AutoSize        =   -1  'True
      Caption         =   "�_�E�����[�h�i�s��"
      Height          =   180
      Index           =   6
      Left            =   90
      TabIndex        =   16
      Top             =   0
      Width           =   1650
   End
   Begin VB.Label lblPercentDown 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   5040
      TabIndex        =   15
      Top             =   0
      Width           =   60
   End
   Begin JVDTLabLibCtl.JVLink axJVLink 
      Left            =   4320
      OleObjectBlob   =   "frmDBUpdate.frx":0000
      Top             =   2100
   End
   Begin VB.Label lblPercentPart 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   5040
      TabIndex        =   6
      Top             =   990
      Width           =   60
   End
   Begin VB.Label lblPercentAll 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   5040
      TabIndex        =   5
      Top             =   480
      Width           =   60
   End
   Begin VB.Label lblFix 
      AutoSize        =   -1  'True
      Caption         =   "�ǂݍ��݌ʐi�s��"
      Height          =   180
      Index           =   4
      Left            =   90
      TabIndex        =   4
      Top             =   990
      Width           =   1800
   End
   Begin VB.Label lblFix 
      AutoSize        =   -1  'True
      Caption         =   "�ǂݍ��ݐi�s��"
      Height          =   180
      Index           =   3
      Left            =   90
      TabIndex        =   3
      Top             =   480
      Width           =   1440
   End
End
Attribute VB_Name = "frmDBUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'   �f�[�^�X�V����
'
'   Other, SLOP, BLOD ,O6H6�̎l��ނ̃��[�h������B
'
Option Explicit

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �����ϐ�
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private mlngGettingMode As Long                  '' BLOD SLOP O6H6�P�ƃ��[�h

Private mblnFinishFlag  As Boolean              '' �I���t���O ����I�����ɐ^�ƂȂ�
Private mblnCancelFlag  As Boolean              '' ���f�t���O ���f���擾���[�v�ɓ`����
Private mblnAfterJVOpen As Boolean              '' JVOpen����I����ɐ^�i�ăZ�b�g�A�b�v�L�����Z���s�j

' �œK�������֘A
Private WithEvents mAsyncCN As ADODB.Connection '' �œK���p�񓯊��R�l�N�V����
Attribute mAsyncCN.VB_VarHelpID = -1
Private mblnOptimizeFinish  As Boolean
Private mstrJobList()       As String
Private mlngJobCount        As Long
Private mstrTargetMDB       As String
Private mstrWorkMDB         As String
Private mstrTableName       As String

' ��ʕ\���X�e�[�^�X�p
Private mstrFree        As String
Private mstrType        As String
Private mlngPercentPart As Long                 '' �p�[�Z���g�i�ʁj
Private mvarStartTime   As Variant              '' �ǂݍ��݊J�n����

Private mlngReadCount       As Long             '' �ǂݍ��ރt�@�C����
Private mlngFileCount       As Long             '' �����t�@�C����
Private mlngDownloadCount   As Long             '' �_�E�����[�h���K�v�ȃt�@�C����
Private mdblReadedByte      As Double           '' �S�̂̓ǂݍ��ݍς݃o�C�g��
Private mlngReadedBytePart  As Long             '' �ʂ̓ǂݍ��ݍς݃o�C�g��
Private mlngCountParSpec    As Long             '' ���R�[�h��ʖ��J�E���^

Private mAV As clsImportAV
Private mBN As clsImportBN
Private mBR As clsImportBR
Private mCH As clsImportCH
Private mDM As clsImportDM
Private mHC As clsImportHC
Private mHN As clsImportHN
Private mHR As clsImportHR
Private mJC As clsImportJC
Private mKS As clsImportKS
Private mO1 As clsImportO1
Private mO2 As clsImportO2
Private mO3 As clsImportO3
Private mO4 As clsImportO4
Private mO5 As clsImportO5
Private mRA As clsImportRA
Private mRC As clsImportRC
Private mSE As clsImportSE
Private mSK As clsImportSK
Private mTK As clsImportTK
Private mUM As clsImportUM
Private mWE As clsImportWE
Private mWH As clsImportWH
Private mYS As clsImportYS
Private mODDS As clsImportODDS
Private mTC As clsImportTC
Private mCC As clsImportCC

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �v���p�e�B
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   �@�\: �f�[�^�X�V���[�h��ݒ肷��
'
'   ���l: �Ȃ�
'
Public Property Let GettingMode(RHS As Long)
    Select Case RHS
    Case 0
        gApp.Log "Set DBUpdateMode: Other"
    Case 1
        gApp.Log "Set DBUpdateMode: SLOP"
    Case 2
        gApp.Log "Set DBUpdateMode: BLOD"
    Case 3
        gApp.Log "Set DBUpdateMode: O6H6"
    Case Else
        gApp.Log "Set DBUpdateMode: Error"
    End Select
    mlngGettingMode = RHS
End Property


'
'   �@�\: �f�[�^�擾���ʃt���O��Ԃ�
'
'   ���l: ����I�� True, �ُ�I�� False
'
Public Property Get Finish() As Boolean
    Finish = mblnFinishFlag
End Property


'
'   �@�\: JVOpen����I����ɐ^
'
'   ���l: �Ȃ�
'
Public Property Get AfterJVOpen() As Boolean
    AfterJVOpen = mblnAfterJVOpen
End Property


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �O���֐�
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   �@�\: �擾���C�����[�v
'
'   ���l: �Ȃ�
'
Public Function FetchJVData() As Boolean
On Error GoTo Errorhandler
    
    Const lngBuffSize         As Long = "102901"   '' JVRead�p�o�b�t�@
    Const lngFileNameSize     As Long = "256"     '' �t�@�C�����o�b�t�@�̃T�C�Y
    Const strDataSpecSetup    As String = "TOKURACEDIFFYSCH"            '' �Z�b�g�A�b�v
    Const strDataSpecUsual    As String = "TOKURACEDIFFYSCH"            '' �ʏ�擾
    Const strDataSpecThisWeek As String = "TOKURACETCOVRCOV"            '' ���T���[�h
    
    Dim lngReturnCode       As Long                  '' JVLink����̖߂�l
    Dim lngPrevReturnCode   As Long                  '' �O���JVLink����̖߂�l
    Dim strDataSpec         As String                '' JVOpen �f�[�^���
    Dim strFromTime         As String                '' �擾�J�n����
    Dim lngOptionFlag       As Long                  '' JVLink�擾�I�v�V����

    Dim strLastTime         As String                '' JVOpen���Ԃ��擾����
    Dim strFileName         As String                '' JVGets���Ԃ��t�@�C�����p�o�b�t�@
    Dim bytBuff()           As Byte                  '' JVGets�p�o�b�t�@
    Dim ImportObj           As clsIImport            '' Import Object Interface
    Dim strRecordIDOld      As String                '' ���R�[�h���ID
    Dim strRecordIDNew      As String                '' ���R�[�h���ID��ۑ�����
    Dim lngCount            As Long                  '' �S�̃J�E���^
    Dim sngTimerStart       As Single                '' �S�̃^�C�}�[
    Dim sngTimerEnd         As Single                '' �S�̃^�C�}�[�I��
    Dim sngSubTimerStart    As Single                '' ���R�[�h��ʖ��^�C�}�[
    Dim sngSubTimerEnd      As Single                '' ���R�[�h��ʖ��^�C�}�[�I��
    Dim lngRecLen           As Long                  '' ���R�[�h��
    Dim strRegSetupCancelLastTime As String
    Dim strCurrentTimeStamp As String                '' ���݂̃t�@�C���̃^�C���X�^���v

    Dim objMP As clsPointer '' �}�E�X�|�C���^����N���X
    
    Dim objDBRAKaiSel As frmDBRAKaiSel

    Set objMP = New clsPointer
    Call objMP.SetBusyPointer(Me)


    
    '-------
    ' JVInit
    
    lngReturnCode = axJVLink.JVInit(gJVLinkSID)
    If lngReturnCode <> 0 Then
        gApp.Log "JVLink - JVInit�G���["
        MsgBox "JVLink - JVInit�G���[", vbExclamation, "�n�g�FJVLink�G���["
        ' �G���[�I��
        FetchJVData = False
        Exit Function
    End If
    
        
    '--------
    ' JVOpen
    
    ' ��x���擾����Ă��Ȃ��ꍇ�̓Z�b�g�A�b�v���s��
    ' �����łȂ���΁A�ʏ�f�[�^�擾������
    If gApp.R_JVLMode = ukjThisWeek Then
        lngOptionFlag = 2 ' ���T�f�[�^���[�h
        strDataSpec = strDataSpecThisWeek
        strFromTime = gApp.R_JVDLastTimeThisWeek ' �ŏI�擾���Ԃ�����̎擾�J�n���ԂƂ���
        Me.Caption = "���T���[�h�f�[�^�擾"
    Else
        Select Case mlngGettingMode
        Case 0 ' Other���[�h
            If val(Left$(gApp.R_JVDLastTime, 2)) = 0 Then
                lngOptionFlag = 3  '�Z�b�g�A�b�v�p�f�[�^���[�h
                strDataSpec = strDataSpecSetup
                Me.Caption = "�f�[�^�Z�b�g�A�b�v"
                If gApp.R_SetupCancelLastTime <> "" Then
                    MsgBox "�Z�b�g�A�b�v���ĊJ���܂��B" & vbCrLf & strFromTime, vbInformation, "�n�g�F�Z�b�g�A�b�v�ĊJ"
                    Me.Caption = "�f�[�^�Z�b�g�A�b�v�ĊJ"
                End If
                strFromTime = Format$(gApp.R_SetupYear, "0000") & String$(10, "0")
            Else
                lngOptionFlag = 1 ' �ʏ�f�[�^���[�h
                strDataSpec = strDataSpecUsual
                strFromTime = gApp.R_JVDLastTime ' �ŏI�擾���Ԃ�����̎擾�J�n���ԂƂ���
                Me.Caption = "�X�V�f�[�^�擾"
            End If
            strDataSpec = strDataSpec & _
                            IIf(gApp.R_JVLGetSLOP, "SLOP", "") & _
                            IIf(gApp.R_JVLGetBLOD, "BLOD", "")
        Case 1 ' SLOP �P�ƃ��[�h
            If val(Left$(gApp.R_JVDLastTimeSLOP, 2)) = 0 Then
                lngOptionFlag = 3  '�Z�b�g�A�b�v�p�f�[�^���[�h
                strDataSpec = "SLOP"
                strFromTime = Format$(gApp.R_SetupYear, "0000") & String$(10, "0")
                Me.Caption = "��H�����f�[�^�Z�b�g�A�b�v"
            Else
                lngOptionFlag = 1 ' �ʏ�f�[�^���[�h
                strDataSpec = "SLOP"
                strFromTime = gApp.R_JVDLastTimeSLOP ' �ŏI�擾���Ԃ�����̎擾�J�n���ԂƂ���
                Me.Caption = "��H�����f�[�^�X�V"
            End If
        Case 2 ' BLOD �P�ƃ��[�h
            If val(Left$(gApp.R_JVDLastTimeBLOD, 2)) = 0 Then
                lngOptionFlag = 3  '�Z�b�g�A�b�v�p�f�[�^���[�h
                strDataSpec = "BLOD"
                strFromTime = Format$(gApp.R_SetupYear, "0000") & String$(10, "0")
                Me.Caption = "�����E�ɐB�n�f�[�^�Z�b�g�A�b�v"
            Else
                lngOptionFlag = 1 ' �ʏ�f�[�^���[�h
                strDataSpec = "BLOD"
                strFromTime = gApp.R_JVDLastTimeBLOD ' �ŏI�擾���Ԃ�����̎擾�J�n���ԂƂ���
                Me.Caption = "�����E�ɐB�n�f�[�^�X�V"
            End If
        End Select
    End If
    
    
    
    lngReturnCode = axJVLink.JVOpen(strDataSpec, _
                                    strFromTime, _
                                    lngOptionFlag, _
                                    mlngReadCount, _
                                    mlngDownloadCount, _
                                    strLastTime)
    
    Select Case lngReturnCode
        Case 0
            ' ���s
        Case -1
            MsgBox "�f�[�^�͍ŐV�ł��B", vbInformation, "�n�g�F�f�[�^�ŐV"
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
    
    ' �_�E�����[�h�i�s�󋵃v���O���X�o�[�Ƀt�@�C������ݒ肷��
    With prgDown
        .Min = 0
        If mlngDownloadCount > 0 Then '�V�K�_�E�����[�h������Ȃ�
            .max = mlngDownloadCount
        End If
    End With
    
    '-------------------
    ' JVRead�̃��[�v����
    FetchJVData = True
    
    If mlngReadCount > 0 Then
        
        If axJVLink.m_TotalReadFilesize > 0 Then
            prgAll.max = axJVLink.m_TotalReadFilesize
        Else
            MsgBox "m_TotalReadFilesize��0�ȉ��ł��B"
            Exit Function
        End If
        
        ' �X�e�[�^�X�`��J�n
        tmrRefresh.Enabled = True

        
        mlngFileCount = 0
        mvarStartTime = Now
        
        ' �S�̃J�E���^������
        lngCount = 0
        
        ' �������Ԍv���p�@�J�n���Ԑݒ�
        sngTimerStart = Timer
        
        '�o�b�t�@�쐬
        strFileName = String$(lngFileNameSize, vbNullChar)
        
        strRegSetupCancelLastTime = gApp.R_SetupCancelLastTime
        
        Do
            DoEvents    ' �o�b�N�O���E���h����
            
            If mblnCancelFlag Then
                ' OtherMode�̃Z�b�g�A�b�v���Ȃ�A���f�ĊJ�p�^�C���X�^���v���L�^����
                If mlngGettingMode = 0 And gApp.R_JVLMode = ukjUsual And lngOptionFlag = 3 Then
                    gApp.R_SetupCancelLastTime = strCurrentTimeStamp
                End If
                mblnFinishFlag = False ' �ُ�I��
                Exit Do ' �擾���~�{�^���������ꂽ��I��
            End If
            
            ' JVGets��1�s�ǂݍ���
            lngReturnCode = axJVLink.JVGets(bytBuff, lngBuffSize, strFileName)
            
            '���^�[���R�[�h�ɂ�菈���𕪊�
            Select Case lngReturnCode
            Case 0      ' �S�t�@�C���ǂݍ��ݏI��
                ' �L�����Z����ԂłȂ�����
                If lngOptionFlag = 3 Then
                    gApp.R_SetupCancelLastTime = ""
                End If
                
                ' �擾������ۑ�����
                If mlngGettingMode = 0 Then
                    If gApp.R_JVLMode = ukjUsual Then
                        gApp.R_JVDLastTime = strLastTime
                        If gApp.R_JVLGetSLOP = True Then
                            gApp.R_JVDLastTimeSLOP = strLastTime
                        End If
                        If gApp.R_JVLGetBLOD = True Then
                            gApp.R_JVDLastTimeBLOD = strLastTime
                        End If
                    ElseIf gApp.R_JVLMode = ukjThisWeek Then
                        gApp.R_JVDLastTimeThisWeek = strLastTime
                    End If
                ElseIf mlngGettingMode = 1 Then ' SLOP
                    gApp.R_JVDLastTimeSLOP = strLastTime
                    gApp.R_SetupCancelLastTime = "" ' SLOP ���[�h�͒��f���T�|�[�g���Ȃ�
                ElseIf mlngGettingMode = 2 Then ' BLOD
                    gApp.R_JVDLastTimeBLOD = strLastTime
                    gApp.R_SetupCancelLastTime = "" ' BLOD ���[�h�͒��f���T�|�[�g���Ȃ�
                End If
                ' ���[�v����E�o����
                mblnFinishFlag = True ' ����I��
                Exit Do
            Case -1     ' �t�@�C���؂�ւ��
                mlngReadedBytePart = prgPart.max
                Call tmrRefresh_Timer
                
                gApp.Log "�t�@�C���؂�ւ��" & strFileName
                
                mlngFileCount = mlngFileCount + 1
                mlngReadedBytePart = 0
                prgPart.max = axJVLink.m_CurrentReadFilesize
                
                If Not ImportObj Is Nothing Then
                    '�C���|�[�g�I�u�W�F�N�g�ɏI��������������
                    gApp.Log ">Close DB"
                    Call ImportObj.CloseDB
                    gApp.Log "<Close DB"
                    ' DB�œK�����K�v�ȂƂ��œK�����s��
                    Call DBOptimize(strRecordIDOld)
                    '�C���|�[�g�I�u�W�F�N�g�ɊJ�n������������
                    gApp.Log ">Open DB"
                    Call ImportObj.OpenDB
                    gApp.Log "<Open DB"
                End If
                
                        
                       
            Case -3     ' �_�E�����[�h��
                If lngReturnCode <> lngPrevReturnCode Then gApp.Log "�_�E�����[�h��"
                mstrType = "�_�E�����[�h�ҋ@"
            Case -201   ' Init����ĂȂ�
                   MsgBox "JVInit���s���Ă��܂���B", vbExclamation, "�n�g�F�G���["
                Exit Do
            Case -203   ' Open����ĂȂ�
                MsgBox "JVOpen���s���Ă��܂���B", vbExclamation, "�n�g�F�G���["
                Exit Do
            Case -402, -403   ' �_�E�����[�h�����t�@�C�����ُ�
                gApp.Log lngReturnCode & "�_�E�����[�h�����t�@�C�����ُ�" & strFileName
                Do While axJVLink.JVFiledelete(strFileName) <> 0
                    Select Case MsgBox("�_�E�����[�h�����t�@�C�����ُ�ȈׁAJVFileDelete�����݂܂��������s���܂����B", vbCritical + vbAbortRetryIgnore)
                    Case VbMsgBoxResult.vbAbort     ' ���~
                        axJVLink.JVClose
                        Exit Function
                    Case VbMsgBoxResult.vbRetry     ' �Ď{�s
                        gApp.Log "JVFiledelete�̍Ď{�s"
                    Case VbMsgBoxResult.vbIgnore    ' ����
                        Exit Do
                    End Select
                Loop
                gApp.Log "JVFileDelete " & strFileName
                axJVLink.JVClose
                If mlngGettingMode = 0 And val(Left$(gApp.R_JVDLastTime, 2)) = 0 Then
                    lngOptionFlag = 4 ' �Z�b�g�A�b�v�̏ꍇ�A�_�C�A���O���o���Ȃ��Z�b�g�A�b�v���[�h�ɂ���
                End If
                Select Case axJVLink.JVOpen(strDataSpec, _
                                            strFromTime, _
                                            lngOptionFlag, _
                                            mlngReadCount, _
                                            mlngDownloadCount, _
                                            strLastTime)
                Case 0
                    ' ���s
                Case -1
                    MsgBox "�f�[�^�͍ŐV�ł��B", vbInformation, "�n�g�F�f�[�^�ŐV"
                    FetchJVData = True  ' �����ŏI��
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
                
                ' �_�E�����[�h�i�s�󋵃v���O���X�o�[�Ƀt�@�C������ݒ肷��
                With prgDown
                    .Min = 0
                    If mlngDownloadCount > 0 Then '�V�K�_�E�����[�h������Ȃ�
                        .max = mlngDownloadCount
                    End If
                End With
                
                ' ���f�ĊJ�p�^�C���X�^���v���L�^����
                gApp.R_SetupCancelLastTime = strCurrentTimeStamp
                strRegSetupCancelLastTime = gApp.R_SetupCancelLastTime
                
                mlngFileCount = 0
                mvarStartTime = Now
                
                ' �S�̃J�E���^������
                lngCount = 0
                
                ' �������Ԍv���p�@�J�n���Ԑݒ�
                sngTimerStart = Timer

            Case -502   ' �_�E�����[�h���s
                MsgBox "�_�E�����[�h���ɃG���[���������܂����B", vbExclamation, "�n�g�F�G���["
                Exit Do
            Case -503   ' �t�@�C�����Ȃ�
                MsgBox "�t�@�C��������܂���B", vbExclamation, "�n�g�F�G���["
                Exit Do
            Case Is > 0 ' ����ǂݍ���
            

                prgPart.max = axJVLink.m_CurrentReadFilesize
                ' �ĊJ�������A�ʏ�擾��
                strCurrentTimeStamp = axJVLink.m_CurrentFileTimeStamp
                If strRegSetupCancelLastTime > strCurrentTimeStamp Then
                    ' �L�����Z���^�C���X�^���v���Â����̂̓X�L�b�v����
                    mdblReadedByte = mdblReadedByte + axJVLink.m_CurrentReadFilesize
                    mlngReadedBytePart = 0
                    mlngFileCount = mlngFileCount + 1
                    mstrType = "�ĊJ����:" & strCurrentTimeStamp
                    gApp.Log "JVSkip : " & strFileName
                    Call tmrRefresh_Timer
                    axJVLink.JVSkip
                Else
                    '���R�[�h���ID���擾
                    strRecordIDNew = StrConv(LeftB(bytBuff, 2), vbUnicode)
                    mstrType = StrConv(LeftB(bytBuff, 20), vbUnicode)
                    
                        '���R�[�h���ID���ύX���ꂽ�ꍇ�i�܂��͏��������j
                        If strRecordIDNew <> strRecordIDOld Then
                        
                            prgPart.max = axJVLink.m_CurrentReadFilesize
        
                            mstrType = strRecordIDNew
                            
                            ' ���茋�ʂ�\��
                            If mlngCountParSpec <> 0 Then
                                sngSubTimerEnd = Timer
                                gApp.Log "���R�[�h��ʂ� " & strRecordIDOld & " ���� " & strRecordIDNew & "�ɕς��܂���"
                                gApp.Log vbTab & strRecordIDOld & "�����s��:" & CStr(mlngCountParSpec)
                                gApp.Log vbTab & strRecordIDOld & "�ǂݍ��ݎ���:" & CStr(sngSubTimerEnd - sngSubTimerStart) & "�b"
                                If (sngSubTimerEnd - sngSubTimerStart) > 0 Then
                                    gApp.Log vbTab & strRecordIDOld & "�������x:" & CStr(mlngCountParSpec / (sngSubTimerEnd - sngSubTimerStart)) & "�s/�b"
                                End If
                            End If
                            
                            If Not ImportObj Is Nothing Then
                                '�C���|�[�g�I�u�W�F�N�g�ɏI��������������
                                gApp.Log "Close DB"
                                Call ImportObj.CloseDB
                            End If
                            
                            '�C���|�[�g�I�u�W�F�N�g��ύX����
                            Set ImportObj = SelectImportObj(strRecordIDNew)

                            If ImportObj Is Nothing Then
                                ' ���m�̃��R�[�h��ʂ̏ꍇDB��Open���Ȃ�
                                mdblReadedByte = mdblReadedByte + axJVLink.m_CurrentReadFilesize
                                mlngReadedBytePart = 0
                                mlngFileCount = mlngFileCount + 1
                                Call tmrRefresh_Timer
                                gApp.Log strRecordIDNew
                                axJVLink.JVSkip
                            Else
                                gApp.Log "Open DB"
                                Call ImportObj.OpenDB
                            End If

                            '���R�[�h���ID�̕ۑ�
                            strRecordIDOld = strRecordIDNew
                            
                            '�J�n������ۑ�
                            sngSubTimerStart = Timer
        
                            '��ʖ��J�E���^������
                            mlngCountParSpec = 0
                        
                        End If
                        
                        If Not ImportObj Is Nothing Then
                            
                                'DB�ǉ�����
                                Do
                                    If Not ImportObj.Add(bytBuff) Then
                                        gApp.Log "���R�[�h�o�^�̎��s" & StrConv(bytBuff, vbUnicode)
                                        Select Case MsgBox("���R�[�h�̓o�^�Ɏ��s���܂����B" & vbCrLf & StrConv(bytBuff, vbUnicode), vbAbortRetryIgnore + vbQuestion, "�n�g�F�G���[")
                                        Case vbAbort
                                            gApp.Log "���~"
                                            mblnCancelFlag = True
                                            Exit Do
                                        Case vbRetry
                                            gApp.Log "�Ď��s"
                
                                        Case vbIgnore
                                            gApp.Log "����"
                                            Exit Do
                                        End Select
                                    Else
                                        Exit Do
                                    End If
                                Loop
                            
                            '�J�E���g�A�b�v
                            mlngCountParSpec = mlngCountParSpec + 1
                            lngCount = lngCount + 1
                            '�ǂݍ��񂾃o�C�g�����v�l
                            mdblReadedByte = mdblReadedByte + lngReturnCode - 1
                            mlngReadedBytePart = mlngReadedBytePart + lngReturnCode - 1
                        End If
    
                End If ' �ĊJ�������ʏ�擾��
                
                ' �o�b�t�@�̊J��
                Erase bytBuff
            
            Case Else
                gApp.Log "�s���ȃ��^�[���R�[�h" & lngReturnCode
            End Select
            
            '���^�[���R�[�h��ۑ�
            lngPrevReturnCode = lngReturnCode
            

        Loop
        
        sngSubTimerEnd = Timer
        sngTimerEnd = Timer
        gApp.Log "���R�[�h��� " & strRecordIDOld & "�̓ǂݍ��ݏ������I�����܂���"
        gApp.Log vbTab & strRecordIDOld & "�����s��:" & CStr(mlngCountParSpec)
        gApp.Log vbTab & strRecordIDOld & "�ǂݍ��ݎ���:" & CStr(sngSubTimerEnd - sngSubTimerStart) & "�b"
        If (sngSubTimerEnd - sngSubTimerStart) > 0 Then
            gApp.Log vbTab & strRecordIDOld & "�������x:" & CStr(mlngCountParSpec / (sngSubTimerEnd - sngSubTimerStart)) & "�s/�b"
        End If
        
        gApp.Log "���ׂẴf�[�^�ǂݍ��ݏ������I�����܂���"
        gApp.Log vbTab & "�S�̏����s��:" & CStr(lngCount)
        gApp.Log vbTab & "�S�̓ǂݍ��ݎ���:" & CStr(sngTimerEnd - sngTimerStart) & "�b"
        If (sngTimerEnd - sngTimerStart) > 0 Then
            gApp.Log "�S�̏������x:" & CStr(lngCount / (sngTimerEnd - sngTimerStart)) & "�s/�b"
        End If
    End If
    
    ' �X�e�[�^�X�`��I��
    mdblReadedByte = prgAll.max
    Call tmrRefresh_Timer
    tmrRefresh.Enabled = False
    
    '---------
    ' JVClose

    axJVLink.JVClose
    
    
    Set objMP = Nothing

    ' �J�ÑI����ʗp�f�[�^�쐬
    If mlngGettingMode = 0 Then
        Call gApp.DeleteAllRAKaiSelCacheFlags
        Set objDBRAKaiSel = New frmDBRAKaiSel
        objDBRAKaiSel.TargetYear = CStr(Year(Now))
        objDBRAKaiSel.Show vbModal, Me
    End If
    
    MsgBox "�擾���I�����܂����", vbOKOnly + vbInformation, "�n�g�F�擾�I��"

    Exit Function
Errorhandler:
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
On Error GoTo Errorhandler
    If (Not mblnFinishFlag) Then
        If MsgBox("�f�[�^�̎擾�𒆎~���܂����H", vbYesNo + vbQuestion, "�n�g�F�f�[�^�擾���~�̊m�F") = vbYes Then
            gApp.Log "�f�[�^�擾�̃L�����Z��"
            mblnCancelFlag = True
        End If
    End If
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �t�H�[�����[�h�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub Form_Load()
On Error GoTo Errorhandler
    Me.Icon = LoadResPicture(100, vbResIcon)

    mblnFinishFlag = False
    
    prgDown.max = 1
    prgAll.max = 1
    prgPart.max = 1
    lblPass.Caption = ""
    lblFinish.Caption = ""
    lblType.Caption = ""
    
    tmrStartTrigger.Interval = 1
    tmrStartTrigger.Enabled = True

    Set mAV = New clsImportAV
    Set mBN = New clsImportBN
    Set mBR = New clsImportBR
    Set mCH = New clsImportCH
    Set mDM = New clsImportDM
    'Set mH1 = New clsImportH1
    Set mHC = New clsImportHC
    Set mHN = New clsImportHN
    Set mHR = New clsImportHR
    Set mJC = New clsImportJC
    Set mKS = New clsImportKS
    Set mO1 = New clsImportO1
    Set mO2 = New clsImportO2
    Set mO3 = New clsImportO3
    Set mO4 = New clsImportO4
    Set mO5 = New clsImportO5
    Set mRA = New clsImportRA
    Set mRC = New clsImportRC
    Set mSE = New clsImportSE
    Set mSK = New clsImportSK
    Set mTK = New clsImportTK
    Set mUM = New clsImportUM
    Set mWE = New clsImportWE
    Set mWH = New clsImportWH
    Set mYS = New clsImportYS
    Set mODDS = New clsImportODDS
    Set mTC = New clsImportTC
    Set mCC = New clsImportCC
            
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �t�H�[���A�����[�h�m�F�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo Errorhandler
    If UnloadMode <> vbFormCode Then
        If MsgBox("�f�[�^�̎擾�𒆎~���܂����H", vbYesNo + vbQuestion, "�n�g�F�f�[�^�擾���~�̊m�F") = vbYes Then
            gApp.Log "�f�[�^�擾�̃L�����Z��"
            Cancel = True
            mblnCancelFlag = True
        Else
            Cancel = True
        End If
    End If
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   �@�\: ���t���b�V���^�C�}�[�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub tmrRefresh_Timer()
On Error GoTo Errorhandler
    Dim sngPersentAll   As Single  ''
    Dim varPass         As Variant '' �o�ߎ���
    Dim lngJVStatus     As Long
    
    
    ' �o�ߎ���
    varPass = Now - mvarStartTime
    
    ' �i�s�x����
    If prgAll.max <> 0 Then
        sngPersentAll = mdblReadedByte / prgAll.max
    End If
    
    ' �_�E�����[�h��
    lngJVStatus = axJVLink.JVStatus
    If lngJVStatus = -502 Then
        Exit Sub
    ElseIf lngJVStatus > 0 Then
        lblPercentDown.Caption = axJVLink.JVStatus & " / " & mlngDownloadCount
        lblPercentDown.Left = ScaleWidth - lblPercentDown.width - 120
        If prgDown.value < mlngDownloadCount Then
            prgDown.value = axJVLink.JVStatus
        End If
    End If
    

    ' �S�̓ǂݍ��ݐi�s�x
    If axJVLink.m_TotalReadFilesize > 0 Then
        lblPercentAll.Caption = Format$(mdblReadedByte / 1024 / prgAll.max * 100, "##0.0") & " %  " & mlngFileCount & "/" & mlngReadCount
        lblPercentAll.Left = ScaleWidth - lblPercentAll.width - 120
        prgAll.value = Smaller(CLng(mdblReadedByte / 1024), prgAll.max)
    Else
        MsgBox "m_TotalReadFilesize��0�ȉ��ł��B"
    End If
    
    ' �ʓǂݍ��ݐi�s�x
    
    lblPercentPart.Caption = mlngCountParSpec & " Rec  " & Format$(mlngReadedBytePart / prgPart.max * 100, "##0.0") & " %  "
    If prgPart.max >= mlngReadedBytePart Then
        prgPart.value = Smaller(mlngReadedBytePart, prgPart.max)
    Else
        lblPercentPart.Caption = "99.9 %"
        gApp.Log prgPart.value & "," & mlngReadedBytePart
    End If
    lblPercentPart.Left = ScaleWidth - lblPercentPart.width - 120
    
    '�o�ߎ���
    lblPass.Caption = Format$(varPass, "Long Time")
    '�c�莞��
    If mlngFileCount > 0 Then
        lblFinish.Caption = Format$((varPass * mlngReadCount / mlngFileCount) - varPass, "Long Time")
    Else
        lblFinish.Caption = ""
    End If
    lblType.Caption = mstrType
    
    Me.Refresh
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �X�^�[�g�^�C�}�[�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub tmrStartTrigger_Timer()
On Error GoTo Errorhandler
    gApp.Log "�f�[�^�擾�̊J�n"
    tmrStartTrigger.Enabled = False
    If Not FetchJVData() Then
        gApp.Log "�f�[�^�擾�̎��s"
    Else
        gApp.Log "�f�[�^�擾�̐���I��"
    End If

    Unload Me
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �J�ÔN�����擾
'
'   ���l: �Ȃ�
'
Private Sub NoteRTRace()
    Dim cn As ADODB.Connection
    Dim RA As ADODB.Recordset
    Dim strOut As String
    Set RA = New ADODB.Recordset
    
    RA.Open "SELECT [Year]&[MonthDay] AS YMD FROM RACE WHERE [DataKubun] <= '6' GROUP BY [Year]&[MonthDay]", gApp.GetCN_RACE, adOpenForwardOnly, adLockReadOnly
    Do While Not RA.EOF
        strOut = strOut & IIf(strOut = "", "", ",") & RA("YMD")
        RA.MoveNext
    Loop
    RA.Close
    
    gApp.R_RTDates = strOut
End Sub


'
'   �@�\: �C���|�[�g�I�u�W�F�N�g�I��
'
'   ���l: �Ȃ�
'
Private Function SelectImportObj(strRecordID As String) As clsIImport
    Select Case strRecordID
    Case "AV"
        Set SelectImportObj = mAV
    Case "BN"
        Set SelectImportObj = mBN
    Case "BR"
        Set SelectImportObj = mBR
    Case "CH"
        Set SelectImportObj = mCH
    Case "DM"
        Set SelectImportObj = mDM
    Case "HC"
        Set SelectImportObj = mHC
    Case "HN"
        Set SelectImportObj = mHN
    Case "HR"
        Set SelectImportObj = mHR
    Case "JC"
        Set SelectImportObj = mJC
    Case "KS"
        Set SelectImportObj = mKS
    Case "O1"
        Set SelectImportObj = mO1
    Case "O2"
        Set SelectImportObj = mO2
    Case "O3"
        Set SelectImportObj = mO3
    Case "O4"
        Set SelectImportObj = mO4
    Case "O5"
        Set SelectImportObj = mO5
    Case "RA"
        Set SelectImportObj = mRA
    Case "RC"
        Set SelectImportObj = mRC
    Case "SE"
        Set SelectImportObj = mSE
    Case "SK"
        Set SelectImportObj = mSK
    Case "TK"
        Set SelectImportObj = mTK
    Case "UM"
        Set SelectImportObj = mUM
    Case "WE"
        Set SelectImportObj = mWE
    Case "WH"
        Set SelectImportObj = mWH
    Case "YS"
        Set SelectImportObj = mYS
    Case "TC"
        Set SelectImportObj = mTC
    Case "CC"
        Set SelectImportObj = mCC
    Case "H1", "H6", "O6"
        Set SelectImportObj = mODDS
    Case Else
        Set SelectImportObj = Nothing
    End Select
End Function


'
'   �@�\: �c�a�œK���̗v�^�s�v�𔻒�
'
'   ���l: �Ȃ�
'
Private Sub DBOptimize(strRecordID As String)
On Error GoTo Errorhandler
    Dim fso As FileSystemObject
    Dim i   As Long

    gApp.Log "�œK���F�m�F"

    Set fso = New FileSystemObject
    
    ReDim mstrJobList(0)
    
    Select Case strRecordID
    Case "HC"
        mstrJobList = Split("subHANRO.mdb", ",")
    Case "O5"
        mstrJobList = Split("subODDS_SANREN0.mdb,subODDS_SANREN1.mdb,subODDS_SANREN2.mdb,subODDS_SANREN3.mdb,subODDS_SANREN4.mdb,subODDS_SANREN5.mdb,subODDS_SANREN6.mdb,subODDS_SANREN7.mdb,subODDS_SANREN8.mdb,subODDS_SANREN9.mdb", ",")
    Case "RA"
        mstrJobList = Split("subRACE.mdb", ",")
    Case "SE"
        mstrJobList = Split("subUMA_RACE_A.mdb,subUMA_RACE_B.mdb", ",")
    Case "UM"
        mstrJobList = Split("subUMA.mdb", ",")
    Case Else
        gApp.Log "�œK���F�s�v"
        Exit Sub
    End Select
        
    mlngJobCount = 0
    tmrOptimize.Interval = 100
    tmrOptimize.Enabled = False
    mblnOptimizeFinish = False
    On Error Resume Next
    For i = 0 To UBound(mstrJobList)
        ' �W���u���X�g���̃t�@�C���T�C�Y���ǂꂩ1.0GB�ȏ�ł���΁AOptimize�^�C�}�[���n������B
        tmrOptimize.Enabled = tmrOptimize.Enabled Or _
            (fso.GetFile(gApp.R_DBPath & "\" & mstrJobList(i)).Size > CLng(1000) * 1000 * 1000)
    Next i
    On Error GoTo Errorhandler
    
    If tmrOptimize.Enabled = False Then
        gApp.Log "�œK���F�s�v"
        Exit Sub
    End If
    
    ' �œK���I���܂Ń��[�v
    gApp.Log "�œK���F�W���u���[�v�J�n"
    Do While Not mblnOptimizeFinish
        DoEvents

        ' �^�C�}�[����~���Ă����ꍇ�I��
        If tmrOptimize.Enabled = False Then
            Exit Do
        End If
    Loop
    tmrOptimize.Enabled = False
    gApp.Log "�œK���F�W���u���[�v�I��"
    
    If mlngJobCount > 0 Then
        gApp.Log "�œK���F�I��"
        mstrType = "�œK���I��"
        Call tmrRefresh_Timer
    Else
        gApp.Log "�œK���F�s�v"
    End If
    
    Set mAsyncCN = Nothing
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �œK���^�C�}�[�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub tmrOptimize_Timer()
On Error GoTo Errorhandler
    
    Dim cat             As ADOX.Catalog
    Dim strSQL          As String
    Dim i               As Long
    
    If mlngJobCount > UBound(mstrJobList) Then
        ' �W���u�J�E���^���ő�l�𒴂��Ă�����I������
        mblnOptimizeFinish = True
        gApp.Log "�œK���F�S�W���u�I��"
    ElseIf mAsyncCN Is Nothing Then
        ' �񓯊��R�l�N�V�����I�u�W�F�N�g���Ȃ���ΐ�������
        Set mAsyncCN = New ADODB.Connection
        gApp.Log "�œK���F�R�l�N�V�����I�u�W�F�N�g����"
    ElseIf mAsyncCN.State = adStateClosed Then
        ' �R�l�N�V���������Ă���΁A�œK�����J�n����
        Set cat = New ADOX.Catalog
        mstrTargetMDB = gApp.R_DBPath & "\" & mstrJobList(mlngJobCount)
        mstrWorkMDB = gApp.R_DBPath & "\" & "Optimizing" & Timer & mstrJobList(mlngJobCount)
        
        ' ���MDB�쐬
        cat.Create "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & mstrWorkMDB
        
        ' �R�l�N�V�����ڑ�
        mAsyncCN.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                        "Data Source=" & mstrTargetMDB
        
        ' �e�[�u�����擾
        Set cat.ActiveConnection = mAsyncCN
        For i = 0 To cat.Tables.count - 1
            If cat.Tables(i).Type = "TABLE" Then
                mstrTableName = cat.Tables(i).Name
            End If
        Next i
        
        ' �e�[�u���R�s�[SQL��
        strSQL = "SELECT " & mstrTableName & ".* INTO " & mstrTableName & _
                " IN '" & mstrWorkMDB & "'" & _
                " FROM " & mstrTableName
                
        mAsyncCN.Execute strSQL, , adAsyncExecute
        mstrType = "�œK����:" & mstrTargetMDB
        Call tmrRefresh_Timer
        gApp.Log "�œK���F" & mstrTargetMDB
    ElseIf mblnCancelFlag = True Then
        ' �L�����Z���{�^���������ꂽ�ꍇ
        If mAsyncCN.State And adStateExecuting <> 0 Then
            ' ���s���Ȃ�L�����Z������
            mAsyncCN.Cancel
        End If
        mAsyncCN.Close
    End If
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �񓯊��R�l�N�V�������s�����C�x���g
'
'   ���l: �Ȃ�
'
Private Sub mAsyncCN_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
On Error GoTo Errorhandler
    Dim IndexCreator    As clsCreateMDB
    Dim fso             As FileSystemObject
    Dim MsgResult       As VbMsgBoxResult
    
    gApp.Log "�œK���FmAsyncCN_ExecuteComplete " & mstrTargetMDB
    
    Set IndexCreator = New clsCreateMDB
    Set fso = New FileSystemObject
    
    ' �R�l�N�V���������
    mAsyncCN.Close
    
    If Not pError Is Nothing Then
        MsgBox "�œK�����s"
    Else
        ' �C���f�b�N�X�쐬
        gApp.Log "�œK���F�C���f�b�N�X�쐬�J�n"
        mstrType = "�C���f�b�N�X�쐬��:" & mstrWorkMDB
        Call tmrRefresh_Timer

        With IndexCreator
            Set .mConnection = New ADODB.Connection
            Call .mConnection.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & mstrWorkMDB)
            Call CallByName(IndexCreator, "createIndex_" & mstrTableName, VbMethod)
            Call .mConnection.Close
        End With
        gApp.Log "�œK���F�C���f�b�N�X�쐬�I��"
        

        ' �I���W�i���t�@�C���폜�����l�[��
        Do
            On Error Resume Next
            fso.DeleteFile mstrTargetMDB
            If Err.Number = 0 Then
                ' �폜�����Ȃ烊�l�[�����ďI��
                gApp.Log "�œK���F���l�[������"
                fso.MoveFile mstrWorkMDB, mstrTargetMDB
                Exit Do
            Else
                MsgResult = MsgBox("�t�@�C�����폜�ł��܂���B", vbAbortRetryIgnore + vbDefaultButton2 + vbExclamation, "�n�g�F�œK�����̃G���[")
                If MsgResult = vbAbort Then
                    ' ���~�Ȃ�A�œK���Ǝ擾�S�����I������
                    gApp.Log "�œK���F���l�[�����s�@���~"
                    fso.DeleteFile mstrWorkMDB
                    mblnFinishFlag = True
                    mblnCancelFlag = True
                    Exit Do
                ElseIf MsgResult = vbIgnore Then
                    gApp.Log "�œK���F���l�[�����s�@����"
                    ' �����Ȃ�A���̍œK����
                    fso.DeleteFile mstrWorkMDB
                    Exit Do
                End If
                gApp.Log "�œK���F���l�[�����s�@���g���C"
            End If
            On Error GoTo Errorhandler
        Loop
    End If
    
    
    ' �W���u�J�E���^��i�߂�
    mlngJobCount = mlngJobCount + 1
    Exit Sub
Errorhandler:
    gApp.ErrLog
    mlngJobCount = mlngJobCount + 1
End Sub

