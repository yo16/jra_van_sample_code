VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmConfig 
   BorderStyle     =   3  '�Œ��޲�۸�
   Caption         =   "�n�g�ݒ�"
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��ʂ̒���
   Begin TabDlg.SSTab mstTab 
      Height          =   7695
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   7785
      _ExtentX        =   13732
      _ExtentY        =   13573
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "�f�[�^�x�[�X"
      TabPicture(0)   =   "frmConfig.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cmdCancel(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdOK(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraDB"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fraJVL"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "�F�̐ݒ�"
      TabPicture(1)   =   "frmConfig.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "dlgColorChoice"
      Tab(1).Control(1)=   "fraTheme"
      Tab(1).Control(2)=   "fraPreview"
      Tab(1).Control(3)=   "cmdOK(1)"
      Tab(1).Control(4)=   "cmdCancel(1)"
      Tab(1).Control(5)=   "cmdReverse"
      Tab(1).ControlCount=   6
      Begin VB.Frame fraJVL 
         Caption         =   "JV-Link�ݒ�"
         Height          =   705
         Left            =   90
         TabIndex        =   11
         Top             =   330
         Width           =   7575
         Begin VB.PictureBox picXPTheme 
            BorderStyle     =   0  '�Ȃ�
            Height          =   375
            Index           =   1
            Left            =   120
            ScaleHeight     =   375
            ScaleWidth      =   6975
            TabIndex        =   25
            Top             =   240
            Width           =   6975
            Begin VB.CommandButton cmdJVSetUIProperties 
               Caption         =   "JVLink�ݒ�_�C�A���O���Ăяo��"
               BeginProperty Font 
                  Name            =   "�l�r �o�S�V�b�N"
                  Size            =   9
                  Charset         =   128
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   0
               TabIndex        =   26
               Top             =   0
               Width           =   3225
            End
            Begin VB.Label lblFix 
               Caption         =   "�T�[�r�X�L�[�̐ݒ��A����I�v�V�����̕ύX������ꍇ�́A���̃{�^���������Ă��������B"
               BeginProperty Font 
                  Name            =   "�l�r �o�S�V�b�N"
                  Size            =   8.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Index           =   3
               Left            =   3420
               TabIndex        =   27
               Top             =   0
               Width           =   3300
            End
         End
      End
      Begin VB.Frame fraDB 
         Caption         =   "�f�[�^�x�[�X�ݒ�"
         Height          =   6075
         Left            =   90
         TabIndex        =   10
         Top             =   1110
         Width           =   7545
         Begin VB.Frame frmJVLMode 
            Caption         =   "JV-Link �擾���[�h"
            Height          =   4125
            Left            =   120
            TabIndex        =   31
            Top             =   1200
            Width           =   7365
            Begin VB.PictureBox picXPThema 
               BorderStyle     =   0  '�Ȃ�
               Height          =   3795
               Index           =   0
               Left            =   60
               ScaleHeight     =   3795
               ScaleWidth      =   7245
               TabIndex        =   32
               Top             =   240
               Width           =   7245
               Begin VB.OptionButton optJVMode 
                  Caption         =   "�ʏ탂�[�h"
                  BeginProperty Font 
                     Name            =   "�l�r �o�S�V�b�N"
                     Size            =   9
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   0
                  Left            =   30
                  TabIndex        =   39
                  Top             =   0
                  Value           =   -1  'True
                  Width           =   1425
               End
               Begin VB.CheckBox chkBLOD 
                  Caption         =   "�Y��E�ɐB�n���܂߂�"
                  BeginProperty Font 
                     Name            =   "�l�r �o�S�V�b�N"
                     Size            =   9
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   2100
                  TabIndex        =   36
                  Top             =   210
                  Value           =   1  '����
                  Width           =   2355
               End
               Begin VB.TextBox txtFix 
                  BackColor       =   &H8000000F&
                  BorderStyle     =   0  '�Ȃ�
                  BeginProperty Font 
                     Name            =   "MS UI Gothic"
                     Size            =   8.25
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   585
                  Index           =   2
                  Left            =   60
                  Locked          =   -1  'True
                  MultiLine       =   -1  'True
                  TabIndex        =   34
                  Text            =   "frmConfig.frx":0038
                  Top             =   3360
                  Width           =   5565
               End
               Begin VB.OptionButton optJVMode 
                  Caption         =   "���T���[�h"
                  BeginProperty Font 
                     Name            =   "�l�r �o�S�V�b�N"
                     Size            =   9
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Index           =   1
                  Left            =   0
                  TabIndex        =   33
                  Top             =   2280
                  Width           =   1305
               End
               Begin VB.TextBox txtFix 
                  BackColor       =   &H8000000F&
                  BorderStyle     =   0  '�Ȃ�
                  BeginProperty Font 
                     Name            =   "�l�r �o�S�V�b�N"
                     Size            =   8.25
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   1095
                  Index           =   1
                  Left            =   390
                  Locked          =   -1  'True
                  MultiLine       =   -1  'True
                  TabIndex        =   35
                  Text            =   "frmConfig.frx":00C3
                  Top             =   2550
                  Width           =   6225
               End
               Begin VB.TextBox txtFix 
                  BackColor       =   &H8000000F&
                  BorderStyle     =   0  '�Ȃ�
                  BeginProperty Font 
                     Name            =   "�l�r �o�S�V�b�N"
                     Size            =   8.25
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   2025
                  Index           =   0
                  Left            =   450
                  Locked          =   -1  'True
                  MultiLine       =   -1  'True
                  TabIndex        =   38
                  Text            =   "frmConfig.frx":01F2
                  Top             =   510
                  Width           =   6765
               End
               Begin VB.CheckBox chkSLOP 
                  Caption         =   "��H�������܂߂�"
                  BeginProperty Font 
                     Name            =   "�l�r �o�S�V�b�N"
                     Size            =   9
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   270
                  TabIndex        =   37
                  Top             =   210
                  Value           =   1  '����
                  Width           =   2085
               End
               Begin VB.Label lblFix 
                  AutoSize        =   -1  'True
                  BackColor       =   &H000000FF&
                  Caption         =   "�Z�b�g�A�b�v�����܂ŕύX�ł��܂���"
                  BeginProperty Font 
                     Name            =   "�l�r �o�S�V�b�N"
                     Size            =   8.25
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H8000000E&
                  Height          =   165
                  Index           =   1
                  Left            =   4530
                  TabIndex        =   40
                  Top             =   240
                  Width           =   2520
               End
            End
         End
         Begin VB.PictureBox picXPTheme 
            BorderStyle     =   0  '�Ȃ�
            Height          =   345
            Index           =   0
            Left            =   150
            ScaleHeight     =   345
            ScaleWidth      =   7095
            TabIndex        =   21
            Top             =   240
            Width           =   7095
            Begin VB.TextBox txtPath 
               Height          =   300
               Left            =   1920
               TabIndex        =   23
               Text            =   "C:\db\"
               Top             =   0
               Width           =   4485
            End
            Begin VB.CommandButton cmdDirRef 
               Caption         =   "�Q��"
               Height          =   300
               Left            =   6390
               TabIndex        =   22
               Top             =   0
               Width           =   645
            End
            Begin VB.Label lblFix 
               AutoSize        =   -1  'True
               Caption         =   "�f�[�^�x�[�X�t�H���_�F"
               BeginProperty Font 
                  Name            =   "�l�r �o�S�V�b�N"
                  Size            =   9
                  Charset         =   128
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   0
               Left            =   0
               TabIndex        =   24
               Top             =   60
               Width           =   1905
            End
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            Caption         =   "�Z�b�g�A�b�v�J�n�N�x�F xxxx�N"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   9
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   1
            Left            =   2070
            TabIndex        =   42
            Top             =   840
            Width           =   2610
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            Caption         =   "�ŏI�f�[�^�擾����"
            BeginProperty Font 
               Name            =   "�l�r �S�V�b�N"
               Size            =   9
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   0
            Left            =   2070
            TabIndex        =   41
            Top             =   600
            Width           =   1620
         End
      End
      Begin VB.CommandButton cmdReverse 
         Caption         =   "�F���t�ɂ���"
         Height          =   375
         Left            =   -70950
         TabIndex        =   9
         Top             =   4800
         Width           =   2175
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "�L�����Z��"
         Height          =   375
         Index           =   1
         Left            =   -70950
         TabIndex        =   8
         Top             =   5430
         Width           =   1905
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Height          =   375
         Index           =   1
         Left            =   -72930
         TabIndex        =   7
         Top             =   5430
         Width           =   1905
      End
      Begin VB.Frame fraPreview 
         Caption         =   "�v���r���["
         Height          =   3255
         Left            =   -70950
         TabIndex        =   4
         Top             =   1440
         Width           =   2175
         Begin VB.Label lblFGDisp 
            BackColor       =   &H80000009&
            Caption         =   "�@�@�@�@�@�@�@�@�@�O�i"
            BeginProperty Font 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   690
            Left            =   360
            TabIndex        =   6
            Top             =   1200
            Width           =   1455
         End
         Begin VB.Label lblBGDisp 
            BackColor       =   &H80000007&
            Caption         =   "�@�@�@�@�@�@�@�@�@�@�@�@�w�i"
            BeginProperty Font 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   2655
            Left            =   120
            TabIndex        =   5
            Top             =   360
            Width           =   1935
         End
      End
      Begin VB.Frame fraTheme 
         Caption         =   "�F�̃e�[�} "
         Height          =   3735
         Left            =   -74430
         TabIndex        =   3
         Top             =   1440
         Width           =   3255
         Begin VB.PictureBox picXPThema 
            BorderStyle     =   0  '�Ȃ�
            Height          =   3285
            Index           =   1
            Left            =   240
            ScaleHeight     =   3285
            ScaleWidth      =   2775
            TabIndex        =   12
            Top             =   360
            Width           =   2775
            Begin VB.OptionButton optTheme 
               Caption         =   "�n���E�B��"
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   20
               Top             =   0
               Width           =   1575
            End
            Begin VB.OptionButton optTheme 
               Caption         =   "�N�[���~���g"
               Height          =   255
               Index           =   1
               Left            =   0
               TabIndex        =   19
               Top             =   360
               Width           =   1575
            End
            Begin VB.OptionButton optTheme 
               Caption         =   "�H"
               Height          =   255
               Index           =   2
               Left            =   0
               TabIndex        =   18
               Top             =   720
               Width           =   1575
            End
            Begin VB.OptionButton optTheme 
               Caption         =   "���̂悤�ɐԂ�"
               Height          =   255
               Index           =   3
               Left            =   0
               TabIndex        =   17
               Top             =   1080
               Width           =   1575
            End
            Begin VB.OptionButton optTheme 
               Caption         =   "�v���e�B�[�s���N"
               Height          =   255
               Index           =   4
               Left            =   0
               TabIndex        =   16
               Top             =   1440
               Width           =   1725
            End
            Begin VB.OptionButton optTheme 
               Caption         =   "�f�B�t�H���g"
               Height          =   255
               Index           =   5
               Left            =   0
               TabIndex        =   15
               Top             =   1800
               Width           =   1575
            End
            Begin VB.OptionButton optTheme 
               Caption         =   "�J�X�^�}�C�Y"
               Height          =   255
               Index           =   6
               Left            =   0
               TabIndex        =   14
               Top             =   2160
               Width           =   2055
            End
            Begin VB.Frame fraAdvanced 
               Caption         =   "�J���[�J�X�^�}�C�Y"
               Height          =   660
               Left            =   120
               TabIndex        =   13
               Top             =   2520
               Width           =   2475
               Begin VB.PictureBox picXPTheme 
                  BorderStyle     =   0  '�Ȃ�
                  Height          =   345
                  Index           =   2
                  Left            =   150
                  ScaleHeight     =   345
                  ScaleWidth      =   2265
                  TabIndex        =   28
                  Top             =   240
                  Width           =   2265
                  Begin VB.CommandButton cmdColor 
                     Caption         =   "�O�i�F"
                     Height          =   300
                     Index           =   1
                     Left            =   1200
                     TabIndex        =   30
                     Top             =   0
                     Width           =   975
                  End
                  Begin VB.CommandButton cmdColor 
                     Caption         =   "�w�i�F"
                     Height          =   300
                     Index           =   0
                     Left            =   0
                     TabIndex        =   29
                     Top             =   0
                     Width           =   975
                  End
               End
            End
         End
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Height          =   375
         Index           =   0
         Left            =   1950
         TabIndex        =   2
         Top             =   7230
         Width           =   1905
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "�L�����Z��"
         Height          =   375
         Index           =   0
         Left            =   3930
         TabIndex        =   1
         Top             =   7230
         Width           =   1905
      End
      Begin MSComDlg.CommonDialog dlgColorChoice 
         Left            =   -74880
         Top             =   420
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'   ���ݒ���
'
Option Explicit

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �����ϐ�
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private mblnNoEdit As Boolean       '' ���ҏW�t���O

Private mlngBGColor As Long         '' �w�i�F
Private mlngFGColor As Long         '' �O�ʐF
Private mlngPrevBGColor As Long     '' �ύX�O�̔w�i�F
Private mlngPrevFGColor As Long     '' �ύX�O�̑O�ʐF

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �O���֐�
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8


'
'   �@�\: �f�[�^�x�[�X�̕s������Ԃ�
'
'   ���l: �Ȃ�
'
Public Function MissingDBNum() As Long
    Dim SpecIDs()   As String
    Dim dbpath      As String
    Dim fso         As FileSystemObject
    Dim count       As Long
    Dim i           As Long

    gApp.Log "FSO New"
    Set fso = New FileSystemObject
    gApp.Log "FSO New finish"
    dbpath = gApp.R_DBPath

    SpecIDs = Split("BANUSI BATAIJYU CHOKYO CHOKYO_SEISEKI HANRO HANSYOKU HARAI KISHU KISHU_CHANGE KISHU_SEISEKI MINING ODDS_SANREN0 ODDS_SANREN1 ODDS_SANREN2 ODDS_SANREN3 ODDS_SANREN4 ODDS_SANREN5 ODDS_SANREN6 ODDS_SANREN7 ODDS_SANREN8 ODDS_SANREN9 ODDS_TANPUKUWAKU ODDS_UMAREN ODDS_UMATAN0 ODDS_UMATAN1 ODDS_UMATAN2 ODDS_UMATAN3 ODDS_UMATAN4 ODDS_UMATAN5 ODDS_UMATAN6 ODDS_UMATAN7 ODDS_UMATAN8 ODDS_UMATAN9 ODDS_WIDE RACE RECORD SANKU SCHEDULE SEISAN TENKO_BABA TOKU TOKU_RACE TORIKESI_JYOGAI UMA UMA_RACE_A UMA_RACE_B")

    For i = 0 To UBound(SpecIDs)
        If Not fso.FileExists(dbpath & "\sub" & SpecIDs(i) & ".mdb") Then
            count = count + 1
        End If
    Next i

    If count = UBound(SpecIDs) + 1 Then
        count = -1
    End If

    MissingDBNum = count
End Function


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �����֐�
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   �@�\: BLOD�f�[�^��������������
'
'   ���l: �Ȃ�
'
Private Sub SetAndSyncBLOD()
On Error GoTo Errorhandler
    Dim DBUpdate As frmDBUpdate
    Dim MsgResult As VbMsgBoxResult

    gApp.Log "Config BLOD option: " & chkBLOD.value
    If gApp.R_JVLGetBLOD = False And chkBLOD.value = 1 And gApp.R_JVDLastTime <> String$(14, "0") Then
        If gApp.R_JVDLastTimeBLOD = String$(14, "0") Then
            MsgResult = MsgBox("�Y��ɐB�n�f�[�^�͂܂��Z�b�g�A�b�v����Ă��܂���B" & vbCrLf & _
                        "�Y��ɐB�n�f�[�^�̃Z�b�g�A�b�v���J�n���܂����H", vbYesNo + vbQuestion, "�n�g�F�Y��ɐB�n�f�[�^�Z�b�g�A�b�v�̊m�F")
        Else
            MsgResult = MsgBox("�s�����Ă��镪�̎Y��ɐB�n�f�[�^���擾���܂��B", vbOKCancel + vbInformation, "�n�g�F�Y��ɐB�n�f�[�^���擾���̎擾�J�n�̊m�F")
        End If
        
        If MsgResult = vbCancel Then
            chkBLOD.value = 0
        Else
            Set DBUpdate = New frmDBUpdate
            DBUpdate.GettingMode = 2 '0:Other 1:SLOP 2:BLOD 3:O6H6 else:err
            DBUpdate.Show vbModal, Me
            If DBUpdate.Finish Then
                MsgResult = MsgBox("����̍X�V����́A�Y��ɐB�n�f�[�^���܂܂��悤�ɂȂ�܂����B", vbInformation, "�n�g�F�ݒ�ύX�̊m�F")
                gApp.R_JVLGetBLOD = True
            Else
                chkBLOD.value = 0
            End If
        End If
    End If
    gApp.R_JVLGetBLOD = (chkBLOD.value = 1)
    Exit Sub
Errorhandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   �@�\: �f�[�^��������������
'
'   ���l: �Ȃ�
'
Private Sub SetAndSyncSLOP()
On Error GoTo Errorhandler
    Dim DBUpdate As frmDBUpdate
    Dim MsgResult As VbMsgBoxResult

    gApp.Log "Config SLOP option: " & chkSLOP.value
    If gApp.R_JVLGetSLOP = False And chkSLOP.value = 1 And gApp.R_JVDLastTime <> String$(14, "0") Then
        If gApp.R_JVDLastTimeSLOP = String$(14, "0") Then
            MsgResult = MsgBox("��H�����f�[�^�͂܂��Z�b�g�A�b�v����Ă��܂���B" & vbCrLf & _
                        "��H�����f�[�^�̃Z�b�g�A�b�v���J�n���܂����H", vbYesNo + vbQuestion, "�n�g�F��H�����f�[�^�Z�b�g�A�b�v�̊m�F")
        Else
            MsgResult = MsgBox("�s�����Ă��镪�̍�H�����f�[�^���擾���܂��B", vbOKCancel + vbInformation, "�n�g�F��H�����f�[�^�Z�b�g�A�b�v�̊m�F")
        End If
        
        If MsgResult = vbCancel Then
            chkSLOP.value = 0
        Else
            Set DBUpdate = New frmDBUpdate
            DBUpdate.GettingMode = 1 '0:Other 1:SLOP 2:BLOD 3:O6H6 else:err
            DBUpdate.Show vbModal, Me
            If DBUpdate.Finish Then
                MsgResult = MsgBox("����̍X�V����́A��H�����f�[�^���܂܂��悤�ɂȂ�܂����B", vbInformation, "�n�g�F��H�����f�[�^")
                gApp.R_JVLGetBLOD = True
            Else
                chkSLOP.value = 0
            End If
        End If
    End If
    gApp.R_JVLGetSLOP = (chkSLOP.value = 1)
    Exit Sub
Errorhandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   �@�\: �L�����Z���{�^���N���b�N�C�x���g
'
'   ���l: �ҏW���Ȃɂ��Ȃ���Ă��Ȃ���Α��I���A�Ȃ���Ă���Ίm�F
'
Private Sub cmdCancel_Click(Index As Integer)
On Error GoTo Errorhandler
    If Index = 1 Then
        Call SetDisplayColors(mlngPrevBGColor, mlngPrevFGColor)
    End If
    
    gApp.R_DBPath = txtPath.tag
    
    Unload Me
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �J���[�J�X�^�}�C�Y�{�^���C�x���g
'
'   ���l: �Ȃ�
'
Private Sub cmdColor_Click(Index As Integer)
On Error GoTo Errorhandler
    dlgColorChoice.CancelError = True
    dlgColorChoice.Flags = cdlCCRGBInit '�t���O�̐ݒ�
    dlgColorChoice.color = IIf(Index = 0, lblBGDisp.BackColor, lblFGDisp.BackColor) '�_�C�A���O�̏����F��ݒ�
    dlgColorChoice.ShowColor
    If Err.Number = 32755 Then Exit Sub
    Select Case Index
    Case 0:
        mlngBGColor = dlgColorChoice.color
    Case 1:
        mlngFGColor = dlgColorChoice.color
    End Select
    Call SetDisplayColors(mlngBGColor, mlngFGColor)

    Exit Sub
Errorhandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   �@�\: �Q�ƃ{�^���C�x���g
'
'   ���l: �Ȃ�
'
Private Sub cmdDirRef_Click()
On Error GoTo Errorhandler
    On Error GoTo Errorhandler
    Dim FolderSelectionDialog As frmDirRef
    
    Set FolderSelectionDialog = New frmDirRef
    With FolderSelectionDialog
        .BeginingPath = txtPath.Text
        .Show vbModal, Me
        txtPath.Text = .ReturnPath
    End With
    Call DBChange
    
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   �@�\: JV-Link�ݒ�_�C�A���O�{�b�N�X�̕\��
'
'   ���l: �Ȃ�
'
Private Sub cmdJVSetUIProperties_Click()
On Error GoTo Errorhandler
    Dim ReturnCode As Long              ''JVLink�Ԓl
    Dim JVlink As frmWrappedJVLink
    
    'JVLink�ݒ��ʕ\��
    Set JVlink = New frmWrappedJVLink
    
    On Error Resume Next
    Load JVlink
    If Err.Number <> "0" Then
        gApp.ErrLog
        MsgBox "JV-Link���C���X�g�[������Ă��܂���B", vbOKOnly + vbExclamation, "�n�g�FJV-Link�G���["
        Exit Sub
    End If
    On Error GoTo Errorhandler
    
    ReturnCode = JVlink.axJVLink.JVSetUIProperties()

    '�G���[����
    If ReturnCode <> 0 Then         ''�G���[
        Call MsgBox("JV-Link�̐ݒ肪�ł��܂���ł����B(" & ReturnCode & ")", vbOKOnly + vbCritical, "�n�g�FJV-Link�G���[")
    End If

    Exit Sub
Errorhandler:
    gApp.ErrLog
    Resume Next
    Exit Sub
End Sub


'
'   �@�\: OK�{�^���N���b�N�C�x���g
'
'   ���l: ���W�X�g���ɏ�������ŏI��
'
Private Sub cmdOK_Click(Index As Integer)
On Error GoTo Errorhandler
    Dim i As Long
    Dim fso As New FileSystemObject
    
    Dim f1 As Folder
    Dim f2 As Folder
    Dim cmdb As New clsCreateMDB
    
    Dim result As Long      '' MsgBox�̖߂�l
    Dim DirRef As frmDirRef '' �t�H���_�Q�ƃ_�C�A���O
    
    Set DirRef = New frmDirRef
    DirRef.Message = "�f�[�^�x�[�X�t�H���_���w�肵�Ă�������"

    gApp.R_DBPath = fso.GetFolder(txtPath.Text).Path
    
    Do While MissingDBNum <> 0

        ' �f�[�^�x�[�X�����邩���ׂ�
        If MissingDBNum = -1 Then
            result = MsgBox(gApp.R_DBPath & "�ɁA�f�[�^�x�[�X��V�K�쐬���܂����H" _
                            , vbYesNoCancel + vbQuestion, "�n�g�F�V�K�f�[�^�x�[�X�쐬�̊m�F")
            If result = vbYes Then
                If cmdb.createMDB = False Then
                    ' �쐬���s�Ȃ�t�H���_�I��
                    DirRef.BeginingPath = gApp.R_DBPath
                    DirRef.Show vbModal
                    
                    gApp.R_DBPath = DirRef.ReturnPath
                End If
            ElseIf result = vbCancel Then
                ' ���ɖ߂�
                gApp.R_DBPath = txtPath.tag
                Exit Sub
            ElseIf result = vbNo Then
                ' �Ȃ���΁A�t�H���_�I��
                DirRef.BeginingPath = gApp.R_DBPath
                DirRef.Show vbModal
                
                gApp.R_DBPath = DirRef.ReturnPath
            End If
        Else
            MsgBox "�f�[�^�x�[�X�����Ă��܂��B���̃t�H���_���w�肵�Ă��������B", vbExclamation, "�n�g:�G���["
            ' �Ȃ���΁A�t�H���_�I��
            DirRef.BeginingPath = gApp.R_DBPath
            DirRef.Show vbModal
            
            gApp.R_DBPath = DirRef.ReturnPath
        End If
    Loop
    
    gApp.R_JVLMode = IIf(optJVMode(0).value, ukJVLMode.ukjUsual, ukJVLMode.ukjThisWeek)

    gApp.R_BackColorDark = mlngFGColor
    gApp.R_BackColorLight = mlngBGColor

    Call SetAndSyncSLOP
    Call SetAndSyncBLOD

    ' ���ҏW��Ԃɂ��ăA�����[�h�iQueryUnload���ɖ��ҏW�Ŗ����ꍇ�j���m�F�����邽�߁j
    mblnNoEdit = True
    
    Unload Me
    Exit Sub
Errorhandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   �@�\: �F���t�ɂ���{�^���C�x���g
'
'   ���l: �Ȃ�
'
Private Sub cmdReverse_Click()
On Error GoTo Errorhandler
    Call SetDisplayColors(mlngFGColor, mlngBGColor)
    mlngBGColor = lblBGDisp.BackColor
    mlngFGColor = lblFGDisp.BackColor
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �t�H�[���̏�����
'
'   ���l: ���W�X�g�������R���g���[���ɐݒ肷��
'
Private Sub Form_Load()
On Error GoTo Errorhandler
    Me.Icon = LoadResPicture(100, vbResIcon)
    mblnNoEdit = True ' ���ҏW���
    
    ' DB�p�X
    txtPath.Text = gApp.R_DBPath
    ' tag�v���p�e�B�ɏ����l��ۑ�����
    txtPath.tag = gApp.R_DBPath
    
    Call setDBStatus
    
    mlngPrevBGColor = gApp.R_BackColorLight
    mlngPrevFGColor = gApp.R_BackColorDark
    mlngBGColor = gApp.R_BackColorLight
    mlngFGColor = gApp.R_BackColorDark

    Call SetDisplayColors(mlngBGColor, mlngFGColor)
    
    fraAdvanced.Enabled = False
    cmdColor(0).Enabled = False
    cmdColor(1).Enabled = False
    
    Call ColorThemeCheck
    
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �c�a�̏�Ԃ�ݒ肷��
'
'   ���l: �Ȃ�
'
Private Sub setDBStatus()
On Error GoTo Errorhandler
    ' JV-Link�擾���[�h
    optJVMode(0).value = (gApp.R_JVLMode = ukjUsual)
    optJVMode(1).value = (gApp.R_JVLMode = ukjThisWeek)
    
    ' ��H�A����
    chkSLOP.value = IIf(gApp.R_JVLGetSLOP, 1, 0)
    chkBLOD.value = IIf(gApp.R_JVLGetBLOD, 1, 0)
    lblFix(1).Visible = (gApp.R_SetupCancelLastTime <> "")
    If gApp.R_SetupCancelLastTime <> "" Then
        chkSLOP.Enabled = False
        chkBLOD.Enabled = False
    Else
        chkSLOP.Enabled = True
        chkBLOD.Enabled = True
    End If
    
    If gApp.R_JVDLastTime <> "00000000000000" Then
        lblInfo(0).Caption = "�ŏI�f�[�^�擾�����F " _
            & Format$(gApp.R_JVDLastTime, "@@@@/@@/@@ @@:@@:@@")
    ElseIf gApp.R_SetupCancelLastTime <> "" Then
        lblInfo(0).Caption = "�Z�b�g�A�b�v���f��"
    Else
        lblInfo(0).Caption = "���擾"
    End If
    
    If gApp.R_JVDLastTime = "00000000000000" And gApp.R_SetupCancelLastTime = "" Then
        lblInfo(1).Caption = ""
    ElseIf gApp.R_SetupYear > 0 Then
        lblInfo(1).Caption = "�Z�b�g�A�b�v�J�n�N�x�F " & gApp.R_SetupYear & "�N"
    Else
        lblInfo(1).Caption = "�Z�b�g�A�b�v�J�n�N�x�F �S��"
    End If
    
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �t�H�[���A�����[�h�L�����Z���m�F
'
'   ���l: �ҏW��j�����ďI�����邩�ǂ����BOK�{�^���ŏI������΁A�j���m�F�͏o�Ȃ��B
'
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo Errorhandler
If Not mblnNoEdit Then
        If MsgBox("�ݒ�̕ύX��j�����A�ύX�O�ɂ��ǂ��āA�ݒ���I�����܂���", vbYesNo + vbQuestion, "�n�g�F�ύX�j���̊m�F") = vbNo Then
            Cancel = 1
        End If
    End If
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �\���F��ݒ肷��
'
'   ���l: �Ȃ�
'
Private Sub SetDisplayColors(BGColor As Long, fgcolor As Long)
    lblBGDisp.BackColor = BGColor
    lblBGDisp.ForeColor = Contrast(BGColor)
    lblFGDisp.BackColor = fgcolor
    lblFGDisp.ForeColor = Contrast(fgcolor)
End Sub


'
'   �@�\: JV-Link�擾���[�h�ݒ�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub optJVMode_Click(Index As Integer)
On Error GoTo Errorhandler
    chkSLOP.Enabled = (optJVMode(0).value And gApp.R_SetupCancelLastTime = "")
    chkBLOD.Enabled = (optJVMode(0).value And gApp.R_SetupCancelLastTime = "")
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �F�̃e�[�}�ݒ�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub optTheme_Click(Index As Integer)
On Error GoTo Errorhandler
    Dim lngBack As Long
    Dim lngFore As Long
    
    Select Case Index
        Case 0: ' �n���E�B��
            lngBack = RGB(238, 102, 51)
            lngFore = RGB(1, 1, 1)
        Case 1: ' �N�[���~���g
            lngBack = RGB(151, 255, 255)
            lngFore = RGB(0, 205, 205)
        Case 2: ' �H
            lngBack = RGB(255, 204, 0)
            lngFore = RGB(255, 69, 0)
        Case 3: ' ���̂悤�ɐԂ�
            lngBack = RGB(238, 51, 51)
            lngFore = RGB(153, 1, 1)
        Case 4: ' �v���[�e�B�[�s���N
            lngBack = RGB(255, 182, 193)
            lngFore = RGB(240, 128, 128)
        Case 5: ' �f�B�t�H���g (UK Default)
            lngBack = RGB(238, 238, 224) ' &HE0EEEE
            lngFore = RGB(204, 204, 192) ' &HC0CCCC
        Case 6: ' �J�X�^�}�C�Y
            fraAdvanced.Enabled = True
            cmdColor(0).Enabled = True
            cmdColor(1).Enabled = True
    End Select
    If optTheme(Index).value = True And Index <> 6 Then
        fraAdvanced.Enabled = False
        cmdColor(0).Enabled = False
        cmdColor(1).Enabled = False
        mlngBGColor = lngBack
        mlngFGColor = lngFore
    End If
    Call SetDisplayColors(mlngBGColor, mlngFGColor)
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �F�̃e�[�}���`�F�b�N
'
'   ���l: �Ȃ�
'
Private Sub ColorThemeCheck()
On Error GoTo Errorhandler
    Dim Index As Long
    Dim BG As Long
    Dim fg As Long
    
    BG = gApp.R_BackColorLight
    fg = gApp.R_BackColorDark
    
    If BG = RGB(238, 102, 51) And _
       fg = RGB(1, 1, 1) Then
        Index = 0 ' �n���E�B��
    ElseIf BG = RGB(151, 255, 255) And _
       fg = RGB(0, 205, 205) Then
        Index = 1  ' �N�[���~���g
    ElseIf BG = RGB(255, 204, 0) And _
       fg = RGB(255, 69, 0) Then
        Index = 2  ' �H
    ElseIf BG = RGB(238, 51, 51) And _
       fg = RGB(153, 1, 1) Then
        Index = 3  ' ���̂悤�ɐԂ�
    ElseIf BG = RGB(255, 182, 193) And _
       fg = RGB(240, 128, 128) Then
        Index = 4  ' �v���[�e�B�[�s���N
    ElseIf BG = &HE0EEEE And _
       fg = &HC0CCCC Then
        Index = 5  ' �f�B�t�H���g (UK Default)
    Else
        Index = 6  ' �J�X�^�}�C�Y
        fraAdvanced.Enabled = True
        cmdColor(0).Enabled = True
        cmdColor(1).Enabled = True
    End If
    optTheme(Index).value = True
    
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �f�[�^�x�[�X�t�H���_�̎Q�Ƃ��`�F�b�N
'
'   ���l: �Ȃ�
'
Private Sub txtPath_Validate(Cancel As Boolean)
On Error GoTo Errorhandler
    Cancel = Not DBChange
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �f�[�^�x�[�X�t�H���_��I������
'
'   ���l: �Ȃ�
'
Private Function DBChange() As Boolean
On Error GoTo Errorhandler
    Dim fso As New FileSystemObject
    
    Dim cmdb As New clsCreateMDB
    
    Dim result As VbMsgBoxResult    '' MsgBox�̖߂�l
    Dim DirRef As frmDirRef         '' �t�H���_�Q�ƃ_�C�A���O
    
    DBChange = False
    
    Set DirRef = New frmDirRef
    DirRef.Message = "�f�[�^�x�[�X�t�H���_���w�肵�Ă�������"
    
    If Not fso.FolderExists(txtPath.Text) Then
        result = MsgBox("�t�H���_������܂���B�쐬���܂����H", vbYesNo + vbQuestion, "�n�g�F�t�H���_�쐬�̊m�F")
        If result = vbYes Then
            On Error Resume Next
                fso.CreateFolder (txtPath.Text)
                If Err.Number <> 0 Then
                    MsgBox "�t�H���_���쐬�ł��܂���ł����B", vbExclamation, "�n�g�F�G���["
                    txtPath.Text = txtPath.tag
                    DBChange = False
                    Exit Function
                End If
            On Error GoTo Errorhandler
        ElseIf result = vbNo Then
            txtPath.Text = txtPath.tag
            DBChange = False
            Exit Function
        End If
    End If
    
    gApp.R_DBPath = fso.GetFolder(txtPath.Text).Path
    ' ���W�X�g���ɕۑ�
    Do While MissingDBNum <> 0

        ' �f�[�^�x�[�X�����邩���ׂ�
        If MissingDBNum = -1 Then
            result = MsgBox(gApp.R_DBPath & "�ɁA�f�[�^�x�[�X��V�K�쐬���܂����H" _
                            , vbYesNoCancel + vbQuestion, "�n�g�F�V�K�f�[�^�x�[�X�쐬�̊m�F")
            If result = vbYes Then
                Call cmdb.createMDB
            ElseIf result = vbCancel Then
                ' ���ɖ߂�
                gApp.R_DBPath = txtPath.tag
                DBChange = False
                Exit Function
            ElseIf result = vbNo Then
                ' �Ȃ���΁A�t�H���_�I��
                DirRef.BeginingPath = gApp.R_DBPath
                DirRef.Show vbModal
                
                gApp.R_DBPath = DirRef.ReturnPath
            End If
        Else
            MsgBox "�f�[�^�x�[�X�����Ă��܂��B���̃t�H���_���w�肵�Ă��������B", vbExclamation, "�n�g:�G���["
            ' �Ȃ���΁A�t�H���_�I��
            DirRef.BeginingPath = gApp.R_DBPath
            DirRef.Show vbModal
            
            gApp.R_DBPath = DirRef.ReturnPath
        End If
    Loop
    
    Call setDBStatus
    
    DBChange = True
    
    Exit Function
Errorhandler:
    gApp.ErrLog
End Function
