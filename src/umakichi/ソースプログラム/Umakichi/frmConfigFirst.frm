VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmConfigFirst 
   BorderStyle     =   3  '�Œ��޲�۸�
   Caption         =   "�f�[�^�Z�b�g�A�b�v�̐ݒ�"
   ClientHeight    =   4770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7470
   Icon            =   "frmConfigFirst.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   7470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows �̊���l
   Begin VB.CommandButton cmdCancel 
      Caption         =   "�L�����Z��(&C)"
      Default         =   -1  'True
      Height          =   375
      Left            =   4050
      TabIndex        =   2
      Top             =   4320
      Width           =   1965
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "�f�[�^�Z�b�g�A�b�v�J�n(&S)"
      Height          =   375
      Left            =   1860
      TabIndex        =   1
      Top             =   4320
      Width           =   2145
   End
   Begin VB.Frame frmJVLMode 
      Caption         =   "JV-Link �擾���[�h"
      Height          =   4245
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   7395
      Begin VB.PictureBox picXPTheme 
         BorderStyle     =   0  '�Ȃ�
         Height          =   4005
         Left            =   60
         ScaleHeight     =   4005
         ScaleWidth      =   7275
         TabIndex        =   3
         Top             =   180
         Width           =   7275
         Begin VB.OptionButton optJVMode 
            Caption         =   "���T���[�h"
            Height          =   240
            Index           =   1
            Left            =   0
            TabIndex        =   7
            Top             =   2160
            Width           =   1305
         End
         Begin MSComCtl2.UpDown updYear 
            Height          =   315
            Left            =   3240
            TabIndex        =   5
            Top             =   480
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   556
            _Version        =   393216
            OrigLeft        =   1260
            OrigTop         =   750
            OrigRight       =   1410
            OrigBottom      =   1065
            Max             =   3000
            Enabled         =   -1  'True
         End
         Begin VB.CheckBox chkBLOD 
            Caption         =   "�Y��E�ɐB�n���܂߂�"
            Height          =   225
            Left            =   2220
            TabIndex        =   4
            Top             =   240
            Value           =   1  '����
            Width           =   2175
         End
         Begin VB.CheckBox chkSLOP 
            Caption         =   "��H�������܂߂�"
            Height          =   255
            Left            =   390
            TabIndex        =   12
            Top             =   210
            Value           =   1  '����
            Width           =   1995
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
            Height          =   435
            Index           =   2
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   11
            Text            =   "frmConfigFirst.frx":000C
            Top             =   3450
            Width           =   5565
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
            Height          =   885
            Index           =   1
            Left            =   360
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   10
            Text            =   "frmConfigFirst.frx":0097
            Top             =   2430
            Width           =   6225
         End
         Begin VB.OptionButton optJVMode 
            Caption         =   "�ʏ탂�[�h"
            Height          =   225
            Index           =   0
            Left            =   0
            TabIndex        =   9
            Top             =   0
            Value           =   -1  'True
            Width           =   1425
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
            Height          =   1365
            Index           =   0
            Left            =   360
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   8
            Text            =   "frmConfigFirst.frx":01C6
            Top             =   870
            Width           =   6765
         End
         Begin VB.TextBox txtYear 
            Alignment       =   1  '�E����
            Height          =   285
            Left            =   2670
            TabIndex        =   6
            Text            =   "1995"
            Top             =   480
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "�w�肵���N�x�ȍ~�̃f�[�^�݂̂��Z�b�g�A�b�v���܂��B"
            BeginProperty Font 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   8.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   190
            Left            =   3540
            TabIndex        =   14
            Top             =   560
            Width           =   3705
         End
         Begin VB.Label lblFix 
            AutoSize        =   -1  'True
            Caption         =   "�Z�b�g�A�b�v�J�n�N�x�F"
            Height          =   180
            Index           =   4
            Left            =   960
            TabIndex        =   13
            Top             =   510
            Width           =   1695
         End
      End
   End
End
Attribute VB_Name = "frmConfigFirst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'   (��)�f�[�^�擾�ݒ� �_�C�A���O
'
Option Explicit

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �����ϐ�
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private mstrButtonType As String   ' �����ꂽ�{�^���̃^�C�v

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �v���p�e�B
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   �@�\: �{�^���^�C�v��Ԃ�
'
'   ���l: �Ȃ�
'
Public Property Get ButtonType() As String
    ButtonType = mstrButtonType
End Property


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �����֐�
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   �@�\: �L�����Z���{�^���I���C�x���g
'
'   ���l: �Ȃ�
'
Private Sub cmdCancel_Click()
On Error GoTo Errorhandler
    mstrButtonType = "Cancel"
    Unload Me
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �n�j�{�^���I���C�x���g
'
'   ���l: �Ȃ�
'
Private Sub cmdOK_Click()
On Error GoTo Errorhandler
    gApp.R_JVLGetSLOP = (chkSLOP.value = 1)
    gApp.R_JVLGetBLOD = (chkBLOD.value = 1)
    If optJVMode(0).value = True Then
        gApp.R_JVLMode = ukjUsual
    Else
        gApp.R_JVLMode = ukjThisWeek
    End If
    gApp.R_SetupYear = Format$(val(txtYear.Text), "0000")
    mstrButtonType = "OK"
    
    Unload Me
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
    optJVMode(0).value = (gApp.R_JVLMode = ukjUsual)
    optJVMode(1).value = (gApp.R_JVLMode = ukjThisWeek)
    chkSLOP.value = IIf(gApp.R_JVLGetSLOP, 1, 0)
    chkBLOD.value = IIf(gApp.R_JVLGetBLOD, 1, 0)
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   �@�\: JV-Link�擾���[�h�I���C�x���g
'
'   ���l: �Ȃ�
'
Private Sub optJVMode_Click(Index As Integer)
On Error GoTo Errorhandler
    chkSLOP.Enabled = (Index = 0)
    chkBLOD.Enabled = (Index = 0)
    txtYear.Enabled = (Index = 0)
    updYear.Enabled = (Index = 0)
    cmdOK.Caption = IIf(Index = 0, "�f�[�^�Z�b�g�A�b�v�J�n", "�擾�J�n")
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �Z�b�g�A�b�v�J�n�N�x�L�[���̓C�x���g
'
'   ���l: �Ȃ�
'
Private Sub txtYear_KeyPress(KeyAscii As Integer)
On Error GoTo Errorhandler
   If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
      KeyAscii = 0      ' �������������܂��B
      Beep            ' �G���[����炵�܂��B
   End If
   Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �Z�b�g�A�b�v�J�n�N�x���X�g�t�H�[�J�X�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub txtYear_LostFocus()
On Error GoTo Errorhandler
    If val(txtYear.Text) > val(Year(Now)) Then  '���N���傫���Ƃ����N�ɒu������
        txtYear.Text = Year(Now)
    ElseIf val(txtYear.Text) <= 1995 Then       '1995�N�ȑO�̂Ƃ�1995�N�ɒu������
        txtYear.Text = "1995"
    End If
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub

'
'   �@�\: �Z�b�g�A�b�v�J�n�N�x�o���f�C�g�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub txtYear_Validate(Cancel As Boolean)
On Error GoTo Errorhandler
    If val(txtYear.Text) > val(Year(Now)) Then
        txtYear.Text = Year(Now)
    ElseIf val(txtYear.Text) <= 1995 Then
        txtYear.Text = "1995"
    End If
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �Z�b�g�A�b�v�J�n�N�x�_�E���N���b�N�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub updYear_DownClick()
On Error GoTo Errorhandler
    If txtYear.Text > 1995 Then
        txtYear.Text = txtYear.Text - 1
    End If
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �Z�b�g�A�b�v�J�n�N�x�A�b�v�N���b�N�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub updYear_UpClick()
On Error GoTo Errorhandler
    If val(txtYear.Text) < CInt(Year(Now)) Then
        txtYear.Text = txtYear.Text + 1
    End If
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub
