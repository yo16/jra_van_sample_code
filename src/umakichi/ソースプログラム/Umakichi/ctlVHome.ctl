VERSION 5.00
Begin VB.UserControl ctlVHome 
   BackColor       =   &H00E0EEEE&
   ClientHeight    =   6480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8340
   ScaleHeight     =   6480
   ScaleMode       =   0  'հ�ް
   ScaleWidth      =   8340
   Begin Umakichi.ctlClickLabel clblCmd 
      Height          =   225
      Index           =   0
      Left            =   570
      TabIndex        =   6
      ToolTipText     =   "�o�n�\��I���A�\�����܂��B"
      Top             =   1500
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   397
      AutoSize        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
      Key             =   "Empty"
      ViewerName      =   "RAKaiSel"
      Caption         =   "�o�n�\"
   End
   Begin VB.Frame fraSub 
      BackColor       =   &H00C0CCCC&
      BorderStyle     =   0  '�Ȃ�
      Caption         =   "Frame1"
      Height          =   345
      Index           =   1
      Left            =   300
      TabIndex        =   4
      Top             =   2640
      Width           =   6915
      Begin VB.Label lblFix 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0CCCC&
         Caption         =   "�}�X�^�[���"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   180
         TabIndex        =   5
         Top             =   60
         Width           =   1245
      End
   End
   Begin VB.Frame fraSub 
      BackColor       =   &H00C0CCCC&
      BorderStyle     =   0  '�Ȃ�
      Caption         =   "Frame1"
      Height          =   345
      Index           =   0
      Left            =   330
      TabIndex        =   2
      Top             =   990
      Width           =   6195
      Begin VB.Label lblFix 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0CCCC&
         Caption         =   "���[�X���"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   180
         TabIndex        =   3
         Top             =   60
         Width           =   1065
      End
   End
   Begin VB.Frame fraTop 
      BackColor       =   &H00C0CCCC&
      BorderStyle     =   0  '�Ȃ�
      Caption         =   "Frame1"
      Height          =   465
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   7800
      Begin VB.Label lblFix 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0CCCC&
         Caption         =   "�������j���["
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   270
         TabIndex        =   1
         Top             =   60
         Width           =   1950
      End
   End
   Begin Umakichi.ctlClickLabel clblCmd 
      Height          =   225
      Index           =   1
      Left            =   570
      TabIndex        =   7
      Top             =   1860
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   397
      AutoSize        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
      Key             =   "Empty"
      ViewerName      =   "TKKaiSel"
      Caption         =   "���ʓo�^�n"
   End
   Begin Umakichi.ctlClickLabel clblCmd 
      Height          =   225
      Index           =   2
      Left            =   2250
      TabIndex        =   8
      Top             =   1860
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   397
      AutoSize        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
      Key             =   "Empty"
      ViewerName      =   "HCSel"
      Caption         =   "��H����"
   End
   Begin Umakichi.ctlClickLabel clblCmd 
      Height          =   225
      Index           =   3
      Left            =   570
      TabIndex        =   9
      Top             =   3150
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   397
      AutoSize        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
      ViewerName      =   "Find"
      Caption         =   "�����n"
   End
   Begin Umakichi.ctlClickLabel clblCmd 
      Height          =   225
      Index           =   4
      Left            =   570
      TabIndex        =   10
      Top             =   3510
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   397
      AutoSize        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
      Key             =   "1"
      ViewerName      =   "Find"
      Caption         =   "�R��"
   End
   Begin Umakichi.ctlClickLabel clblCmd 
      Height          =   225
      Index           =   5
      Left            =   570
      TabIndex        =   11
      Top             =   3870
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   397
      AutoSize        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
      Key             =   "2"
      ViewerName      =   "Find"
      Caption         =   "�����t"
   End
   Begin Umakichi.ctlClickLabel clblCmd 
      Height          =   225
      Index           =   6
      Left            =   2250
      TabIndex        =   12
      Top             =   3150
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   397
      AutoSize        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
      Key             =   "3"
      ViewerName      =   "Find"
      Caption         =   "�n��"
   End
   Begin Umakichi.ctlClickLabel clblCmd 
      Height          =   225
      Index           =   7
      Left            =   2250
      TabIndex        =   13
      Top             =   3510
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   397
      AutoSize        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
      Key             =   "4"
      ViewerName      =   "Find"
      Caption         =   "���Y��"
   End
   Begin Umakichi.ctlClickLabel clblCmd 
      Height          =   225
      Index           =   8
      Left            =   2250
      TabIndex        =   14
      Top             =   3870
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   397
      AutoSize        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
      Key             =   "5"
      ViewerName      =   "Find"
      Caption         =   "�ɐB�n"
   End
   Begin Umakichi.ctlClickLabel clblCmd 
      Height          =   225
      Index           =   9
      Left            =   2280
      TabIndex        =   15
      Top             =   1530
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   397
      AutoSize        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
      ViewerName      =   "RCSel"
      Caption         =   "�R�[�X���R�[�h"
   End
   Begin VB.Menu mnuRight 
      Caption         =   "�E�N���b�N"
      Visible         =   0   'False
      Begin VB.Menu mnuNewWin 
         Caption         =   "�V�����E�C���h�E�ŊJ��"
         Index           =   0
      End
   End
End
Attribute VB_Name = "ctlVHome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'   �z�[����� �z�[���{�^���������ƁA���̉�ʂ��\�������
'   �f�t�H���g�̋N������ʂ����̉�ʂł���
'

Option Explicit

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �O���֐�(�C�x���g)
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Public Event ChangeTo(strViewerName As String, strKey As String)        '' Vierer�ύX�C�x���g
Public Event LinkContextMenu(strViewerName As String, strKey As String) '' �����N�E�N���b�N�C�x���g

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �����ϐ�
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private mViewerState    As clsVSNothing         '' ���
Private mblnNoData As Boolean           '' �f�[�^�����t���O

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �v���p�e�B
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8


'
'   �@�\: �L�[�ݒ�v���p�e�B
'
'   ���l: Viewer�K�{�v���p�e�B
'
Public Property Let Key(strKey As String)
    ' Empty
End Property


'
'   �@�\: �L�[�擾�v���p�e�B
'
'   ���l: Viewer�K�{�v���p�e�B
'
Public Property Get Key() As String
    ' Empty
End Property


'
'   �@�\: �^�C�g���擾�v���p�e�B
'
'   ���l: Viewer�K�{�v���p�e�B�A�@Browser ���Ăт܂�
'
Public Property Get Title() As String
    Title = "�z�[��"
End Property


'
'   �@�\: Viewer��Ԓ�
'
'   ���l: �Ȃ�
'
Public Property Get ViewerState() As clsVSNothing
    Set ViewerState = mViewerState
End Property


'
'   �@�\: Viewer��Ԏ󂯎��
'
'   ���l: �Ȃ�
'
Public Property Set ViewerState(RHS As clsVSNothing)
    Set mViewerState = RHS
End Property


'
'   �@�\: �f�[�^�������u���E�U�ɓ`����
'
'   ���l:�@Viewer�K�{�v���p�e�B
'
Public Property Get NoData() As Boolean
    NoData = mblnNoData
End Property


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �����֐�
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8


'
'   �@�\: ��ʕύX�C�x���g
'
'   ���l: �u���E�U�ɃC�x���g���X���[����
'
Private Sub clblCmd_ChangeViewer(Index As Integer)
On Error GoTo ErrorHandler
    RaiseEvent ChangeTo(clblCmd(Index).ViewerName, clblCmd(Index).Key)
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �E�N���b�N�����C�x���g
'
'   ���l: �u���E�U�̃|�b�v�A�b�v�C�x���g�ɃX���[����
'
Private Sub clblCmd_RightMouseDown(Index As Integer)
On Error GoTo ErrorHandler
    RaiseEvent LinkContextMenu(clblCmd(Index).ViewerName, clblCmd(Index).Key)
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: ���[�U�R���g���[��������
'
'   ���l: �Ȃ�
'
Private Sub UserControl_Initialize()
On Error GoTo ErrorHandler
    gApp.InitLog Me
    
    Dim i As Integer
    
    Set mViewerState = New clsVSNothing

    ' Color Assign
    BackColor = gApp.ColBG
    fraTop.BackColor = gApp.ColDarkBG
    For i = 0 To fraSub.UBound
        fraSub(i).BackColor = gApp.ColDarkBG
    Next i
    
    For i = lblFix.LBound To lblFix.UBound
        lblFix(i).BackColor = gApp.ColDarkBG
        lblFix(i).ForeColor = Contrast(gApp.ColDarkBG)
    Next i
    
    For i = clblCmd.LBound To clblCmd.UBound
        clblCmd(i).BackColor = gApp.ColBG
    Next i
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: ���[�U�R���g���[���̃}�E�X�A�b�v�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrorHandler
    Call UserControl.Parent.ShowPopupMenu(Button, Shift, X, Y)
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: ���[�U�R���g���[���̃��T�C�Y�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub UserControl_Resize()
On Error GoTo ErrorHandler
    With UserControl
        .width = Bigger(8000, .width)
        .Height = Bigger(5000, .Height)
    End With
    
    fraTop.width = ScaleWidth - fraTop.Left * 2
    fraSub(0).width = ScaleWidth - fraSub(0).Left * 2
    fraSub(1).width = ScaleWidth - fraSub(1).Left * 2
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: ���[�U�R���g���[���I���C�x���g
'
'   ���l: �Ȃ�
'
Private Sub UserControl_Terminate()
On Error GoTo ErrorHandler
    gApp.TermLog Me
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �O���֐�
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   �@�\: �I������
'
'   ���l: �Ȃ�
'
Public Sub Free()
End Sub

