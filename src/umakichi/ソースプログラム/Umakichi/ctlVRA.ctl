VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.UserControl ctlVRA 
   BackColor       =   &H00E0EEEE&
   ClientHeight    =   6690
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9345
   ScaleHeight     =   6690
   ScaleWidth      =   9345
   Begin VB.Timer tmrKako 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   8310
      Top             =   690
   End
   Begin VB.Frame fraHeader 
      BackColor       =   &H00E0EEEE&
      BorderStyle     =   0  '�Ȃ�
      Caption         =   "fraHeader"
      Height          =   585
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Width           =   8655
      Begin VB.Timer tmrTBS 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   7530
         Top             =   90
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0EEEE&
         Caption         =   "."
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   9
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   8
         Left            =   2835
         TabIndex        =   38
         Top             =   390
         Width           =   105
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0EEEE&
         Caption         =   "."
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   7
         Left            =   2520
         TabIndex        =   11
         Top             =   360
         Width           =   90
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0EEEE&
         Caption         =   "."
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   6
         Left            =   4050
         TabIndex        =   10
         Top             =   210
         Width           =   90
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0EEEE&
         Caption         =   "."
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   5
         Left            =   1980
         TabIndex        =   9
         Top             =   360
         Width           =   90
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0EEEE&
         Caption         =   "."
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   4
         Left            =   1980
         TabIndex        =   8
         Top             =   210
         Width           =   90
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0EEEE&
         Caption         =   "."
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   3
         Left            =   1980
         TabIndex        =   7
         Top             =   30
         Width           =   90
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0EEEE&
         Caption         =   "."
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   2
         Left            =   60
         TabIndex        =   5
         Top             =   30
         Width           =   90
      End
   End
   Begin MSComctlLib.ImageList ilsTbrSmall 
      Left            =   1410
      Top             =   5100
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin TabDlg.SSTab mstTab 
      Height          =   4815
      Left            =   180
      TabIndex        =   2
      Top             =   1230
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   8493
      _Version        =   393216
      Style           =   1
      Tabs            =   7
      TabsPerRow      =   9
      TabHeight       =   520
      BackColor       =   14741230
      TabCaption(0)   =   "��{���"
      TabPicture(0)   =   "ctlVRA.ctx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "paneTab(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "����"
      TabPicture(1)   =   "ctlVRA.ctx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "paneTab(1)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "�ߋ� ��"
      TabPicture(2)   =   "ctlVRA.ctx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "paneTab(2)"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "�}�C�j���O"
      TabPicture(3)   =   "ctlVRA.ctx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "paneTab(3)"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "�����ʐ���"
      TabPicture(4)   =   "ctlVRA.ctx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "paneTab(4)"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "�����^�C��"
      TabPicture(5)   =   "ctlVRA.ctx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "paneTab(5)"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "����"
      TabPicture(6)   =   "ctlVRA.ctx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "paneTab(6)"
      Tab(6).Control(1)=   "hsbSE"
      Tab(6).Control(2)=   "vsbSE"
      Tab(6).ControlCount=   3
      Begin Umakichi.ctlPane paneTab 
         Height          =   1575
         Index           =   0
         Left            =   60
         TabIndex        =   21
         Top             =   360
         Width           =   3585
         _ExtentX        =   6324
         _ExtentY        =   2778
         Begin Umakichi.ctlWrappedGrid flexTab 
            Height          =   795
            Index           =   0
            Left            =   0
            TabIndex        =   27
            Top             =   0
            Width           =   2955
            _ExtentX        =   0
            _ExtentY        =   0
         End
      End
      Begin Umakichi.ctlPane paneTab 
         Height          =   2025
         Index           =   6
         Left            =   -74940
         TabIndex        =   14
         Top             =   360
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   3572
         Begin VB.PictureBox picIPane 
            Appearance      =   0  '�ׯ�
            BorderStyle     =   0  '�Ȃ�
            ForeColor       =   &H80000008&
            Height          =   1875
            Left            =   0
            ScaleHeight     =   1875
            ScaleWidth      =   6555
            TabIndex        =   15
            Top             =   0
            Width           =   6555
            Begin Umakichi.ctlWrappedGrid flexTab 
               Height          =   735
               Index           =   6
               Left            =   0
               TabIndex        =   16
               Top             =   0
               Width           =   2475
               _ExtentX        =   0
               _ExtentY        =   0
            End
            Begin Umakichi.ctlWrappedGrid flexTab 
               Height          =   735
               Index           =   7
               Left            =   0
               TabIndex        =   17
               Top             =   930
               Width           =   2475
               _ExtentX        =   0
               _ExtentY        =   0
            End
            Begin Umakichi.ctlWrappedGrid flexTab 
               Height          =   735
               Index           =   8
               Left            =   2580
               TabIndex        =   18
               Top             =   930
               Width           =   2475
               _ExtentX        =   0
               _ExtentY        =   0
            End
            Begin VB.Label lblFix 
               AutoSize        =   -1  'True
               Caption         =   "����"
               Height          =   180
               Index           =   3
               Left            =   0
               TabIndex        =   20
               Top             =   750
               Width           =   360
            End
            Begin VB.Label lblFix 
               AutoSize        =   -1  'True
               Caption         =   "���b�v�^�C��"
               Height          =   180
               Index           =   2
               Left            =   2580
               TabIndex        =   19
               Top             =   750
               Width           =   915
            End
         End
      End
      Begin VB.HScrollBar hsbSE 
         Height          =   345
         Left            =   -74880
         TabIndex        =   13
         Top             =   2580
         Visible         =   0   'False
         Width           =   7305
      End
      Begin VB.VScrollBar vsbSE 
         Height          =   2325
         Left            =   -67410
         TabIndex        =   12
         Top             =   390
         Visible         =   0   'False
         Width           =   495
      End
      Begin Umakichi.ctlPane paneTab 
         Height          =   2085
         Index           =   1
         Left            =   -74970
         TabIndex        =   22
         Top             =   330
         Width           =   4125
         _ExtentX        =   4736
         _ExtentY        =   1561
         Begin Umakichi.ctlWrappedGrid flexTab 
            Height          =   735
            Index           =   1
            Left            =   0
            TabIndex        =   28
            Top             =   0
            Width           =   2475
            _ExtentX        =   0
            _ExtentY        =   0
         End
      End
      Begin Umakichi.ctlPane paneTab 
         Height          =   2085
         Index           =   2
         Left            =   -74940
         TabIndex        =   23
         Top             =   360
         Width           =   4155
         _ExtentX        =   4736
         _ExtentY        =   1561
         Begin VB.TextBox txtKako 
            Alignment       =   1  '�E����
            Height          =   270
            Left            =   360
            TabIndex        =   29
            Text            =   "5"
            Top             =   0
            Width           =   375
         End
         Begin MSComCtl2.UpDown UpDown1 
            Height          =   270
            Left            =   736
            TabIndex        =   30
            Top             =   0
            Width           =   240
            _ExtentX        =   318
            _ExtentY        =   476
            _Version        =   393216
            Value           =   1
            BuddyControl    =   "txtKako"
            BuddyDispid     =   196621
            OrigLeft        =   750
            OrigRight       =   990
            OrigBottom      =   255
            Max             =   99
            Min             =   1
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin Umakichi.ctlWrappedGrid flexTab 
            Height          =   735
            Index           =   2
            Left            =   0
            TabIndex        =   31
            Top             =   300
            Width           =   2475
            _ExtentX        =   0
            _ExtentY        =   0
         End
         Begin VB.Label lblFix 
            AutoSize        =   -1  'True
            Caption         =   "�ߋ�"
            Height          =   180
            Index           =   0
            Left            =   30
            TabIndex        =   33
            Top             =   45
            Width           =   360
         End
         Begin VB.Label lblFix 
            AutoSize        =   -1  'True
            Caption         =   "��"
            Height          =   180
            Index           =   1
            Left            =   1050
            TabIndex        =   32
            Top             =   45
            Width           =   180
         End
      End
      Begin Umakichi.ctlPane paneTab 
         Height          =   2205
         Index           =   3
         Left            =   -74940
         TabIndex        =   24
         Top             =   360
         Width           =   4185
         _ExtentX        =   4736
         _ExtentY        =   1561
         Begin Umakichi.ctlWrappedGrid flexTab 
            Height          =   735
            Index           =   3
            Left            =   0
            TabIndex        =   34
            Top             =   180
            Width           =   2475
            _ExtentX        =   0
            _ExtentY        =   0
         End
         Begin VB.Label lblDMKubun 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            Caption         =   "�}�C�j���O�敪"
            Height          =   180
            Left            =   0
            TabIndex        =   35
            Top             =   0
            Width           =   1155
         End
      End
      Begin Umakichi.ctlPane paneTab 
         Height          =   1875
         Index           =   4
         Left            =   -74940
         TabIndex        =   25
         Top             =   360
         Width           =   3825
         _ExtentX        =   4736
         _ExtentY        =   1561
         Begin Umakichi.ctlWrappedGrid flexTab 
            Height          =   735
            Index           =   4
            Left            =   0
            TabIndex        =   36
            Top             =   0
            Width           =   2475
            _ExtentX        =   0
            _ExtentY        =   0
         End
      End
      Begin Umakichi.ctlPane paneTab 
         Height          =   1875
         Index           =   5
         Left            =   -74940
         TabIndex        =   26
         Top             =   360
         Width           =   4125
         _ExtentX        =   7276
         _ExtentY        =   3307
         Begin Umakichi.ctlWrappedGrid flexTab 
            Height          =   735
            Index           =   5
            Left            =   0
            TabIndex        =   37
            Top             =   0
            Width           =   2475
            _ExtentX        =   0
            _ExtentY        =   0
         End
      End
   End
   Begin VB.Frame fraTop 
      BackColor       =   &H00C0CCCC&
      BorderStyle     =   0  '�Ȃ�
      Caption         =   "fraTop"
      Height          =   405
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   8955
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "."
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   1
         Left            =   8310
         TabIndex        =   6
         Top             =   120
         Width           =   30
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0CCCC&
         Caption         =   "���\�[�X�s���ł��B�s�v�ȉ�ʂ���Ă�������"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   180
         TabIndex        =   1
         Top             =   90
         Width           =   5010
      End
   End
   Begin VB.Label lblMakeDate 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0EEEE&
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   5190
      TabIndex        =   3
      Top             =   30
      Width           =   30
   End
End
Attribute VB_Name = "ctlVRA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'   �o�n�\ �\���R���g���[��
'

Option Explicit

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �O���֐�(�C�x���g)
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Public Event ChangeTo(strViewerName As String, strKey As String)    '' Vierer�ύX�C�x���g
Public Event NewWindow(strViewerName As String, strKey As String)   '' Vierer�ύX�C�x���g
Public Event WindowTitle(strKey As String)                          '' �E�C���h�E�^�C�g���ύX�C�x���g
Public Event LinkContextMenu(strViewerName As String, strKey As String)
Public Event Reload()                                               '' �ēǂݍ���
Public Event StatusBarTextChange(strMessage As String)              '' �X�e�[�^�X�o�[�e�L�X�g�ύX�C�x���g
Public Event Progression()

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �����ϐ�
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private WithEvents mToolBar As ctlToolBars
Attribute mToolBar.VB_VarHelpID = -1
Private mVB As clsViewerBase
Private mViewerState As clsVSTabOnly
Private mblnNoData As Boolean

Private mstrTitle As String              '' �E�C���h�E�^�C�g��
Private mKey As clsKeyRA                 '' �L�[
Private WithEvents mData As clsDataRA    '' �f�[�^�擾�I�u�W�F�N�g
Attribute mData.VB_VarHelpID = -1

Private mRecKey As clsKeyRC
Private mG1RecKey As clsKeyRC

Private blnInsertedLapData As Boolean

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �����萔
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Const MINIMUMWIDTH  As Long = 7000
Const MINIMUMHEIGHT As Long = 5000


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �v���p�e�B
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8


'
'   �@�\: �L�[�ݒ�v���p�e�B
'
'   ���l: Viewer�K�{�v���p�e�B
'
Public Property Let Key(strKey As String)
    mKey.str = strKey
    Call Update
End Property


'
'   �@�\: �^�C�g���擾�v���p�e�B
'
'   ���l: Viewer�K�{�v���p�e�B�A�@Browser ���Q��
'
Public Property Get Title() As String
    Title = mstrTitle
End Property


'
'   �@�\: �^�C�g���ݒ�v���p�e�B
'
'   ���l: �u���E�U�ɕύX�ʒm�̃C�x���g����
'
Public Property Let Title(strTitle As String)
    mstrTitle = strTitle
    RaiseEvent WindowTitle(mstrTitle)
End Property


'
'   �@�\: �c�[���o�[��ݒ肷��
'
'   ���l: �u���E�U����c�[���o�[���󂯎��A�c�[���o�[���Z�b�g�A�b�v����
' �@�@�@  RA, OD �̂݁A�K�{�v���p�e�B
'
Public Property Set ToolBar(RHS As ctlToolBars)
    Dim rc   As New clsRCSearch
    Dim rcG1 As New clsRCSearch
    Dim p    As Long
    
    Set mToolBar = RHS

     rc.CurrentRecordKeyStr = mData.RCKey
     rcG1.CurrentRecordKeyStr = mData.RCG1Key
    
    With ilsTbrSmall
        .ListImages.Clear
        .ImageHeight = 16
        .ImageWidth = 16
        .ListImages.Add 1, , LoadResPicture(201, vbResIcon)
        .ListImages.Add 2, , LoadResPicture(413, vbResIcon)
        .ListImages.Add 3, , LoadResPicture(411, vbResIcon)
    End With

    With mToolBar.ToolBar(1)
        .Buttons.Clear
        .ImageList = ilsTbrSmall
        
        p = 1
        
        .Buttons.Add p, "ODDS", "�I�b�Y", tbrDefault, 1
        .Buttons.item(p).ToolTipText = "���̃��[�X�̃I�b�Y�\���J���܂�"
        .Buttons.item(p).Image = 1

        
        p = p + 1
        .Buttons.Add p, "HYO", "�[��", tbrDefault, 1
        .Buttons.item(p).ToolTipText = "���̃��[�X�̕[���\���J���܂�"
        .Buttons.item(p).Image = 1
        
        p = p + 1
        .Buttons.Add p, "HENKO", "�ύX���", tbrDefault, 1
        .Buttons.item(p).ToolTipText = "���̓��̕ύX�����J���܂�"
        .Buttons.item(p).Image = 3
        p = p + 1
        .Buttons.Add p, "RECORD", "���R�[�h", tbrDefault, 1
        .Buttons.item(p).ToolTipText = "���̏����̃��R�[�h���J���܂�"
        .Buttons.item(p).Image = 2
        Set mRecKey = rc.PreviousRecordKey(True)
        If mRecKey Is Nothing Then
            .Buttons.item(p).Enabled = False
        Else
            .Buttons.item(p).Enabled = True
        End If
        
        p = p + 1
        .Buttons.Add p, "G1RECORD", "GI���R�[�h", tbrDefault, 1
        .Buttons.item(p).ToolTipText = "���̏�����G�T���R�[�h���J���܂�"
        .Buttons.item(p).Visible = mData.IsG1()
        .Buttons.item(p).Image = 2
        If mData.IsG1() Then
            Set mG1RecKey = rcG1.PreviousRecordKey(True)
            If mG1RecKey Is Nothing Then
                .Buttons.item(p).Enabled = False
            Else
                .Buttons.item(p).Enabled = True
            End If
        End If
        
    End With
    
    With mToolBar.ToolBar(2)
        .Buttons(1).Caption = "�J�Ï��擾"
        .Visible = mData.IsPrompt()
    End With
    
End Property


'
'   �@�\: Viewer��Ԓ�
'
'   ���l: �Ȃ�
'
Public Property Get ViewerState() As clsVSTabOnly
    Set ViewerState = mViewerState
End Property


'
'   �@�\: Viewer��Ԏ󂯎��
'
'   ���l: �Ȃ�
'
Public Property Set ViewerState(RHS As clsVSTabOnly)
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
'   �@�\: �\�[�g�O�C�x���g�B����\�[�g�𐧌䂵�܂��
'
'   ���l: ���уO���b�h�̒����\�[�g(�f�t�H���g)�̓J�X�^���\�[�g�ł��B
'
Private Sub flexTab_BeforeSort(Index As Integer, ByVal col As Long, Order As Integer)
On Error GoTo EH:
    Dim i As Long
    
    ' ���уO���b�h�ŁA�����J�����̏ꍇ�A��{�\�[�g���L�����Z�����A����\�[�g�����s����B
    If Index = 6 And col = 2 Then
        With flexTab(Index).Grid
            flexTab(Index).SortOrder(2) = Not flexTab(Index).SortOrder(2)
        End With
        Order = 0 ' �W���̃\�[�g�̓L�����Z������B
        Call flexTab_BeforeSort(6, 0, 2)
    ElseIf Index = 6 And col = 3 Then
        With flexTab(Index).Grid
            flexTab(Index).SortOrder(3) = Not flexTab(Index).SortOrder(3)
        End With
        Order = 0 ' �W���̃\�[�g�̓L�����Z������B
        Call flexTab_BeforeSort(6, 0, 2)
    End If
    Exit Sub
EH:
    gApp.ErrLog
End Sub


'
'   �@�\: �N���b�N�C�x���g
'
'   ���l: �����N���ʂ֐؂�ւ���
'
Private Sub flexTab_Click(Index As Integer)
On Error GoTo ErrorHandler
    Dim msrow As Long             '' �}�E�X���E
    Dim mscol As Long             '' �}�E�X�J����
    Dim item As clsGridItem     '' �O���b�h�A�C�e��
    
    
    ' �}�E�X�ʒu�̃O���b�h���W���擾
    With flexTab(Index).Grid
        msrow = .MouseRow
        mscol = .MouseCol
    End With
    
    ' �O���b�h�A�C�e�����Z��������o��
    Call SetItem(item, flexTab(Index), msrow, mscol)
    
    ' �A�C�e���������N�������Ă���ꍇ
    If item.HasAKey Then
        ' ��ʐ؂�ւ��C�x���g���M
        RaiseEvent ChangeTo(item.Link, item.Key)
    End If
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �}�E�X�̉��������N�\�ȃO���b�h�Ȃ�Δ�������ׂ̃C�x���g
'
'   ���l: �W���I�Ȕ����́AclsGridData.MouseMoveDriven�v���V�[�W���ɔC����
'
Private Sub flexTab_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrorHandler
    Select Case Index
    Case 1
        Call flexTab(Index).ReflexiveMouseMoveDriven
    Case Else
        Call flexTab(Index).ReflexiveMouseMoveDriven(True)
    End Select
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �E�N���b�N�ŁA�R���e�L�X�g���j���[���o��
'
'   ���l: �Ȃ�
'
Private Sub flexTab_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrorHandler
    Dim msrow As Long
    Dim mscol As Long
    Dim item As clsGridItem
    
    ' �}�E�X�̎����O���b�h���W���擾
    msrow = flexTab(Index).Grid.MouseRow
    mscol = flexTab(Index).Grid.MouseCol
    
    If mscol < 0 Or msrow < 0 Then
        Exit Sub
    End If
    
    Call SetItem(item, flexTab(Index), msrow, mscol)
    
    ' �f�[�^�������N�L�[�������Ă���
    If item.HasAKey Then
        ' ���A�E�N���b�N�ł���
        If Button = vbRightButton Then
            RaiseEvent LinkContextMenu(item.Link, item.Key)
        End If
    End If
    Exit Sub
    
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   �@�\: �����X�N���[���o�[�ύX�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub hsbSE_Change()
On Error GoTo ErrorHandler
    picIPane.Left = -hsbSE.value
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �����X�N���[���o�[�t�H�[�J�X�擾�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub hsbSE_GotFocus()
On Error Resume Next
    picIPane.SetFocus
End Sub


'
'   �@�\: �����X�N���[���o�[�X�N���[���C�x���g
'
'   ���l: �Ȃ�
'
Private Sub hsbSE_Scroll()
On Error GoTo ErrorHandler
    picIPane.Left = -hsbSE.value
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �f�[�^�擾���\��
'
'   ���l: �Ȃ�
'
Private Sub paneTab_Progression(Index As Integer)
On Error GoTo ErrorHandler
    RaiseEvent Progression
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �ߋ�n���̕\���́An�̐؂�ւ����ɗ��̎擾�𒆎~���A�V�����l�ł̎擾�����܂��B
'
'   ���l: �Ȃ�
'
Private Sub tmrKako_Timer()
On Error GoTo ErrorHandler
    Call mData.CancelKakoFetching
    If Not mData.NowKakoFetching Then
        tmrKako.Enabled = False
        mData.FetchKako
    End If
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �^�C�}�[�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub tmrTBS_Timer()
On Error GoTo ErrorHandler
    tmrTBS.Enabled = False
    Select Case tmrTBS.tag
    Case "ODDS"
        RaiseEvent NewWindow("OD", mKey.str & "0")
    Case "HYO"
        RaiseEvent NewWindow("OD", mKey.str & "1")
    Case "RECORD"
        RaiseEvent NewWindow("RC", mRecKey.str)
    Case "G1RECORD"
        RaiseEvent NewWindow("RC", mG1RecKey.str)
    Case "RTOPEN"
        Call gApp.DBPrompt(ukpRA, Left$(mKey.str, 8))
    Case "HENKO"
        RaiseEvent NewWindow("HK", Left$(mKey.str, 8))
    End Select
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �����X�N���[���o�[�ύX�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub vsbSE_Change()
On Error GoTo ErrorHandler
    picIPane.Top = -vsbSE.value
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �����X�N���[���o�[�t�H�[�J�X�擾�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub vsbSE_GotFocus()
On Error Resume Next
    picIPane.SetFocus
End Sub


'
'   �@�\: �����X�N���[���o�[�X�N���[���C�x���g
'
'   ���l: �Ȃ�
'
Private Sub vsbSE_Scroll()
On Error GoTo ErrorHandler
    picIPane.Top = -vsbSE.value
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �ߋ����L�[���̓C�x���g
'
'   ���l: �Ȃ�
'
Private Sub txtKako_KeyPress(KeyAscii As Integer)
On Error GoTo ErrorHandler
    If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
        KeyAscii = 0
    End If

    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �ߋ����ύX�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub txtKako_Change()
On Error GoTo ErrorHandler
    If (txtKako.Text) = 0 Then
        txtKako.Enabled = False
        txtKako.Text = "5"
        txtKako.Enabled = True
    End If
    If Len(txtKako.Text) > 2 Then
        txtKako.Enabled = False
        txtKako.Text = Right$(txtKako.Text, 2)
        txtKako.Enabled = True
    End If
    If txtKako.Enabled Then
        gApp.Log "�ߋ����\�����ύX"
        ' ���W�X�g���ɋL��
        gApp.R_KakoNum = val(txtKako.Text)
        mstTab.TabCaption(2) = "�ߋ�" & gApp.R_KakoNum & "��"
        Call mData.CancelKakoFetching
        tmrKako.Enabled = True
    End If
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �^�u�N���b�N�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub mstTab_Click(PreviousTab As Integer)
On Error GoTo ErrorHandler
    Dim i As Integer
    
    ' �I�����ꂽ�^�u�ɑΉ�����paneTab�݂̂�����
    For i = 0 To paneTab.count - 1
        paneTab(i).Visible = (i = mstTab.Tab)
    Next i
    
    mViewerState.LastTabNumber = mstTab.Tab
    
    Call UserControl_Resize
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   �@�\: �c�[���o�[�̃{�^���N���b�N�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub mToolBar_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
On Error GoTo ErrorHandler
    tmrTBS.tag = Button.Key
    tmrTBS.Enabled = True
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
    
    Set mKey = Nothing
    Set mData = Nothing
    Set mVB = Nothing
    
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
    
    blnInsertedLapData = False
    
    Dim i As Long
    Set mKey = New clsKeyRA
    Set mData = New clsDataRA
    Set mVB = New clsViewerBase
    Set mViewerState = New clsVSTabOnly
    
    mstrTitle = "�o�n�\"
    
    ' �ߋ����̎擾�����A�^�u�^�C�g���ƃe�L�X�g�{�b�N�X�ɐݒ肷��
    mstTab.TabCaption(2) = "�ߋ�" & gApp.R_KakoNum & "��"
    With txtKako
        .Enabled = False ' �C�x���g�𔭐������Ȃ�
        .Text = gApp.R_KakoNum
        .Enabled = True
    End With
    
    ' �ŏ����ݒ�
    With UserControl
        .width = Bigger(.width, MINIMUMWIDTH)
        .Height = Bigger(.Height, MINIMUMHEIGHT)
    End With
    
    ' ����UI�ݒ�
    Call mVB.InitGUI(UserControl.width, UserControl.Height, fraTop, lblMakeDate, fraHeader, mstTab)
    
   
    ' �X�N���[���o�[�̕���ݒ肷��
    vsbSE.width = gApp.vsbWidth
    hsbSE.Height = gApp.hsbHeight
    
    ' Font Asign
    Call mVB.FraTopFontType1(lblInfo(0).Font)
    
    ' FlexGrid�ݒ�
    For i = flexTab.LBound To flexTab.UBound
        Call mVB.FlexGridCommonSetting(flexTab(i).Grid)
        flexTab(i).Grid.FixedCols = 0
    Next i
    With flexTab(6).Grid
        .ScrollBars = flexScrollBarNone
        .FixedCols = 0
    End With
    With flexTab(7).Grid
        .ScrollBars = flexScrollBarNone
        .FixedCols = 0
    End With
    With flexTab(8).Grid
        .ScrollBars = flexScrollBarNone
        .FixedCols = 0
        .FixedRows = 0
        Call flexTab(8).FlexDisable
    End With
    
    ' Color Asign
    UserControl.BackColor = gApp.ColBG
    lblInfo(0).BackColor = gApp.ColDarkBG
    lblInfo(0).ForeColor = Contrast(gApp.ColDarkBG)
    ' skip lblInfo(1)
    For i = 2 To lblInfo.UBound
        lblInfo(i).BackColor = gApp.ColBG
        lblInfo(i).ForeColor = Contrast(gApp.ColBG)
    Next i
    
    ' ���ׂẴy�C�����A�f�[�^�擾���ɐݒ肷��B
    For i = paneTab.LBound To paneTab.UBound
        paneTab(i).Mode = 0
    Next i
    
    ' ���ׂẴy�C�����A�f�[�^�擾���ɐݒ肷��B
    For i = 0 To mstTab.Tabs - 1
        mstTab.TabEnabled(i) = False
    Next i
    
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
    Dim i As Integer
    
    ' �ŏ����ݒ�
    With UserControl
        .width = Bigger(.width, MINIMUMWIDTH)
        .Height = Bigger(.Height, MINIMUMHEIGHT)
    End With
    
    ' ����UI���T�C�Y
    
    Call mVB.ResizeGUI(UserControl.width, UserControl.Height, fraTop, lblMakeDate, fraHeader, mstTab)
    
    ' Viewer���LUI���T�C�Y
    
    
    For i = 0 To 6
        With paneTab(i)
            .Top = mstTab.TabHeight + 60
            .Left = 60
            .width = Bigger(1, mstTab.width - .Left * 2)
            .Height = Bigger(1, (mstTab.Height - .Top) - (.Top - mstTab.TabHeight))
        End With ' paneTab(mstTab.Index)
        
        ' ���у^�u�ӊO�́A�O���b�h���ő啝�ɂ���
        If i <> 6 Then
            With flexTab(i)
                .width = Bigger(1, paneTab(i).width - .Left)
                .Height = Bigger(1, paneTab(i).Height - .Top)
            End With ' flexTab(i)
        Else
            ' ���у^�u�́A�O�̃O���b�h�ƃ��x���𓮓I�ɐ���
            
            ' �O���b�h�̑傫�����Z���̓��e�Ƀt�B�b�g������
            Call FitGrid(flexTab(6))
            Call FitGrid(flexTab(7))
            
            If blnInsertedLapData Then
                With flexTab(8)
                    Call .AutoSize(5, .Grid.Cols - 1)
                    Dim r As Long, c As Long
    
                    For r = 4 To .Grid.Rows - 1
                        If LenB(.Grid.TextMatrix(r, 0)) > 30 Then
                            For c = 1 To .Grid.Cols - 1
                                .Grid.TextMatrix(r, c) = .Grid.TextMatrix(r, 0)
                            Next c
                        End If
                    Next r

                    .Grid.MergeCells = flexMergeRestrictRows
                    For r = 0 To ((.Grid.Rows - 1) / 2)
                        .Grid.MergeRow(r) = False
                    Next r
                    
                End With
                blnInsertedLapData = False
            End If
            Call FitGrid(flexTab(8))
            
            
            ' ���ꂼ��̃O���b�h�𐮗񂷂�

            flexTab(6).Top = 0
            flexTab(6).Left = 0
            lblFix(3).Top = flexTab(6).Height
            flexTab(7).Top = flexTab(6).Height + lblFix(3).Height
            flexTab(7).Left = 0
            lblFix(2).Top = lblFix(3).Top
            lblFix(2).Left = flexTab(7).width + lblFix(2).Height
            flexTab(8).Top = flexTab(7).Top
            flexTab(8).Left = lblFix(2).Left
            
            ' ���A�������A�ő�ɑ�����
            If flexTab(6).width < flexTab(8).Left + flexTab(8).width Then
                flexTab(6).width = flexTab(8).Left + flexTab(8).width
            Else
                flexTab(8).width = flexTab(6).width - flexTab(8).Left
            End If
            gApp.Log flexTab(7).Height & " " & flexTab(8).Height
            If flexTab(7).Height < flexTab(8).Height Then
                flexTab(7).Height = flexTab(8).Height
            Else
                flexTab(8).Height = flexTab(7).Height
            End If
            
            ' �����y�C���𐮗񂵂��O���b�h������悤�Ƀt�B�b�g������
            picIPane.width = flexTab(6).width
            picIPane.Height = flexTab(7).Top + flexTab(7).Height
            
            
            
            ' �X�N���[���o�[�̊Ǘ�
            Call ScrollBarManage
        End If
    Next i
    
    With lblInfo(1)
        .Left = fraTop.width - .width - 100
    End With
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   �@�\: ��{�f�[�^�擾�����ʒm�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub mData_FetchedKihon(GridData As clsGridData)
On Error GoTo ErrorHandler
    gApp.Log "Fetch Kihon"
    Call flexTab(0).InsertGrid(GridData)
    
    With flexTab(0).Grid
        .FixedCols = 0
    End With
    Call flexTab(0).AutoSize(0, flexTab(0).Grid.Cols - 1)
    paneTab(0).Mode = 2
    mstTab.TabEnabled(0) = True
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �����f�[�^�擾�����ʒm�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub mData_FetchedKetto(GridData As clsGridData)
On Error GoTo ErrorHandler
    gApp.Log "Fetch Ketto"
    Call flexTab(1).InsertGrid(GridData)
    Call flexTab(1).AutoSize(0, flexTab(1).Grid.Cols - 1)
    
    ' �����O���b�h
    With flexTab(1).Grid
        .MergeCol(0) = True
        .MergeCol(1) = True
        .MergeCol(2) = True
        .MergeCol(3) = True
        .MergeCol(4) = True
        .MergeCol(5) = True
        .MergeCells = flexMergeFree
    End With
    paneTab(1).Mode = 2
    mstTab.TabEnabled(1) = True
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �ߋ����f�[�^�擾�����ʒm�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub mData_FetchedKako(GridData As clsGridData)
On Error GoTo ErrorHandler
    gApp.Log "Fetch Kako"
    Call flexTab(2).InsertGrid(GridData)

    Call flexTab(2).AutoSize(0, flexTab(2).Grid.Cols - 1, False, True)
    ' �ߋ�N���^�u
    With flexTab(2).Grid
        .FixedCols = 0
        .FixedRows = 1
        .WordWrap = True

        ' �Z���������l��
        Dim i As Integer
        For i = 0 To 1
            .ColWidth(i) = 360
        Next
        .ColWidth(2) = 1800
        Dim r As Long, c As Long
        Dim newColWid As Long
        newColWid = 700
        For r = 1 To .Rows - 1
            If LenB(.TextMatrix(r, 3)) > 10 Then
                newColWid = 2880
                Exit For
            End If
        Next r
            
        For c = 3 To .Cols - 1
            .ColWidth(c) = newColWid
        Next c
        
    End With
    
    paneTab(2).Mode = 2
    mstTab.TabEnabled(2) = True
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �}�C�j���O�f�[�^�擾�����ʒm�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub mData_FetchedMining(GridData As clsGridData, DMKubun As String)
On Error GoTo ErrorHandler
    gApp.Log "Fetch Mining"
    Dim i As Long
    Dim strBeforeRow As String
    Dim iOld As Long
    
    Call flexTab(3).InsertGrid(GridData)
    
    '�^�C������\�z���ʂ�}��
    With flexTab(3).Grid
        .col = 6
        .Sort = flexSortStringAscending
        
        For i = 1 To .Rows - 1
            .row = i
            .col = 6
            If Mid(.Text, 2, 1) = ":" Then
                '�Z���Ƀ^�C���������Ă��鎞�A�^�C�������ʂɓ���ւ�
                If strBeforeRow = .Text And i <> 1 Then
                    .Text = iOld
                Else
                    strBeforeRow = .Text
                    .Text = i
                    iOld = i
                End If
            End If
        Next
        
        .col = 1
        .Sort = flexSortStringAscending
        
    End With

    Call flexTab(3).AutoSize(0, flexTab(3).Grid.Cols - 1)
    flexTab(3).Grid.FixedCols = 0
    lblDMKubun.Caption = DMKubun
    lblDMKubun.Visible = (DMKubun <> "")
    
    paneTab(3).Mode = 2
    mstTab.TabEnabled(3) = True
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �����ʃf�[�^�擾�����ʒm�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub mData_FetchedJokenBetu(GridData As clsGridData)
On Error GoTo ErrorHandler
    gApp.Log "Fetch JokenBetu"

    Call flexTab(4).InsertGrid(GridData)
    
    Call flexTab(4).AutoSize(0, flexTab(4).Grid.Cols - 1, False, False, 1)
    
    ' �����ʃO���b�h
    With flexTab(4).Grid
        flexTab(4).Grid.TextMatrix(0, 0) = " "
        flexTab(4).Grid.TextMatrix(0, 1) = " "
        flexTab(4).Grid.TextMatrix(0, 2) = " "
        
        .MergeRow(0) = True
        .MergeCells = flexMergeFree
        .FixedCols = 0
        .FixedRows = 2
    End With
    paneTab(4).Mode = 2
    mstTab.TabEnabled(4) = True
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �����^�C���f�[�^�擾�����ʒm�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub mData_FetchedMotiTIme(GridData As clsGridData)
On Error GoTo ErrorHandler
    gApp.Log "Fetch MochiTime"
    
    Call flexTab(5).InsertGrid(GridData)

    Call flexTab(5).AutoSize(0, flexTab(5).Grid.Cols - 1, False, False, 1)
    ' �����^�C��
    With flexTab(5).Grid
        flexTab(5).Grid.TextMatrix(0, 0) = " "
        flexTab(5).Grid.TextMatrix(0, 1) = " "
        flexTab(5).Grid.TextMatrix(0, 2) = " "
            
        .MergeRow(0) = True
        .MergeCells = flexMergeFree
        .FixedRows = 2
        .FixedCols = 0
    End With
    paneTab(5).Mode = 2
    mstTab.TabEnabled(5) = True
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: ���уf�[�^�擾�����ʒm�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub mData_FetchedSeiseki(GridData As clsGridData, GridHarai As clsGridData, GridLap As clsGridData, flag As Long)
On Error GoTo ErrorHandler
    Dim i           As Long
    Dim j           As Long
    Dim maxWidth    As Long
    Dim sumWidth    As Long
    
    ' ���у^�u
    Call flexTab(6).InsertGrid(GridData)
    
    ' �\�[�g���{
    flexTab(6).Grid.col = 0
    flexTab(6).Grid.Sort = flexSortGenericAscending
    Call SortFlexGrid(flexTab(6), 0)
    
    ' ���߃^�u
    Call flexTab(7).InsertGrid(GridHarai)

    ' �}�[�W
    For i = mData.HenkanRow To flexTab(7).Grid.Rows - 1
        flexTab(7).Grid.MergeRow(i) = True
    Next i
    flexTab(7).Grid.MergeCol(0) = True
    flexTab(7).Grid.MergeCells = flexMergeFree
    ' ���b�v�^�C���^�u
    Call flexTab(8).InsertGrid(GridLap)
    blnInsertedLapData = True
    
    ' ���Ń\�[�g
    Call flexTab(6).AutoSize(0, flexTab(6).Grid.Cols - 1)
    flexTab(6).Grid.ColWidth(0) = 0
    Call flexTab_BeforeSort(6, 2, 2)
    Call flexTab(7).AutoSize(0, flexTab(7).Grid.Cols - 1)

    ' �����߂��O���b�h
    With flexTab(7).Grid
        .Visible = (.Cols >= 2)
        lblFix(3).Visible = (.Cols >= 2)
        If .Rows > 1 Then
            For i = 0 To .Rows - 1
                .RowSel = i
                .ColSel = 1
                If .Text = "�s����" Then
                    .MergeRow(i) = True
                End If
            Next i
        End If
    End With

    ' ���b�v�^�C���O���b�h
    With flexTab(8).Grid
        For i = 0 To 3
            For j = 1 To .Cols - 1
                .RowSel = mData.CornerRow + i
                .ColSel = j
                .Text = ""
            Next j
            .MergeRow(mData.CornerRow + i) = True
        Next i
        .MergeCells = flexMergeFree
        .GridColor = &HFFFFFF
        
        maxWidth = .ColWidth(0)
        
        sumWidth = 0
        For i = 0 To .Cols - 1
            .RowSel = 0
            .ColSel = i
            sumWidth = sumWidth + .width
        Next i
        
        If sumWidth < maxWidth Then
            .ColWidth(.Cols - 1) = maxWidth - (sumWidth - .ColWidth(.Cols - 1))
        End If
        
    End With
    
    

    Call UserControl_Resize
    
    paneTab(6).Mode = 2
    paneTab(6).BorderStyle = ebscThin
    mstTab.TabEnabled(6) = True
    Exit Sub
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   �@�\: �����n���[�X�̃f�[�^������
'
'   ���l: �Ȃ�
'
Private Sub mData_NoUMARACE()
On Error GoTo ErrorHandler
    Dim i As Long
    
    For i = 0 To 6
        paneTab(i).Mode = 1
        mstTab.TabEnabled(i) = True
    Next i
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �X�N���[���o�[
'
'   ���l: �Ȃ�
'
Private Sub ScrollBarManage()
    Dim hsbIsVisible As Boolean
    
    
    vsbSE.Visible = False
    hsbSE.Visible = False
    
    ' ����
    hsbIsVisible = False
    If picIPane.width > paneTab(6).width Then
        paneTab(6).Height = paneTab(6).Height - hsbSE.Height
        hsbIsVisible = True
        hsbSE.Visible = (6 = mstTab.Tab)
    End If
    
    ' ����
    If picIPane.Height > paneTab(6).Height Then
        paneTab(6).width = paneTab(6).width - vsbSE.width
        vsbSE.Visible = (6 = mstTab.Tab)
    End If
    
    ' �����X�N���[���o�[�\���ɂ�萅���X�N���[���o�[���K�v�ɂȂ����Ƃ�
    If hsbIsVisible = False And picIPane.width > paneTab(6).width Then
        paneTab(6).Height = paneTab(6).Height - hsbSE.Height
        hsbSE.Visible = (6 = mstTab.Tab)
    End If
    
    With hsbSE
        .Top = paneTab(6).Top + paneTab(6).Height
        .Left = paneTab(6).Left
        .width = paneTab(6).width
    End With
    
    With vsbSE
        .Top = paneTab(6).Top
        .Left = paneTab(6).Left + paneTab(6).width
        .Height = paneTab(6).Height
    End With
    
    hsbSE.max = picIPane.width - paneTab(6).width
    hsbSE.LargeChange = paneTab(6).width
    hsbSE.SmallChange = vsbSE.width
    
    vsbSE.max = picIPane.Height - paneTab(6).Height
    vsbSE.LargeChange = paneTab(6).Height
    vsbSE.SmallChange = hsbSE.Height
    
End Sub


'
'   �@�\: �O���b�h����
'
'   ���l: �Ȃ�
'
Private Sub FitGrid(wgd As Object)
    Dim i As Long
    Dim w As Long
    Dim h As Long
    Dim ctlGrid As Control
    
    For i = 0 To wgd.Grid.Cols - 1
        w = w + wgd.Grid.ColWidth(i)
    Next i
    For i = 0 To wgd.Grid.Rows - 1
        h = h + wgd.Grid.RowHeight(i)
    Next i

    wgd.width = w + wgd.Grid.GridLineWidth * (wgd.Grid.Cols + 1) * Screen.TwipsPerPixelY
    wgd.Height = h + wgd.Grid.GridLineWidth * (wgd.Grid.Rows + 1) * Screen.TwipsPerPixelX
    
    With wgd.Grid
        wgd.width = .ColPos(.Cols - 1) + .ColWidth(.Cols - 1) + 1 * Screen.TwipsPerPixelX
        wgd.Height = .RowPos(.Rows - 1) + .RowHeight(.Rows - 1) + 1 * Screen.TwipsPerPixelY
    End With
End Sub


'
'   �@�\: �f�[�^���擾����
'
'   ���l: �Ȃ�
'
Private Sub Update()
On Error GoTo ErrorHandler
    Dim i As Integer
    Dim j As Integer
    Dim LastTabNumber As Long
    

    ' �f�[�^���擾���Ă��炤
    gApp.Log "RA Fetch Start"
    If Not mData.Fetch(mKey) Then
        gApp.Log "RA Fetch End - NODATA"
        mblnNoData = True
    End If
    gApp.Log "RA Fetch End"

    ' �E�C���h�E�^�C�g���̕ύX
    Me.Title = mData.Title
    
    ' ��Q���[�X�̏ꍇ�́A�����^�C���^�u��\�����Ȃ�
    mstTab.TabVisible(5) = Not mData.IsShougai

    
    ' ���x�����擾
    For i = 0 To 7
        lblInfo(i).Caption = ReplaceAmpersand(mData.Labels(i))
    Next i
    lblMakeDate = mData.Labels(8)
    lblInfo(8).Caption = mData.Labels(9) ' ���R�[�h
    
    ' ���x���𐮗񂳂���
    lblInfo(2).Left = 0
    lblInfo(2).Top = 30
    lblInfo(3).Left = lblInfo(2).Left + lblInfo(2).width
    lblInfo(3).Top = lblInfo(2).Top
    lblInfo(4).Left = lblInfo(3).Left
    lblInfo(4).Top = lblInfo(3).Top + lblInfo(3).Height
    lblInfo(5).Left = lblInfo(4).Left
    lblInfo(5).Top = lblInfo(4).Top + lblInfo(4).Height
    lblInfo(6).Left = Bigger(lblInfo(4).Left + lblInfo(4).width, lblInfo(5).Left + lblInfo(5).width)
    lblInfo(6).Top = lblInfo(4).Top
    lblInfo(7).Left = lblInfo(6).Left
    lblInfo(7).Top = lblInfo(6).Top + lblInfo(6).Height
    lblInfo(8).Left = lblInfo(7).Left + lblInfo(7).width
    lblInfo(8).Top = lblInfo(7).Top
    lblInfo(8).ForeColor = vbRed
    
    ' �ŏ��ɕ\������^�u��ݒ肷��
    If mViewerState.IsNoTouch Then
        mstTab.Tab = 2
        mstTab.Tab = mData.FirstTabNumber
    Else
        LastTabNumber = mViewerState.LastTabNumber
        mstTab.Tab = (mViewerState.LastTabNumber + 1) Mod mstTab.Tabs
        mstTab.Tab = LastTabNumber
    End If
    If mData.FirstTabNumber = 0 Then
        mstTab.TabEnabled(6) = False
        mstTab.TabVisible(6) = False
    End If
    
    gApp.Log "RA Update Finish"
    Exit Sub
ErrorHandler:
    gApp.ErrLog
    Resume Next
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
    gApp.Log "Free"
    If Not mData Is Nothing Then
        mData.CancelKakoFetching
        mData.CancelFetching
    End If
    Dim i As Integer
    For i = flexTab.LBound To flexTab.UBound
        Call DestroyFlexGrid(flexTab(i))
    Next i
    Set mData = Nothing
End Sub

