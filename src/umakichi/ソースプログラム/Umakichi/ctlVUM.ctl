VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.UserControl ctlVUM 
   BackColor       =   &H00E0EEEE&
   ClientHeight    =   6885
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11280
   ScaleHeight     =   6885
   ScaleWidth      =   11280
   Begin VB.Frame fraHeader 
      BackColor       =   &H00E0EEEE&
      BorderStyle     =   0  '�Ȃ�
      Caption         =   "fraHeader"
      Height          =   1665
      Left            =   330
      TabIndex        =   4
      Top             =   600
      Width           =   10185
      Begin Umakichi.ctlClickLabel clblInfo 
         Height          =   180
         Index           =   2
         Left            =   3690
         TabIndex        =   11
         Top             =   675
         Width           =   90
         _ExtentX        =   159
         _ExtentY        =   318
         AutoSize        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �S�V�b�N"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "."
      End
      Begin Umakichi.ctlClickLabel clblInfo 
         Height          =   180
         Index           =   1
         Left            =   3690
         TabIndex        =   10
         Top             =   480
         Width           =   90
         _ExtentX        =   159
         _ExtentY        =   318
         AutoSize        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �S�V�b�N"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "."
      End
      Begin Umakichi.ctlClickLabel clblInfo 
         Height          =   180
         Index           =   0
         Left            =   3690
         TabIndex        =   12
         Top             =   270
         Width           =   90
         _ExtentX        =   159
         _ExtentY        =   318
         AutoSize        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �S�V�b�N"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "."
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  '�ׯ�
         BackColor       =   &H00E0EEEE&
         BorderStyle     =   0  '�Ȃ�
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   1
         Left            =   3000
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   9
         Text            =   "ctlVUM.ctx":0000
         Top             =   300
         Width           =   825
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  '�ׯ�
         BackColor       =   &H00E0EEEE&
         BorderStyle     =   0  '�Ȃ�
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   4
         Left            =   5880
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   8
         Text            =   "ctlVUM.ctx":001F
         Top             =   900
         Width           =   1545
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  '�ׯ�
         BackColor       =   &H00E0EEEE&
         BorderStyle     =   0  '�Ȃ�
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   3
         Left            =   2880
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   7
         Text            =   "ctlVUM.ctx":0054
         Top             =   900
         Width           =   2685
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  '�ׯ�
         BackColor       =   &H00E0EEEE&
         BorderStyle     =   0  '�Ȃ�
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Index           =   0
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   6
         Text            =   "ctlVUM.ctx":0088
         Top             =   90
         Width           =   5265
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  '�ׯ�
         BackColor       =   &H00E0EEEE&
         BorderStyle     =   0  '�Ȃ�
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   2
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   5
         Text            =   "ctlVUM.ctx":00EB
         Top             =   900
         Width           =   2685
      End
   End
   Begin VB.Frame fraTop 
      BackColor       =   &H00C0CCCC&
      BorderStyle     =   0  '�Ȃ�
      Caption         =   "Frame1"
      Height          =   465
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8955
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
         Left            =   240
         TabIndex        =   1
         Top             =   120
         Width           =   5010
      End
   End
   Begin TabDlg.SSTab mstTab 
      Height          =   3615
      Left            =   360
      TabIndex        =   2
      Top             =   2520
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   6376
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      BackColor       =   14741230
      TabCaption(0)   =   "����"
      TabPicture(0)   =   "ctlVUM.ctx":011F
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "paneTab(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "�ߋ�����"
      TabPicture(1)   =   "ctlVUM.ctx":013B
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "paneTab(1)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "�����ʐ���"
      TabPicture(2)   =   "ctlVUM.ctx":0157
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "vsbSE"
      Tab(2).Control(1)=   "hsbSE"
      Tab(2).Control(2)=   "paneTab(2)"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "�����^�C��"
      TabPicture(3)   =   "ctlVUM.ctx":0173
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "paneTab(3)"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "��H����"
      TabPicture(4)   =   "ctlVUM.ctx":018F
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "paneTab(4)"
      Tab(4).ControlCount=   1
      Begin Umakichi.ctlPane paneTab 
         Height          =   2205
         Index           =   2
         Left            =   -74940
         TabIndex        =   15
         Top             =   360
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   3889
         Begin VB.PictureBox picIPane 
            Appearance      =   0  '�ׯ�
            BackColor       =   &H8000000C&
            ForeColor       =   &H80000008&
            Height          =   1815
            Left            =   0
            ScaleHeight     =   1785
            ScaleWidth      =   6945
            TabIndex        =   16
            Top             =   0
            Width           =   6975
            Begin Umakichi.ctlWrappedGrid flexTab 
               Height          =   1455
               Index           =   2
               Left            =   0
               TabIndex        =   21
               Top             =   0
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   2566
            End
            Begin Umakichi.ctlWrappedGrid flexTab 
               Height          =   1455
               Index           =   5
               Left            =   1920
               TabIndex        =   26
               Top             =   0
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   2566
            End
            Begin Umakichi.ctlWrappedGrid flexTab 
               Height          =   1455
               Index           =   6
               Left            =   3960
               TabIndex        =   27
               Top             =   0
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   2566
            End
         End
      End
      Begin VB.HScrollBar hsbSE 
         Height          =   285
         Left            =   -74340
         TabIndex        =   14
         Top             =   3060
         Width           =   5295
      End
      Begin VB.VScrollBar vsbSE 
         Height          =   1995
         Left            =   -67440
         TabIndex        =   13
         Top             =   360
         Width           =   285
      End
      Begin Umakichi.ctlPane paneTab 
         Height          =   2370
         Index           =   0
         Left            =   0
         TabIndex        =   17
         Top             =   600
         Width           =   5580
         _ExtentX        =   9843
         _ExtentY        =   4180
         Begin Umakichi.ctlWrappedGrid flexTab 
            Height          =   1455
            Index           =   0
            Left            =   0
            TabIndex        =   18
            Top             =   0
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   2566
         End
      End
      Begin Umakichi.ctlPane paneTab 
         Height          =   2160
         Index           =   1
         Left            =   -75000
         TabIndex        =   19
         Top             =   840
         Width           =   3870
         _ExtentX        =   6826
         _ExtentY        =   3810
         Begin Umakichi.ctlWrappedGrid flexTab 
            Height          =   1455
            Index           =   1
            Left            =   0
            TabIndex        =   20
            Top             =   0
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   2566
         End
         Begin VB.Label lblFix 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  '����
            Caption         =   "��Q���[�X�ɂ��ẮA[��3�n����]��""��3F�^�C��""�łȂ��A""���Y���[�X���j�^�C����1F���σ^�C��""��\�����Ă��܂��B"
            BeginProperty Font 
               Name            =   "�l�r �o�S�V�b�N"
               Size            =   8.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   165
            Left            =   0
            TabIndex        =   28
            Top             =   1560
            Width           =   8190
         End
      End
      Begin Umakichi.ctlPane paneTab 
         Height          =   1695
         Index           =   3
         Left            =   -75000
         TabIndex        =   22
         Top             =   480
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   2990
         Begin Umakichi.ctlWrappedGrid flexTab 
            Height          =   1455
            Index           =   3
            Left            =   0
            TabIndex        =   23
            Top             =   0
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   2566
         End
      End
      Begin Umakichi.ctlPane paneTab 
         Height          =   1695
         Index           =   4
         Left            =   -75000
         TabIndex        =   24
         Top             =   600
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   2990
         Begin Umakichi.ctlWrappedGrid flexTab 
            Height          =   1455
            Index           =   4
            Left            =   0
            TabIndex        =   25
            Top             =   0
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   2566
         End
      End
   End
   Begin VB.Label lblMakeDate 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0EEEE&
      Caption         =   "�f�[�^�쐬��: 9999�N99��99��"
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
      Left            =   6060
      TabIndex        =   3
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "ctlVUM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'   �����n  �\���R���g���[��
'

Option Explicit

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �O���֐�(�C�x���g)
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Public Event ChangeTo(strViewerName As String, strKey As String)        '' Vierer�ύX�C�x���g
Public Event WindowTitle(strKey As String)                              '' �E�C���h�E�^�C�g���ύX�C�x���g
Public Event LinkContextMenu(strViewerName As String, strKey As String) '' �E�N���b�N���j���[�\���C�x���g

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �����ϐ�
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private mVB As clsViewerBase            '' Viewer Base
Private mViewerState As clsVSTabOnly    '' Viewer State

Private WithEvents mData As clsDataUM   '' �f�[�^�擾�I�u�W�F�N�g
Attribute mData.VB_VarHelpID = -1
Private mstrTitle As String             '' �E�C���h�E�^�C�g��
Private mKey As clsKeyUM                '' �L�[
Private mblnNoData As Boolean           '' �f�[�^�����t���O

Private mSortAscending As Boolean       '' �\�[�g�����t���O

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �����萔
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

' ��ʍŏ����l
Const MINIMUMWIDTH  As Long = 4000
Const MINIMUMHEIGHT As Long = 4000

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �v���p�e�B
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8


'
'   �@�\: �L�[�ݒ�v���p�e�B
'
'   ���l: Viewer�K�{�v���p�e�B
'
Public Property Let Key(strKey As String)
    gApp.Log "UM: " & strKey
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
Private Sub clblinfo_ChangeViewer(Index As Integer)
On Error GoTo ErrorHandler
    RaiseEvent ChangeTo(clblinfo(Index).ViewerName, clblinfo(Index).Key)
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �E�N���b�N�����C�x���g
'
'   ���l: �u���E�U�̃|�b�v�A�b�v�C�x���g�ɃX���[����
'
Private Sub clblinfo_RightMouseDown(Index As Integer)
On Error GoTo ErrorHandler
    RaiseEvent LinkContextMenu(clblinfo(Index).ViewerName, clblinfo(Index).Key)
    Exit Sub
ErrorHandler:
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
    
    ' �O���b�h�A�C�e�����Z��������o��
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
'   �@�\: �}�E�X�̉��������N�\�ȃO���b�h�Ȃ�Δ�������ׂ̃C�x���g
'
'   ���l: �W���I�Ȕ����́AclsGridData.MouseMoveDriven�v���V�[�W���ɔC����
'
Private Sub flexTab_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrorHandler
    
    If 0 = Index Then
        Call flexTab(Index).MouseMoveDriven
    ElseIf 1 = Index Or 3 = Index Then
        Call flexTab(Index).ReflexiveMouseMoveDriven(True)
    Else
    End If
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �\�[�g�O�C�x���g�B����\�[�g�𐧌䂵�܂��
'
'   ���l: �O���b�h����J������sort�֎~ & �B���J�����Ń\�[�g�B
'
Private Sub flexTab_BeforeSort(Index As Integer, ByVal col As Long, Order As Integer)
On Error GoTo ErrorHandler
    If Index = 1 And col = 0 Then
        Order = 0
        With flexTab(Index).Grid
            If mSortAscending Then
                mSortAscending = False
            Else
                mSortAscending = True
            End If
        End With
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
        
    ' �I�����ꂽ�^�u�ɑΉ�����fraTab�݂̂�����
    For i = 0 To paneTab.count - 1
        paneTab(i).Visible = (i = mstTab.Tab)
    Next i
    
    mViewerState.LastTabNumber = mstTab.Tab
    
    Call Tab_Resize
    
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

    Dim i As Long
    Set mKey = New clsKeyUM
    Set mData = New clsDataUM
    Set mVB = New clsViewerBase
    Set mViewerState = New clsVSTabOnly
    
    mstrTitle = "�����n"

    ' �ŏ����ݒ�
    With UserControl
        .width = Bigger(.width, MINIMUMWIDTH)
        .Height = Bigger(.Height, MINIMUMHEIGHT)
    End With
    
    ' ����UI�ݒ�
    Call mVB.InitGUI(UserControl.width, UserControl.Height, fraTop, lblMakeDate, fraHeader, mstTab)
    
    ' FlexGrid�ݒ�
    For i = flexTab.LBound To flexTab.UBound
        Call mVB.FlexGridCommonSetting(flexTab(i).Grid)
    Next i
    
    ' �X�N���[���o�[������
    vsbSE.width = gApp.vsbWidth
    hsbSE.Height = gApp.hsbHeight
    
    ' Color Assign
    BackColor = gApp.ColBG
    
    lblInfo(0).BackColor = gApp.ColDarkBG
    lblInfo(0).ForeColor = Contrast(gApp.ColDarkBG)
    For i = txtInfo.LBound To txtInfo.UBound
        txtInfo(i).BackColor = gApp.ColBG
        txtInfo(i).ForeColor = Contrast(gApp.ColBG)
    Next i
    For i = clblinfo.LBound To clblinfo.UBound
        clblinfo(i).BackColor = gApp.ColBG
        clblinfo(i).ForeColor = Contrast(gApp.ColBG)
    Next i
        
    Call mstTab_Click(0)
    
    ' ���ׂẴy�C�����A�f�[�^�擾���ɐݒ肷��B
    For i = paneTab.LBound To paneTab.UBound
        paneTab(i).Mode = 0
    Next i
    
    ' ���ׂẴ^�u���A�f�[�^�擾���ɐݒ肷��B
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
    
    Call Tab_Resize
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub
    

'
'   �@�\: �^�u�̃��T�C�Y�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub Tab_Resize()
On Error GoTo ErrorHandler
    Dim i As Integer
    
    For i = 0 To 4
        With paneTab(i)
            .Top = mstTab.TabHeight + 60
            .Left = 60
            .width = Bigger(1, mstTab.width - .Left * 2)
            .Height = Bigger(1, (mstTab.Height - .Top) - (.Top - mstTab.TabHeight))
        End With ' fraTab(mstTab.Index)

        Select Case i
        Case 0, 3, 4
        ' �ߋ�����, �����ʐ��у^�u�ӊO�́A�O���b�h���ő��
            With flexTab(i)
                .width = Bigger(1, paneTab(i).width - .Left)
                .Height = Bigger(1, paneTab(i).Height - .Top)
            End With ' flexTab(i)
        Case 1
        ' �ߋ����у^�u
            With flexTab(i)
                .Top = lblFix.Height
                .width = Bigger(1, paneTab(i).width - .Left)
                .Height = Bigger(1, paneTab(i).Height - .Top)
                lblFix.Top = 0
                lblFix.Left = 0
            End With
        Case 2
        ' �����ʐ��у^�u
            Call ScrollBarManage

            flexTab(2).Height = 2100
            flexTab(2).width = 6650
            flexTab(5).Height = 3400
            flexTab(5).width = 6650
            flexTab(6).Height = 2100
            flexTab(6).width = 6650
            
            flexTab(2).Top = 0
            flexTab(2).Left = 0
            flexTab(5).Top = 0
            flexTab(5).Left = flexTab(2).width + 300
            flexTab(6).Top = flexTab(2).Height + 300
            flexTab(6).Left = 0

            With picIPane
                .width = Bigger(MINIMUMWIDTH, flexTab(5).Left + flexTab(5).width + 200)
                .Height = Bigger(MINIMUMHEIGHT, flexTab(6).Top + flexTab(6).Height)
                .width = Bigger(.width, paneTab(2).width - .Left)
                .Height = Bigger(.Height, paneTab(2).Height - .Top)
            End With
        End Select
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
    If picIPane.width > paneTab(2).width + vsbSE.width Then
        paneTab(2).Height = paneTab(2).Height - hsbSE.Height
        hsbIsVisible = True
        hsbSE.Visible = (2 = mstTab.Tab)
    End If

    ' ����
    If picIPane.Height > paneTab(2).Height + hsbSE.Height Then
        paneTab(2).width = paneTab(2).width - vsbSE.width
        vsbSE.Visible = (2 = mstTab.Tab)
    End If

    ' �����X�N���[���o�[�ɂ�萅���X�N���[���o�[���K�v�ɂȂ����Ƃ�
    If hsbIsVisible = False And picIPane.width > paneTab(2).width + vsbSE.width Then
        paneTab(2).Height = paneTab(2).Height - hsbSE.Height
        hsbSE.Visible = (2 = mstTab.Tab)
    End If

    With hsbSE
        .Top = paneTab(2).Top + paneTab(2).Height
        .Left = paneTab(2).Left
        .width = paneTab(2).width
    End With

    With vsbSE
        .Top = paneTab(2).Top
        .Left = paneTab(2).Left + paneTab(2).width
        .Height = paneTab(2).Height
    End With

    hsbSE.max = picIPane.width - paneTab(2).width
    hsbSE.LargeChange = paneTab(2).width
    hsbSE.SmallChange = vsbSE.width

    vsbSE.max = picIPane.Height - paneTab(2).Height
    vsbSE.LargeChange = paneTab(2).Height
    vsbSE.SmallChange = hsbSE.Height
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
'   �@�\: �f�[�^���Ȃ�
'
'   ���l: �Ȃ�
'
Private Sub mData_NoData()
On Error GoTo ErrorHandler
    gApp.Log "d:�Y�����R�[�h�͂���܂���BUM�̑��݂���key���w�肵�Ă��������B" & vbCrLf _
            & "�Ăяo�������`�F�b�N���܂��傤�����J����"
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �^�u�f�[�^���Ȃ�
'
'   ���l: �Ȃ�
'
Private Sub mData_NoTabData(Index As Long)
On Error GoTo ErrorHandler
    paneTab(Index).Mode = 1
    mstTab.TabEnabled(Index) = True
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �����N���x���f�[�^�̎擾����
'
'   ���l: �Ȃ�
'
Private Sub mData_FetchedLinkLabels()
On Error GoTo ErrorHandler
    Dim i As Long

    gApp.Log "FetchedLinkLabels"
    Set clblinfo(0).LinkItem = mData.LinkLabels(0)
    Set clblinfo(1).LinkItem = mData.LinkLabels(1)
    Set clblinfo(2).LinkItem = mData.LinkLabels(2)
    
    For i = 0 To 2
        With clblinfo(i)
            If .Key <> "" Then
                .ForeColor = ColorLinkExist
                .Font.Underline = True
            Else
                .ForeColor = Contrast(gApp.ColBG)
                .Font.Underline = False
            End If
        End With
    Next i
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �����ʃf�[�^�擾�����ʒm�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub mData_FetchedJokenBetu(gd2 As clsGridData, gd5 As clsGridData, gd6 As clsGridData)
On Error GoTo ErrorHandler
    gApp.Log "FetchedJokenBetu"
    Dim i As Long
    Dim r As Long, c As Long
    i = 2
    Call flexTab(i).InsertGrid(gd2)
    With flexTab(i).Grid
        .ScrollBars = flexScrollBarNone
        .BorderStyle = flexBorderNone
        Call flexTab(i).AutoSize(0, .Cols - 1)
        .FixedCols = 2
        .MergeCol(0) = True
        .MergeRow(0) = True
        .MergeCells = flexMergeFree
        .Enabled = False
                
        For r = 0 To .Rows - 1
            For c = 0 To .Cols - 1
                .row = r
                .col = c
                .CellAlignment = flexAlignRightCenter
            Next c
        Next r
    End With
    
    i = 5
    Call flexTab(i).InsertGrid(gd5)
    With flexTab(i).Grid
        .ScrollBars = flexScrollBarNone
        .BorderStyle = flexBorderNone
        Call flexTab(i).AutoSize(0, .Cols - 1)
        .FixedCols = 2
        .MergeCol(0) = True
        .MergeRow(0) = True
        .MergeCells = flexMergeFree
        .Enabled = False
        
        For r = 0 To .Rows - 1
            For c = 0 To .Cols - 1
                .row = r
                .col = c
                .CellAlignment = flexAlignRightCenter
            Next c
        Next r
    End With
    
    i = 6
    Call flexTab(i).InsertGrid(gd6)
    With flexTab(i).Grid
        .ScrollBars = flexScrollBarNone
        .BorderStyle = flexBorderNone
        Call flexTab(i).AutoSize(0, .Cols - 1)
        .FixedCols = 2
        .MergeCol(0) = True
        .MergeRow(0) = True
        .MergeCells = flexMergeFree
        .Enabled = False
        
        For r = 0 To .Rows - 1
            For c = 0 To .Cols - 1
                .row = r
                .col = c
                .CellAlignment = flexAlignRightCenter
            Next c
        Next r
    End With
        
    Call UserControl_Resize
    
    paneTab(2).Mode = 2
    paneTab(2).BorderStyle = ebscThin
    mstTab.TabEnabled(2) = True
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �����f�[�^�擾�����ʒm�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub mData_FetchedKetto(gd As clsGridData)
On Error GoTo ErrorHandler
    gApp.Log "FetchedKetto"
    Dim i As Long
    
    i = 0
    Call flexTab(i).InsertGrid(gd)
    With flexTab(i).Grid
        Call flexTab(i).AutoSize(0, .Cols - 1)
        .ColWidth(0) = 0
        .MergeCol(1) = True
        .MergeCol(2) = True
        .MergeCol(3) = True
        .MergeCells = flexMergeRestrictColumns
    End With
    paneTab(i).Mode = 2
    mstTab.TabEnabled(i) = True
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: ��H�����f�[�^�擾�����ʒm�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub mData_FetchedHanro(gd As clsGridData)
On Error GoTo ErrorHandler
    gApp.Log "FetchedHanro"
    Dim i As Long
    
    i = 4
    Call flexTab(i).InsertGrid(gd)
    If flexTab(i).Grid.Rows > 2 Then
        With flexTab(i).Grid
            flexTab(i).Grid.TextMatrix(0, 0) = " "
            flexTab(i).Grid.TextMatrix(0, 1) = " "
            flexTab(i).Grid.TextMatrix(0, 2) = " "
            
            Call flexTab(i).AutoSize(0, .Cols - 1, False, False, 0)
            
            .FixedRows = 2
            .FixedCols = 0
            .MergeRow(0) = True
            .MergeCells = flexMergeFree
            
            Dim r As Long, c As Long
            'For r = 0 To 1
            
        End With
        paneTab(i).Mode = 2
    End If
    mstTab.TabEnabled(i) = True
        
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �ߋ����f�[�^�擾�����ʒm�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub mData_FetchedKako(gd As clsGridData)
On Error GoTo ErrorHandler
    gApp.Log "FetchedKako"
    Dim i As Long
    
    i = 1
    Call flexTab(i).InsertGrid(gd)
    With flexTab(i).Grid
        Call flexTab(i).AutoSize(0, .Cols - 1)
        .FixedCols = 0
        .MergeCells = flexMergeRestrictColumns
        
        .col = 0
        .Sort = flexSortStringDescending
        Call SortFlexGrid(flexTab(i), .Cols - 1)
        
        .ColWidth(.Cols - 1) = 0
    End With
    paneTab(i).Mode = 2
    mstTab.TabEnabled(i) = True
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �����^�C���f�[�^�擾�����ʒm�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub mData_FetchedTime(gd As clsGridData)
On Error GoTo ErrorHandler
    gApp.Log "FetchedTime"
    Dim i As Long
    Dim r As Long, c As Long
    Dim strTemp As String
    
    i = 3
    Call flexTab(i).InsertGrid(gd)
    With flexTab(i).Grid
        Call flexTab(i).AutoSize(0, .Cols - 1)
        .FixedCols = 0
        .MergeCells = flexMergeRestrictColumns
        .row = 0
        
        .col = .Cols - 1
        .Sort = flexSortStringAscending
        Call SortFlexGrid(flexTab(i), .Cols - 1)
        
        For r = 1 To .Rows - 1
            .row = r
            .col = .Cols - 1
            If strTemp <> Left(.Text, 5) Then
                strTemp = Left(.Text, 5)
            Else
                .RowHeight(r) = 0
            End If
        Next r
        
        .ColWidth(.Cols - 2) = 0
        .ColWidth(.Cols - 1) = 0
    End With
    paneTab(i).Mode = 2
    mstTab.TabEnabled(i) = True
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �f�[�^���擾����
'
'   ���l: �Ȃ�
'
Private Sub Update()
    Dim i As Integer
    Dim j As Integer
    
    ' �f�[�^���擾
    If Not mData.Fetch(mKey) Then
        mblnNoData = True
        Exit Sub
    End If
    
    ' ���x�����擾
    lblMakeDate = mData.Labels(0)
    lblInfo(0) = ReplaceAmpersand(mData.Labels(1))
    txtInfo(0) = mData.Labels(2)
    txtInfo(1) = mData.Labels(3)
    txtInfo(2) = mData.Labels(4)
    txtInfo(3) = mData.Labels(5)
    txtInfo(4) = mData.Labels(6)
    Set clblinfo(0).LinkItem = mData.LinkLabels(0)
    Set clblinfo(1).LinkItem = mData.LinkLabels(1)
    Set clblinfo(2).LinkItem = mData.LinkLabels(2)
    
    '����p������ǉ�
    mstrTitle = mstrTitle & " " & mData.Labels(1)
    
    ' �ŏ��ɕ\������^�u��ݒ肷��
    If mViewerState.IsNoTouch Then
        mstTab.Tab = 0
    Else
        mstTab.Tab = mViewerState.LastTabNumber
    End If
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


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �O���֐�
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   �@�\: �I������
'
'   ���l: �Ȃ�
'
Public Sub Free()
    If Not (mData Is Nothing) Then mData.CancelFetching
    Dim i As Integer
    For i = flexTab.LBound To flexTab.UBound
        Call DestroyFlexGrid(flexTab(i))
    Next i
    Set mData = Nothing
End Sub

