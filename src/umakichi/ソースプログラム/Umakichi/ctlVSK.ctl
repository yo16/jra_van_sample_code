VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.UserControl ctlVSK 
   BackColor       =   &H00E0EEEE&
   ClientHeight    =   6390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9360
   ScaleHeight     =   6390
   ScaleWidth      =   9360
   Begin VB.Frame fraHeader 
      BackColor       =   &H00E0EEEE&
      BorderStyle     =   0  '�Ȃ�
      Caption         =   "fraHeader"
      Height          =   1035
      Left            =   180
      TabIndex        =   3
      Top             =   810
      Width           =   7185
      Begin Umakichi.ctlClickLabel clblInfo 
         Height          =   180
         Index           =   1
         Left            =   720
         TabIndex        =   4
         Top             =   210
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
         Height          =   195
         Index           =   1
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   $"ctlVSK.ctx":0000
         Top             =   210
         Visible         =   0   'False
         Width           =   730
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
         Height          =   945
         Index           =   0
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   5
         Text            =   "ctlVSK.ctx":000E
         Top             =   30
         Width           =   3795
      End
   End
   Begin VB.Frame fraTop 
      BackColor       =   &H00C0CCCC&
      BorderStyle     =   0  '�Ȃ�
      Caption         =   "�ߓc�C�j"
      Height          =   465
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8955
      Begin Umakichi.ctlClickLabel clblInfo 
         Height          =   240
         Index           =   0
         Left            =   270
         TabIndex        =   6
         Top             =   120
         Width           =   4995
         _ExtentX        =   8811
         _ExtentY        =   423
         AutoSize        =   -1  'True
         BackColor       =   12635340
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "���\�[�X�s���ł��B�s�v�ȉ�ʂ���Ă�������"
      End
   End
   Begin TabDlg.SSTab mstTab 
      Height          =   2985
      Left            =   240
      TabIndex        =   1
      Top             =   1950
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   5265
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   14741230
      TabCaption(0)   =   "����"
      TabPicture(0)   =   "ctlVSK.ctx":0046
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "paneTab(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin Umakichi.ctlPane paneTab 
         Height          =   1695
         Index           =   0
         Left            =   60
         TabIndex        =   7
         Top             =   360
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   2990
         Begin Umakichi.ctlWrappedGrid flexTab 
            Height          =   1455
            Index           =   0
            Left            =   0
            TabIndex        =   8
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
      Left            =   6540
      TabIndex        =   2
      Top             =   30
      Width           =   2175
   End
End
Attribute VB_Name = "ctlVSK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'   �Y��}�X�^ �\���R���g���[��
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
Private mViewerState    As clsVSNothing '' Viewer State

Private WithEvents mData As clsDataSK   '' �f�[�^�擾�I�u�W�F�N�g
Attribute mData.VB_VarHelpID = -1
Private mstrTitle As String             '' �E�C���h�E�^�C�g��
Private mKey As clsKeySK                '' �L�[
Private mblnNoData As Boolean           '' �f�[�^�����t���O

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
    gApp.Log "SK: " & strKey
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
'   ���l: Viewer�K�{�v���p�e�B
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
    
    Call flexTab(Index).MouseMoveDriven
    
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
    With paneTab.item(mstTab.Tab)
        .Top = mstTab.TabHeight + 60
        .Left = 60
        .width = Bigger(1, mstTab.width - .Left * 2)
        .Height = Bigger(1, (mstTab.Height - .Top) - (.Top - mstTab.TabHeight))
    End With ' paneTab.Item(mstTab.Index)
    For i = flexTab.LBound To flexTab.UBound
        With flexTab(i)
            .width = Bigger(1, paneTab(i).width - .Left)
            .Height = Bigger(1, paneTab(i).Height - .Top)
        End With ' flexTab(i)
    Next i
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
    Set mKey = New clsKeySK
    Set mData = New clsDataSK
    Set mVB = New clsViewerBase
    Set mViewerState = New clsVSNothing
    
    mstrTitle = "�Y��"

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
    
    ' Color Assign
    BackColor = gApp.ColBG
    
    clblinfo(0).BackColor = gApp.ColDarkBG
    clblinfo(0).ForeColor = Contrast(gApp.ColBG)
    clblinfo(1).BackColor = gApp.ColBG
    clblinfo(1).ForeColor = Contrast(gApp.ColBG)
    txtInfo(0).BackColor = gApp.ColBG
    txtInfo(0).ForeColor = Contrast(gApp.ColBG)
    txtInfo(1).BackColor = gApp.ColBG
    txtInfo(1).ForeColor = Contrast(gApp.ColBG)
    
    ' ctlPane �̏�����
    paneTab(0).Mode = 0 ' �f�[�^�擾����
    mstTab.TabEnabled(0) = False
    
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
    
    With paneTab.item(mstTab.Tab)
        .width = Bigger(1, mstTab.width - .Left * 2)
        .Height = Bigger(1, (mstTab.Height - .Top) - (.Top - mstTab.TabHeight))
    End With ' paneTab.Item(mstTab.Index)
    
    For i = flexTab.LBound To flexTab.UBound
        With flexTab(i)
            .width = Bigger(1, paneTab(i).width - .Left)
            .Height = Bigger(1, paneTab(i).Height - .Top)
        End With ' flexTab(i)
    Next i
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
    gApp.Log "d:�Y�����R�[�h�͂���܂���BSK�̑��݂���key���w�肵�Ă��������B" & vbCrLf _
            & "�Ăяo�������`�F�b�N���܂��傤�����J����"
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �����N���x���f�[�^�擾�����ʒm�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub mData_FetchCompleteLinklabels()
On Error GoTo ErrorHandler
    Set clblinfo(1).LinkItem = mData.LinkLabels(1)
    With clblinfo(1)
        If .Key <> "" Then
            .ForeColor = ColorLinkExist
            .Font.Underline = True
        Else
            .ForeColor = Contrast(gApp.ColBG)
            .Font.Underline = False
        End If
    End With
    txtInfo(1).Visible = True
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub

'
'   �@�\: �����^�u�̃f�[�^���������
'
'   ���l: �Ȃ�
'
Private Sub mData_FetchCompleteKetto(gd As clsGridData)
On Error GoTo ErrorHandler
    
    Call flexTab(0).InsertGrid(gd)
    
    With flexTab(0).Grid
        .FixedCols = 0
        Call flexTab(0).AutoSize(0, .Cols - 1)
        .MergeCol(0) = True
        .MergeCol(1) = True
        .MergeCells = flexMergeRestrictColumns
    End With
    paneTab(0).Mode = 2
    mstTab.TabEnabled(0) = True
    Exit Sub
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   �@�\:�@�����^�u�̃f�[�^���Ȃ�
'
'   ���l: �Ȃ�
'
Private Sub mData_NoDataKetto()
On Error GoTo ErrorHandler
    paneTab(0).Mode = 1
    mstTab.TabEnabled(0) = True
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

    ' �f�[�^���擾���Ă��炤
    If Not mData.Fetch(mKey) Then
        mblnNoData = True
        Exit Sub
    End If

    ' ���x�����擾
    lblMakeDate = mData.Labels(0)
    txtInfo(0) = mData.Labels(1)
    Set clblinfo(0).LinkItem = mData.LinkLabels(0)
    Set clblinfo(1).LinkItem = mData.LinkLabels(1)
    For i = 0 To 1
        Set clblinfo(i).LinkItem = mData.LinkLabels(i)
        With clblinfo(i)
            If .Key <> "" Then
                .ForeColor = ColorLinkExist
                .Font.Underline = True
            Else
                .ForeColor = Contrast(IIf(i = 0, gApp.ColDarkBG, gApp.ColBG))
                .Font.Underline = False
            End If
        End With
    Next i
    
    '����p������ǉ�
    mstrTitle = mstrTitle & " " & clblinfo(0).Caption
    
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

