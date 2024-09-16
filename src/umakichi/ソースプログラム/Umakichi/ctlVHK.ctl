VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.UserControl ctlVHK 
   BackColor       =   &H00E0EEEE&
   ClientHeight    =   4215
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7920
   ScaleHeight     =   4215
   ScaleWidth      =   7920
   Begin VB.Frame fraTop 
      BackColor       =   &H00C0CCCC&
      BorderStyle     =   0  '�Ȃ�
      Height          =   405
      Left            =   90
      TabIndex        =   0
      Top             =   150
      Width           =   7305
      Begin VB.Timer tmrTBS 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   6060
         Top             =   0
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   285
         Left            =   1680
         TabIndex        =   4
         Top             =   60
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         Format          =   103284736
         CurrentDate     =   37890
         MinDate         =   31048
      End
      Begin VB.Timer tmrUpdateTrigger 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   6450
         Top             =   -30
      End
      Begin VB.Label lblFix 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0CCCC&
         Caption         =   "�ύX���"
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
         Left            =   270
         TabIndex        =   1
         Top             =   90
         Width           =   1020
      End
   End
   Begin TabDlg.SSTab mstTab 
      Height          =   3105
      Left            =   180
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   660
      Width           =   6555
      _ExtentX        =   11562
      _ExtentY        =   5477
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   14741230
      TabCaption(0)   =   "�ύX�ꗗ"
      TabPicture(0)   =   "ctlVHK.ctx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "paneTab"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin Umakichi.ctlPane paneTab 
         Height          =   2055
         Left            =   60
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   360
         Width           =   5325
         _ExtentX        =   9393
         _ExtentY        =   3625
         Begin Umakichi.ctlWrappedGrid flexTab 
            Height          =   975
            Left            =   0
            TabIndex        =   5
            Top             =   0
            Width           =   1965
            _ExtentX        =   3466
            _ExtentY        =   1720
         End
      End
   End
End
Attribute VB_Name = "ctlVHK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'   �ύX���\�����[�U�[�R���g���[��
'

Option Explicit

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �O���֐�(�C�x���g)
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Public Event ChangeTo(strViewerName As String, strKey As String)
Public Event WindowTitle(strKey As String)
Public Event LinkContextMenu(strViewerName As String, strKey As String)

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �����ϐ�
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private mVB As clsViewerBase

Private mstrTitle    As String
Private mKey         As clsKeyRASel
Private mGridData    As clsGridData
Private mViewerState As clsVSDate
Private mblnNoData As Boolean           '' �f�[�^�����t���O

Private WithEvents mData As clsDataHK '' �f�[�^�擾�I�u�W�F�N�g
Attribute mData.VB_VarHelpID = -1
Private WithEvents mToolBar As ctlToolBars
Attribute mToolBar.VB_VarHelpID = -1


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �v���p�e�B
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8


'
'   �@�\: �L�[�ݒ�v���p�e�B
'
'   ���l: Viewer�K�{�v���p�e�B
'
Public Property Let Key(RHS As String)
    ' �������̃L�[��8�����łȂ���΁A���݂̔N�ƑS�ꏊ�ɐݒ�
    If Len(RHS) = 8 Then
        mKey.str = RHS
    Else
        mKey.str = Format$(Now, "YYYYMMDD")
    End If

    If Not mViewerState.IsNoTouch Then
        mKey.str = mViewerState.YMD
    End If
    
    dtpDate.Enabled = False
    dtpDate.value = Format$(mKey.str, "##/##/##")
    dtpDate.Enabled = True

    tmrUpdateTrigger.Enabled = True
End Property


'
'   �@�\: �^�C�g���擾�v���p�e�B
'
'   ���l: Viewer�K�{�v���p�e�B�A�@Browser ���Ăт܂�
'
Public Property Get Title() As String
    Title = mstrTitle
End Property


'
'   �@�\: �^�C�g���ݒ�v���p�e�B
'
'   ���l: �u���E�U�ɕύX�ʒm�̃C�x���g�𓊂��܂�
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
Public Property Get ViewerState() As clsVSDate
    Set ViewerState = mViewerState
End Property


'
'   �@�\: Viewer��Ԏ󂯎��
'
'   ���l: �Ȃ�
'
Public Property Set ViewerState(RHS As clsVSDate)
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


'
'   �@�\: �c�[���o�[��ݒ肷��
'
'   ���l: RA, OD �̂݁A�K�{�v���p�e�B
'         �u���E�U����c�[���o�[�����炤�ƁA
'         �c�[���o�[���Z�b�g�A�b�v����
'
Public Property Set ToolBar(RHS As ctlToolBars)
    Set mToolBar = RHS

    With mToolBar.ToolBar(2)
        .Buttons(1).Caption = "�J�Ï��擾"
    End With
End Property


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �����֐�
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   �@�\: �N���b�N�C�x���g
'
'   ���l: �����N���ʂ֐؂�ւ���
'
Private Sub flexTab_Click()
On Error GoTo ErrorHandler
    Dim msrow As Long             '' �}�E�X���E
    Dim mscol As Long             '' �}�E�X�J����
    Dim item As clsGridItem       '' �O���b�h�A�C�e��

    ' �}�E�X�ʒu�̃O���b�h���W���擾
    With flexTab.Grid
        msrow = .MouseRow
        mscol = .MouseCol
    End With

    ' �O���b�h�A�C�e�����Z��������o��
    Call SetItem(item, flexTab, msrow, mscol)

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
Private Sub flexTab_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrorHandler
    Dim msrow As Long
    Dim mscol As Long
    Dim item As clsGridItem

    ' �}�E�X�̎����O���b�h���W���擾
    msrow = flexTab.Grid.MouseRow
    mscol = flexTab.Grid.MouseCol

    If mscol < 0 Or msrow < 0 Then
        Exit Sub
    End If

    Call SetItem(item, flexTab, msrow, mscol)

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
Private Sub flexTab_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrorHandler

    Call flexTab.ReflexiveMouseMoveDriven(True)

    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �c�[���o�[�̃{�^���N���b�N�C�x���g
'
'   ���l: �^�C�}�[���C�l�[�u���ɂ���
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
'   �@�\: �^�C�}�[�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub tmrTBS_Timer()
On Error GoTo ErrorHandler
    tmrTBS.Enabled = False
    Select Case tmrTBS.tag
    Case "RTOPEN"
        Call gApp.DBPrompt(ukpRA, Left$(mKey.str, 8))
    End Select
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �A�b�v�f�[�g�^�C�}�[�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub tmrUpdateTrigger_Timer()
On Error GoTo ErrorHandler
    paneTab.Mode = 0
    mData.CancelFetching
    If Not mData.NowFetching Then
        tmrUpdateTrigger.Enabled = False
        Call mData.Fetch(mKey)
    End If
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub



'
'   �@�\: �N�����̕ύX�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub dtpDate_Change()
On Error GoTo ErrorHandler

    If dtpDate.Enabled Then
        mKey.str = Format$(dtpDate.value, "YYYYMMDD")
        mViewerState.YMD = mKey.Year & mKey.MonthDay
        tmrUpdateTrigger.Enabled = True
    End If
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �N�����̃L�[���̓C�x���g
'
'   ���l: �Ȃ�
'
Private Sub dtpDate_KeyPress(KeyAscii As Integer)
On Error GoTo ErrorHandler
   If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
      KeyAscii = 0      ' �������������܂��B
      Beep            ' �G���[����炵�܂��B
   End If
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

    Set mVB = New clsViewerBase     '' ViewerBase �I�u�W�F�N�g
    Set mData = New clsDataHK       '' �f�[�^�擾�I�u�W�F�N�g
    Set mKey = New clsKeyRASel
    Set mViewerState = New clsVSDate

    gApp.InitLog Me
    mstrTitle = "�ύX���"

    ' FlexGrid�ݒ�
    Call mVB.FlexGridCommonSetting(flexTab.Grid)
    With flexTab.Grid
        .FixedCols = 0
        .FixedRows = 1
    End With

    ' Color Assign
    BackColor = gApp.ColBG
    mstTab.BackColor = gApp.ColBG
    fraTop.BackColor = gApp.ColDarkBG

    lblFix.BackColor = gApp.ColDarkBG
    lblFix.ForeColor = Contrast(gApp.ColDarkBG)

    paneTab.Mode = 0

    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: ���[�U�R���g���[���̂�T�C�Y�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub UserControl_Resize()
On Error GoTo ErrorHandler
    Dim i As Integer

    fraTop.width = Bigger(1, ScaleWidth - fraTop.Left * 2)
    With mstTab
        .width = Bigger(1, ScaleWidth - .Left * 2)
        .Height = Bigger(1, ScaleHeight - .Top - .Left)
    End With ' mstTab

    With paneTab
        .width = Bigger(1, mstTab.width - .Left * 2)
        .Height = Bigger(1, (mstTab.Height - .Top) - (.Top - mstTab.TabHeight))
    End With ' fraTab.Item(mstTab.Index)

    With flexTab
        .width = Bigger(1, paneTab.width - .Left)
        .Height = Bigger(1, paneTab.Height - .Top)
    End With
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

    Set mVB = Nothing
    
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

    paneTab.Mode = 0

    Refresh

    Call mData.Fetch(mKey)
End Sub


'
'   �@�\: �f�[�^���������
'
'   ���l: �Ȃ�
'
Private Sub mData_FetchComplete(gd As clsGridData)
On Error GoTo ErrorHandler
    Dim i As Long

    Call flexTab.Grid.Clear
    Call flexTab.InsertGrid(gd)
    Call flexTab.AutoSize(0, flexTab.Grid.Cols - 1)

    With flexTab.Grid
        .col = 2
        .Sort = flexSortGenericAscending
        .MergeRow(0) = True
        .MergeCells = flexMergeFree
    End With

    paneTab.Mode = 2
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �f�[�^���Ȃ�����
'
'   ���l: �Ȃ�
'
Private Sub mData_NoData()
On Error GoTo ErrorHandler
    paneTab.Mode = 1
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
'   ���l: �u���E�U���A�����j������܂��ɌĂт܂�
'
Public Sub Free()
    gApp.Log "Free"
        
    Call DestroyFlexGrid(flexTab)
    
    Set mData = Nothing
End Sub

