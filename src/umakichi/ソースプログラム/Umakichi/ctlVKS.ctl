VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.UserControl ctlVKS 
   BackColor       =   &H00E0EEEE&
   ClientHeight    =   5535
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9120
   ScaleHeight     =   5535
   ScaleWidth      =   9120
   Begin VB.Frame fraHeader 
      BackColor       =   &H00E0EEEE&
      BorderStyle     =   0  '�Ȃ�
      Caption         =   "fraHeader"
      Height          =   1005
      Left            =   330
      TabIndex        =   4
      Top             =   810
      Width           =   7545
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
         Text            =   "ctlVKS.ctx":0000
         Top             =   30
         Width           =   3195
      End
      Begin Umakichi.ctlClickLabel clblInfo 
         Height          =   180
         Index           =   0
         Left            =   4350
         TabIndex        =   18
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
         Height          =   660
         Index           =   1
         Left            =   3270
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   6
         Text            =   "ctlVKS.ctx":0047
         Top             =   30
         Width           =   3255
      End
   End
   Begin VB.Frame fraTop 
      BackColor       =   &H00C0CCCC&
      BorderStyle     =   0  '�Ȃ�
      Caption         =   "Frame1"
      Height          =   645
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8955
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0CCCC&
         Caption         =   "."
         Height          =   180
         Index           =   1
         Left            =   270
         TabIndex        =   7
         Top             =   60
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
         Left            =   210
         TabIndex        =   1
         Top             =   300
         Width           =   5010
      End
   End
   Begin TabDlg.SSTab mstTab 
      Height          =   2985
      Left            =   210
      TabIndex        =   2
      Top             =   2040
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   5265
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      BackColor       =   14741230
      TabCaption(0)   =   "��������"
      TabPicture(0)   =   "ctlVKS.ctx":0073
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "paneTab(3)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "paneTab(2)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "paneTab(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "paneTab(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "�����ʐ���"
      TabPicture(1)   =   "ctlVKS.ctx":008F
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "���n��ʐ���"
      TabPicture(2)   =   "ctlVKS.ctx":00AB
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "���R��"
      TabPicture(3)   =   "ctlVKS.ctx":00C7
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "paneTab(4)"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "��v����"
      TabPicture(4)   =   "ctlVKS.ctx":00E3
      Tab(4).ControlEnabled=   0   'False
      Tab(4).ControlCount=   0
      Begin Umakichi.ctlPane paneTab 
         Height          =   1695
         Index           =   0
         Left            =   60
         TabIndex        =   8
         Top             =   360
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   2990
         Begin Umakichi.ctlWrappedGrid flexTab 
            Height          =   1455
            Index           =   0
            Left            =   0
            TabIndex        =   9
            Top             =   0
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   2566
         End
      End
      Begin Umakichi.ctlPane paneTab 
         Height          =   1695
         Index           =   1
         Left            =   60
         TabIndex        =   10
         Top             =   360
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   2990
         Begin Umakichi.ctlWrappedGrid flexTab 
            Height          =   1455
            Index           =   1
            Left            =   0
            TabIndex        =   11
            Top             =   0
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   2566
         End
      End
      Begin Umakichi.ctlPane paneTab 
         Height          =   1695
         Index           =   2
         Left            =   60
         TabIndex        =   12
         Top             =   360
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   2990
         Begin Umakichi.ctlWrappedGrid flexTab 
            Height          =   1455
            Index           =   2
            Left            =   0
            TabIndex        =   13
            Top             =   0
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   2566
         End
      End
      Begin Umakichi.ctlPane paneTab 
         Height          =   1695
         Index           =   3
         Left            =   60
         TabIndex        =   14
         Top             =   360
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   2990
         Begin Umakichi.ctlWrappedGrid flexTab 
            Height          =   1455
            Index           =   3
            Left            =   0
            TabIndex        =   15
            Top             =   0
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   2566
         End
      End
      Begin Umakichi.ctlPane paneTab 
         Height          =   1695
         Index           =   4
         Left            =   -74940
         TabIndex        =   16
         Top             =   360
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   2990
         Begin Umakichi.ctlWrappedGrid flexTab 
            Height          =   1455
            Index           =   4
            Left            =   0
            TabIndex        =   17
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
      Left            =   6570
      TabIndex        =   3
      Top             =   60
      Width           =   2175
   End
End
Attribute VB_Name = "ctlVKS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'   �R��}�X�^ �\���R���g���[��
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

Private WithEvents mData As clsDataKS   '' �f�[�^�擾�I�u�W�F�N�g
Attribute mData.VB_VarHelpID = -1
Private mstrTitle As String             '' �E�C���h�E�^�C�g��
Private mKey As clsKeyKS                '' �L�[
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
    gApp.Log "KS: " & strKey
    mKey.str = strKey
    Call Update
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
    
    If mscol < 0 Or msrow < 0 Then
        Exit Sub
    End If
    
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
    
    If mscol < 0 Or msrow < 0 Then
        Exit Sub
    End If
    
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
    
    If 4 = Index Then
        Call flexTab(Index).ReflexiveMouseMoveDriven(True)
    ElseIf 3 = Index Then
        Call flexTab(Index).ReflexiveMouseMoveDriven(True)
    Else
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
    
    With paneTab.item(mstTab.Tab)
        .Top = mstTab.TabHeight + 60
        .Left = 60
        .width = Bigger(1, mstTab.width - .Left * 2)
        .Height = Bigger(1, (mstTab.Height - .Top) - (.Top - mstTab.TabHeight))
    End With ' fraTab.Item(mstTab.Index)
    For i = flexTab.LBound To flexTab.UBound
        With flexTab(i)
            .width = Bigger(1, paneTab(i).width - .Left)
            .Height = Bigger(1, paneTab(i).Height - .Top)
        End With ' flexTab(i)
    Next i
    
    mViewerState.LastTabNumber = mstTab.Tab
    
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
    Set mKey = New clsKeyKS
    Set mData = New clsDataKS
    Set mVB = New clsViewerBase
    Set mViewerState = New clsVSTabOnly
    
    gApp.Log "KISYU Initialize start"
    
    mstrTitle = "�R��"

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
    
    With clblinfo(0)
        .BackColor = gApp.ColBG
    End With
    For i = lblInfo.LBound To lblInfo.UBound
        lblInfo(i).BackColor = gApp.ColDarkBG
        lblInfo(i).ForeColor = Contrast(gApp.ColDarkBG)
    Next i
    
    For i = txtInfo.LBound To txtInfo.UBound
        txtInfo(i).BackColor = gApp.ColBG
        txtInfo(i).ForeColor = Contrast(gApp.ColBG)
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
    
    gApp.Log "KISYU Resize start"
    
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
    End With ' fraTab.Item(mstTab.Index)
    
    For i = flexTab.LBound To flexTab.UBound
        With flexTab(i)
            .width = Bigger(1, paneTab(i).width - .Left)
            .Height = Bigger(1, paneTab(i).Height - .Top)
        End With ' flexTab(i)
    Next i
    
    gApp.Log "KISYU Resize end"
    
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
'   �@�\: �f�[�^���Ȃ�
'
'   ���l: �Ȃ�
'
Private Sub mData_NoData()
On Error GoTo ErrorHandler
    gApp.Log "d:�Y�����R�[�h�͂���܂���BKS�̑��݂���key���w�肵�Ă��������B" & vbCrLf _
            & "�Ăяo�������`�F�b�N���܂��傤�����J����"
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

    ' �f�[�^���擾���Ă��炤
    If Not mData.Fetch(mKey) Then
        mblnNoData = True
        Exit Sub
    End If

    ' ���x�������炤
    lblMakeDate = mData.Labels(0)
    lblInfo(0) = ReplaceAmpersand(mData.Labels(1))
    lblInfo(1) = ReplaceAmpersand(mData.Labels(2))
    txtInfo(0) = mData.Labels(3)
    txtInfo(1) = mData.Labels(4)
    If Not mData.LinkLabels(0) Is Nothing Then
        Set clblinfo(0).LinkItem = mData.LinkLabels(0)
        With clblinfo(0)
            If .Key <> "" Then
                .ForeColor = ColorLinkExist
                .Font.Underline = True
            Else
                .ForeColor = Contrast(gApp.ColBG)
                .Font.Underline = False
            End If
        End With
    Else
        clblinfo(0).Caption = ""
    End If
    
    '����p������ǉ�
    mstrTitle = mstrTitle & " " & mData.Labels(1)
    
    '�O���b�h�̑}��
    For i = 0 To 4
        Call flexTab(i).InsertGrid(mData.GridDatas(i))
    Next i
    
    '�}�[�W�A�񂹁A���̒���
    Dim r As Long, c As Long
    With flexTab(0).Grid
        Call flexTab(0).AutoSize(0, .Cols - 1, False, False, 0)
        .FixedRows = 1
        For c = 0 To .Cols - 1
            .row = 0
            .col = c
            .CellAlignment = flexAlignCenterCenter
        Next c
        For r = 1 To .Rows - 1
            For c = 0 To .Cols - 1
                .row = r
                .col = c
                .CellAlignment = flexAlignRightCenter
            Next c
        Next r
    End With
    
    With flexTab(1).Grid
        Call flexTab(1).AutoSize(1, .Cols - 1, False, False, 1)
        .FixedRows = 2
        .MergeRow(0) = True
        .MergeRow(1) = True
        .MergeCol(0) = True
        .MergeCells = flexMergeFree
        
        For r = 0 To 1
            For c = 0 To .Cols - 1
                .row = r
                .col = c
                .CellAlignment = flexAlignCenterCenter
            Next c
        Next r
        For r = 2 To .Rows - 1
            For c = 0 To .Cols - 1
                .row = r
                .col = c
                .CellAlignment = flexAlignRightCenter
            Next c
        Next r
    End With
    
    With flexTab(2).Grid
        Call flexTab(2).AutoSize(0, .Cols - 1, False, False, 2)
        .FixedRows = 3
        .MergeRow(0) = True
        .MergeRow(1) = True
        .MergeCol(0) = True
        .MergeCells = flexMergeFree
        For r = 0 To 1
            For c = 0 To .Cols - 1
                .row = r
                .col = c
                .CellAlignment = flexAlignCenterCenter
            Next c
        Next r
        For r = 2 To .Rows - 1
            For c = 0 To .Cols - 1
                .row = r
                .col = c
                .CellAlignment = flexAlignRightCenter
            Next c
        Next r
    End With
    
    With flexTab(3).Grid
        Call flexTab(3).AutoSize(0, .Cols - 1, False, False, 0)
        .FixedCols = 2
        .MergeRow(0) = True
        .MergeCol(0) = True
        .MergeCells = flexMergeFree
        For r = 1 To .Rows - 1
            .row = r
            .col = 4
            .CellAlignment = flexAlignRightCenter
        Next r
    End With
    
    With flexTab(4).Grid
        Call flexTab(4).AutoSize(0, .Cols - 1, False, False, 0)
        .FixedCols = 0
        For r = 1 To .Rows - 1
            .row = r
            .col = 2
            .CellAlignment = flexAlignRightCenter
        Next r
    End With
    
    ' �ŏ��ɕ\������^�u��ݒ肷��
    If mViewerState.IsNoTouch Then
        mstTab.Tab = 0
    Else
        mstTab.Tab = mViewerState.LastTabNumber
    End If
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

