VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.UserControl ctlVHCSel 
   BackColor       =   &H00E0EEEE&
   ClientHeight    =   5475
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9240
   ScaleHeight     =   5475
   ScaleWidth      =   9240
   Begin VB.Frame fraTop 
      BackColor       =   &H00C0CCCC&
      BorderStyle     =   0  '�Ȃ�
      Caption         =   "Frame1"
      Height          =   465
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   8955
      Begin VB.Timer tmrTrigger 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   6690
         Top             =   30
      End
      Begin VB.ComboBox cmbBasho 
         Height          =   300
         ItemData        =   "ctlVHCSel.ctx":0000
         Left            =   4020
         List            =   "ctlVHCSel.ctx":000D
         Style           =   2  '��ۯ���޳� ؽ�
         TabIndex        =   4
         Top             =   90
         Width           =   855
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   285
         Left            =   1980
         TabIndex        =   2
         Top             =   90
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   503
         _Version        =   393216
         CalendarBackColor=   14741230
         CalendarTitleBackColor=   12635340
         Format          =   103219200
         CurrentDate     =   37809
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  '�E����
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "��"
         Height          =   180
         Index           =   1
         Left            =   8715
         TabIndex        =   5
         Top             =   120
         Width           =   180
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0CCCC&
         Caption         =   "��H����"
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
         Top             =   120
         Width           =   1035
      End
   End
   Begin TabDlg.SSTab mstTab 
      Height          =   2985
      Left            =   120
      TabIndex        =   3
      Top             =   690
      Width           =   8145
      _ExtentX        =   14367
      _ExtentY        =   5265
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   14741230
      TabCaption(0)   =   "��H�����ꗗ"
      TabPicture(0)   =   "ctlVHCSel.ctx":0023
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "paneTab"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin Umakichi.ctlPane paneTab 
         Height          =   2055
         Left            =   60
         TabIndex        =   6
         Top             =   360
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   3625
         Begin Umakichi.ctlWrappedGrid flexTab 
            Height          =   1095
            Left            =   0
            TabIndex        =   7
            Top             =   0
            Width           =   2595
            _ExtentX        =   4577
            _ExtentY        =   1931
         End
      End
   End
End
Attribute VB_Name = "ctlVHCSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'   ��H�ꗗ�I�� �\���R���g���[��
'

Option Explicit

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �O���֐�(�C�x���g)
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Public Event ChangeTo(strViewerName As String, strKey As String)    '' Vierer�ύX�C�x���g
Public Event WindowTitle(strKey As String)                          '' �E�C���h�E�^�C�g���ύX�C�x���g
Public Event LinkContextMenu(strViewerName As String, strKey As String)

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �����ϐ�
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private mVB             As clsViewerBase
Private mViewerState    As clsVSDateJyo

Private mstrTitle As String              '' �E�C���h�E�^�C�g��
Private mKey      As clsKeyHCSel         '' �L�[
Private WithEvents mData     As clsDataHCSel        '' �f�[�^�擾�I�u�W�F�N�g
Attribute mData.VB_VarHelpID = -1
Private mblnNoData As Boolean           '' �f�[�^�����t���O


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �����萔
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Const strTitle As String = "��H����"


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �v���p�e�B
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8


'
'   �@�\: �L�[�ݒ�v���p�e�B
'
'   ���l: Viewer�K�{�v���p�e�B
'
Public Property Let Key(strKey As String)
On Error GoTo ErrorHandler
    If strKey <> "Empty" Then
        mKey.str = strKey
    Else
        mKey.str = Format$(Now, "YYYYMMDD") & "0"
    End If
    
    If Not mViewerState.IsNoTouch Then
        mKey.str = mViewerState.YMD & mViewerState.Jyo
    End If
    
    cmbBasho.tag = "Disenable"
    dtpDate.tag = "Disenable"
        dtpDate.value = Format$(mKey.Year & mKey.MonthDay, "@@@@/@@/@@")
        cmbBasho.ListIndex = val(mKey.Basho)
    dtpDate.tag = ""
    cmbBasho.tag = ""
    
    tmrTrigger.Enabled = True
    Exit Property
ErrorHandler:
    gApp.ErrLog
    Resume Next
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
Public Property Get ViewerState() As clsVSDateJyo
    Set ViewerState = mViewerState
End Property


'
'   �@�\: Viewer��Ԏ󂯎��
'
'   ���l: �Ȃ�
'
Public Property Set ViewerState(RHS As clsVSDateJyo)
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
'   �@�\: �f�[�^���擾����
'
'   ���l: �Ȃ�
'
Private Sub Update()
    Dim mp As New clsPointer
    
    paneTab.Mode = 0
    
    ' �f�[�^���擾����
    Call mData.Fetch(mKey)

End Sub


'
'   �@�\: �J�Ï�ύX
'
'   ���l: �Ȃ�
'
Private Sub cmbBasho_Change()
On Error GoTo ErrorHandler
    
    If cmbBasho.tag = "" Then
        tmrTrigger.Enabled = True
    End If

    ' ���������x���ɕ\������
    lblInfo(1).Caption = ""
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �J�Ï�I��
'
'   ���l: �Ȃ�
'
Private Sub cmbBasho_Click()
On Error GoTo ErrorHandler
    mKey.str = Format$(dtpDate.value, "yyyymmdd") & cmbBasho.ListIndex
    If cmbBasho.tag = "" Then
        tmrTrigger.Enabled = True
    End If
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �J�Ó��ύX
'
'   ���l: �Ȃ�
'
Private Sub dtpDate_Change()
On Error GoTo ErrorHandler
    mKey.str = Format$(dtpDate.value, "yyyymmdd") & mKey.Basho
    If dtpDate.tag = "" Then
        tmrTrigger.Enabled = True
    End If

    ' ���������x���ɕ\������
    lblInfo(1).Caption = ""
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �X�N���[����̍s���v���p�e�B�ɐݒ�
'
'   ���l: �Ȃ�
'
Private Sub flexTab_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
On Error GoTo ErrorHandler
    mViewerState.row = NewTopRow
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
    paneTab.Mode = 1
    dtpDate.Enabled = True
    ' ���������x���ɕ\������
    lblInfo(1).Caption = "0��"
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �g���K�^�C�}�[�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub tmrTrigger_Timer()
On Error GoTo ErrorHandler
    paneTab.Mode = 0
    
    mData.CancelFetching
    If Not mData.NowFetching Then
        tmrTrigger.Enabled = False
        mViewerState.YMD = mKey.Year & mKey.MonthDay
        mViewerState.Jyo = mKey.Basho
        mViewerState.row = 0

        Call Update
    End If
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
    Resume Next
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
    
    Set mKey = New clsKeyHCSel
    Set mData = New clsDataHCSel
    Set mVB = New clsViewerBase
    Set mViewerState = New clsVSDateJyo
    
    mstrTitle = strTitle
    

    ' ���������x���ɕ\������
    lblInfo(1).Caption = ""

    ' FlexGrid�ݒ�
    Call mVB.FlexGridCommonSetting(flexTab.Grid)
    
    ' Color Assign
    BackColor = gApp.ColBG
    mstTab.BackColor = gApp.ColBG
    fraTop.BackColor = gApp.ColDarkBG
       
    For i = lblInfo.LBound To lblInfo.UBound
        lblInfo(i).BackColor = gApp.ColDarkBG
        lblInfo(i).ForeColor = Contrast(gApp.ColDarkBG)
    Next i
    
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
    End With ' fraTab
    
    With flexTab
        .width = Bigger(1, paneTab.width - .Left)
        .Height = Bigger(1, paneTab.Height - .Top)
    End With ' flexTab(i)
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �N���b�N�C�x���g
'
'   ���l: �����N���ʂ֐؂�ւ���
'
Private Sub flexTab_Click()
On Error GoTo ErrorHandler
    Dim msrow As Long             '' �}�E�X���E
    Dim mscol As Long             '' �}�E�X�J����
    Dim item As clsGridItem     '' �O���b�h�A�C�e��
    
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
    
    Call flexTab.MouseMoveDriven
    
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
'   �@�\: �f�[�^���������
'
'   ���l: �Ȃ�
'
Private Sub mData_FetchComplete(gd As clsGridData)
On Error GoTo ErrorHandler
    If mKey.str = "" Then
        dtpDate.value = mData.FetchDate
    End If
    dtpDate.Enabled = True
        
    ' ���������x���ɕ\������
    lblInfo(1).Caption = mData.NumRecord & "��"
        
    ' �O���b�h�f�[�^���󂯎��
    ' �O���b�h�f�[�^���R���g���[���ɔ��f����
    If gd.Cols > 2 Then
        flexTab.Grid.Redraw = False
        flexTab.Grid.Clear
        
        Call flexTab.InsertGrid(gd)
        
        With flexTab.Grid
            flexTab.Grid.TextMatrix(0, 0) = " "
            flexTab.Grid.TextMatrix(0, 1) = " "
            flexTab.Grid.TextMatrix(0, 2) = " "

            .FixedRows = 2
            .FixedCols = 0
            .MergeRow(0) = True
            .MergeRow(1) = False
            
            .MergeCells = flexMergeRestrictRows
            .ColWidth(0) = 870
            .ColWidth(1) = 555
            .ColWidth(2) = 1890
            
            Dim i As Integer
            For i = 3 To .Cols - 3 Step 2
                .ColWidth(i) = 585
                .ColWidth(i + 1) = 840
            Next
            
            .Redraw = True
            .Visible = True
        End With
        
        '�O���b�h��\������
        paneTab.Mode = 2
        
        If Not mViewerState.IsNoTouch And flexTab.Grid.Rows >= mViewerState.row Then
            If mViewerState.row > 0 Then
                flexTab.Grid.TopRow = mViewerState.row
            End If
        End If
    Else
        flexTab.Grid.Visible = False
        '�f�[�^������܂���
        paneTab.Mode = 1
    End If
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
'
Public Sub Free()
    gApp.Log "Free"
    
    Call DestroyFlexGrid(flexTab)
    
    Set mData = Nothing
End Sub

