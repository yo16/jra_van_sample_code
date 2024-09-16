VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.UserControl ctlVRAKaiSel 
   BackColor       =   &H00E0EEEE&
   ClientHeight    =   5100
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7020
   ScaleHeight     =   5100
   ScaleWidth      =   7020
   Begin VB.Frame fraTop 
      BackColor       =   &H00C0CCCC&
      BorderStyle     =   0  '�Ȃ�
      Caption         =   "Frame1"
      Height          =   465
      Left            =   90
      TabIndex        =   0
      Top             =   150
      Width           =   6615
      Begin VB.ComboBox cmbYear 
         Height          =   300
         Left            =   420
         TabIndex        =   6
         Text            =   "2000"
         Top             =   90
         Width           =   795
      End
      Begin VB.Timer tmrUpdateTrigger 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   5820
         Top             =   120
      End
      Begin VB.ComboBox cmbJyoCD 
         Height          =   300
         ItemData        =   "ctlVRAKaiSel.ctx":0000
         Left            =   1590
         List            =   "ctlVRAKaiSel.ctx":0026
         Style           =   2  '��ۯ���޳� ؽ�
         TabIndex        =   2
         Top             =   90
         Width           =   1005
      End
      Begin VB.Label lblFix 
         Appearance      =   0  '�ׯ�
         AutoSize        =   -1  'True
         BackColor       =   &H00C0CCCC&
         Caption         =   "�N"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   1260
         TabIndex        =   3
         Top             =   150
         Width           =   180
      End
   End
   Begin TabDlg.SSTab mstTab 
      Height          =   3225
      Left            =   90
      TabIndex        =   1
      Top             =   780
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   5689
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   14741230
      TabCaption(0)   =   "�J�Èꗗ"
      TabPicture(0)   =   "ctlVRAKaiSel.ctx":006C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "paneTab"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin Umakichi.ctlPane paneTab 
         Height          =   2475
         Left            =   60
         TabIndex        =   4
         Top             =   360
         Width           =   4785
         _ExtentX        =   8440
         _ExtentY        =   4366
         Begin Umakichi.ctlWrappedGrid flexTab 
            Height          =   1095
            Left            =   0
            TabIndex        =   5
            Top             =   0
            Width           =   2265
            _ExtentX        =   3995
            _ExtentY        =   1931
         End
      End
   End
End
Attribute VB_Name = "ctlVRAKaiSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'   �o�n�\�J�ÑI�� �\���R���g���[��
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
Private mViewerState As clsVSYearJyo

Private mstrTitle As String
Private mKey As clsKeyRAKaiSel
Private mblnNoData As Boolean           '' �f�[�^�����t���O

Private WithEvents mData As clsDataRAKaiSel '' �f�[�^�擾�I�u�W�F�N�g
Attribute mData.VB_VarHelpID = -1


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �v���p�e�B
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8


'
'   �@�\: �L�[�ݒ�v���p�e�B
'
'   ���l: Viewer�K�{�v���p�e�B
'
Public Property Let Key(RHS As String)
    ' �������̃L�[��6�����łȂ���΁A���݂̔N�ƑS�ꏊ�ɐݒ�
    If Len(RHS) = 6 Then
        mKey.Str = RHS
    Else
        mKey.Str = Format$(Now, "YYYY") & "00"
    End If
    
    If Not mViewerState.IsNoTouch Then
        mKey.Year = mViewerState.Year
        mKey.JyoCD = mViewerState.Jyo
    End If
    
    cmbJyoCD.Enabled = False
    cmbYear.Enabled = False
        cmbJyoCD.ListIndex = val(mKey.JyoCD)
        cmbYear.Text = mKey.Year
    cmbYear.Enabled = True
    cmbJyoCD.Enabled = True
    
    tmrUpdateTrigger.Enabled = True
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
Public Property Get ViewerState() As clsVSYearJyo
    Set ViewerState = mViewerState
End Property


'
'   �@�\: Viewer��Ԏ󂯎��
'
'   ���l: �Ȃ�
'
Public Property Set ViewerState(RHS As clsVSYearJyo)
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
'   �@�\: �ꏊ�R���{�{�b�N�X�N���b�N�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub cmbJyoCD_Click()
On Error GoTo ErrorHandler
    If cmbJyoCD.Enabled Then
        mKey.JyoCD = Right$("0" & cmbJyoCD.ItemData(cmbJyoCD.ListIndex), 2)
        mViewerState.Jyo = mKey.JyoCD
        tmrUpdateTrigger.Enabled = True
    End If
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �J�ÔN�R���{�{�b�N�X�N���b�N�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub cmbYear_Click()
On Error GoTo ErrorHandler
    If cmbYear.Enabled Then
        mKey.Year = Format$(val(cmbYear.Text), "0000")
        mViewerState.Year = mKey.Year
        tmrUpdateTrigger.Enabled = True
    End If
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �J�ÔN�R���{�{�b�N�X�L�[���̓C�x���g
'
'   ���l: �Ȃ�
'
Private Sub cmbYear_KeyPress(KeyAscii As Integer)
On Error GoTo ErrorHandler
    If KeyAscii = 13 Then
        Call cmbYear_Click
    ElseIf (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And Not KeyAscii = 8 Then
        KeyAscii = 0      ' �������������܂��B
        Beep            ' �G���[����炵�܂��B
    End If
    cmbYear.Text = Format$(val(cmbYear.Text), "0000")
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �J�ÔN�R���{�{�b�N�X���X�g�t�H�[�J�X�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub cmbYear_LostFocus()

End Sub


'
'   �@�\: �X�N���[����C�x���g
'
'   ���l: �s�̕ۑ�
'
Private Sub flexTab_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
On Error GoTo ErrorHandler
    mViewerState.row = NewTopRow
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �\�[�g��C�x���g
'
'   ���l: �s�̕ۑ�
'
Private Sub flexTab_AfterSort(ByVal col As Long, Order As Integer)
On Error GoTo ErrorHandler
    Dim i As Long
    Dim item As clsGridItem
    
    With flexTab.Grid
        For i = 0 To .Rows - 1
            Call SetItem(item, flexTab, i, 0)
            If item.Key <> "" Then
                .CellBackColor = IIf(i Mod 2 = 0, RGB(240, 240, 255), RGB(223, 223, 255))
            End If
        Next i
    End With
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
'   �@�\: �A�b�v�f�[�g�^�C�}�[
'
'   ���l: �Ȃ�
'
Private Sub tmrUpdateTrigger_Timer()
On Error GoTo ErrorHandler
    paneTab.Mode = 0
    If Not mData Is Nothing Then
        mData.CancelFetching
        If Not mData.NowFetching Then
            tmrUpdateTrigger.Enabled = False
            Call mData.Fetch(mKey)
            Call setCmbYearItems
        End If
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
    Set mKey = New clsKeyRAKaiSel   '' �L�[�I�u�W�F�N�g
    Set mData = New clsDataRAKaiSel '' �f�[�^�擾�I�u�W�F�N�g
    Set mViewerState = New clsVSYearJyo
    
    gApp.InitLog Me
    mstrTitle = "�o�n�\�J�ÑI��"
    
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
'   �@�\: ���[�U�R���g���[���̃��T�C�Y�C�x���g
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
    End With ' flexTab.Grid(i)
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
    Set mKey = Nothing
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �����֐�
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   �@�\: �f�[�^���擾����
'
'   ���l: �Ȃ�
'
Private Sub Update()
    
    paneTab.Mode = 0
    
    Refresh
    
    Call mData.Fetch(mKey)

    Call setCmbYearItems
End Sub


'
'   �@�\: �J�ÔN�R���{�{�b�N�X�Ƀf�[�^���Z�b�g
'
'   ���l: �Ȃ�
'
Private Sub setCmbYearItems()
    Dim Y As String
    Dim i As Long
    Dim yl() As String
        
    yl = mData.YearList
    
    Y = cmbYear.Text
    cmbYear.Clear
    cmbYear.Text = Y
    For i = 0 To UBound(yl)
        cmbYear.AddItem yl(i)
    Next i
End Sub


'
'   �@�\: �f�[�^�擾�����ʒm�C�x���g
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
        .col = 0
        .MergeRow(0) = True
        .MergeCells = flexMergeFree
        
        ' �s�̈ړ�
        If mViewerState.IsNoTouch Then
            For i = 0 To .Rows - 1
                If flexTab.HasKey(i, 0) Then Exit For
            Next i
        Else
            i = IIf(mViewerState.row > 0, mViewerState.row, 1)
        End If
        
        If i < .Rows Then
            .TopRow = i
            .col = 0
            .row = i
        End If
    End With

    paneTab.Mode = 2
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �f�[�^������
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
    
    Set mKey = Nothing
    Set mData = Nothing
    Set mData = Nothing
End Sub

