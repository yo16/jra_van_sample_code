VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.UserControl ctlVRASel 
   BackColor       =   &H00E0EEEE&
   ClientHeight    =   6855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9495
   ScaleHeight     =   6855
   ScaleWidth      =   9495
   Begin VB.Frame fraTop 
      BackColor       =   &H00C0CCCC&
      BorderStyle     =   0  '�Ȃ�
      Caption         =   "Frame1"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4425
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0CCCC&
         Caption         =   "�o�n�\�I�����"
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
         Left            =   180
         TabIndex        =   1
         Top             =   120
         Width           =   1800
      End
   End
   Begin TabDlg.SSTab mstTab 
      Height          =   2745
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   4842
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   14741230
      TabCaption(0)   =   "�J�Ï��"
      TabPicture(0)   =   "ctlVRASel.ctx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "paneTab"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin Umakichi.ctlPane paneTab 
         Height          =   1815
         Left            =   60
         TabIndex        =   3
         Top             =   360
         Width           =   5325
         _ExtentX        =   9393
         _ExtentY        =   3201
         Begin VB.PictureBox fraScroll 
            Appearance      =   0  '�ׯ�
            BackColor       =   &H80000003&
            ForeColor       =   &H80000008&
            Height          =   1035
            Left            =   0
            ScaleHeight     =   1005
            ScaleWidth      =   2265
            TabIndex        =   6
            Top             =   0
            Width           =   2295
            Begin Umakichi.ctlWrappedGrid flexTab 
               Height          =   405
               Left            =   0
               TabIndex        =   7
               Top             =   0
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   714
            End
         End
         Begin VB.VScrollBar vsbSel 
            Height          =   1395
            Left            =   4320
            TabIndex        =   5
            Top             =   30
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.HScrollBar hsbSel 
            Height          =   255
            Left            =   1170
            TabIndex        =   4
            Top             =   1440
            Visible         =   0   'False
            Width           =   3105
         End
      End
   End
End
Attribute VB_Name = "ctlVRASel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'   �o�n�\�I�� �\���R���g���[��
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

Private mstrTitle       As String
Private mKey            As clsKeyRASel
Private mVB             As clsViewerBase
Private mViewerState    As clsVSNothing         '' ���
Private mblnNoData As Boolean           '' �f�[�^�����t���O

Private WithEvents mData As clsDataRASel
Attribute mData.VB_VarHelpID = -1

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
'   �@�\: �����X�N���[���o�[�ύX�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub hsbSel_Change()
On Error GoTo ErrorHandler
    flexTab.Left = -hsbSel.value
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �����X�N���[���o�[�t�H�[�J�X�擾�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub hsbSel_GotFocus()
On Error GoTo ErrorHandler
    paneTab.SetFocus
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �����X�N���[���o�[�X�N���[���C�x���g
'
'   ���l: �Ȃ�
'
Private Sub hsbSel_Scroll()
On Error GoTo ErrorHandler
    flexTab.Left = -hsbSel.value
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �����X�N���[���o�[�ύX�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub vsbSel_Change()
On Error GoTo ErrorHandler
    flexTab.Top = -vsbSel.value
    paneTab.SetFocus
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �����X�N���[���o�[�t�H�[�J�X�擾�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub vsbSel_GotFocus()
On Error GoTo ErrorHandler
    paneTab.SetFocus
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �����X�N���[���o�[�t�H�[�J�X�擾�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub vsbSel_Scroll()
On Error GoTo ErrorHandler
    flexTab.Top = -vsbSel.value
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
    
    mstrTitle = "�o�n�\�I��"
    Set mData = New clsDataRASel
    Set mKey = New clsKeyRASel
    Set mVB = New clsViewerBase
    Set mViewerState = New clsVSNothing
    
    Call mVB.FlexGridCommonSetting(flexTab.Grid)

    ' Color Assign
    BackColor = gApp.ColBG
    mstTab.BackColor = gApp.ColBG
    fraTop.BackColor = gApp.ColDarkBG
    
    lblInfo.BackColor = gApp.ColDarkBG
    lblInfo.ForeColor = Contrast(gApp.ColDarkBG)
    
    flexTab.Grid.ScrollBars = flexScrollBarNone

    paneTab.Mode = 0
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
'   �@�\: ���[�U�R���g���[���̃��T�C�Y�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub UserControl_Resize()
On Error GoTo ErrorHandler
    Dim i As Integer
    Dim blnVSBVisible As Boolean
    Dim blnHSBVisible As Boolean
    
    ' ���[�U�[�R���g���[���̍Œᕝ�����߂�
    With UserControl
        .width = Bigger(8000, .width)
        .Height = Bigger(5000, .Height)
    End With
    
    
    fraTop.width = ScaleWidth - fraTop.Left * 2
    With mstTab
        .width = Bigger(1, ScaleWidth - .Left * 2)
        .Height = Bigger(1, ScaleHeight - .Top - .Left)
    End With ' mstTab
    
    With paneTab
        .width = Bigger(1, mstTab.width - .Left * 2)
        .Height = Bigger(1, (mstTab.Height - .Top) - (.Top - mstTab.TabHeight))
    End With ' fraTab.Item(mstTab.Index)
    
    With fraScroll
        .width = paneTab.width
        .Height = paneTab.Height
    End With
    
    With flexTab
        .width = .Grid.ColPos(.Grid.Cols - 1) + .Grid.ColWidth(.Grid.Cols - 1) + Screen.TwipsPerPixelX
        .Height = .Grid.RowPos(.Grid.Rows - 1) + .Grid.RowHeight(.Grid.Rows - 1) + Screen.TwipsPerPixelY
    End With ' flexTab(i)
    
    If flexTab.width > fraScroll.width Then
        blnHSBVisible = True
        fraScroll.Height = paneTab.Height - gApp.hsbHeight
    End If
    
    If flexTab.Height > fraScroll.Height Then
        blnVSBVisible = True
        fraScroll.width = paneTab.width - gApp.vsbWidth
    End If
    
    If flexTab.width > fraScroll.width Then
        blnHSBVisible = True
        fraScroll.Height = paneTab.Height - gApp.hsbHeight
    End If
    
    With fraScroll
        hsbSel.Move .Left, .Top + .Height, .width, gApp.hsbHeight
        vsbSel.Move .Left + .width, .Top, gApp.vsbWidth, .Height
    End With
    
    flexTab.Grid.BorderStyle = flexBorderNone
    hsbSel.Min = 0
    hsbSel.max = flexTab.width - fraScroll.width
    hsbSel.LargeChange = flexTab.width
    hsbSel.Visible = blnHSBVisible
    vsbSel.Min = 0
    vsbSel.max = flexTab.Height - fraScroll.Height
    vsbSel.LargeChange = flexTab.Height
    vsbSel.Visible = blnVSBVisible
    
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
    Dim sc As New clsStringConverter
    
    mblnNoData = Not mData.Fetch(mKey)
    
    lblInfo.Caption = mData.FraTopStr
    
    '����p������ǉ�
    mstrTitle = mstrTitle & " " & lblInfo.Caption
End Sub


'
'   �@�\: �f�[�^�擾�����ʒm�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub mData_FetchComplete(gd As clsGridData)
On Error GoTo ErrorHandler
    Call flexTab.InsertGrid(gd)
    
    If gd.Cols = 3 Then
        
        With flexTab.Grid
            .FixedCols = 0
            .WordWrap = True
            .Visible = True
        End With
        Call flexTab.AutoSize(0, flexTab.Grid.Cols - 1)
        
        ' �Z�������Œ�l��
        Dim i As Integer
        For i = 0 To flexTab.Grid.Cols - 1
            flexTab.Grid.ColWidth(i) = 3960
        Next

    Else
        flexTab.Grid.Visible = False
    End If
    paneTab.Mode = 2
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
    gApp.Log "Free VRASel"
    
    Call DestroyFlexGrid(flexTab)
    
    Set mKey = Nothing
    Set mData = Nothing
End Sub

