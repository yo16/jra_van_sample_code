VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmBrowser 
   Caption         =   "�n�g"
   ClientHeight    =   8130
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11775
   Icon            =   "frmBrowser.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8130
   ScaleWidth      =   11775
   StartUpPosition =   3  'Windows �̊���l
   Begin MSComDlg.CommonDialog dlgHelpFile 
      Left            =   5940
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ilsTbrCmd 
      Left            =   1710
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.HScrollBar hsbPane 
      Height          =   285
      Left            =   540
      TabIndex        =   4
      Top             =   4350
      Width           =   4155
   End
   Begin VB.VScrollBar vsbPane 
      Height          =   2175
      Left            =   4890
      TabIndex        =   3
      Top             =   1920
      Width           =   255
   End
   Begin VB.Frame fraScrollPane 
      Appearance      =   0  '�ׯ�
      BackColor       =   &H00C0CCCC&
      BorderStyle     =   0  '�Ȃ�
      ForeColor       =   &H80000008&
      Height          =   2355
      Left            =   600
      TabIndex        =   2
      Top             =   1860
      Width           =   4065
   End
   Begin VB.Timer tmrToolbarBug 
      Interval        =   1
      Left            =   810
      Top             =   1140
   End
   Begin MSComctlLib.ImageList ilsSmallIcons 
      Left            =   90
      Top             =   1050
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar staStatusBar 
      Align           =   2  '������
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   7845
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   20258
            MinWidth        =   176
            Text            =   "�@"
            TextSave        =   "�@"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.Width           =   706
            MinWidth        =   706
            Text            =   "�b�^�|�_�|"
            TextSave        =   "�b�^�|�_�|"
            Object.ToolTipText     =   "�A�j��"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComCtl3.CoolBar cbrTop 
      Align           =   1  '�㑵��
      Height          =   915
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   1614
      _CBWidth        =   11775
      _CBHeight       =   915
      _Version        =   "6.7.9782"
      BandBackColor1  =   16711935
      Child1          =   "tbsBrowser"
      MinWidth1       =   1005
      MinHeight1      =   330
      Width1          =   2970
      NewRow1         =   0   'False
      MinWidth2       =   1005
      MinHeight2      =   525
      Width2          =   1410
      NewRow2         =   0   'False
      Visible2        =   0   'False
      Child3          =   "tbdTitleBand"
      MinHeight3      =   300
      Width3          =   495
      NewRow3         =   -1  'True
      Begin Umakichi.ctlToolBars tbsBrowser 
         Height          =   330
         Left            =   165
         TabIndex        =   6
         Top             =   120
         Width           =   2775
         _ExtentX        =   0
         _ExtentY        =   582
      End
      Begin Umakichi.ctlTitleBand tbdTitleBand 
         Height          =   300
         Left            =   165
         TabIndex        =   5
         Top             =   585
         Width           =   11520
         _ExtentX        =   8916
         _ExtentY        =   529
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "�t�@�C��(&F)"
      Begin VB.Menu mnuFileSub 
         Caption         =   "�V�K�E�C���h�E(&N)"
         Index           =   0
         Begin VB.Menu mnuFileNewSub 
            Caption         =   "�z�[�����j���[(&H)"
            Index           =   0
         End
         Begin VB.Menu mnuFileNewSub 
            Caption         =   "�o�n�\(&D)"
            Index           =   1
         End
         Begin VB.Menu mnuFileNewSub 
            Caption         =   "���ʓo�^�n(&T)"
            Index           =   2
         End
         Begin VB.Menu mnuFileNewSub 
            Caption         =   "��H����(&C)"
            Index           =   3
         End
         Begin VB.Menu mnuFileNewSub 
            Caption         =   "���R�[�h(&R)"
            Index           =   4
         End
      End
      Begin VB.Menu mnuFileSub 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuFileSub 
         Caption         =   "����(&C)"
         Index           =   2
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuFileSub 
         Caption         =   "�n�g�̏I��(&X)"
         Index           =   3
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "�ݒ�(&V)"
      Begin VB.Menu mnuConfig 
         Caption         =   "�n�g�ݒ�_�C�A���O(&C)"
      End
      Begin VB.Menu mnuViewSub 
         Caption         =   "�W���̃{�^��(&S)"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuViewSub 
         Caption         =   "���j���[�p���b�g(&P)"
         Checked         =   -1  'True
         Index           =   1
      End
   End
   Begin VB.Menu mnuDB 
      Caption         =   "�f�[�^�x�[�X(&D)"
      Begin VB.Menu mnuDBSub 
         Caption         =   "�X�V(&G)"
         Index           =   0
      End
      Begin VB.Menu mnuDBSub 
         Caption         =   "�œK��(&O)"
         Index           =   1
      End
      Begin VB.Menu mnuDBSub 
         Caption         =   "�f�[�^�Z�b�g�A�b�v(&S)"
         Index           =   2
      End
   End
   Begin VB.Menu mnuBrowser 
      Caption         =   "BrowserContext���j���["
      Visible         =   0   'False
      Begin VB.Menu mnuBrowserSub 
         Caption         =   "�߂�(&B)"
         Index           =   0
      End
      Begin VB.Menu mnuBrowserSub 
         Caption         =   "�i��(&N)"
         Index           =   1
      End
      Begin VB.Menu mnuBrowserSub 
         Caption         =   "�z�[��(&H)"
         Index           =   2
      End
   End
   Begin VB.Menu mnuToolBar 
      Caption         =   "�u���E�U�c�[���o�[Context���j���|"
      Visible         =   0   'False
      Begin VB.Menu mnuToolBarSubText 
         Caption         =   "�e�L�X�g�̕\��"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuLink 
      Caption         =   "LinkContext���j���["
      Visible         =   0   'False
      Begin VB.Menu mnuLinkSub 
         Caption         =   "�V�����E�C���h�E�ŊJ��(&W)"
         Index           =   0
      End
      Begin VB.Menu mnuLinkSub 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuLinkSub 
         Caption         =   "�߂�(&B)"
         Index           =   2
      End
      Begin VB.Menu mnuLinkSub 
         Caption         =   "�i��(&N)"
         Index           =   3
      End
   End
   Begin VB.Menu mnuHelpMenu 
      Caption         =   "�w���v(&H)"
      Begin VB.Menu mnuHelpFile 
         Caption         =   "�n�g�w���v"
         HelpContextID   =   1
      End
      Begin VB.Menu mnuBorder 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpUma 
         Caption         =   "�n�g�ɂ���(&U)"
      End
   End
End
Attribute VB_Name = "frmBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'   �u���E�U�[�t�H�[��
'
'   Viewer���悹��R���e�i
'   WebBrowser�Ɏ����C���^�[�t�F�C�X�����B
'   Viewer���͂ݏo���ꍇ�X�N���[���o�[�Ő���B
'
Option Explicit

Private WithEvents mextViewer As VBControlExtender ' Viewer�R���g���[���Q�Ɓ@�C�x���g�擾�p
Attribute mextViewer.VB_VarHelpID = -1

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �����ϐ�
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private mobjViewer As Object                       ' Viewer�R���g���[���Q�Ɓ@���\�b�h�R�[���p
Attribute mobjViewer.VB_VarHelpID = -1

Private mblnDoubleGameFlag As Boolean
Private mHistoryMgr As clsHistoryMgr               ' �����Ǘ��I�u�W�F�N�g

Private mstrViewerContextMenuViewerName As String  ' Viewer�R���e�L�X�g���j���[�̃����N��Viewer��
Private mstrViewerContextMenuKey As String         ' Viewer�R���e�L�X�g���j���[�̃����N��L�[

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
' �v���p�e�B
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   �@�\: ���j���[�p���b�g�̕\���̃`�F�b�N��ݒ肷��
'
'   ���l: �Ȃ�
'
Public Property Let ShowMenuPalette(RHS As Boolean)
    mnuViewSub(1).Checked = RHS
End Property

'
'   �@�\: ���j���[�p���b�g�̕\���̃`�F�b�N��Ԃ�Ԃ�
'
'   ���l: �Ȃ�
'
Public Property Get ShowMenuPalette() As Boolean
    ShowMenuPalette = mnuViewSub(1).Checked
End Property


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �O���֐�
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   �@�\: �����\����ʂ̐ݒ�
'
'   ���l: ���� strViewerName - Viewer��, strKey - Viewer���f�[�^����肷��ׂ̃L�[
'
Public Sub FirstPage(strViewerName As String, strKey As String)
    Dim newHistory As clsHistoryItem
    
    Set newHistory = New clsHistoryItem
    
    Set mextViewer = Controls.Add("Umakichi.ctlV" & strViewerName, VName)
    Set mobjViewer = mextViewer
    mobjViewer.key = strKey
    
    With newHistory
        .key = strKey
        .ViewerName = strViewerName
        .Title = mobjViewer.Title
    End With
    mHistoryMgr.Add newHistory
    
    Call FitViewer
    
    mextViewer.Visible = True
        
    Call TitleChange(mobjViewer.Title, strViewerName)
    Call SetHistoryToToolbar
    Call ChangeToolBar(strViewerName, strKey)

End Sub


'
'   �@�\: �|�b�v�A�b�v���j���[�\������
'
'   ���l: �Ȃ�
'
Public Sub ShowPopupMenu(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuBrowser, vbPopupMenuRightButton
    End If
End Sub


'
'   �@�\: Viewer�̃����[�h
'
'   ���l: �Ȃ�
'
Public Sub Reload()
    With mHistoryMgr.Current
        Call ChangeViewer(.ViewerName, .key)
    End With
End Sub


'
'   �@�\: Viewer�Ƀz�[����\��
'
'   ���l: �Ȃ�
'
Public Sub GotoHome()
    Call GoToNextViewer("Home", "Empty")
End Sub


'
'   �@�\: Viewer���P�O�ɖ߂�
'
'   ���l: �Ȃ�
'
Public Sub BackOne()
    Call historyBack(1)
End Sub


'
'   �@�\: Viewer�̗������J������
'
'   ���l: �Ȃ�
'
Public Sub FreeViewer()
    Call mobjViewer.Free
End Sub


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �����֐�
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   �@�\: �N�[���o�[�̍����ύX�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub cbrTop_HeightChanged(ByVal newHeight As Single)
On Error GoTo Errorhandler

    Call FitViewer(newHeight)
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �N�[���o�[�}�E�X�_�E���C�x���g
'
'   ���l: �Ȃ�
'
Private Sub cbrTop_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Errorhandler
    If Button = vbRightButton Then
    
    End If
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �t�H�[���������C�x���g
'
'   ���l: �Ȃ�
'
Private Sub Form_Initialize()
On Error GoTo Errorhandler

    Set mHistoryMgr = New clsHistoryMgr
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

    mnuViewSub(1).Checked = gApp.R_MenuVisible
    
    With ilsSmallIcons
        .ImageHeight = 16
        .ImageWidth = 16
        .ListImages.Add 1, , LoadResPicture(101, vbResIcon)
        .ListImages.Add 2, , LoadResPicture(102, vbResIcon)
        .ListImages.Add 3, , LoadResPicture(103, vbResIcon)
        .ListImages.Add 4, , LoadResPicture(104, vbResIcon)
        .ListImages.Add 5, , LoadResPicture(105, vbResIcon)
        .ListImages.Add 6, , LoadResPicture(106, vbResIcon)
        .ListImages.Add 7, , LoadResPicture(107, vbResIcon)
        Me.Icon = LoadResPicture(100, vbResIcon)
    End With
    With tbsBrowser
        .num = 3
        With .ToolBar(0)
            .ImageList = ilsSmallIcons
            .Buttons.Add 1, "BACK", "�߂�", tbrDropdown
            .Buttons.item(1).Image = 1
            .Buttons.Add 2, "NEXT", "�i��", tbrDropdown
            .Buttons.item(2).Image = 2
            .Buttons.Add 3, "HOME", "�z�[��"
            .Buttons.item(3).Image = 3
            .Buttons.Add 4, "UPDT", "�X�V"
            .Buttons.item(4).Image = 4
            .Buttons.Add 5, "CONF", "�ݒ�"
            .Buttons.item(5).Image = 5
            .width = .Buttons.item(1).width + .Buttons.item(2).width + .Buttons.item(3).width + _
                    .Buttons.item(4).width + .Buttons.item(5).width
        End With
        With .ToolBar(2)
            .ImageList = ilsSmallIcons
            .Buttons.Add 1, "RTOPEN", "����擾"
            .Buttons.item(1).Image = 7
        End With
        Call .fit
    End With
    cbrTop.Bands(1).MinWidth = tbsBrowser.ToolBar(0).width
    cbrTop.Bands(1).width = tbsBrowser.ToolBar(0).width
    
    vsbPane.width = gApp.vsbWidth
    hsbPane.Height = gApp.hsbHeight

    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �t�H�[���A�����[�h�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Errorhandler
    Call mobjViewer.Free
    If Err.Number <> 0 Then
        gApp.ErrLog
        gApp.Log "frmBrowser.Form_Unload() " & TypeName(mobjViewer) & "��Free()���������ĂȂ���������܂���>�J����"
    End If
    gApp.BrowserUnregist Me
    Exit Sub
Errorhandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   �@�\: ���j���[�̐ݒ�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub mnuConfig_Click()
    gApp.Configulation
End Sub


'
'   �@�\: ���j���[�̃w���v�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub mnuHelpFile_Click()
        Call ShowHtmlHelp
End Sub


'
'   �@�\: ���j���[�̃w���v�|�n�g�ɂ��ăC�x���g
'
'   ���l: �Ȃ�
'
Private Sub mnuHelpUma_Click()
On Error GoTo Errorhandler

    Dim aboutWindow As New frmAbout
    aboutWindow.Show vbModal, Me
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �c�[���o�[�I���C�x���g
'
'   ���l: �Ȃ�
'
Private Sub mnuToolBarSubText_Click()
On Error GoTo Errorhandler

    Dim i As Long
    mnuToolBarSubText.Checked = Not mnuToolBarSubText.Checked
    If mnuToolBarSubText.Checked Then
        For i = 1 To tbsBrowser.ToolBar(0).Buttons.count
            tbsBrowser.ToolBar(0).Buttons(i).Caption = tbsBrowser.ToolBar(0).Buttons(i).Description
        Next i
    Else
        For i = 1 To tbsBrowser.ToolBar(0).Buttons.count
            tbsBrowser.ToolBar(0).Buttons(i).Caption = ""
        Next i
    End If
    cbrTop.Bands(1).MinHeight = tbsBrowser.ToolBar(0).ButtonHeight
    cbrTop.Bands(1).MinWidth = tbsBrowser.ToolBar(0).width
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �t�H�[�����T�C�Y�C�x���g
'
'   ���l: �I�u�W�F�N�g���ŏ��ɕ\�����ꂽ�Ƃ��ɔ�������ق��A
'         �ő剻�A�ŏ����A���̃T�C�Y�ɖ߂��ȂǁA
'         �I�u�W�F�N�g�̃E�B���h�E��Ԃ��ω������Ƃ��ɂ������B
'
Private Sub Form_Resize()
On Error GoTo Errorhandler
    Call FitViewer
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   �@�\: Viewer���t�H�[���S�̂Ƀt�B�b�g������
'
'   ���l: Viewer�̑傫������AScrollPane�����傫���ꍇ�A�X�N���[���o�[��\������
'
Private Sub FitViewer(Optional newHeight As Single)
    Dim blnHSBVisible As Boolean
    Dim blnVSBVisible As Boolean
    
    blnHSBVisible = False
    blnVSBVisible = False
    
    cbrTop.Height = IIf(newHeight = 0, cbrTop.Height, newHeight)
    
    ' �X�N���[���y�C����\���̈�ő���ɍ��킹��
    With fraScrollPane
        .Left = 0                                                               ' ���[
        .Top = cbrTop.Height                                                    ' �N�[���o�[�̉�
        .width = Bigger(ScaleWidth, 1)                                          ' �g�܂ŕ������ς�
        .Height = Bigger(ScaleHeight - cbrTop.Height - staStatusBar.Height, 1)  ' �g�܂ō����҂�����
    End With
    
    Set mextViewer.Container = fraScrollPane
    
    ' Viewer ���X�N���[���y�C���̑傫���ɍ��킹��
    With mextViewer
        .Left = 0                       ' �X�N���[���y�C���̍��[
        .Top = 0                        ' �X�N���[���y�C���̉E�[
        .width = fraScrollPane.width    ' �ő啝
        .Height = fraScrollPane.Height  ' �ő卂
    End With
    
    ' Viewer�̕����X�N���[���y�C���̕����傫���ꍇ
    If mextViewer.width > fraScrollPane.width Then
        ' ����SB ���Ɍ���
        blnHSBVisible = True
        ' ����SB�̕��A�X�N���[���y�C������������
        fraScrollPane.Height = Bigger(fraScrollPane.Height - hsbPane.Height, 1)
        ' �X�N���[���y�C���̍�����Viewer�����킹��
        mextViewer.Height = fraScrollPane.Height
    End If
    
    ' Viewer�̍������X�N���[���y�C���̍������傫���ꍇ
    If mextViewer.Height > fraScrollPane.Height Then
        ' ���� SB ���Ɍ���
        blnVSBVisible = True
        ' �����X�N���[���o�[�̕��A�X�N���[���y�C������������
        fraScrollPane.width = Bigger(fraScrollPane.width - vsbPane.width, 1)
        ' �X�N���[���y�C���̕���Viewer�����킹��
        mextViewer.width = fraScrollPane.width
    End If
    
    ' �����X�N���[���o�[�̏o���ŕ��������Ȃ�������
    ' Viewer�̕����X�N���[���y�C���̕����傫���ꍇ
    If blnHSBVisible = False And mextViewer.width > fraScrollPane.width Then
        ' ����SB ���Ɍ���
        blnHSBVisible = True
        ' �����X�N���[���o�[�̕��A�X�N���[���y�C������������
        fraScrollPane.Height = Bigger(fraScrollPane.Height - hsbPane.Height, 1)
        ' �X�N���[���y�C���̍�����Viewer�����킹��
        mextViewer.Height = fraScrollPane.Height
    End If
    
    ' ����SB ��z�u
    If blnHSBVisible Then
        With hsbPane
            .Left = fraScrollPane.Left
            .Top = fraScrollPane.Top + fraScrollPane.Height
            .width = fraScrollPane.width
        End With
    End If
    hsbPane.Visible = blnHSBVisible
    
    ' ����SB ��z�u
    If blnVSBVisible Then
        With vsbPane
            .Left = fraScrollPane.Left + fraScrollPane.width
            .Top = fraScrollPane.Top
            .Height = fraScrollPane.Height
        End With
    End If
    vsbPane.Visible = blnVSBVisible
    
    hsbPane.max = mextViewer.width - fraScrollPane.width
    hsbPane.LargeChange = mextViewer.width
    hsbPane.SmallChange = mextViewer.width / 10
    vsbPane.max = mextViewer.Height - fraScrollPane.Height
    vsbPane.LargeChange = mextViewer.Height
    vsbPane.SmallChange = mextViewer.Height / 10

End Sub


'
'   �@�\: �����X�N���[���o�[�ύX�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub hsbPane_Change()
On Error GoTo Errorhandler
    mextViewer.Left = -hsbPane.value
    mextViewer.SetFocus                 ' Viewer�Ƀt�H�[�J�X���Z�b�g
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �����X�N���[���o�[�h���b�N���C�x���g
'
'   ���l: �Ȃ�
'
Private Sub hsbPane_Scroll()
On Error GoTo Errorhandler
    mextViewer.Left = -hsbPane.value
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   �@�\: ���[�X�`�F���W���[�I���C�x���g
'
'   ���l: �Ȃ�
'
Private Sub tbdTitleBand_Change(key As clsKeyRA)
On Error GoTo EH
    With mHistoryMgr.Current
        Call GoToNextViewer(.ViewerName, key.str)
    End With
    Exit Sub
EH:
    gApp.ErrLog
End Sub


'
'   �@�\: �c�[���o�[�\���^�C�}�[�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub tmrToolbarBug_Timer()
On Error GoTo Errorhandler:
    If Me.Visible = True Then
        tmrToolbarBug.Enabled = False
        cbrTop.Bands(1).Child.fit
    Else
        '
    End If
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �����X�N���[���o�[�ύX�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub vsbPane_Change()
On Error GoTo Errorhandler
    mextViewer.Top = -vsbPane.value
    mextViewer.SetFocus             ' Viewer�Ƀt�H�[�J�X���Z�b�g
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �����X�N���[���o�[�h���b�N���C�x���g
'
'   ���l: �Ȃ�
'
Private Sub vsbPane_Scroll()
On Error GoTo Errorhandler
    mextViewer.Top = -vsbPane.value
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �t�@�C���T�u���j���[�I���C�x���g
'
'   ���l: �Ȃ�
'
Private Sub mnuFileSub_Click(Index As Integer)
On Error GoTo Errorhandler
Select Case Index
    Case 0  ' �V�K�E�C���h�E
    Case 1 ' �{�[�_�[
    Case 2 ' ����
        Unload Me
    Case 3 ' �n�g�S�̂̏I��
        If vbYes = MsgBox("�n�g���I�����܂����H", vbYesNo + vbQuestion, "�n�g�F�I���̊m�F") Then
            gApp.ExitApp
        End If
    End Select
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �t�@�C���V�K�T�u���j���[�I���C�x���g
'
'   ���l: �Ȃ�
'
Private Sub mnuFileNewSub_Click(Index As Integer)
On Error GoTo Errorhandler
    Select Case Index
    Case 0  ' �z�[��
        Call gApp.NewWindow("Home", "Empty")
    Case 1  ' �o�n�\
        Call gApp.NewWindow("RAKaiSel", "Empty")
    Case 2  ' ���ʓo�^�n
        Call gApp.NewWindow("TKKaiSel", "Empty")
    Case 3  ' �̘H����
        Call gApp.NewWindow("HCSel", "Empty")
    Case 4  ' ���R�[�h
        Call gApp.NewWindow("RCSel", "Empty")
    End Select
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �c�a�T�u���j���[�I���C�x���g
'
'   ���l: �Ȃ�
'
Private Sub mnuDBSub_Click(Index As Integer)
On Error GoTo Errorhandler
    Dim MsgResult       As VbMsgBoxResult
    
    Select Case Index
    Case 0  ' �X�V
        If MsgBox("�f�[�^�̍X�V�����܂���", vbYesNo + vbQuestion, "�n�g�F�f�[�^�X�V�����J�n�̊m�F") = vbYes Then
            Call gApp.DBUpdate
            With mHistoryMgr.Current
                Call ChangeViewer(.ViewerName, .key)
            End With
            
        End If
    ' �œK��
    Case 1
        Call gApp.DBCompact
    ' �Z�b�g�A�b�v
    Case 2
        Call setup

    End Select
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �f�[�^�Z�b�g�A�b�v
'
'   ���l: �Ȃ�
'
Private Sub setup()
On Error GoTo Errorhandler
    Dim LastTimeBackup  As String
    Dim ConfigFirst     As frmConfigFirst
    Dim DBUpdateForm    As frmDBUpdate
    Dim MsgResult       As VbMsgBoxResult
    
    ' �ăZ�b�g�A�b�v���ǂ���
    If gApp.R_JVDLastTime <> String$(14, "0") Then
        MsgResult = MsgBox(gApp.R_DBPath & _
            "�f�[�^�x�[�X�͂��łɃZ�b�g�A�b�v����Ă��܂��B" & vbCrLf & _
            "�Z�b�g�A�b�v����蒼���܂����H", vbExclamation + vbYesNo + vbDefaultButton2, "�n�g�F�ăZ�b�g�A�b�v�̊m�F")
        If MsgResult = vbYes Then
            ' �ŏI�擾������ޔ�
            LastTimeBackup = gApp.R_JVDLastTime
            ' �ŏI�擾���������Z�b�g
            gApp.R_JVDLastTime = String$(14, "0")
            ' �f�[�^�Z�b�g�A�b�v�̐ݒ��ʂ��o��
            Set ConfigFirst = New frmConfigFirst
            ConfigFirst.Show vbModal
            If ConfigFirst.ButtonType <> "OK" Then
                ' �ŏI�擾�����𕜋A
                gApp.R_JVDLastTime = LastTimeBackup
                MsgBox "�ăZ�b�g�A�b�v�͍s���܂���ł����B", vbInformation, "�n�g�F�ăZ�b�g�A�b�v�L�����Z��"
                Exit Sub
            End If
            ' �擾���s��
            Set DBUpdateForm = New frmDBUpdate
            DBUpdateForm.Show vbModal
            If DBUpdateForm.AfterJVOpen = False Then
                gApp.R_JVDLastTime = LastTimeBackup
            End If
            gApp.AllReload
        End If
    Else
        Call gApp.DBUpdate
    End If
    
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �����N�T�u���j���[�I���C�x���g
'
'   ���l: �E�N���b�N���j���[�ł��B
'
Private Sub mnuLinkSub_Click(Index As Integer)
On Error GoTo Errorhandler
Select Case Index
    Case 0
        Call gApp.NewWindow(mstrViewerContextMenuViewerName _
                            , mstrViewerContextMenuKey)
    Case 1 ' �{�[�_�[
    
    Case 2 ' �߂�
        Call historyBack(1)
    Case 3 ' �i��
        Call historyNext(1)
    End Select
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �\���T�u���j���[�I���C�x���g
'
'   ���l: �Ȃ�
'
Private Sub mnuViewSub_Click(Index As Integer)
On Error GoTo Errorhandler
    Select Case Index
    Case 0 ' �W���̃{�^��
        With mnuViewSub(Index)
            .Checked = Not .Checked
            cbrTop.Bands(1).Visible = .Checked
            tbsBrowser.fit
        End With
    Case 1 ' ���j���[�p���b�g
        Call gApp.ShowMenuPalette(Not mnuViewSub(1).Checked)
    End Select
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �u���E�U���j���[�I���C�x���g
'
'   ���l: �Ȃ�
'
Private Sub mnuBrowserSub_Click(Index As Integer)
On Error GoTo Errorhandler
    
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub

'
'   �@�\: Viewer�̃C�x���g����
'
'   ���l: �^�C�g���ύX�A��ʐ؂�ւ���
'
Private Sub mextViewer_ObjectEvent(Info As EventInfo)
On Error GoTo Errorhandler
    Dim newViewer As Control
    Dim obj As Object
    
    Dim ViewerName As String
    Dim key As String

    Select Case Info.Name
        Case "WindowTitle"
            ' �E�C���h�E�^�C�g���ύX
            ViewerName = Mid(TypeName(mextViewer), 5)
            Call TitleChange(Info.EventParameters.item(0), ViewerName)
        
        Case "ChangeTo"
            ' ��ʕύX�C�x���g
            
            Call GoToNextViewer(Info.EventParameters.item(0), Info.EventParameters.item(1))
            
        Case "NewWindow"
            ' �V�K�E�C���h�E
            
            ' �C�x���g���������󂯎��
            With Info.EventParameters
                ViewerName = .item(0)
                key = .item(1)
            End With
            
            Call gApp.NewWindow(ViewerName, key)
        
        Case "LinkContextMenu"
            ' �E�N���b�N���j���[
            
            ' �����N��������W���[���ϐ��ɕۑ��A���j���[�C�x���g�ŏE��
            With Info.EventParameters
                mstrViewerContextMenuViewerName = .item(0)
                mstrViewerContextMenuKey = .item(1)
            End With
                        
            Me.PopupMenu mnuLink, vbPopupMenuRightButton
            
        Case "Reload"
            Call Reload
            
        
        Case "StatusBarTextChange"
            ' �i��
            
            staStatusBar.Panels(1).Text = Info.EventParameters(0)
            
        Case "Progression"
            
        Case Else
            gApp.Log "Unknown Viewer Event (" & Info.Name & ")"
    
    End Select ' Info.Name
    
    Exit Sub

Errorhandler:
    gApp.ErrLog

End Sub


'
'   �@�\: Viewer��؂�ւ��q�X�g����i�߃c�[���o�[�̍X�V
'
'   ���l: �Ȃ�
'
Private Sub GoToNextViewer(ViewerName As String, key As String)
On Error GoTo Errorhandler
    Dim newHistory As clsHistoryItem
    
    ' �ŏI��ԃf�[�^���擾
    On Error Resume Next
        Set mHistoryMgr.Current.ViewerState = mobjViewer.ViewerState
        If Err.Number <> 0 Then
            gApp.ErrLog
            gApp.Log TypeName(mobjViewer) & "�ɂ�ViewerStatus ����������Ă��Ȃ��Ǝv���܂��B"
        End If
    On Error GoTo Errorhandler
    
    
    ' �؂�ւ�
    If ChangeViewer(ViewerName, key) Then
        ' �����Ȃ�㏈��
        ' ����ۑ�
        Set newHistory = New clsHistoryItem
        With newHistory
            .ViewerName = ViewerName
            .key = key
            .Title = mobjViewer.Title
            .DateTime = Timer
        End With
        mHistoryMgr.Add newHistory
        
        ' �������c�[���o�[�ɔ��f
        Call SetHistoryToToolbar

    Else
        ' �؂�ւ����s�̏ꍇ
        MsgBox "�J���܂���B", vbExclamation, "�n�g�F�G���["
    End If
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �E�C���h�E�^�C�g����ݒ�
'
'   ���l: �\�ʕ����ƈ�������A�����Đݒ肷��
'
Private Sub TitleChange(strTitle As String, ViewerName As String)
On Error GoTo Errorhandler

    Me.Caption = strTitle & " : " & cAppName
    
    tbdTitleBand.Caption = strTitle

    cbrTop.Bands(3).MinWidth = tbdTitleBand.width
    Exit Sub

Errorhandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �\Viewer����Ԃ�
'
'   ���l: �_�u���o�b�t�@�����O�̂悤�ȏ����̈�
'
Private Function VName() As String
    VName = "Viewer" & IIf(mblnDoubleGameFlag, "1", "0")
End Function


'
'   �@�\: ��Viewer����Ԃ�
'
'
'   ���l: �Ԃ�l ��Viewer��, �_�u���o�b�t�@�����O�̂悤�ȏ����̈�
'
Private Function VNameBack() As String
    VNameBack = "Viewer" & IIf(mblnDoubleGameFlag, "0", "1")
End Function


'
'   �@�\: Viewer�̐؂�ւ�
'
'   ���l: ������ strViewerName - Viewer��
'                strKey        - Viewer���\���f�[�^����肷��ׂ̃L�[
'                viewerState   - Viewer���
'         ��Viewer�Ƃ��Đ��������㤕\Viewer��j��������ė���\�Ƃ��Ă���
'         ����́A�������s�������̂��ߋ�Viewer��ێ����Ă�����
'         �R���g���[�������d���o���Ȃ����߁B
'         �����܂ŁAViewer�𒣂�ւ��邾���ŁA�q�X�g���̐i�s�����̑��̍�Ƃ͍s��Ȃ�
'         ���̉�ʂփ����N���ԏꍇ�́AGoToNextViewer�����
'
Private Function ChangeViewer(strViewerName As String, strKey As String, _
                              Optional ViewerState As clsIViewerState) As Boolean
On Error GoTo Errorhandler
    Dim mp           As New clsPointer '' �}�E�X�|�C���^
    Dim objNewViewer As Object         '' �V�KViewer
    Dim faze         As Long
    
    
    faze = 1
    
    Call mp.SetBusyPointer(Me)
    staStatusBar.Panels(1).Text = "�ǂݍ��ݒ�..."
    
    faze = 2
    
    ' ��Viewer�쐬
    gApp.Log "CreateControl: " & VNameBack
    Set objNewViewer = Controls.Add("Umakichi.ctlV" & strViewerName, VNameBack)
    
    faze = 3
    
    ' �ŏI��ԃf�[�^���Z�b�g
    On Error Resume Next
        If Not ViewerState Is Nothing Then
            Set objNewViewer.ViewerState = ViewerState
            If Err.Number <> 0 Then
                gApp.ErrLog
                gApp.Log TypeName(objNewViewer) & "�ɂ�ViewerStatus ����������Ă��Ȃ��Ǝv���܂��B"
            End If
        End If
    On Error GoTo Errorhandler
    
    ' ��Viewer�ɃL�[�̐ݒ� = �擾�J�n
    gApp.Log "ChangeViewer SetKey"
    objNewViewer.key = strKey
    
    ' �\Viewer�̃C�x���g���E��Ȃ��悤�ɂ���
    Set mextViewer = Nothing
    
    ' �f�[�^�������ꍇ�A�z�[���ֈړ�������
    If objNewViewer.NoData Then
        Call Controls.Remove(VNameBack)
        Set objNewViewer = Controls.Add("Umakichi.ctlVHome", VNameBack)
        strViewerName = "Home"
        strKey = Empty
    End If
    
    ' �\Viewer�̏I�������v���V�[�W�����Ă�
    gApp.Log "ChangeViewer Free"
    Call mobjViewer.Free
    
    ' �\Viewer�폜
    gApp.Log "Unload: " & VName
    Call Controls.Remove(VName)
    
    ' �E�C���h�E�^�C�g����Viewer����擾
    gApp.Log "ChangeViewer TitleCange"
    Call TitleChange(objNewViewer.Title, strViewerName)
    
    ' �C�x���g���󂯎��׃R���g���[���G�N�X�e���_�Ƃ��ĕێ�
    gApp.Log "ChangeViewer Set mextViewer"
    Set mextViewer = objNewViewer
    
    ' �A�N�Z�X�̈׃I�u�W�F�N�g�Ƃ��ĕێ�
    gApp.Log "ChangeViewer Set mobjViewer"
    Set mobjViewer = objNewViewer
    
    ' ��Viewer�̃T�C�Y���E�C���h�E�ɖ�������悤�ύX
    gApp.Log "ChangeViewer Fit Viewer"
    FitViewer
    
    ' ��Viewer�̕\��
    mextViewer.Visible = True
    
    ' �c�[���o�[�ύX
    gApp.Log "ChangeViewer ToolBar Setting"
    Call ChangeToolBar(strViewerName, strKey)
    
    ' ���\�؂�ւ�
    gApp.Log "DoubleGameFlag Before: " & mblnDoubleGameFlag
    mblnDoubleGameFlag = Not mblnDoubleGameFlag
    gApp.Log "DoubleGameFlag After : " & mblnDoubleGameFlag
    
    staStatusBar.Panels(1).Text = ""
    
    ChangeViewer = True
    Exit Function

Errorhandler:
    gApp.ErrLog
    If faze < 3 Then
    Call Controls.Remove(VNameBack)
    ChangeViewer = False
    Else
    Resume Next
    End If
End Function


'
'   �@�\: Viewer�ɉ����āAViewer�p�c�[���o�[��\��
'
'   ���l: ������ strViewerName - Viewer��, strKey - �Ăяo���L�[(RaceChanger�������p)
'
Private Sub ChangeToolBar(strViewerName As String, strKey As String)
On Error GoTo Errorhandler
    Dim i As Long
    Dim max As Long
    Dim raceChnagerKey As New clsKeyRASel
    
    Select Case strViewerName
    Case "RA", "OD"
        Set mobjViewer.ToolBar = tbsBrowser  ' �Ǝ��{�^���ݒ�
        tbsBrowser.ToolBar(1).Visible = True ' �Ǝ��{�^���\��
        tbsBrowser.ToolBar(2).Visible = True ' ����擾�{�^���\��
        tbsBrowser.fit
        raceChnagerKey.str = strKey
        
        Call tbdTitleBand.ShowRaceChanger(raceChnagerKey)
        cbrTop.Bands(1).MinWidth = tbsBrowser.MinWidth
        cbrTop.Bands(1).width = tbsBrowser.MinWidth
    
    Case "HK"
        Set mobjViewer.ToolBar = tbsBrowser
        With tbsBrowser
            .ToolBar(1).Visible = False  ' �Ǝ��{�^����\��
            .ToolBar(2).Visible = True   ' ����擾�{�^���\��
            .fit
        End With
        Call tbdTitleBand.ShowRaceChanger(Nothing)
    Case "TK"
        With tbsBrowser
            .ToolBar(1).Visible = False   ' �Ǝ��{�^����\��
            .ToolBar(2).Visible = False   ' ����擾�{�^����\��
            .fit
        End With
        raceChnagerKey.str = strKey
        Call tbdTitleBand.ShowRaceChanger(raceChnagerKey, True) ' TKmode ON
    Case Else
        With tbsBrowser
            .ToolBar(1).Visible = False   ' �Ǝ��{�^����\��
            .ToolBar(2).Visible = False   ' ����擾�{�^����\��
            .fit
        End With
        Call tbdTitleBand.ShowRaceChanger(Nothing)
    End Select
    
    Exit Sub
Errorhandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   �@�\: �߂�A�i�ރ{�^���̃h���b�v�_�E�����j���[�ɗ�����o�^����
'
'   ���l: �ő�10���܂łɐ������Ă���
'
Private Sub SetHistoryToToolbar()
    Dim i As Long
    Dim tmpHistory As clsHistoryItem
    
    tbsBrowser.ToolBar(0).Buttons(1).ButtonMenus.Clear
    tbsBrowser.ToolBar(0).Buttons(2).ButtonMenus.Clear
    
    ' �߂�{�^��
    For i = 1 To 10
        Set tmpHistory = mHistoryMgr.Current(-i)
        If tmpHistory Is Nothing Then
            Exit For
        End If
        tbsBrowser.ToolBar(0).Buttons(1).ButtonMenus.Add , , mHistoryMgr.CurrentNum - i & tmpHistory.Title
    Next i
    ' �擪�Ȃ疳���ɂ���
    tbsBrowser.ToolBar(0).Buttons(1).Enabled = Not mHistoryMgr.IsFirst
    
'    ' �i�ރ{�^��
    For i = 1 To 10
        Set tmpHistory = mHistoryMgr.Current(i)
        If tmpHistory Is Nothing Then
            Exit For
        End If
        tbsBrowser.ToolBar(0).Buttons(2).ButtonMenus.Add , , mHistoryMgr.CurrentNum + i & tmpHistory.Title
    Next i
    ' �ŏI�Ȃ疳���ɂ���
    tbsBrowser.ToolBar(0).Buttons(2).Enabled = Not mHistoryMgr.IsLast

End Sub


'
'   �@�\: �u���E�U�p�c�[���o�[�̃N���b�N�C�x���g
'
'   ���l: ������߂�A�i�ނ̏���
'
Private Sub tbsBrowser_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
On Error GoTo Errorhandler
    gApp.Log "Browser Catch TBSClick"
    Select Case Button.key
    ' �߂�{�^��
    Case "BACK"
        Call historyBack(1)
    ' �i�ރ{�^��
    Case "NEXT"
        Call historyNext(1)
    ' �z�[���{�^��
    Case "HOME"
        Call GoToNextViewer("Home", "Empty")
    ' �X�V�{�^��
    Case "UPDT"
        If MsgBox("�f�[�^�̍X�V�����܂���", vbYesNo + vbQuestion, "�n�g�F�f�[�^�X�V�����J�n�̊m�F") = vbYes Then
            Call gApp.DBUpdate
            With mHistoryMgr.Current
                Call ChangeViewer(.ViewerName, .key)
            End With
        End If
    ' �ݒ�{�^��
    Case "CONF"
        Call gApp.Configulation
        Call Reload
    End Select
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �u���E�U�p�c�[���o�[�̃h���b�v�_�E�����j���[�̃N���b�N�C�x���g
'
'   ���l: ��������x�ɉ��i�K���߂�A�i�ނ̏���
'
Private Sub tbsBrowser_ButtonMenuClick(Index As Integer, ByVal ButtonMenu As MSComctlLib.ButtonMenu)
On Error GoTo Errorhandler
Select Case ButtonMenu.Parent.key
    Case "BACK"
        gApp.Log ButtonMenu.Index
        Call historyBack(ButtonMenu.Index)
    Case "NEXT"
        Call historyNext(ButtonMenu.Index)
    End Select
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �������P�i�K�܂��͉��i�K���߂�Viewer��ύX����
'
'   ���l: �c�[���o�[�̕ύX���s��
'
Private Sub historyBack(lngStep As Long)
On Error GoTo Errorhandler
    gApp.Log "�߂�"
    
    ' �ŏI��ԃf�[�^���擾
    On Error Resume Next
        Set mHistoryMgr.Current.ViewerState = mobjViewer.ViewerState
        If Err.Number <> 0 Then
            gApp.ErrLog
            gApp.Log TypeName(mobjViewer) & "�ɂ�ViewerStatus ����������Ă��Ȃ��Ǝv���܂��B"
        End If
    On Error GoTo Errorhandler
    
    If Not mHistoryMgr.IsFirst Then
'        ' �q�X�g���|�C���^�̓�O���O���
        Call mHistoryMgr.Move(-lngStep)
        With mHistoryMgr.Current
            Call ChangeViewer(.ViewerName, .key, .ViewerState)
        End With
        Call SetHistoryToToolbar
    End If
    Exit Sub
Errorhandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   �@�\: �������P�i�K�܂��͉��i�K���i��Viewer��ύX����
'
'   ���l: �c�[���o�[�̕ύX���s��
'
Private Sub historyNext(lngStep As Long)
On Error GoTo Errorhandler

    gApp.Log "�i��"
        
    ' �ŏI��ԃf�[�^���擾
    On Error Resume Next
        Set mHistoryMgr.Current.ViewerState = mobjViewer.ViewerState
        If Err.Number <> 0 Then
            gApp.ErrLog
            gApp.Log TypeName(mobjViewer) & "�ɂ�ViewerStatus ����������Ă��Ȃ��Ǝv���܂��B"
        End If
    On Error GoTo Errorhandler
    
    Call mHistoryMgr.Move(lngStep)
    With mHistoryMgr.Current
        Call ChangeViewer(.ViewerName, .key, .ViewerState)
    End With
    
    ' �c�[���o�[���X�V
    Call SetHistoryToToolbar
    
    Exit Sub
Errorhandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   �@�\: �u���E�U�p�c�[���o�[�̃}�E�X�_�E���C�x���g
'
'   ���l: �c�[���o�[�̕ύX���s��
'
Private Sub tbrBrowser_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Errorhandler
    If Button = vbRightButton Then
        PopupMenu mnuToolBar, vbPopupMenuRightButton
    End If
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub
