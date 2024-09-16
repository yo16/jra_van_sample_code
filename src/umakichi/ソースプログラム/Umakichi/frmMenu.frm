VERSION 5.00
Begin VB.Form frmMenu 
   BackColor       =   &H000000FF&
   BorderStyle     =   4  '�Œ�°� ����޳
   Caption         =   "�n�g���j���["
   ClientHeight    =   4545
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   2010
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   2010
   StartUpPosition =   3  'Windows �̊���l
   Begin Umakichi.ctlMenu ctlMenu 
      Height          =   4350
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2010
      _ExtentX        =   3545
      _ExtentY        =   7673
   End
   Begin VB.Menu mnuFile 
      Caption         =   "̧��(&F)"
      Begin VB.Menu mnuBorder 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "�I��(&X)"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuConfig 
      Caption         =   "�ݒ�(&C)"
      Begin VB.Menu mnuUmakichiConfig 
         Caption         =   "�n�g�ݒ�_�C�A���O(&C)"
      End
      Begin VB.Menu mnuBorder2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAlwaysTop 
         Caption         =   "��Ɏ�O�ɕ\��(&T)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuCanAlone 
         Caption         =   "���j���[���c��(&A)"
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'   ���j���[�p���b�g�t�H�[��
'

Option Explicit


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �����ϐ�
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private NowResizing As Boolean


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   API�֐��錾
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
' SetWindowPos API�֐��p�萔
Private Const HWND_NOTOPMOST = -2
Private Const HWND_TOPMOST = -1
Private Const SWP_NOMOVE = 2
Private Const SWP_NOSIZE = 1

Private Declare Function SetWindowPos Lib "user32" _
            (ByVal hwindow As Long _
            , ByVal hinsertafter As Long _
            , ByVal X As Long _
            , ByVal Y As Long _
            , ByVal cx As Long _
            , ByVal cy As Long _
            , ByVal flag As Long) As Long


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �O���֐�
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8


'
'   �@�\: �t�H�[������Ɏ�O�ɕ\������
'
'   ���l: �Ȃ�
'
Public Sub AlwaysTop(Mode As Boolean)
    If Mode Then
        SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE
    Else
        SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE
    End If
End Sub


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �����֐�
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8


'
'   �@�\: �t�H�[�����[�h�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub Form_Load()
On Error GoTo ErrorHandler
    Call AlwaysTop(gApp.R_MenuAlwaysTopFlag)
    mnuAlwaysTop.Checked = gApp.R_MenuAlwaysTopFlag
    mnuCanAlone.Checked = gApp.R_MenuCanAlone
    
    Me.Icon = LoadResPicture(100, vbResIcon)
    
    DoEvents
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �t�H�[�������Ɠ����ɁA���j���[�\���I�v�V������؂�
'
'   ���l: �Ȃ�
'
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrorHandler
    If UnloadMode = vbFormControlMenu Then
        Call gApp.ShowMenuPalette(False)
    End If
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �t�H�[�����T�C�Y�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub Form_Resize()
On Error GoTo ErrorHandler
    Dim FX1 As Long
    Dim FY1 As Long
    Dim FX2 As Long
    Dim FY2 As Long
    
    Dim CX1 As Long
    Dim CY1 As Long
    Dim CX2 As Long
    Dim CY2 As Long
    
    Call GetWindowRect(Me.hwnd, FX1, FY1, FX2, FY2)
    Call ctlMenu.WindowRect(CX1, CY1, CX2, CY2)
    Me.Height = ((CY2 - CY1) + (CY1 - FY1) + 2) * Screen.TwipsPerPixelY
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: ���j���[�I���C�x���g�u��Ɏ�O�ɕ\���v
'
'   ���l: �Ȃ�
'
Private Sub mnuAlwaysTop_Click()
On Error GoTo ErrorHandler
    gApp.R_MenuAlwaysTopFlag = Not gApp.R_MenuAlwaysTopFlag
    mnuAlwaysTop.Checked = gApp.R_MenuAlwaysTopFlag
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: ���j���[�I���C�x���g�u���j���[���c���v
'
'   ���l: �Ȃ�
'
Private Sub mnuCanAlone_Click()
On Error GoTo ErrorHandler
    mnuCanAlone.Checked = Not mnuCanAlone.Checked
    
    gApp.R_MenuCanAlone = mnuCanAlone.Checked
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: ���j���[�I���C�x���g�u�A�v���P�[�V�����I���v
'
'   ���l: �Ȃ�
'
Private Sub mnuExit_Click()
On Error GoTo ErrorHandler
    Call gApp.ExitApp
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: ���j���[�I���C�x���g�u�n�g�ݒ�_�C�A���O�v
'
'   ���l: �Ȃ�
'
Private Sub mnuUmakichiConfig_Click()
On Error GoTo ErrorHandler
    Call gApp.Configulation
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �c�[���o�[�{�^�������C�x���g
'
'   ���l: �Ȃ�
'
Private Sub ctlMenu_Click(tag As String)
On Error GoTo ErrorHandler
    Select Case tag
        Case "Find0"
            Call gApp.NewWindow("Find", 0)
        Case "Find1"
            Call gApp.NewWindow("Find", 1)
        Case "Find2"
            Call gApp.NewWindow("Find", 2)
        Case "Find3"
            Call gApp.NewWindow("Find", 3)
        Case "Find4"
            Call gApp.NewWindow("Find", 4)
        Case "Find5"
            Call gApp.NewWindow("Find", 5)
        Case "CONF"
            Call gApp.Configulation
        Case "HELP"
            Call ShowHtmlHelp
        Case "UPDT"
            Call gApp.DBUpdate
        Case "SOKU"
            Call gApp.DBPrompt(ukpPALLET, gApp.R_RTDates)
        Case Else
            Call gApp.NewWindow(tag, "Empty")
    End Select
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub

