VERSION 5.00
Begin VB.UserControl ctlPane 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Timer tmrAnim 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3900
      Top             =   2940
   End
   Begin VB.Label lblAnim 
      AutoSize        =   -1  'True
      Caption         =   "[��������������������������������]"
      Height          =   180
      Left            =   930
      TabIndex        =   2
      Top             =   1260
      Width           =   3000
   End
   Begin VB.Label lblMode1 
      AutoSize        =   -1  'True
      Caption         =   "�f�[�^������܂���B"
      Height          =   180
      Left            =   1380
      TabIndex        =   1
      Top             =   1770
      Width           =   1620
   End
   Begin VB.Label lblMode0 
      AutoSize        =   -1  'True
      Caption         =   "�f�[�^���擾���ł��B"
      Height          =   180
      Left            =   1440
      TabIndex        =   0
      Top             =   750
      Width           =   1650
   End
End
Attribute VB_Name = "ctlPane"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'   �ǂݍ��ݒ��A�f�[�^������܂���A�L���A�̂R��Ԃ����R���e�i�̈�
'

Option Explicit


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �O���֐�(�C�x���g)
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Public Event Progression()

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   API�֐��錾
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const SWP_FRAMECHANGED = &H20        ''  The frame changed: send WM_NCCALCSIZE
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOOWNERZORDER = &H200      ''  Don't do owner Z ordering
Private Const SWP_NOREDRAW = &H8
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4
Private Const SWP_SHOWWINDOW = &H40

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Const GWL_STYLE = (-16)
Private Const WS_THICKFRAME = &H40000
Private Const WS_BORDER = &H800000
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_WINDOWEDGE = &H100&
Private Const WS_EX_CLIENTEDGE = &H200&
Private Const WS_EX_STATICEDGE = &H20000

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �O���萔
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Public Enum EAppearanceConstants
   eacFlat
   eac3D
End Enum

Public Enum EBorderStyleConstants
   ebscNone
   ebscFixedSingle
   ebscThin
   ebscRaised
End Enum

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �����ϐ�
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private mlngMode As ukCtlPaneMode
Private mlngPos  As Long

Private m_eAppearance As EAppearanceConstants
Private m_eBorderStyle As EBorderStyleConstants

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �v���p�e�B
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   �@�\: Appearance�v���p�e�B�̎擾
'
'   ���l: �Ȃ�
'
Public Property Get Appearance() As EAppearanceConstants
   Appearance = m_eAppearance
End Property

'
'   �@�\: Appearance�v���p�e�B�̃Z�b�g
'
'   ���l: �Ȃ�
'
Public Property Let Appearance(ByVal eStyle As EAppearanceConstants)
   m_eAppearance = eStyle
   pSetBorder
   PropertyChanged "Appearance"
End Property

'
'   �@�\: BorderStyle�v���p�e�B�̎擾
'
'   ���l: �Ȃ�
'
Public Property Get BorderStyle() As EBorderStyleConstants
   BorderStyle = m_eBorderStyle
End Property

'
'   �@�\: BorderStyle�v���p�e�B�̃Z�b�g
'
'   ���l: �Ȃ�
'
Public Property Let BorderStyle(ByVal eStyle As EBorderStyleConstants)
   m_eBorderStyle = eStyle
   pSetBorder
   PropertyChanged "BorderStyle"
End Property

'
'   �@�\: Mode�v���p�e�B�̃Z�b�g
'
'   ���l: �Ȃ�
'
Public Property Let Mode(RHS As ukCtlPaneMode)
    mlngMode = RHS
    
    lblMode0.Visible = (RHS = ukcpNowFetching)
    lblAnim.Visible = (RHS = ukcpNowFetching)
    lblMode1.Visible = (RHS = ukcpNoData)
    Call VisibleAllContained(RHS = ukcpShowControls)
    tmrAnim.Enabled = (RHS = ukcpNowFetching)
End Property

'
'   �@�\: Mode�v���p�e�B�̎擾
'
'   ���l: �Ȃ�
'
Public Property Get Mode() As ukCtlPaneMode
    Mode = mlngMode
End Property


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �����֐�
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   �@�\: �E�B���h�E�X�^�C���̐ݒ�
'
'   ���l: �Ȃ�
'
Private Sub pSetBorder()
Dim lS As Long
   
   UserControl.BorderStyle() = 0
   If m_eAppearance = eacFlat Then
      ' Flat border
      pSetWinStyle GWL_EXSTYLE, 0, WS_EX_STATICEDGE Or WS_EX_CLIENTEDGE Or WS_EX_WINDOWEDGE
      If m_eBorderStyle > ebscNone Then
         pSetWinStyle GWL_STYLE, WS_BORDER, 0
      Else
         pSetWinStyle GWL_STYLE, 0, WS_BORDER
      End If
   Else
      ' 3d border
      Select Case m_eBorderStyle
      Case ebscNone
         ' No borders
         pSetWinStyle GWL_STYLE, 0, WS_BORDER Or WS_THICKFRAME
         pSetWinStyle GWL_EXSTYLE, 0, WS_EX_STATICEDGE Or WS_EX_CLIENTEDGE Or WS_EX_WINDOWEDGE
      Case ebscFixedSingle
         ' Default border:
         pSetWinStyle GWL_STYLE, 0, WS_BORDER Or WS_THICKFRAME
         pSetWinStyle GWL_EXSTYLE, WS_EX_CLIENTEDGE, WS_EX_STATICEDGE Or WS_EX_WINDOWEDGE
      Case ebscThin
         ' Thin style
         pSetWinStyle GWL_STYLE, 0, WS_BORDER Or WS_THICKFRAME
         pSetWinStyle GWL_EXSTYLE, WS_EX_STATICEDGE, WS_EX_CLIENTEDGE Or WS_EX_WINDOWEDGE
      Case ebscRaised
         pSetWinStyle GWL_STYLE, WS_BORDER Or WS_THICKFRAME, 0
         pSetWinStyle GWL_EXSTYLE, WS_EX_WINDOWEDGE, WS_EX_STATICEDGE Or WS_EX_CLIENTEDGE
      End Select
   End If
   
End Sub

'
'   �@�\: �E�B���h�E�X�^�C�����Z�b�g����
'
'   ���l: �Ȃ�
'
Private Sub pSetWinStyle(ByVal lType As Long, ByVal lStyle As Long, ByVal lStyleNot As Long)
Dim lS As Long
Dim lhWNd As Long
   lhWNd = UserControl.hwnd
   lS = GetWindowLong(lhWNd, lType)
   lS = lS And Not lStyleNot
   lS = lS Or lStyle
   SetWindowLong lhWNd, lType, lS
   SetWindowPos lhWNd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_FRAMECHANGED
End Sub


'
'   �@�\: �����R���g���[����Visible���Z�b�g����
'
'   ���l: �Ȃ�
'
Private Sub VisibleAllContained(blnVisible As Boolean)
    Dim i As Long
    
    For i = 0 To UserControl.ContainedControls.count - 1
        UserControl.ContainedControls.item(i).Visible = blnVisible
    Next i
End Sub

'
'   �@�\: ���[�U�R���g���[���̏�����
'
'   ���l: �Ȃ�
'
Private Sub UserControl_Initialize()
On Error GoTo ErrorHandler
    tmrAnim.Interval = 100
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub

'
'   �@�\: ���[�U�R���g���[���̃}�E�X���[�u�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrorHandler
    With Extender
    End With
    
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
    With lblMode0
        .Left = (ScaleWidth / 2) - (.width / 2)
        .Top = (ScaleHeight / 2) - (.Height / 2) - lblAnim.Height
    End With
    With lblAnim
        .Left = (ScaleWidth / 2) - (.width / 2)
        .Top = (ScaleHeight / 2) - (.Height / 2) + lblMode0.Height
    End With
    With lblMode1
        .Left = (ScaleWidth / 2) - (.width / 2)
        .Top = (ScaleHeight / 2) - (.Height / 2)
    End With
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub

'
'   �@�\: Anim�^�C�}�[�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub tmrAnim_Timer()
On Error GoTo ErrorHandler
    Dim numBlock As Long
    Dim i As Long
    Dim strOut As String
    
    numBlock = 16
    mlngPos = mlngPos + 1
    If mlngPos > numBlock + 3 Then
        mlngPos = 0
    End If
    strOut = "["
    For i = 0 To numBlock - 1
        strOut = strOut & IIf(i < mlngPos And i > mlngPos - 4, "��", "��")
    Next i
    strOut = strOut & "]"
    lblAnim = strOut
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub

