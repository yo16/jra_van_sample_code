VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ctlToolBars 
   ClientHeight    =   1365
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2265
   ScaleHeight     =   1365
   ScaleWidth      =   2265
   Begin MSComctlLib.Toolbar tbrInner 
      Height          =   330
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      _Version        =   393216
   End
End
Attribute VB_Name = "ctlToolBars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'   �c�[���o�[
'

Option Explicit

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �O���֐�(�C�x���g)
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Public Event ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
Public Event ButtonMenuClick(Index As Integer, ByVal ButtonMenu As MSComctlLib.ButtonMenu)

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �����ϐ�
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private msngMinWidth As Single

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �v���p�e�B
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   �@�\: ToolBar�v���p�e�B�̎擾
'
'   ���l: �Ȃ�
'
Public Property Get ToolBar(Index As Long) As ToolBar
    Set ToolBar = tbrInner(Index)
End Property

'
'   �@�\: num�v���p�e�B�̃Z�b�g
'
'   ���l: �Ȃ�
'
Public Property Let num(RHS As Long)
    Dim i As Long
    ' ����Ȃ��������[�h����
    For i = tbrInner.UBound + 1 To RHS - 1
        Load tbrInner(i)
        tbrInner(i).Visible = True
    Next i
End Property

'
'   �@�\: MinWidth�v���p�e�B�̎擾
'
'   ���l: �Ȃ�
'
Public Property Get MinWidth() As Single
    MinWidth = msngMinWidth
End Property

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �O���֐�
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   �@�\: ���[�U�R���g���[���̍����𒲐�
'
'   ���l: �Ȃ�
'
Public Sub fit()
    Call StandInLine
    UserControl.Height = tbrInner(0).ButtonHeight
End Sub


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �����֐�
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'
'   �@�\: MinWidth�v���p�e�B�l�̃Z�b�g
'
'   ���l: �Ȃ�
'
Private Sub StandInLine()
    Dim i As Long
    Dim p As Single
    
    For i = 0 To tbrInner.UBound
        tbrInner(i).Left = p
        tbrInner(i).width = CalcButtonsWidth(tbrInner(i))
        If tbrInner(i).Visible = True Then
            p = p + tbrInner(i).width
        End If
    Next i
    msngMinWidth = p
End Sub

'
'   �@�\: �{�^���̕���Ԃ�
'
'   ���l: �Ȃ�
'
Private Function CalcButtonsWidth(tbrTarget As ToolBar) As Single
    Dim p As Single
    Dim i As Long
    For i = 1 To tbrTarget.Buttons.count
        p = p + tbrTarget.Buttons(i).width
    Next i
    CalcButtonsWidth = p
End Function

'
'   �@�\: �����c�[���o�[�̃{�^���N���b�N�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub tbrInner_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
On Error GoTo ErrorHandler
    RaiseEvent ButtonClick(Index, Button)
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub

'
'   �@�\: �����c�[���o�[�̃{�^�����j���[�N���b�N�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub tbrInner_ButtonMenuClick(Index As Integer, ByVal ButtonMenu As MSComctlLib.ButtonMenu)
On Error GoTo ErrorHandler
    RaiseEvent ButtonMenuClick(Index, ButtonMenu)
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
    Call StandInLine
    UserControl.Height = tbrInner(0).ButtonHeight
    UserControl.width = msngMinWidth * 10
    Exit Sub
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Sub
