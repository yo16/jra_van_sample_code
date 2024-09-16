VERSION 5.00
Begin VB.Form frmDirRef 
   Caption         =   "[ Message ]"
   ClientHeight    =   4590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   ScaleHeight     =   4590
   ScaleWidth      =   5880
   StartUpPosition =   1  '��Ű ̫�т̒���
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   300
      Left            =   3270
      TabIndex        =   2
      Top             =   4200
      Width           =   1230
   End
   Begin VB.TextBox txtPath 
      Height          =   300
      Left            =   60
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   3810
      Width           =   4110
   End
   Begin VB.CommandButton cmdNewBox 
      Caption         =   "�V�K�t�H���_�쐬"
      Height          =   300
      Left            =   4260
      TabIndex        =   1
      Top             =   3810
      Width           =   1560
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "�L�����Z��"
      Height          =   300
      Left            =   4590
      TabIndex        =   3
      Top             =   4200
      Width           =   1230
   End
   Begin VB.DriveListBox drvChoice 
      Height          =   300
      Left            =   60
      TabIndex        =   4
      Top             =   90
      Width           =   2700
   End
   Begin VB.DirListBox dirChoice 
      Height          =   3240
      Left            =   60
      TabIndex        =   5
      Top             =   480
      Width           =   5760
   End
End
Attribute VB_Name = "frmDirRef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'   �t�H���_��I������_�C�A���O�{�b�N�X
'
'   �N�����ɑI������Ă���p�X���A�N���O�� BeginingPath Property �ŁA�ݒ�\�B
'   �N���O�ɁAString�� Message Property �ɐݒ�\�B
'   ���̃t�H�[���́A�K�� .Show vbModal �ŌĂяo���B
'   �N����� ReturnPath Property �őI������FolderPath��ǂ߂�B

Option Explicit

Private mstrBeginingPath As String
Private mstrReturnPath As String
Private mstrMessage As String
Private mstrCurrentPath As String
Private mintCurWindowHeight As Integer
Private mintCurWindowWidth As Integer


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �v���p�e�B
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   �@�\: ���b�Z�[�W��ݒ肷��
'
'   ���l: �Ȃ�
'
Public Property Let Message(RHS As String)
    mstrMessage = RHS
End Property


'
'   �@�\: �N������ Path ��ݒ肷��
'
'   ���l: �Ȃ�
'
Public Property Let BeginingPath(RHS As String)
    mstrBeginingPath = RHS
End Property


'
'   �@�\: �I�����ꂽ�p�X��Ԃ�
'
'   ���l: �Ȃ�
'
Public Property Get ReturnPath() As String
    ReturnPath = mstrReturnPath
End Property


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �����֐�
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   �@�\: �L�����Z���{�^���C�x���g
'
'   ���l: �Ȃ�
'
Private Sub cmdCancel_Click()
On Error GoTo ErrorHandler
    mstrReturnPath = mstrBeginingPath
    Unload Me
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub

'
'   �@�\: �V�K�쐬�{�^���C�x���g
'
'   ���l: �Ȃ�
'
Private Sub cmdNewBox_Click()
On Error GoTo ErrorHandler
    Dim NewFolderDialog As frmNewFolder
    
    Set NewFolderDialog = New frmNewFolder
    NewFolderDialog.PathParam = txtPath.Text
    NewFolderDialog.Show vbModal, Me
    
    txtPath.Text = NewFolderDialog.ReturnPath
    dirChoice.Path = NewFolderDialog.ReturnPath
    dirChoice.Refresh
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �n�j�{�^���C�x���g
'
'   ���l: �Ȃ�
'
Private Sub cmdOK_Click()
On Error GoTo ErrorHandler
    mstrReturnPath = txtPath.Text
    Unload Me
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �f�B���N�g���I�����X�g�{�b�N�X�ύX�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub dirChoice_Change()
On Error GoTo ErrorHandler
    txtPath.Text = dirChoice.Path
    'dirChoice.Refresh
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �h���C�u�I�����X�g�{�b�N�X�ύX�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub drvChoice_Change()
On Error GoTo ErrorHandler
    dirChoice.Path = drvChoice.Drive
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �t�H�[�����[�h�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub Form_Load()
On Error GoTo ErrorHandler
    gApp.Log "DirRef"
    Me.Icon = LoadResPicture(100, vbResIcon)
    Dim fso As FileSystemObject
    
    gApp.Log "DirRef FSO NEW >"
    Set fso = New FileSystemObject
    gApp.Log "DirRef FSO NEW <"
   
    mintCurWindowHeight = Me.Height
    mintCurWindowWidth = Me.width
       
    If PathExistence_Check(mstrBeginingPath) Then
        drvChoice.Drive = fso.GetFolder(mstrBeginingPath).Drive
        dirChoice.Path = fso.GetFolder(mstrBeginingPath).Path
        mstrCurrentPath = mstrBeginingPath
    Else
        drvChoice.Drive = fso.GetFolder(App.Path).Drive
        dirChoice.Path = fso.GetFolder(App.Path).Path
        mstrCurrentPath = App.Path
        
    End If
    txtPath.Text = mstrCurrentPath
    
    ' Message �ݒ�
    Me.Caption = IIf(mstrMessage <> "", mstrMessage, "�t�H���_��I�����Ă�������")
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �p�X�̗L�������`�F�b�N
'
'   ���l: �Ȃ�
'
Private Function PathExistence_Check(Path As String) As Boolean
    Dim fso As New FileSystemObject
    
    If fso.FolderExists(Path) = 0 Then
        PathExistence_Check = False
    Else
        PathExistence_Check = True
    End If

End Function


'
'   �@�\: �t�H�[�����T�C�Y�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub Form_Resize()
On Error GoTo ErrorHandler
    
    If Me.WindowState = 0 Then
        If Me.Height < 4995 Or Me.width < 6000 Then
            SetToSmallestWindow
        Else
            SetObjectDimensions
        End If
    End If
    
    If Me.WindowState = 2 Then
        SetObjectDimensions
    End If
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �e�L�X�g�p�X�ύX�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub txtPath_Change()
On Error GoTo ErrorHandler
    mstrCurrentPath = txtPath.Text
    dirChoice.Refresh
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �e�L�X�g���͏���
'
'   ���l: �Ȃ�
'
Private Sub txtPath_KeyPress(KeyAscii As Integer)
On Error GoTo ErrorHandler
    Dim fso As New FileSystemObject
    
    
    If KeyAscii = 13 Then
        If fso.FolderExists(txtPath.Text) Then
            drvChoice.Drive = fso.GetFolder(txtPath.Text).Drive
            dirChoice.Path = fso.GetFolder(txtPath.Text).Path
        Else
            MsgBox "�t�H���_��������܂���B" & vbCrLf, vbInformation, "�n�g�F�t�H���_�̐ݒ�G���["
        End If
    End If

    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �E�B���h�E�̍ŏ��l���Z�b�g
'
'   ���l: �Ȃ�
'
Private Sub SetToSmallestWindow()
    With Me
        .Height = 4995
        .width = 6000
    End With
    
    With drvChoice
        .width = 2700
        .Left = 60
        .Top = 90
    End With
    
    With dirChoice
        .Height = 3240
        .width = 5760
        .Left = 60
        .Top = 480
    End With
    
    With txtPath
        .Height = 300
        .width = 4110
        .Left = 60
        .Top = 3810
    End With
    
    With cmdNewBox
        .Height = 300
        .width = 1560
        .Left = 4260
        .Top = 3810
    End With
    
    With cmdOK
        .Height = 300
        .width = 1230
        .Left = 3270
        .Top = 4200
    End With
    
    With cmdCancel
        .Height = 300
        .width = 1230
        .Left = 4590
        .Top = 4200
    End With
End Sub

'
'   �@�\: �ő剻�����E�B���h�E�̒l���Z�b�g
'
'   ���l: �Ȃ�
'
Private Sub SetObjectDimensions()
    
    With drvChoice
        .width = CInt(Me.width * 0.45)
        .Left = 60
        .Top = 90
    End With
    
    With dirChoice
        .Height = Me.Height - 1755
        .width = Me.width - 240
        .Left = 60
        .Top = 480
    End With

    With txtPath
        .Height = 300
        .width = Me.width - 1890
        .Left = 60
        .Top = dirChoice.Height + 570
    End With

    With cmdNewBox
        .Height = 300
        .width = 1560
        .Left = txtPath.width + 150
        .Top = dirChoice.Height + 570
    End With

    With cmdOK
        .Height = 300
        .width = 1230
        .Left = Me.width - 2730
        .Top = txtPath.Top + 390
    End With

    With cmdCancel
        .Height = 300
        .width = 1230
        .Left = Me.width - 1410
        .Top = txtPath.Top + 390
    End With
End Sub
