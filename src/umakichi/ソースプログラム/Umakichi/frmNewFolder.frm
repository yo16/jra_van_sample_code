VERSION 5.00
Begin VB.Form frmNewFolder 
   BorderStyle     =   3  '�Œ��޲�۸�
   Caption         =   "�V�����t�H���_�쐬"
   ClientHeight    =   2100
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '��Ű ̫�т̒���
   Begin VB.TextBox txtNewFolder 
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   1320
      Width           =   4695
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "�L�����Z��"
      Height          =   285
      Left            =   4560
      TabIndex        =   2
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   285
      Left            =   3240
      TabIndex        =   1
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "�t�H���_��:"
      Height          =   180
      Left            =   120
      TabIndex        =   4
      Top             =   1380
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "�V�����t�H���_������͂��Ă��������B"
      Height          =   180
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2805
   End
End
Attribute VB_Name = "frmNewFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'   �t�H���_�쐬 �_�C�A���O
'

Option Explicit

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �����ϐ�
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private mstrPathParam As String
Private mstrReturnPath As String
Dim mfso As New FileSystemObject

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �v���p�e�B
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   �@�\: �N������ Path ��ݒ肷��
'
'   ���l: �Ȃ�
'
Public Property Let PathParam(CurrentPath As String)
    mstrPathParam = CurrentPath
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
'   �@�\: �I�����ꂽ�p�X��Ԃ�
'
'   ���l: �Ȃ�
'
Private Sub cmdCancel_Click()
On Error GoTo ErrorHandler
    mstrReturnPath = mstrPathParam
    Unload Me
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �I�����ꂽ�p�X��Ԃ�
'
'   ���l: �Ȃ�
'
Private Sub cmdOK_Click()
On Error GoTo ErrorHandler
    Dim strPath As String
    
    If Right$(mstrPathParam, 1) <> "\" Then mstrPathParam = mstrPathParam & "\"
    strPath = mstrPathParam & txtNewFolder
    
    If Not mfso.FolderExists(strPath) Then
        mfso.CreateFolder strPath
        mstrReturnPath = strPath
        Unload Me
    Else
        MsgBox "�����̃t�H���_�����łɑ��݂��Ă��܂��B", vbExclamation, "�n�g�F�V�K�t�H���_�̍쐬�G���["
        
        txtNewFolder.SetFocus
    End If
    
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
    Me.Icon = LoadResPicture(100, vbResIcon)
    cmdOK.Enabled = False
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �t�H�[���A�����[�h�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrorHandler
    Set mfso = Nothing
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �t�H���_���ύX�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub txtNewFolder_Change()
On Error GoTo ErrorHandler
    
    If txtNewFolder.Text = "" Then
        cmdOK.Enabled = False
    End If
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �t�H���_�����̓C�x���g
'
'   ���l: �Ȃ�
'
Private Sub txtNewFolder_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrorHandler
    
    If txtNewFolder.Text <> "" Then
        cmdOK.Enabled = True
    End If
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub
