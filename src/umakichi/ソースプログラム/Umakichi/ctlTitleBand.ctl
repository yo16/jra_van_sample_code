VERSION 5.00
Begin VB.UserControl ctlTitleBand 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.ComboBox cboRace 
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1080
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   2
      Top             =   600
      Width           =   2520
   End
   Begin VB.ComboBox cboLocation 
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1095
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   1
      Top             =   120
      Width           =   1425
   End
   Begin VB.Label lblWrapped 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
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
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   750
   End
End
Attribute VB_Name = "ctlTitleBand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'   �^�C�g���o���h
'

Option Explicit

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �O���֐�(�C�x���g)
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

' RaceChanger �� ���[�U�[�ɂ���ĕύX���ꂽ���ɔ�������C�x���g
Public Event Change(Key As clsKeyRA)

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �����ϐ�
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private WithEvents mData As clsDataRaceChanger
Attribute mData.VB_VarHelpID = -1
Private mKey As clsKeyRASel
Private mblnTKMode As Boolean

Private mlngPreLocationIndex As Long    'cboLocation�̒��O��ListIndex
Private mlngPreRaceIndex As Long

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �v���p�e�B
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   �@�\: Caption�v���p�e�B�̎擾
'
'   ���l: �Ȃ�
'
Public Property Get Caption() As String
    Caption = lblWrapped.Caption
End Property

'
'   �@�\: Caption�v���p�e�B�̃Z�b�g
'
'   ���l: �Ȃ�
'
Public Property Let Caption(ByVal RHS As String)
    lblWrapped.Visible = False
    lblWrapped.Caption = ReplaceAmpersand(RHS)
End Property


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �O���֐�
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   �@�\: RaceChanger�̕\�����N�G�X�g
'
'   ���l: ������ Key - ���[�X�I����ʗp�̃L�[�N���X
'                blnTKMode - ���ʓo�^�ꃂ�[�h�A���ʓo�^���ʂ̎���True�ŌĂ΂��
'
'         Key �� Nothing�̎��͔�\���ɂ���
'
Public Sub ShowRaceChanger(Key As clsKeyRASel, Optional blnTKMode As Boolean = False)
On Error GoTo ErrorHandler
    Dim i As Long   '' LocationIndex
    Dim j As Long   '' RaceIndex
    Dim lngLocationIndex As Long    '' mdata.RaceKey�̃C���f�b�N�X
    Dim lngRaceIndex As Long        '' mdata.RaceKey�̃C���f�b�N�X
    
    
    If Key Is Nothing Then
        ' �o�n�\�ȊO�̉�ʂł́ANothing�ŌĂ΂��B
        cboLocation.Visible = False '��\��
        cboRace.Visible = False
    Else
        cboLocation.Visible = False  '�\��
        cboRace.Visible = False
        cboLocation.Enabled = False 'Disable    ListIndex�v���p�e�B�ɑ�������Click�C�x���g����������̂�
        cboRace.Enabled = False     '           �C�x���g���ł�Enabled�v���p�e�B��False�̂Ƃ��̓C�x���g�������Ȃ�
        If (Key.JyoCD > "10") Then      'JRA�ȊO��Disable�̂܂�
            ' �n���A�C�O���[�X�ł́A�R���{�{�b�N�X�͕\�����邪�󗓂őI��s�Ƃ���B
            cboLocation.Clear
            cboRace.Clear
            Set mKey = New clsKeyRASel
        Else                            'JRA��cbo���ăZ�b�g�E�I������Enable�ɂ���
        
            ' ��I���R���{�{�b�N�X��\������Ƃ�
            ' ���Ƃ��ƃ^�C�g���o���h�ɕ\�����Ă��������Əd�������
            ' �o�n�\�^�C�g��������̗j���� ")" �ȍ~���폜����
            ' �@���o�n�\�^�C�g��������̓E�C���h�E�^�C�g���A�����ɂ��p���Ă����
            ' �@�@�^�C�g���o���h�ł݂̗̂�O�����Ƃ��āA�����ō폜���Ă���B
            Me.Caption = Left(Me.Caption, InStr(Me.Caption, ")"))

            If (Left(Key.str, 8) <> Left(mKey.str, 8)) Or mblnTKMode <> blnTKMode Then '�J�Ó����ς������Ď擾
                Set mData = New clsDataRaceChanger
                Call mData.Fetch(Key, blnTKMode)
            Else                                            '�J�Ó��������Ȃ�
                Call GetRaceKeyIndex(Key, lngLocationIndex, lngRaceIndex)
                If Key.JyoCD <> mKey.JyoCD Then             '�J�Ïꂪ�ς������
                    cboLocation.ListIndex = lngLocationIndex    'cboLocation�I��
                    Call SetCboRace(lngLocationIndex)           'cboRace�����ւ�
                End If
                cboRace.ListIndex = lngRaceIndex    'cboRace�I��
            End If
            Set mKey = Key  'Key��ۑ�
            mblnTKMode = blnTKMode
            cboLocation.Enabled = True  'Enable
            cboRace.Enabled = True
        End If
        cboLocation.Visible = True  '�\��
        cboRace.Visible = True
    End If

    Call resize
    lblWrapped.Visible = True
    Exit Sub

ErrorHandler:
    gApp.ErrLog
End Sub

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �����֐�
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   �@�\: Key�ƈ�v����RaceChanger�N���X��Key�z��v���p�e�B�̃C���f�b�N�X���擾����
'
'   ���l: �Ȃ�
'
Private Sub GetRaceKeyIndex(Key As clsKeyRASel, lngLocationIndex, lngRaceIndex)
    Dim i As Long
    Dim j As Long
    
    For i = 0 To mData.LocationCount - 1
        For j = 0 To mData.RaceCount(i) - 1
            If mData.RaceKey(i, j) = Key.str Then
                lngLocationIndex = i
                lngRaceIndex = j
                Exit Sub
            End If
        Next j
    Next i

End Sub


'
'   �@�\: ��R���{��RaceChanger�N���X����A�C�e�����Z�b�g����
'
'   ���l: �Ȃ�
'
Private Sub SetCboLocation()
    Dim i As Long
    
    cboLocation.Clear
    For i = 0 To mData.LocationCount - 1
        cboLocation.AddItem mData.LocationName(i)
    Next

End Sub


'
'   �@�\: ���[�X�R���{��RaceChanger�N���X����A�C�e�����Z�b�g����
'
'   ���l: �Ȃ�
'
Private Sub SetCboRace(lngLocationIndex)
    Dim i As Long
    
    cboRace.Clear
    For i = 0 To mData.RaceCount(lngLocationIndex) - 1
        cboRace.AddItem mData.RaceName(lngLocationIndex, i)
    Next
    
End Sub

'
'   �@�\: �R���{�A�C�e���I���C�x���g
'
'   ���l: ���[�X�R���{�̃A�C�e����I��������ɓ���ւ�
'
Private Sub cboLocation_Click()
On Error GoTo ErrorHandler
    Dim i As Long
    Dim lngNewLocationIndex As Long
    Dim strPreRaceNum As String                     '���݂̃��[�X�ԍ�
    Dim lngNewRaceIndex As Long                     'cboRace�ɃZ�b�g����C���f�b�N�X
    Dim blnFlag As Boolean
    Dim Key As clsKeyRA
    Dim lngSmallerRaceNumIndex As Long
    Dim lngLargerRaceNumIndex As Long
    
    If cboLocation.Enabled = False Then             '�����Ȃǂő��삳�ꂽ�Ƃ�
        mlngPreLocationIndex = cboLocation.ListIndex
        Exit Sub
    Else                                            '�R���{�{�b�N�X�𑀍삳�ꂽ�Ƃ�
        
        lngNewLocationIndex = cboLocation.ListIndex '�I��������
        
        If (lngNewLocationIndex <> mlngPreLocationIndex) Then   '���ύX����
            
            strPreRaceNum = Left(cboRace.Text, 2)       '�I�����Ă������[�X�ԍ�
            
            Call SetCboRace(lngNewLocationIndex)        '���[�X�R���{�����ւ�
        
            '���ύX�����Ƃ��������[�X�ԍ����Ȃ��ꍇ�A�ł��߂����[�X�ԍ���I������
            lngSmallerRaceNumIndex = -1
            lngLargerRaceNumIndex = -1
            blnFlag = False
            For i = 0 To mData.RaceCount(lngNewLocationIndex) - 1
                If (strPreRaceNum = Left(cboRace.List(i), 2)) Then  '�������[�X�ԍ���������
                    lngNewRaceIndex = i     '�������[�X�ԍ�
                    blnFlag = True
                    Exit For
                ElseIf (strPreRaceNum > Left(cboRace.List(i), 2)) Then  '�O�̃��[�X
                    lngSmallerRaceNumIndex = i                          '
                ElseIf (strPreRaceNum < Left(cboRace.List(i), 2)) Then  '��̃��[�X
                    lngLargerRaceNumIndex = i                           '
                    Exit For                                            '
                End If
            Next
            If blnFlag = False Then         '�������[�X�ԍ����Ȃ��Ƃ�
                If (lngSmallerRaceNumIndex = -1) Then   '�O�̃��[�X������
                    If (lngLargerRaceNumIndex = -1) Then    '��̃��[�X������
                        lngNewRaceIndex = 0                     '�G���[�A�ŏ��̃��[�X���Z�b�g����
                    Else                                    '��̃��[�X������
                        lngNewRaceIndex = lngLargerRaceNumIndex '��̃��[�X�ԍ�
                    End If
                Else                                    '�O�̃��[�X������
                    If (lngLargerRaceNumIndex = -1) Then    '��̃��[�X������
                        lngNewRaceIndex = lngSmallerRaceNumIndex    '�O�̃��[�X�ԍ�
                    Else                                    '���̃��[�X������
                        If strPreRaceNum - Left(cboRace.List(lngSmallerRaceNumIndex), 2) >= Left(cboRace.List(lngLargerRaceNumIndex), 2) - strPreRaceNum Then   '�ǂ��炪�߂���
                            lngNewRaceIndex = lngLargerRaceNumIndex     '�������O�������Ƃ��͌�̃��[�X�ԍ�
                        Else
                            lngNewRaceIndex = lngSmallerRaceNumIndex    '�O���߂��Ƃ��͑O�̃��[�X�ԍ�
                        End If
                    End If
                End If
            End If
            cboRace.ListIndex = lngNewRaceIndex     'cboRace_Click����
        End If
    End If
    Exit Sub

ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �R���{�A�C�e���I���C�x���g
'
'   ���l: �I���������[�X�ŉ�ʂ��X�V
'
Private Sub cboRace_Click()
On Error GoTo ErrorHandler
    Dim Key As clsKeyRA
    
    If cboRace.Enabled = False Then             '�����Ȃǂő��삳�ꂽ�Ƃ�
        mlngPreRaceIndex = cboRace.ListIndex
    Else                                        '�R���{�{�b�N�X�𑀍삳�ꂽ�Ƃ�
        If (mlngPreLocationIndex <> cboLocation.ListIndex) Or (mlngPreRaceIndex <> cboRace.ListIndex) Then
            Set Key = New clsKeyRA
            Key.str = mData.RaceKey(cboLocation.ListIndex, cboRace.ListIndex)
    
            RaiseEvent Change(Key)              '�o�n�\��\��
            
            mlngPreLocationIndex = cboLocation.ListIndex    '�R���{�̃C���f�b�N�X��ۑ�
            mlngPreRaceIndex = cboRace.ListIndex
        End If
    End If
    Exit Sub
    
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �f�[�^�擾�I���C�x���g�A�R���{�Ɏ擾�����l���Z�b�g����
'
'   ���l: �Ȃ�
'
'
Private Sub mData_FetchComplete(locationIndex As Long, raceIndex As Long)
On Error GoTo ErrorHandler
    
    cboLocation.Enabled = False     'Disable
    cboRace.Enabled = False
    
    Call SetCboLocation             '�A�C�e�����Z�b�g
    Call SetCboRace(locationIndex)
    
    cboLocation.ListIndex = locationIndex   '�A�C�e����I��
    cboRace.ListIndex = raceIndex
    
    cboLocation.Enabled = True      'Enable
    cboRace.Enabled = True
    Exit Sub

ErrorHandler:
    gApp.ErrLog
End Sub

'
'   �@�\: ���[�U�R���g���[���̏�����
'
'   ���l: �Ȃ�
'
'
Private Sub UserControl_Initialize()

    Set mKey = New clsKeyRASel
    Set mData = New clsDataRaceChanger
    
End Sub


'
'   �@�\: ���[�U�R���g���[���̃v���p�e�B�擾
'
'   ���l: �Ȃ�
'
'
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error GoTo ErrorHandler
    lblWrapped.Caption = ReplaceAmpersand(PropBag.ReadProperty("Caption", "Label1"))
    Call resize
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub

'
'   �@�\: ���[�U�R���g���[���̃��T�C�Y�C�x���g
'
'   ���l: �Ȃ�
'
'
Private Sub UserControl_Resize()
On Error GoTo ErrorHandler
    Call resize
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: ���[�U�R���g���[���̃v���p�e�B�Z�b�g
'
'   ���l: �Ȃ�
'
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
On Error GoTo ErrorHandler
    Call PropBag.WriteProperty("Caption", lblWrapped.Caption, "Label1")
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �R���g���[���̃��C�A�E�g
'
'   ���l: �L���v�V�������ݒ肳��A���x����AutoSize�ŕύX���ꂽ�Ƃ�
'
Private Sub resize()
    Dim p As Single
    
    p = lblWrapped.Left + lblWrapped.width
    
    '���x���ƃR���{�̋���
    p = p + 200
    
    ' ��R���{�̍Ĕz�u
    cboLocation.Move p, 0
    p = p + cboLocation.width
    
    '��R���{��Race�R���{�̋���
    p = p + 100
    
    ' Race�R���{�̍Ĕz�u
    cboRace.Move p, 0
    p = p + cboRace.width
    
    UserControl.width = p
    UserControl.Height = lblWrapped.Height + lblWrapped.Top * 2
End Sub
