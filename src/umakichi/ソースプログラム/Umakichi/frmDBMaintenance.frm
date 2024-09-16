VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDBMaintenance 
   Caption         =   "DB�����e�i���X"
   ClientHeight    =   3285
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7650
   LinkTopic       =   "Form1"
   ScaleHeight     =   3285
   ScaleWidth      =   7650
   StartUpPosition =   3  'Windows �̊���l
   Begin VB.Timer tmrMaintenance 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1560
      Top             =   1680
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "�L�����Z��"
      Height          =   405
      Left            =   210
      TabIndex        =   1
      Top             =   1680
      Width           =   1185
   End
   Begin MSComctlLib.ProgressBar prgBar 
      Height          =   345
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   330
      Width           =   7245
      _ExtentX        =   12779
      _ExtentY        =   609
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.ProgressBar prgBar 
      Height          =   345
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   1110
      Width           =   7245
      _ExtentX        =   12779
      _ExtentY        =   609
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   60
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      Caption         =   "�o�ߎ���: 0:00:00"
      Height          =   180
      Left            =   240
      TabIndex        =   3
      Top             =   2280
      Width           =   1320
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   60
      Width           =   60
   End
End
Attribute VB_Name = "frmDBMaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'   �f�[�^�x�[�X�œK���i�����
'

Option Explicit

Private WithEvents mAsync_Cn As ADODB.Connection
Attribute mAsync_Cn.VB_VarHelpID = -1

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �����ϐ�
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private mCn As ADODB.Connection
Private mCat As New ADOX.Catalog
Private mfso As New FileSystemObject
Private mDBFolder As Folder
Private mstrSQL As String
Private mstrTableName(0 To 45) As String
Private msngTimerStart As Single
Private mblnCancelFlag As Boolean
Private mblnNoTask As Boolean
Private i As Long

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �����֐�
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   �@�\: �t�H�[��������
'
'   ���l: �Ȃ�
'
Private Sub Form_Initialize()
On Error GoTo ErrorHandler
    Dim strFolder As String
    
    Me.Icon = LoadResPicture(100, vbResIcon)
    
    msngTimerStart = Timer
    
    mblnNoTask = False
    
    mblnCancelFlag = False

    mstrTableName(0) = "BANUSI"
    mstrTableName(1) = "BATAIJYU"
    mstrTableName(2) = "CHOKYO"
    mstrTableName(3) = "CHOKYO_SEISEKI"
    mstrTableName(4) = "HANRO"
    mstrTableName(5) = "HANSYOKU"
    mstrTableName(6) = "HARAI"
    mstrTableName(7) = "KISHU"
    mstrTableName(8) = "KISHU_CHANGE"
    mstrTableName(9) = "KISHU_SEISEKI"
    mstrTableName(10) = "MINING"
    mstrTableName(11) = "ODDS_SANREN"
    mstrTableName(12) = "ODDS_SANREN"
    mstrTableName(13) = "ODDS_SANREN"
    mstrTableName(14) = "ODDS_SANREN"
    mstrTableName(15) = "ODDS_SANREN"
    mstrTableName(16) = "ODDS_SANREN"
    mstrTableName(17) = "ODDS_SANREN"
    mstrTableName(18) = "ODDS_SANREN"
    mstrTableName(19) = "ODDS_SANREN"
    mstrTableName(20) = "ODDS_SANREN"
    mstrTableName(21) = "ODDS_TANPUKUWAKU"
    mstrTableName(22) = "ODDS_UMAREN"
    mstrTableName(23) = "ODDS_UMATAN"
    mstrTableName(24) = "ODDS_UMATAN"
    mstrTableName(25) = "ODDS_UMATAN"
    mstrTableName(26) = "ODDS_UMATAN"
    mstrTableName(27) = "ODDS_UMATAN"
    mstrTableName(28) = "ODDS_UMATAN"
    mstrTableName(29) = "ODDS_UMATAN"
    mstrTableName(30) = "ODDS_UMATAN"
    mstrTableName(31) = "ODDS_UMATAN"
    mstrTableName(32) = "ODDS_UMATAN"
    mstrTableName(33) = "ODDS_WIDE"
    mstrTableName(34) = "RACE"
    mstrTableName(35) = "RECORD"
    mstrTableName(36) = "SANKU"
    mstrTableName(37) = "SCHEDULE"
    mstrTableName(38) = "SEISAN"
    mstrTableName(39) = "TENKO_BABA"
    mstrTableName(40) = "TOKU"
    mstrTableName(41) = "TOKU_RACE"
    mstrTableName(42) = "TORIKESI_JYOGAI"
    mstrTableName(43) = "UMA"
    mstrTableName(44) = "UMA_RACE_A"
    mstrTableName(45) = "UMA_RACE_B"
    
    strFolder = "UK" & Format$(Date, "YYYYMMDD") & Format$(Now, "HHMM")
    
    Set mAsync_Cn = New ADODB.Connection

    With prgBar(0)
        .Min = 0
        .max = 45
        .value = 0
    End With
    
    ' Create temporary folder
    If Not mfso.FolderExists(gApp.R_DBPath & "\" & strFolder) Then
        Set mDBFolder = mfso.CreateFolder(gApp.R_DBPath & "\" & strFolder)
    End If
    Set mDBFolder = mfso.GetFolder(gApp.R_DBPath & "\" & strFolder)
    If mDBFolder.Files.count <> 0 Then kill mDBFolder.Path & "\*.*"
        
    i = 0
    mblnNoTask = True
    
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
    tmrMaintenance.Enabled = True
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �L�����Z���{�^���C�x���g
'
'   ���l: �Ȃ�
'
Private Sub cmdCancel_Click()
On Error GoTo ErrorHandler
    tmrMaintenance.Enabled = False
    Enabled = False
    mblnCancelFlag = True
    gApp.Log gstrMDBName(i) & " Maintenace was Cancelled: " & Now
    If mAsync_Cn.State <> adStateClosed Then
        Do
            If mAsync_Cn.State <> adStatusUnwantedEvent Then Exit Do
        Loop
        
        Select Case mAsync_Cn.State
        Case adStateOpen
            mAsync_Cn.Cancel
            mAsync_Cn.Close
        Case adStateConnecting
            mAsync_Cn.Cancel
            mAsync_Cn.Close
        Case adStateExecuting
            mAsync_Cn.Cancel
            mAsync_Cn.Close
        Case adStateFetching
            mAsync_Cn.Cancel
            mAsync_Cn.Close
        End Select

    End If
    Set mAsync_Cn = Nothing
    Unload Me
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �œK���^�C�}�[�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub tmrMaintenance_Timer()
On Error GoTo ErrorHandler
    Dim sngTimeElapsed As Single
    Dim strSeconds As String
    Dim strMinutes As String
    Dim strHours As String
    Dim strTimeElapsed As String
    
    If i < 46 Then
        prgBar(0).value = i
        lblInfo(0).Caption = gstrMDBName(i) & " (" & i & "/45)"
        Me.Refresh
    End If
    
    ' start block: Time Compute and Display
        sngTimeElapsed = Timer - msngTimerStart
        
        ' seconds format
        If (CInt(sngTimeElapsed) Mod 60) < 10 Then
            strSeconds = "0" & (CInt(sngTimeElapsed) Mod 60)
        Else
            strSeconds = CInt(sngTimeElapsed) Mod 60
        End If
        
        ' minutes format
        If ((CInt(sngTimeElapsed) \ 60) Mod 60) < 10 Then
            strMinutes = "0" & (CInt(sngTimeElapsed) \ 60) Mod 60
        Else
            strMinutes = (CInt(sngTimeElapsed) \ 60) Mod 60
        End If
        
        ' hours format
        strHours = CInt(sngTimeElapsed) \ 3600
        
        strTimeElapsed = strHours & ":" & strMinutes & ":" & strSeconds
        lblTime.Caption = "�o�ߎ���: " & strTimeElapsed
        
        Me.Refresh
    ' end block: Time Compute and Display
    
    If mblnNoTask Then
        If i < 46 Then
            ' Create mdb File
            mCat.Create "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & mDBFolder.Path & "\" & gstrMDBName(i)
            
            ' Set Connection
            mAsync_Cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                           "Data Source=" & gApp.R_DBPath & "\" & gstrMDBName(i)
                    
            ' Copy Table
            mstrSQL = "SELECT " & mstrTableName(i) & ".* INTO " & mstrTableName(i) & _
                      " IN '" & mDBFolder.Path & "\" & gstrMDBName(i) & "'" & _
                      " FROM " & mstrTableName(i)
            
            mAsync_Cn.Execute mstrSQL, , adAsyncExecute
            gApp.Log gstrMDBName(i) & " Connection's execution initiated: " & Now
                           
            mblnNoTask = False
            
        Else
            Set mCat = Nothing
            Set mAsync_Cn = Nothing
            Unload Me
        End If
    End If
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
    Set mCat = Nothing
    
    If mblnCancelFlag Then
        MsgBox "���f���܂����B", vbInformation, "�n�g�F�f�[�^�x�[�X�����e�i���X"
        gApp.Log "���f���܂����B"
        If mfso.FileExists(mDBFolder.Path & "\" & gstrMDBName(i)) Then
            Call mfso.DeleteFile(mDBFolder.Path & "\" & gstrMDBName(i), True)
        End If
    End If
           
    Call mfso.DeleteFolder(mDBFolder.Path, True)
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �񓯊��R�l�N�V�������s�����C�x���g
'
'   ���l: �Ȃ�
'
Private Sub mAsync_Cn_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
On Error GoTo ErrorHandler
    Dim IndexMaker As clsCreateMDB
    
    Dim lngReturnCode As Long
    gApp.Log gstrMDBName(i) & " Connection's execution completed: " & Now
    mAsync_Cn.Close
    gApp.Log gstrMDBName(i) & " Connection's connection closed: " & Now
    
    Set mCat = Nothing
    
    Set IndexMaker = New clsCreateMDB
    Set IndexMaker.mConnection = New ADODB.Connection
    
    With prgBar(1)
        .max = i + 1
        .Min = 0
        .value = i
    End With

    
    If Not pError Is Nothing Then
        lblInfo(0).Caption = gstrMDBName(i) & " (" & i & "/45): �G���[������܂����B"
        lblInfo(1).Caption = mstrTableName(i) & "�e�[�u���C���f�b�N�X�̍쐬���X�L�b�v���܂��B"
        If pError.Number = -2147467259 Then
            MsgBox "�f�B�X�N�̋󂫗e�ʂ��s�����Ă��܂��B", vbCritical, "�n�g�F�f�B�X�N�X�y�[�X�G���["
            lblInfo(0).Caption = gstrMDBName(i) & " (" & i & "/45): �f�B�X�N�X�y�[�X�G���["
        End If
        
        Me.Refresh
        mfso.DeleteFile mDBFolder.Path & "\" & gstrMDBName(i)
        gApp.Log "Error " & pError.Number & "; " & pError.Description
    ElseIf mblnCancelFlag Then
        mfso.DeleteFile mDBFolder.Path & "\" & gstrMDBName(i)
    Else
        lblInfo(1).Caption = mstrTableName(i) & "�e�[�u���C���f�b�N�X���쐬���ł��B"
        
        ' Creat Keys and Indexes
        IndexMaker.mConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & mDBFolder.Path & "\" & gstrMDBName(i)
        Call CallByName(IndexMaker, "createIndex_" & mstrTableName(i), VbMethod)
        IndexMaker.mConnection.Close
        
        lblInfo(1).Caption = mstrTableName(i) & "�e�[�u���C���f�b�N�X���쐬���܂����B"
        
        Set IndexMaker.mConnection = Nothing
        Me.Refresh
        
        Do
            On Error Resume Next
            mfso.DeleteFile gApp.R_DBPath & "\" & gstrMDBName(i)
            If Err.Number = 70 Then
                lngReturnCode = MsgBox("�f�[�^�x�[�X�ɏ������߂܂���", vbAbortRetryIgnore + vbCritical, "�n�g�F�œK���G���[")
                Select Case lngReturnCode
                Case vbAbort
                    mblnCancelFlag = True
                    Exit Do
                Case vbRetry
                    '
                Case vbIgnore
                    mfso.DeleteFile mDBFolder.Path & "\" & gstrMDBName(i)
                    Exit Do
                End Select
            Else
                mfso.MoveFile mDBFolder.Path & "\" & gstrMDBName(i), gApp.R_DBPath & "\" & gstrMDBName(i)
                Exit Do
            End If
            On Error GoTo ErrorHandler
        Loop
        
    End If
    Me.Refresh
    
    i = i + 1
    If mblnCancelFlag Then
        Unload Me
    Else
        mblnNoTask = True
    End If
        
    Exit Sub
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   �@�\: �񓯊��R�l�N�V������񃁃b�Z�[�W�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub mAsync_Cn_InfoMessage(ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    gApp.Log pError.Number & pError.Description
    Resume Next
End Sub


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �O���֐�
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   �@�\: DB��TOKU_RACE(���ʃ��[�X)�e�[�u���C���f�b�N�X�����
'
'   ���l: �Ȃ�
'
Public Function SetKeys_Indexes_TOKU_RACE() As Boolean
On Error GoTo ErrorHandler
    Dim strSQL As String

    strSQL = "CREATE INDEX PrimaryKey ON TOKU_RACE ("
    strSQL = strSQL & "[Year],"
    strSQL = strSQL & "[MonthDay],"
    strSQL = strSQL & "[JyoCD],"
    strSQL = strSQL & "[Kaiji],"
    strSQL = strSQL & "[Nichiji],"
    strSQL = strSQL & "[RaceNum]"
    strSQL = strSQL & ") WITH PRIMARY"

    mCn.Execute strSQL

    SetKeys_Indexes_TOKU_RACE = True
    Exit Function
ErrorHandler:
    gApp.ErrLog
    SetKeys_Indexes_TOKU_RACE = False
End Function ' SetKeys_Indexes_TOKU_RACE


'
'   �@�\: DB��TOKU(���ʓo�^�n)�e�[�u���C���f�b�N�X�����
'
'   ���l: �Ȃ�
'
Public Function SetKeys_Indexes_TOKU() As Boolean
On Error GoTo ErrorHandler
    Dim strSQL As String

    strSQL = "CREATE INDEX PrimaryKey ON TOKU ("
    strSQL = strSQL & "[Year],"
    strSQL = strSQL & "[MonthDay],"
    strSQL = strSQL & "[JyoCD],"
    strSQL = strSQL & "[Kaiji],"
    strSQL = strSQL & "[Nichiji],"
    strSQL = strSQL & "[RaceNum],"
    strSQL = strSQL & "[Num]"
    strSQL = strSQL & ") WITH PRIMARY"

    mCn.Execute strSQL

    SetKeys_Indexes_TOKU = True
    Exit Function
ErrorHandler:
    gApp.ErrLog
    SetKeys_Indexes_TOKU = False
End Function ' SetKeys_Indexes_TOKU


'
'   �@�\: DB��RACE(���[�X�ڍ�)�e�[�u���C���f�b�N�X�����
'
'   ���l: �Ȃ�
'
Public Function SetKeys_Indexes_RACE() As Boolean
On Error GoTo ErrorHandler
    Dim strSQL As String

    strSQL = "CREATE INDEX PrimaryKey ON RACE ("
    strSQL = strSQL & "[Year],"
    strSQL = strSQL & "[MonthDay],"
    strSQL = strSQL & "[JyoCD],"
    strSQL = strSQL & "[Kaiji],"
    strSQL = strSQL & "[Nichiji],"
    strSQL = strSQL & "[RaceNum]"
    strSQL = strSQL & ") WITH PRIMARY"

    mCn.Execute strSQL


    strSQL = "CREATE INDEX Kyori ON RACE ( [Kyori] )"
    mCn.Execute strSQL
    
    SetKeys_Indexes_RACE = True
    Exit Function
ErrorHandler:
    gApp.ErrLog
    SetKeys_Indexes_RACE = False
End Function ' SetKeys_Indexes_RACE


'
'   �@�\: DB��UMA_RACE_A(�n�����[�X���_�O��)�e�[�u���C���f�b�N�X�����
'
'   ���l: �Ȃ�
'
Public Function SetKeys_Indexes_UMA_RACE_A() As Boolean
On Error GoTo ErrorHandler
    Dim strSQL As String

    strSQL = "CREATE INDEX PrimaryKey ON UMA_RACE_A ("
    strSQL = strSQL & "[Year],"
    strSQL = strSQL & "[MonthDay],"
    strSQL = strSQL & "[JyoCD],"
    strSQL = strSQL & "[Kaiji],"
    strSQL = strSQL & "[Nichiji],"
    strSQL = strSQL & "[RaceNum],"
    strSQL = strSQL & "[Umaban],"
    strSQL = strSQL & "[KettoNum]"
    strSQL = strSQL & ") WITH PRIMARY"

    mCn.Execute strSQL

    strSQL = "CREATE INDEX KettoNum ON UMA_RACE_A ( [KettoNum] )"
    mCn.Execute strSQL

    SetKeys_Indexes_UMA_RACE_A = True
    Exit Function
ErrorHandler:
    gApp.ErrLog
    SetKeys_Indexes_UMA_RACE_A = False
End Function ' SetKeys_Indexes_UMA_RACE_A


'
'   �@�\: DB��UMA_RACE_B(�n�����[�X���_�㔼)�e�[�u���C���f�b�N�X�����
'
'   ���l: �Ȃ�
'
Public Function SetKeys_Indexes_UMA_RACE_B() As Boolean
On Error GoTo ErrorHandler
    Dim strSQL As String

    strSQL = "CREATE INDEX PrimaryKey ON UMA_RACE_B ("
    strSQL = strSQL & "[B_Year],"
    strSQL = strSQL & "[B_MonthDay],"
    strSQL = strSQL & "[B_JyoCD],"
    strSQL = strSQL & "[B_RaceNum],"
    strSQL = strSQL & "[B_Umaban],"
    strSQL = strSQL & "[B_KettoNum]"
    strSQL = strSQL & ") WITH PRIMARY"

    mCn.Execute strSQL
    
    strSQL = "CREATE INDEX [Time] ON UMA_RACE_B ( [Time] )"
    mCn.Execute strSQL

    SetKeys_Indexes_UMA_RACE_B = True
    Exit Function
ErrorHandler:
    gApp.ErrLog
    SetKeys_Indexes_UMA_RACE_B = False
End Function ' SetKeys_Indexes_UMA_RACE_B


'
'   �@�\: DB��HARAI(����)�e�[�u���C���f�b�N�X�����
'
'   ���l: �Ȃ�
'
Public Function SetKeys_Indexes_HARAI() As Boolean
On Error GoTo ErrorHandler
    Dim strSQL As String

    strSQL = "CREATE INDEX PrimaryKey ON HARAI ("
    strSQL = strSQL & "[Year],"
    strSQL = strSQL & "[MonthDay],"
    strSQL = strSQL & "[JyoCD],"
    strSQL = strSQL & "[Kaiji],"
    strSQL = strSQL & "[Nichiji],"
    strSQL = strSQL & "[RaceNum]"
    strSQL = strSQL & ") WITH PRIMARY"

    mCn.Execute strSQL

    SetKeys_Indexes_HARAI = True
    Exit Function
ErrorHandler:
    gApp.ErrLog
    SetKeys_Indexes_HARAI = False
End Function ' SetKeys_Indexes_HARAI


'
'   �@�\: DB��ODDS_TANPUKUWAKU(�I�b�Y_�P���g)�e�[�u���C���f�b�N�X�����
'
'   ���l: �Ȃ�
'
Public Function SetKeys_Indexes_ODDS_TANPUKUWAKU() As Boolean
On Error GoTo ErrorHandler
    Dim strSQL As String

    strSQL = "CREATE INDEX PrimaryKey ON ODDS_TANPUKUWAKU ("
    strSQL = strSQL & "[Year],"
    strSQL = strSQL & "[MonthDay],"
    strSQL = strSQL & "[JyoCD],"
    strSQL = strSQL & "[Kaiji],"
    strSQL = strSQL & "[Nichiji],"
    strSQL = strSQL & "[RaceNum]"
    strSQL = strSQL & ") WITH PRIMARY"

    mCn.Execute strSQL

    SetKeys_Indexes_ODDS_TANPUKUWAKU = True
    Exit Function
ErrorHandler:
    gApp.ErrLog
    SetKeys_Indexes_ODDS_TANPUKUWAKU = False
End Function ' SetKeys_Indexes_ODDS_TANPUKUWAKU


'
'   �@�\: DB��ODDS_UMAREN(�I�b�Y_�n�A)�e�[�u���C���f�b�N�X�����
'
'   ���l: �Ȃ�
'
Public Function SetKeys_Indexes_ODDS_UMAREN() As Boolean
On Error GoTo ErrorHandler
    Dim strSQL As String

    strSQL = "CREATE INDEX PrimaryKey ON ODDS_UMAREN ("
    strSQL = strSQL & "[Year],"
    strSQL = strSQL & "[MonthDay],"
    strSQL = strSQL & "[JyoCD],"
    strSQL = strSQL & "[Kaiji],"
    strSQL = strSQL & "[Nichiji],"
    strSQL = strSQL & "[RaceNum]"
    strSQL = strSQL & ") WITH PRIMARY"

    mCn.Execute strSQL

    SetKeys_Indexes_ODDS_UMAREN = True
    Exit Function
ErrorHandler:
    gApp.ErrLog
    SetKeys_Indexes_ODDS_UMAREN = False
End Function ' SetKeys_Indexes_ODDS_UMAREN


'
'   �@�\: DB��ODDS_WIDE(�I�b�Y_���C�h)�e�[�u���C���f�b�N�X�����
'
'   ���l: �Ȃ�
'
Public Function SetKeys_Indexes_ODDS_WIDE() As Boolean
On Error GoTo ErrorHandler
    Dim strSQL As String

    strSQL = "CREATE INDEX PrimaryKey ON ODDS_WIDE ("
    strSQL = strSQL & "[Year],"
    strSQL = strSQL & "[MonthDay],"
    strSQL = strSQL & "[JyoCD],"
    strSQL = strSQL & "[Kaiji],"
    strSQL = strSQL & "[Nichiji],"
    strSQL = strSQL & "[RaceNum]"
    strSQL = strSQL & ") WITH PRIMARY"

    mCn.Execute strSQL

    SetKeys_Indexes_ODDS_WIDE = True
    Exit Function
ErrorHandler:
    gApp.ErrLog
    SetKeys_Indexes_ODDS_WIDE = False
End Function ' SetKeys_Indexes_ODDS_WIDE


'
'   �@�\: DB��ODDS_UMATAN(�I�b�Y_�n�P)�e�[�u���C���f�b�N�X�����
'
'   ���l: �Ȃ�
'
Public Function SetKeys_Indexes_ODDS_UMATAN() As Boolean
On Error GoTo ErrorHandler
    Dim strSQL As String

    strSQL = "CREATE INDEX PrimaryKey ON ODDS_UMATAN ("
    strSQL = strSQL & "[Year],"
    strSQL = strSQL & "[MonthDay],"
    strSQL = strSQL & "[JyoCD],"
    strSQL = strSQL & "[Kaiji],"
    strSQL = strSQL & "[Nichiji],"
    strSQL = strSQL & "[RaceNum]"
    strSQL = strSQL & ") WITH PRIMARY"

    mCn.Execute strSQL

    SetKeys_Indexes_ODDS_UMATAN = True
    Exit Function
ErrorHandler:
    gApp.ErrLog
    SetKeys_Indexes_ODDS_UMATAN = False
End Function ' SetKeys_Indexes_ODDS_UMATAN


'
'   �@�\: DB��ODDS_SANREN(�I�b�Y_3�A��)�e�[�u���C���f�b�N�X�����
'
'   ���l: �Ȃ�
'
Public Function SetKeys_Indexes_ODDS_SANREN() As Boolean
On Error GoTo ErrorHandler
    Dim strSQL As String

    strSQL = "CREATE INDEX PrimaryKey ON ODDS_SANREN ("
    strSQL = strSQL & "[Year],"
    strSQL = strSQL & "[MonthDay],"
    strSQL = strSQL & "[JyoCD],"
    strSQL = strSQL & "[Kaiji],"
    strSQL = strSQL & "[Nichiji],"
    strSQL = strSQL & "[RaceNum]"
    strSQL = strSQL & ") WITH PRIMARY"

    mCn.Execute strSQL

    SetKeys_Indexes_ODDS_SANREN = True
    Exit Function
ErrorHandler:
    gApp.ErrLog
    SetKeys_Indexes_ODDS_SANREN = False
End Function ' SetKeys_Indexes_ODDS_SANREN


'
'   �@�\: DB��UMA(�����n�}�X�^)�e�[�u���C���f�b�N�X�����
'
'   ���l: �Ȃ�
'
Public Function SetKeys_Indexes_UMA() As Boolean
On Error GoTo ErrorHandler
    Dim strSQL As String

    strSQL = "CREATE INDEX PrimaryKey ON UMA ("
    strSQL = strSQL & "[KettoNum]"
    strSQL = strSQL & ") WITH PRIMARY"

    mCn.Execute strSQL
    
    strSQL = "CREATE INDEX Bamei ON UMA ( [Bamei] )"
    mCn.Execute strSQL
    
    strSQL = "CREATE INDEX BameiKana ON UMA ( [BameiEng] )"
    mCn.Execute strSQL
    
    strSQL = "CREATE INDEX BameiEng ON UMA ( [BameiKana] )"
    mCn.Execute strSQL
    
    strSQL = "CREATE INDEX ChokyosiCode ON UMA ( [ChokyosiCode] )"
    mCn.Execute strSQL
    
    strSQL = "CREATE INDEX BreederCode ON UMA ( [BreederCode] )"
    mCn.Execute strSQL

    strSQL = "CREATE INDEX BanusiCode ON UMA ( [BanusiCode] )"
    mCn.Execute strSQL

    SetKeys_Indexes_UMA = True
    Exit Function
ErrorHandler:
    gApp.ErrLog
    SetKeys_Indexes_UMA = False
End Function ' SetKeys_Indexes_UMA


'
'   �@�\: DB��KISHU(�R��}�X�^)�e�[�u���C���f�b�N�X�����
'
'   ���l: �Ȃ�
'
Public Function SetKeys_Indexes_KISHU() As Boolean
On Error GoTo ErrorHandler
    Dim strSQL As String

    strSQL = "CREATE INDEX PrimaryKey ON KISHU ("
    strSQL = strSQL & "[KisyuCode]"
    strSQL = strSQL & ") WITH PRIMARY"

    mCn.Execute strSQL

    strSQL = "CREATE INDEX KisyuName ON KISHU ( [KisyuName] )"
    mCn.Execute strSQL
    strSQL = "CREATE INDEX KisyuNameKana ON KISHU ( [KisyuNameKana] )"
    mCn.Execute strSQL
    strSQL = "CREATE INDEX KisyuRyakusyo ON KISHU ( [KisyuRyakusyo] )"
    mCn.Execute strSQL
    strSQL = "CREATE INDEX KisyuNameEng ON KISHU ( [KisyuNameEng] )"
    mCn.Execute strSQL

    SetKeys_Indexes_KISHU = True
    Exit Function
ErrorHandler:
    gApp.ErrLog
    SetKeys_Indexes_KISHU = False
End Function ' SetKeys_Indexes_KISHU


'
'   �@�\: DB��KISHU_SEISEKI(�R��}�X�^_����)�e�[�u���C���f�b�N�X�����
'
'   ���l: �Ȃ�
'
Public Function SetKeys_Indexes_KISHU_SEISEKI() As Boolean
On Error GoTo ErrorHandler
    Dim strSQL As String

    strSQL = "CREATE INDEX PrimaryKey ON KISHU_SEISEKI ("
    strSQL = strSQL & "[KisyuCode],"
    strSQL = strSQL & "[Num]"
    strSQL = strSQL & ") WITH PRIMARY"

    mCn.Execute strSQL

    SetKeys_Indexes_KISHU_SEISEKI = True
    Exit Function
ErrorHandler:
    gApp.ErrLog
    SetKeys_Indexes_KISHU_SEISEKI = False
End Function ' SetKeys_Indexes_KISHU_SEISEKI


'
'   �@�\: DB��CHOKYO(�����t�}�X�^)�e�[�u���C���f�b�N�X�����
'
'   ���l: �Ȃ�
'
Public Function SetKeys_Indexes_CHOKYO() As Boolean
On Error GoTo ErrorHandler
    Dim strSQL As String

    strSQL = "CREATE INDEX PrimaryKey ON CHOKYO ("
    strSQL = strSQL & "[ChokyosiCode]"
    strSQL = strSQL & ") WITH PRIMARY"
    mCn.Execute strSQL
    
    strSQL = "CREATE INDEX ChokyosiNameKey ON CHOKYO ("
    strSQL = strSQL & "[ChokyosiName]"
    strSQL = strSQL & ") "
    mCn.Execute strSQL

    strSQL = "CREATE INDEX ChokyosiNameKanaKey ON CHOKYO ("
    strSQL = strSQL & "[ChokyosiNameKana]"
    strSQL = strSQL & ") "
    mCn.Execute strSQL

    strSQL = "CREATE INDEX ChokyosiRyakusyoKey ON CHOKYO ("
    strSQL = strSQL & "[ChokyosiRyakusyo]"
    strSQL = strSQL & ") "
    mCn.Execute strSQL

    strSQL = "CREATE INDEX ChokyosiNameEngKey ON CHOKYO ("
    strSQL = strSQL & "[ChokyosiNameEng]"
    strSQL = strSQL & ") "
    mCn.Execute strSQL

    SetKeys_Indexes_CHOKYO = True
    Exit Function
ErrorHandler:
    gApp.ErrLog
    SetKeys_Indexes_CHOKYO = False
End Function ' SetKeys_Indexes_CHOKYO


'
'   �@�\: DB��CHOKYO_SEISEKI(�����t�}�X�^_����)�e�[�u���C���f�b�N�X�����
'
'   ���l: �Ȃ�
'
Public Function SetKeys_Indexes_CHOKYO_SEISEKI() As Boolean
On Error GoTo ErrorHandler
    Dim strSQL As String

    strSQL = "CREATE INDEX PrimaryKey ON CHOKYO_SEISEKI ("
    strSQL = strSQL & "[ChokyosiCode],"
    strSQL = strSQL & "[Num]"
    strSQL = strSQL & ") WITH PRIMARY"

    mCn.Execute strSQL

    SetKeys_Indexes_CHOKYO_SEISEKI = True
    Exit Function
ErrorHandler:
    gApp.ErrLog
    SetKeys_Indexes_CHOKYO_SEISEKI = False
End Function ' SetKeys_Indexes_CHOKYO_SEISEKI


'
'   �@�\: DB��SEISAN(���Y�҃}�X�^)�e�[�u���C���f�b�N�X�����
'
'   ���l: �Ȃ�
'
Public Function SetKeys_Indexes_SEISAN() As Boolean
On Error GoTo ErrorHandler
    Dim strSQL As String

    strSQL = "CREATE INDEX PrimaryKey ON SEISAN ("
    strSQL = strSQL & "[BreederCode]"
    strSQL = strSQL & ") WITH PRIMARY"

    mCn.Execute strSQL

    strSQL = "CREATE INDEX BreederName_CoKey ON SEISAN ([BreederName_Co]) "
    mCn.Execute strSQL
    
    strSQL = "CREATE INDEX BreederNameKey ON SEISAN ([BreederName]) "
    mCn.Execute strSQL
    
    strSQL = "CREATE INDEX BreederNameKanaKey ON SEISAN ([BreederNameKana]) "
    mCn.Execute strSQL
    
    strSQL = "CREATE INDEX BreederNameEngKey ON SEISAN ([BreederNameEng]) "
    mCn.Execute strSQL
    
    SetKeys_Indexes_SEISAN = True
    Exit Function
ErrorHandler:
    gApp.ErrLog
    SetKeys_Indexes_SEISAN = False
End Function ' SetKeys_Indexes_SEISAN


'
'   �@�\: DB��BANUSI(�n��}�X�^)�e�[�u���C���f�b�N�X�����
'
'   ���l: �Ȃ�
'
Public Function SetKeys_Indexes_BANUSI() As Boolean
On Error GoTo ErrorHandler
    Dim strSQL As String

    strSQL = "CREATE INDEX PrimaryKey ON BANUSI ("
    strSQL = strSQL & "[BanusiCode]"
    strSQL = strSQL & ") WITH PRIMARY"
    mCn.Execute strSQL
    
    strSQL = "CREATE INDEX BanusiName_CoKey ON BANUSI ("
    strSQL = strSQL & "[BanusiName_Co]"
    strSQL = strSQL & ") "
    mCn.Execute strSQL
    
    strSQL = "CREATE INDEX BanusiNameKey ON BANUSI ("
    strSQL = strSQL & "[BanusiName]"
    strSQL = strSQL & ") "
    mCn.Execute strSQL
    
    strSQL = "CREATE INDEX BanusiNameKanaKey ON BANUSI ("
    strSQL = strSQL & "[BanusiNameKana]"
    strSQL = strSQL & ") "
    mCn.Execute strSQL
    
    strSQL = "CREATE INDEX BanusiNameEngKey ON BANUSI ("
    strSQL = strSQL & "[BanusiNameEng]"
    strSQL = strSQL & ") "
    
    mCn.Execute strSQL

    SetKeys_Indexes_BANUSI = True
    Exit Function
ErrorHandler:
    gApp.ErrLog
    SetKeys_Indexes_BANUSI = False
End Function ' SetKeys_Indexes_BANUSI


'
'   �@�\: DB��HANSYOKU(�ɐB�n�}�X�^)�e�[�u���C���f�b�N�X�����
'
'   ���l: �Ȃ�
'
Public Function SetKeys_Indexes_HANSYOKU() As Boolean
On Error GoTo ErrorHandler
    Dim strSQL As String

    strSQL = "CREATE INDEX PrimaryKey ON HANSYOKU ("
    strSQL = strSQL & "[HansyokuNum]"
    strSQL = strSQL & ") WITH PRIMARY"

    mCn.Execute strSQL
    
    strSQL = "CREATE INDEX BameiKey ON HANSYOKU ("
    strSQL = strSQL & "[Bamei]"
    strSQL = strSQL & ") "
    mCn.Execute strSQL
    
    strSQL = "CREATE INDEX BameiKanaKey ON HANSYOKU ("
    strSQL = strSQL & "[BameiKana]"
    strSQL = strSQL & ") "
    mCn.Execute strSQL
    
    strSQL = "CREATE INDEX BameiEngKey ON HANSYOKU ("
    strSQL = strSQL & "[BameiEng]"
    strSQL = strSQL & ") "
    mCn.Execute strSQL
    
    SetKeys_Indexes_HANSYOKU = True
    Exit Function
ErrorHandler:
    gApp.ErrLog
    SetKeys_Indexes_HANSYOKU = False
End Function ' SetKeys_Indexes_HANSYOKU


'
'   �@�\: DB��SANKU(�Y��}�X�^)�e�[�u���C���f�b�N�X�����
'
'   ���l: �Ȃ�
'
Public Function SetKeys_Indexes_SANKU() As Boolean
On Error GoTo ErrorHandler
    Dim strSQL As String

    strSQL = "CREATE INDEX PrimaryKey ON SANKU ("
    strSQL = strSQL & "[KettoNum]"
    strSQL = strSQL & ") WITH PRIMARY"

    mCn.Execute strSQL

    strSQL = "CREATE INDEX BreederCodeKey ON SANKU ("
    strSQL = strSQL & "[BreederCode]"
    strSQL = strSQL & ") "

    mCn.Execute strSQL
    
    strSQL = "CREATE INDEX FNumKey ON SANKU ("
    strSQL = strSQL & "[FNum]"
    strSQL = strSQL & ") "

    mCn.Execute strSQL
    
    strSQL = "CREATE INDEX MNumKey ON SANKU ("
    strSQL = strSQL & "[MNum]"
    strSQL = strSQL & ") "

    mCn.Execute strSQL

    SetKeys_Indexes_SANKU = True
    Exit Function
ErrorHandler:
    gApp.ErrLog
    SetKeys_Indexes_SANKU = False
End Function ' SetKeys_Indexes_SANKU


'
'   �@�\: DB��RECORD(���R�[�h�}�X�^)�e�[�u���C���f�b�N�X�����
'
'   ���l: �Ȃ�
'
Public Function SetKeys_Indexes_RECORD() As Boolean
On Error GoTo ErrorHandler
    Dim strSQL As String

    strSQL = "CREATE INDEX PrimaryKey ON RECORD ("
    strSQL = strSQL & "[RecInfoKubun],"
    strSQL = strSQL & "[Year],"
    strSQL = strSQL & "[MonthDay],"
    strSQL = strSQL & "[JyoCD],"
    strSQL = strSQL & "[Kaiji],"
    strSQL = strSQL & "[Nichiji],"
    strSQL = strSQL & "[RaceNum],"
    strSQL = strSQL & "[TokuNum_SyubetuCD],"
    strSQL = strSQL & "[Kyori],"
    strSQL = strSQL & "[TrackCD]"
    strSQL = strSQL & ") WITH PRIMARY"

    mCn.Execute strSQL

    SetKeys_Indexes_RECORD = True
    Exit Function
ErrorHandler:
    gApp.ErrLog
    SetKeys_Indexes_RECORD = False
End Function ' SetKeys_Indexes_RECORD


'
'   �@�\: DB��HANRO(��H����)�e�[�u���C���f�b�N�X�����
'
'   ���l: �Ȃ�
'
Public Function SetKeys_Indexes_HANRO() As Boolean
On Error GoTo ErrorHandler
    Dim strSQL As String

    strSQL = "CREATE INDEX PrimaryKey ON HANRO ("
    strSQL = strSQL & "[TresenKubun],"
    strSQL = strSQL & "[ChokyoDate],"
    strSQL = strSQL & "[ChokyoTime],"
    strSQL = strSQL & "[KettoNum]"
    strSQL = strSQL & ") WITH PRIMARY"

    mCn.Execute strSQL

    strSQL = "CREATE INDEX KettoNumKey ON HANRO ( [KettoNum] )"
    mCn.Execute strSQL
    
    SetKeys_Indexes_HANRO = True
    Exit Function
ErrorHandler:
    gApp.ErrLog
    SetKeys_Indexes_HANRO = False
End Function ' SetKeys_Indexes_HANRO


'
'   �@�\: DB��BATAIJYU(�n�̏d)�e�[�u���C���f�b�N�X�����
'
'   ���l: �Ȃ�
'
Public Function SetKeys_Indexes_BATAIJYU() As Boolean
On Error GoTo ErrorHandler
    Dim strSQL As String

    strSQL = "CREATE INDEX PrimaryKey ON BATAIJYU ("
    strSQL = strSQL & "[Year],"
    strSQL = strSQL & "[MonthDay],"
    strSQL = strSQL & "[JyoCD],"
    strSQL = strSQL & "[Kaiji],"
    strSQL = strSQL & "[Nichiji],"
    strSQL = strSQL & "[RaceNum]"
    strSQL = strSQL & ") WITH PRIMARY"

    mCn.Execute strSQL

    SetKeys_Indexes_BATAIJYU = True
    Exit Function
ErrorHandler:
    gApp.ErrLog
    SetKeys_Indexes_BATAIJYU = False
End Function ' SetKeys_Indexes_BATAIJYU


'
'   �@�\: DB��TENKO_BABA(�V��n����)�e�[�u���C���f�b�N�X�����
'
'   ���l: �Ȃ�
'
Public Function SetKeys_Indexes_TENKO_BABA() As Boolean
On Error GoTo ErrorHandler
    Dim strSQL As String

    strSQL = "CREATE INDEX PrimaryKey ON TENKO_BABA ("
    strSQL = strSQL & "[Year],"
    strSQL = strSQL & "[MonthDay],"
    strSQL = strSQL & "[JyoCD],"
    strSQL = strSQL & "[Kaiji],"
    strSQL = strSQL & "[Nichiji],"
    strSQL = strSQL & "[HappyoTime],"
    strSQL = strSQL & "[HenkoID]"
    strSQL = strSQL & ") WITH PRIMARY"

    mCn.Execute strSQL

    SetKeys_Indexes_TENKO_BABA = True
    Exit Function
ErrorHandler:
    gApp.ErrLog
    SetKeys_Indexes_TENKO_BABA = False
End Function ' SetKeys_Indexes_TENKO_BABA


'
'   �@�\: DB��TORIKESI_JYOGAI(�o������E�������O)�e�[�u���C���f�b�N�X�����
'
'   ���l: �Ȃ�
'
Public Function SetKeys_Indexes_TORIKESI_JYOGAI() As Boolean
On Error GoTo ErrorHandler
    Dim strSQL As String

    strSQL = "CREATE INDEX PrimaryKey ON TORIKESI_JYOGAI ("
    strSQL = strSQL & "[Year],"
    strSQL = strSQL & "[MonthDay],"
    strSQL = strSQL & "[JyoCD],"
    strSQL = strSQL & "[Kaiji],"
    strSQL = strSQL & "[Nichiji],"
    strSQL = strSQL & "[RaceNum],"
    strSQL = strSQL & "[Umaban]"
    strSQL = strSQL & ") WITH PRIMARY"

    mCn.Execute strSQL

    SetKeys_Indexes_TORIKESI_JYOGAI = True
    Exit Function
ErrorHandler:
    gApp.ErrLog
    SetKeys_Indexes_TORIKESI_JYOGAI = False
End Function ' SetKeys_Indexes_TORIKESI_JYOGAI


'
'   �@�\: DB��KISHU_CHANGE(�R��ύX)�e�[�u���C���f�b�N�X�����
'
'   ���l: �Ȃ�
'
Public Function SetKeys_Indexes_KISHU_CHANGE() As Boolean
On Error GoTo ErrorHandler
    Dim strSQL As String

    strSQL = "CREATE INDEX PrimaryKey ON KISHU_CHANGE ("
    strSQL = strSQL & "[Year],"
    strSQL = strSQL & "[MonthDay],"
    strSQL = strSQL & "[JyoCD],"
    strSQL = strSQL & "[Kaiji],"
    strSQL = strSQL & "[Nichiji],"
    strSQL = strSQL & "[RaceNum],"
    strSQL = strSQL & "[HappyoTime],"
    strSQL = strSQL & "[Umaban]"
    strSQL = strSQL & ") WITH PRIMARY"

    mCn.Execute strSQL

    SetKeys_Indexes_KISHU_CHANGE = True
    Exit Function
ErrorHandler:
    gApp.ErrLog
    SetKeys_Indexes_KISHU_CHANGE = False
End Function ' SetKeys_Indexes_KISHU_CHANGE


'
'   �@�\: DB��MINING(�f�[�^�}�C�j���O�\�z)�e�[�u���C���f�b�N�X�����
'
'   ���l: �Ȃ�
'
Public Function SetKeys_Indexes_MINING() As Boolean
On Error GoTo ErrorHandler
    Dim strSQL As String

    strSQL = "CREATE INDEX PrimaryKey ON MINING ("
    strSQL = strSQL & "[Year],"
    strSQL = strSQL & "[MonthDay],"
    strSQL = strSQL & "[JyoCD],"
    strSQL = strSQL & "[Kaiji],"
    strSQL = strSQL & "[Nichiji],"
    strSQL = strSQL & "[RaceNum]"
    strSQL = strSQL & ") WITH PRIMARY"

    mCn.Execute strSQL

    SetKeys_Indexes_MINING = True
    Exit Function
ErrorHandler:
    gApp.ErrLog
    SetKeys_Indexes_MINING = False
End Function ' SetKeys_Indexes_MINING


'
'   �@�\: DB��SCHEDULE(�J�ÃX�P�W���[��)�e�[�u���C���f�b�N�X�����
'
'   ���l: �Ȃ�
'
Public Function SetKeys_Indexes_SCHEDULE() As Boolean
On Error GoTo ErrorHandler
    Dim strSQL As String
    
    strSQL = "CREATE INDEX PrimaryKey ON SCHEDULE ("
    strSQL = strSQL & "[Year],"
    strSQL = strSQL & "[MonthDay],"
    strSQL = strSQL & "[JyoCD],"
    strSQL = strSQL & "[Kaiji],"
    strSQL = strSQL & "[Nichiji]"
    strSQL = strSQL & ") WITH PRIMARY"

    mCn.Execute strSQL

    SetKeys_Indexes_SCHEDULE = True
    Exit Function
ErrorHandler:
    gApp.ErrLog
    SetKeys_Indexes_SCHEDULE = False
End Function ' SetKeys_Indexes_SCHEDULE


'
'   �@�\: DB��RAKaiSel(�o�n�\�J�ÑI��)�e�[�u���C���f�b�N�X�����
'
'   ���l: �Ȃ�
'
Public Function SetKeys_Indexes_RAKaiSel() As Boolean
On Error GoTo ErrorHandler
    Dim strSQL As String

    strSQL = "CREATE INDEX PrimaryKey ON RAKaiSel ("
    strSQL = strSQL & "[Year] DESC,"
    strSQL = strSQL & "[MonthDay] DESC,"
    strSQL = strSQL & "[JyoCD]"
    strSQL = strSQL & ") WITH PRIMARY"

    mCn.Execute strSQL

    SetKeys_Indexes_RAKaiSel = True
    Exit Function
ErrorHandler:
    gApp.ErrLog
    SetKeys_Indexes_RAKaiSel = False
End Function ' SetKeys_Indexes_RAKaiSel

