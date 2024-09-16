VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDBRAKaiSel 
   BorderStyle     =   3  '�Œ��޲�۸�
   Caption         =   "�J�Èꗗ�̍쐬��"
   ClientHeight    =   375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   375
   ScaleWidth      =   4680
   StartUpPosition =   1  '��Ű ̫�т̒���
   Begin VB.Timer tmrTrigger 
      Enabled         =   0   'False
      Left            =   3780
      Top             =   -30
   End
   Begin MSComctlLib.ProgressBar prgBar 
      Height          =   315
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   4635
      _ExtentX        =   8176
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   1
   End
End
Attribute VB_Name = "frmDBRAKaiSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'   �J�Ï��쐬 �_�C�A���O
'
Option Explicit

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �����ϐ�
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private mRTDates As Dictionary

Private mTargetYear As String

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �v���p�e�B
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   �@�\: �J�ÔN��ݒ肷��
'
'   ���l: �Ȃ�
'
Public Property Let TargetYear(RHS As String)
    mTargetYear = RHS
End Property


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �����֐�
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   �@�\: �J�Ï��쐬����
'
'   ���l: �Ȃ�
'
Private Function MakeRAKaiSel() As Boolean
On Error GoTo ErrorHandler
    gApp.Log "MakeRAKaiSel"
    
    Dim CN    As ADODB.Connection
    Dim rsRA  As ADODB.Recordset    '' RACE ���R�[�h�Z�b�g
    Dim rsSC  As ADODB.Recordset    '' SCHEDULE ���R�[�h�Z�b�g
    Dim rsOut As ADODB.Recordset    '' �o�͐�
    Dim rs    As ADODB.Recordset    '' �J�����g���R�[�h�Z�b�g
    
    Dim cc As clsCodeConverter
    
    Dim gd As clsGridData
    Dim lngCP As Long               '' �J�����|�C���^
    Dim lngRP As Long               '' ���E�|�C���^
    Dim blnWriteFlag As Boolean     '' �㏑���t���O
    Dim blnNewRecordFlag As Boolean '' �V���R�[�h�t���O
    Dim p As Long                   '' �������݃G���A
    Dim i As Long
    Dim strPrevJyokenCD As String '' ��r�p�̑O������R�[�h
    
    Dim strRTDate As String
    
    Set CN = New ADODB.Connection
    
    Set rsRA = New ADODB.Recordset
    Set rsSC = New ADODB.Recordset
    Set rsOut = New ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    Set cc = New clsCodeConverter
    Set mRTDates = New Dictionary
    
    CN.Open "Provider=Microsoft.Jet.OLEDB.4.0;" _
        & "Data Source=" & gApp.R_DBPath & "\" & gstrMDBName(47)
    On Error Resume Next
    CN.Execute "DELETE * FROM RAKaiSel WHERE [Year]='" & mTargetYear & "'", , adExecuteNoRecords
    If Err.Number <> 0 Then
        gApp.ErrLog
        gApp.Log "RAKaiSel�p��ǂ݃e�[�u���̍폜�G���["
        MakeRAKaiSel = False
        Exit Function
    End If
    On Error GoTo ErrorHandler

    rsRA.Open "SELECT * FROM RACE WHERE [Year] = '" & mTargetYear & "' ORDER BY MonthDay, JyoCD, Kaiji, Nichiji, RaceNum", _
                gApp.GetCN_RACE, adOpenForwardOnly, adLockReadOnly, adCmdText
    rsSC.Open "SELECT * FROM SCHEDULE WHERE [Year] = '" & mTargetYear & "' ORDER BY MonthDay, JyoCD, Kaiji, Nichiji", _
                gApp.GetCN_SCHEDULE, adOpenForwardOnly, adLockReadOnly, adCmdText
    rsOut.Open "RAKaiSel", CN, adOpenKeyset, adLockOptimistic, adCmdTable
    
    lngRP = 1
    blnWriteFlag = True
    
    Set rs = ActiveRecordset(rsRA, rsSC)
    Do While Not (rsRA.EOF And rsSC.EOF)
    
        DoEvents
        
        If rs("JyoCD") > "10" Then  ' �������n�ȊO�͏������܂Ȃ�
            blnNewRecordFlag = False
            blnWriteFlag = False
        ElseIf rsOut.BOF Then       ' �ŏ��͏�������
            blnNewRecordFlag = True
            blnWriteFlag = True
        Else                        ' ���ڈȍ~�́A��r���ď������ނ����߂�
            ' ���t�����������ׂ�
            If rsOut("Year") = rs("Year") And rsOut("MonthDay") = rs("MonthDay") Then
                ' ���t�������ꍇ�́A
                blnNewRecordFlag = False
                ' �ꏊ�����������ׂ�
                If rs("JyoCD") = rsOut("JyoCD" & p) Then
                    ' �ꏊ�������ꍇ��
                    ' �O���[�h����荂����
                    ' ����v���[�X�����ׂ�
                    If CompareMajorRace(rs, rsOut("GradeCD" & p).value & strPrevJyokenCD) Then
                        ' �O���[�h�������ꍇ��
                        blnWriteFlag = True  ' ����͏�������
                    Else
                        ' �O���[�h���Ⴂ�ꍇ��
                        blnWriteFlag = False ' ����͏������܂Ȃ�
                    End If
                Else
                    ' �ꏊ���قȂ�ꍇ��
                    p = p + 1 ' �E�ɂ��炷
                    If p >= 3 Then ' �R�ӏ��ȏ���f�[�^���������ꍇ
                        gApp.Log "�J�Â��R�ӏ��ȏ㓯���ɑ��݂��܂�" & ":" & rs("Year") & rs("MonthDay") & rs("JyoCD")
                        p = 2 ' �R�Ԗڂ̏ꏊ�ŏ�������
                    End If
                    blnWriteFlag = True ' ����͏�������
                End If
            Else
                ' ���t���قȂ�ꍇ��
                p = 0            ' �P�Ԗڂ̏ꏊ�ɃZ�b�g
                blnNewRecordFlag = True
                blnWriteFlag = True ' ����͏�������
            End If

        End If
        
        If blnNewRecordFlag Then  ' �V�����s�ɏ������ޏꍇ
            Call rsOut.AddNew
            prgBar.value = 100 * (Left$(rs("MonthDay"), 2)) / 12
        End If
    
        If blnWriteFlag Then
            With gd
                ' ���t�J����
                rsOut("Year") = rs("Year")
                rsOut("MonthDay") = rs("MonthDay")
                rsOut("JyoCD") = "00"
                rsOut("YoubiCD") = rs("YoubiCD")

                ' ���R�[�h�����[�X���Ȃ烊���N�A�J�ÃX�P�W���[���Ȃ烊���N���Ȃ�
                If rs("RecordSpec") = "RA" Then
                    rsOut("CanLink") = "1" ' �����N����
                Else
                    If IsNull(rsOut("CanLink")) Then  ' �����N�ݒ�������Ȃ�
                        rsOut("CanLink") = "0" ' �����N���Ȃ�
                    End If
                End If

                ' �ꏊ�J���� (p=�������݃G���A)
                rsOut("JyoCD" & p) = rs("JyoCD")
                
                ' �d�ܖ��J���� (p=�������݃G���A)
                If GetRyakusyo6(rs) <> "" Then
                    rsOut("Dai" & p) = GetRyakusyo6(rs)
                    rsOut("DaiToolTip" & p) = GetHondai(rs)
                    rsOut("GradeCD" & p) = GetGradeCD(rs)
                    strPrevJyokenCD = GetJokenCD(rs)
                End If
                    
            End With
        End If
                
        ' ����擾�p�̕K�v���t�R���N�V�����̍쐬
        strRTDate = rs("Year") & rs("MonthDay")
        If rs("RecordSpec") = "RA" And rs("DataKubun") <= "6" And Not mRTDates.Exists(strRTDate) Then
            mRTDates.Add strRTDate, strRTDate
        End If
        
        rs.MoveNext
        Set rs = ActiveRecordset(rsRA, rsSC)
        
    Loop
    
    If Not rsOut.EOF Then
        rsOut.Update
    End If
    
    rsRA.Close
    rsSC.Close
    MakeRAKaiSel = True
    
    ' ����擾�p�̕K�v���t�R���N�V�������L�^����
    If Join(mRTDates.Keys, ",") <> "" Then
        gApp.R_RTDates = Join(mRTDates.Keys, ",")
    End If
    
    Exit Function
ErrorHandler:
    gApp.ErrLog
    MakeRAKaiSel = False
End Function

'
'   �@�\: ��v���[�X�̔���
'
'   ���l: ��ȃ��[�X�ɂ��ӂ��킵�����̔�r
'         ���ӂ̕����ӂ��킵�����True
'
'       �@ �O���[�h�R�[�h������������
'       �A �O���[�h�R�[�h�������ꍇ�͏����R�[�h����������
'       �B �O���[�h�R�[�h������R�[�h�������ꍇ�̓��[�X�ԍ��̏���������
'       �A����O���[�h�R�[�h�̏��Ԃ͈ȉ��̂Ƃ���
'       �D��x
'       1    A G1(���n����)
'       3�@�@F J�EG1�i��Q�����j
'       3    B G2(���n����)
'       4�@�@G J�EG2�i��Q�����j
'       5    C G3(���n����)
'       6�@�@H J�EG3�i��Q�����j
'       7    D �O���[�h�̂Ȃ��d��
'       8    E �d�܈ȊO�̓��ʋ���
'       9      ���̑�
'
'       ���������R�[�h
'       9 701 �V�n
'       8 702 ���o��
'       7 703 ������
'       6 001 �P�O�O���~�ȉ�
'       5 002 �Q�O�O���~�ȉ�
'       4 003 �R�O�O���~�ȉ�
'        .          .
'        .          .
'        .          .
'       3 099 �X�X�O�O���~�ȉ�
'       2 100 �P���~�ȉ�
'       1 999 �I�[�v��
'
Private Function CompareMajorRace(LHSrs As ADODB.Recordset, RHS As String) As Boolean
On Error GoTo ErrorHandler
    Dim lngLHSLevel As Long
    Dim lngRHSLevel As Long
    
    lngLHSLevel = LevelOfGrace(GetGradeCD(LHSrs))
    lngRHSLevel = LevelOfGrace(Left(RHS, 1))
    
    ' �O���[�h�R�[�h��r
    If lngLHSLevel < lngRHSLevel Then
        CompareMajorRace = True
        Exit Function
    ElseIf lngLHSLevel > lngRHSLevel Then
        CompareMajorRace = False
        Exit Function
    End If
    
    lngLHSLevel = LevelOfJyokenCD(GetJokenCD(LHSrs))
    lngRHSLevel = LevelOfJyokenCD(Right(RHS, 3))
    
    '  ����������r
    If lngLHSLevel < lngRHSLevel Then
        CompareMajorRace = True
        Exit Function
    ElseIf lngLHSLevel > lngRHSLevel Then
        CompareMajorRace = False
        Exit Function
    End If
    
    ' ���[�X�ԍ���r�́A�Ⴂ���ɓ����Ă���͂��Ȃ̂�LHS�����邱�Ƃ͖���
    
    CompareMajorRace = False
    Exit Function
ErrorHandler:
    gApp.ErrLog
    Debug.Assert False
End Function


'
'   �@�\: �����R�[�h�̗D�揇��
'
'   ���l: �Ȃ�
'
Private Function LevelOfJyokenCD(JyokenCD As String) As Long
    Select Case JyokenCD
    Case "999"
        LevelOfJyokenCD = 0
    Case "001" To "100"
        LevelOfJyokenCD = 100 - CLng(JyokenCD)
    Case "703"
        LevelOfJyokenCD = 101
    Case "702"
        LevelOfJyokenCD = 102
    Case "701"
        LevelOfJyokenCD = 103
    Case "---"
        LevelOfJyokenCD = 104
    Case Else
        LevelOfJyokenCD = 105
    End Select
End Function


'
'   �@�\: �O���[�h�R�[�h�̗D�揇��
'
'   ���l: �Ȃ�
'
Private Function LevelOfGrace(GradeCD As String) As Long
    Select Case GradeCD
    Case "A"
        LevelOfGrace = 1
    Case "F"
        LevelOfGrace = 2
    Case "B"
        LevelOfGrace = 3
    Case "G"
        LevelOfGrace = 4
    Case "C"
        LevelOfGrace = 5
    Case "H"
        LevelOfGrace = 6
    Case "D"
        LevelOfGrace = 7
    Case "E"
        LevelOfGrace = 8
    Case Else
        LevelOfGrace = 9
    End Select
End Function


'
'   �@�\: �L���ȃ��R�[�h�Z�b�g��Ԃ�
'
'   ���l: �Ȃ�
'
Private Function ActiveRecordset(RA As ADODB.Recordset, sc As ADODB.Recordset) As ADODB.Recordset
On Error GoTo ErrorHandler
    Dim strRADate As String
    Dim strSCDate As String
    
    If Not RA.EOF Then
        strRADate = RA("Year") & RA("MonthDay")
    End If
    If Not sc.EOF Then
        strSCDate = sc("Year") & sc("MonthDay")
    End If
    
    If RA.EOF Then
        Set ActiveRecordset = sc
    ElseIf sc.EOF Then
        Set ActiveRecordset = RA
    ElseIf strRADate < strSCDate Then
        Set ActiveRecordset = RA
    ElseIf strRADate > strSCDate Then
        Set ActiveRecordset = sc
    ElseIf RA("JyoCD") > sc("JyoCD") Then
        Set ActiveRecordset = sc
    ElseIf RA("JyoCD") < sc("JyoCD") Then
        Set ActiveRecordset = RA
    Else
        sc.MoveNext
        Set ActiveRecordset = RA
    End If
    Exit Function
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Function


'
'   �@�\: �����R�[�h���擾
'
'   ���l: �Ȃ�
'
Private Function GetJokenCD(rs As ADODB.Recordset) As String
    If rs("RecordSpec") = "RA" Then
        GetJokenCD = rs("JyokenCD5")
    ElseIf rs("RecordSpec") = "YS" Then
        GetJokenCD = "---"
    End If
End Function


'
'   �@�\: �O���[�h�R�[�h���擾
'
'   ���l: �Ȃ�
'
Private Function GetGradeCD(rs As ADODB.Recordset) As String
    If rs("RecordSpec") = "RA" Then
        GetGradeCD = rs("GradeCD")
    ElseIf rs("RecordSpec") = "YS" Then
        GetGradeCD = rs("Jyusyo1GradeCD")
    End If
End Function


'
'   �@�\: ����6���擾
'
'   ���l: �Ȃ�
'
Private Function GetRyakusyo6(rs As ADODB.Recordset) As String
    If rs("RecordSpec") = "RA" Then
        GetRyakusyo6 = rs("Ryakusyo6")
    ElseIf rs("RecordSpec") = "YS" Then
        GetRyakusyo6 = rs("Jyusyo1Ryakusyo6")
    End If
End Function


'
'   �@�\: �{����擾
'
'   ���l: �Ȃ�
'
Private Function GetHondai(rs As ADODB.Recordset) As String
    If rs("RecordSpec") = "RA" Then
        GetHondai = rs("Hondai")
    ElseIf rs("RecordSpec") = "YS" Then
        GetHondai = rs("Jyusyo1Hondai")
    End If
End Function


'
'   �@�\: �t�H�[�����[�h�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub Form_Load()
On Error GoTo ErrorHandler
    prgBar.max = 100
    prgBar.Min = 0
    tmrTrigger.Interval = 100
    tmrTrigger.Enabled = True
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �g���K�[�^�C�}�[�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub tmrTrigger_Timer()
On Error GoTo ErrorHandler
    tmrTrigger.Enabled = False
    gApp.R_RAKaiSelCacheExist(mTargetYear) = MakeRAKaiSel
    Unload Me
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub
