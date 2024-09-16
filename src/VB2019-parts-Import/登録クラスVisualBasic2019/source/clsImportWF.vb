Option Strict Off
Option Explicit On
Option Compare Binary
Friend Class clsImportWF
	' @(h) clsReadWF.cls
	' @(s)
	' JVData "WF" �f�[�^�x�[�X�o�^�N���X
	'
	
    Private mBuf As JV_WF_INFO     '' �d����(WIN5)�\����

	Private mRS1 As ADODB.Recordset
	Private mRS2 As ADODB.Recordset
	
	
	' @(f)
	'
	' �@�\      : ��������
	'
	' ������    :
	'
	' �Ԃ�l    :
	'
	' �@�\����  :
	'

    Private Sub Class_Initialize_Renamed()
        On Error GoTo ErrorHandler

        Dim strSql As String ''SQL��

        '���R�[�h�Z�b�g�I�[�v��
        strSql = "SELECT * FROM JYUSYOSIKI_HEAD"
        mRS1 = New ADODB.Recordset()
        mRS1.Open(strSql, gCon, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockPessimistic)

        strSql = "SELECT * FROM JYUSYOSIKI"
        mRS2 = New ADODB.Recordset()
        mRS2.Open(strSql, gCon, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockPessimistic)

ExitHandler:
        Exit Sub
ErrorHandler:
        System.Diagnostics.Debug.WriteLine(Err.Description)
        Resume ExitHandler
    End Sub
    Public Sub New()
        MyBase.New()
        Class_Initialize_Renamed()
    End Sub
    ' @(f)
    '
    ' �@�\      : Close�̃R�[�f�B���O
    '
    ' �@�\����  : �K�[�x�b�W�R���N�V������Close���Ă΂���Ƃǂ��ŌĂ΂�邩������Ȃ��ׁA
    '           �@�����I�ɌĂяo���K�v������B
    '
    Public Sub Close()
        '���R�[�h�Z�b�g�N���[�Y
        mRS1.Close()

        mRS1 = Nothing
        mRS2.Close()

        mRS2 = Nothing

    End Sub

    ' @(f)
    '
    ' �@�\      : �I������
    '
    ' ������    :
    '
    ' �Ԃ�l    :
    '
    ' �@�\����  : ���R�[�h�Z�b�g�N���[�Y
    '

    Private Sub Class_Terminate_Renamed()
        On Error GoTo ErrorHandler

ExitHandler:
        Exit Sub
ErrorHandler:
        System.Diagnostics.Debug.WriteLine(Err.Description)
        Resume ExitHandler
    End Sub
    Protected Overrides Sub Finalize()
        Class_Terminate_Renamed()
        MyBase.Finalize()
    End Sub


    ' @(f)
    '
    ' �@�\      : Add�v���V�[�W�����Ă�
    '
    ' ������    : lBuf - JVData ���ʎq"WF" �̂P�s
    '
    ' �Ԃ�l    : True - ����, False - ���s
    '
    ' �@�\����  : clsIRead�C���^�[�t�F�C�XAdd�v���V�[�W���̎���
    '
    Public Function Add(ByRef strBuf As String, ByVal lngBufSize As Integer) As Boolean
        On Error GoTo ErrorHandler

        Dim strMakeDate As String '' �o�^����f�[�^�̍쐬�N����

        '�\���̂Ƀf�[�^�Z�b�g
        mBuf.SetData(strBuf)


        With mBuf.head.MakeDate

            strMakeDate = .Year & .Month & .Day
        End With

        'INSERT����
        If Not InsertDB() Then
            'UPDATE�����iINSERT�����s�����ꍇ�j
            If Not UpdateDB(strMakeDate) Then System.Diagnostics.Debug.WriteLine("�X�V�Ɏ��s���܂����B" & Left(strBuf, 2))
        End If

        Add = True

ExitHandler:
        Exit Function
ErrorHandler:
        Add = False
        System.Diagnostics.Debug.WriteLine(Err.Description)
        Resume ExitHandler
    End Function

    ' @(f)
    '
    ' �@�\      : �f�[�^�x�[�X�ɒǉ�����
    '
    ' ������    :
    '
    ' �Ԃ�l    : True - ����, False - ���s
    '
    ' �@�\����  :
    '
    Public Function InsertDB() As Boolean
        On Error GoTo ErrorHandler
        Dim i As Short '' ���[�v�J�E���^

        gCon.BeginTrans()
        System.Diagnostics.Debug.WriteLine("BeginTrans")

        mRS1.AddNew()

        With mBuf
            With .head
                mRS1.Fields("RecordSpec").Value = .RecordSpec             '' ���R�[�h���
                mRS1.Fields("DataKubun").Value = .DataKubun               '' �f�[�^�敪
                With .MakeDate
                    mRS1.Fields("MakeDate").Value = .Year & .Month & .Day '' �N����
                End With ' MakeDate
            End With ' head
            With .KaisaiDate
                mRS1.Fields("Year").Value = .Year                         '' �J�ÔN
                mRS1.Fields("MonthDay").Value = .Month & .Day             '' �J�Ì���
            End With ' KaisaiDate
            mRS1.Fields("reserved1").Value = .reserved1                   '' �\��1
            For i = 0 To 4
                With .WFRaceInfo(i)
                    mRS1.Fields("JyoCD" & i + 1).Value = .JyoCD           '' ���n��R�[�h
                    mRS1.Fields("Kaiji" & i + 1).Value = .Kaiji           '' �J�É�
                    mRS1.Fields("Nichiji" & i + 1).Value = .Nichiji       '' �J�Ó���
                    mRS1.Fields("RaceNum" & i + 1).Value = .RaceNum       '' ���[�X�ԍ�
                End With ' WFRaceInfo()
            Next i
            mRS1.Fields("reserved2").Value = .reserved2                   '' �\��2
            For i = 0 To 4
                With .WFYukoHyoInfo(i)
                    mRS1.Fields("YukoHyosu" & i + 1).Value = .Yuko_Hyo    '' �L���[��
                End With ' WFYukoHyoInfo()
            Next
            mRS1.Fields("HenkanFlag").Value = .HenkanFlag                 '' �Ԋ҃t���O
            mRS1.Fields("FuseirituFlag").Value = .FuseiritsuFlag          '' �s�����t���O
            mRS1.Fields("TekichunashiFlag").Value = .TekichunashiFlag     '' �I�����t���O
            mRS1.Fields("CarryoverSyoki").Value = .COShoki                '' �L�����[�I�[�o�[���z����
            mRS1.Fields("CarryoverZandaka").Value = .COZanDaka            '' �L�����[�I�[�o�[���z�c��

            With .KaisaiDate
                System.Diagnostics.Debug.WriteLine("Insert JYUSYOSIKI_HEAD : " & .Year & .Month & .Day)
            End With ' KaisaiDate
        End With

        mRS1.Update()

        For i = 0 To 242
            If mBuf.WFPayInfo(i).Kumiban <> "          " Then
                mRS2.AddNew()
                With mBuf
                    With .head.MakeDate
                        mRS2.Fields("MakeDate").Value = .Year & .Month & .Day '' �N����
                    End With ' MakeDate
                    With .KaisaiDate
                        mRS2.Fields("Year").Value = .Year                     '' �J�ÔN
                        mRS2.Fields("MonthDay").Value = .Month & .Day         '' �J�Ì���
                    End With ' KaisaiDate
                    With .WFPayInfo(i)
                        mRS2.Fields("Kumi").Value = .Kumiban                  '' �g��
                        mRS2.Fields("PayJyushosiki").Value = .Pay             '' �d�������ߋ�
                        mRS2.Fields("TekichuHyo").Value = .Tekichu_Hyo        '' �I���[��
                    End With ' WFPayInfo()

                    With .KaisaiDate
                        System.Diagnostics.Debug.WriteLine("Insert JYUSYOSIKI : " & .Year & .Month & .Day & mBuf.WFPayInfo(i).Kumiban)
                    End With ' KaisaiDate

                    mRS2.Update()
                End With
            End If
        Next i

        gCon.CommitTrans()
        System.Diagnostics.Debug.WriteLine("CommitTrans")

        InsertDB = True

ExitHandler:
        Exit Function
ErrorHandler:
        InsertDB = False
        mRS1.CancelUpdate()
        mRS2.CancelUpdate()
        gCon.RollbackTrans()
        System.Diagnostics.Debug.WriteLine(Err.Description)
        System.Diagnostics.Debug.WriteLine("RollbackTrans")
        Resume ExitHandler
    End Function


    ' @(f)
    '
    ' �@�\      : �f�[�^�x�[�X���X�V����
    '
    ' ������    :
    '
    ' �Ԃ�l    : True - ����, False - ���s
    '
    ' �@�\����  :
    '
    Public Function UpdateDB(ByRef strMakeDate As String) As Boolean
        On Error GoTo ErrorHandler
        Dim strSql As String '' SQL��
        Dim i As Short '' ���[�v�J�E���^

        gCon.BeginTrans()
        System.Diagnostics.Debug.WriteLine("BeginTrans")

        strSql = "UPDATE JYUSYOSIKI_HEAD SET "
        With mBuf
            With .head
                strSql = strSql & SS & "RecordSpec" & SE & "='" & Replace(.RecordSpec, "'", "''") & "',"   '' ���R�[�h���
                strSql = strSql & SS & "DataKubun" & SE & "='" & Replace(.DataKubun, "'", "''") & "',"     '' �f�[�^�敪
                strSql = strSql & SS & "MakeDate" & SE & "='" & Replace(strMakeDate, "'", "''") & "',"     '' �N����
            End With ' head
            With .KaisaiDate
                strSql = strSql & SS & "Year" & SE & "='" & Replace(.Year, "'", "''") & "',"               '' �J�ÔN
                strSql = strSql & SS & "MonthDay" & SE & "='" & Replace(.Month & .Day, "'", "''") & "',"   '' �J�Ì���
            End With ' KaisaiDate
            strSql = strSql & SS & "reserved1" & SE & "='" & Replace(.reserved1, "'", "''") & "',"         '' �\��1
            For i = 0 To 4
                With .WFRaceInfo(i)
                    strSql = strSql & SS & "JyoCD" & i + 1 & SE & "='" & Replace(.JyoCD, "'", "''") & "',"      '' ���n��R�[�h
                    strSql = strSql & SS & "Kaiji" & i + 1 & SE & "='" & Replace(.Kaiji, "'", "''") & "',"      '' �J�É�
                    strSql = strSql & SS & "Nichiji" & i + 1 & SE & "='" & Replace(.Nichiji, "'", "''") & "',"  '' �J�Ó���
                    strSql = strSql & SS & "RaceNum" & i + 1 & SE & "='" & Replace(.RaceNum, "'", "''") & "',"  '' ���[�X�ԍ�
                End With ' WFRaceInfo()
            Next i
            strSql = strSql & SS & "reserved2" & SE & "='" & Replace(.reserved2, "'", "''") & "',"  '' �\��2
            For i = 0 To 4
                With .WFYukoHyoInfo(i)
                    strSql = strSql & SS & "YukoHyosu" & i + 1 & SE & "='" & Replace(.Yuko_Hyo, "'", "''") & "',"  '' �L���[��
                End With ' WFYukoHyoInfo()
            Next i
            strSql = strSql & SS & "HenkanFlag" & SE & "='" & Replace(.HenkanFlag, "'", "''") & "',"              '' �Ԋ҃t���O
            strSql = strSql & SS & "FuseirituFlag" & SE & "='" & Replace(.FuseiritsuFlag, "'", "''") & "',"       '' �s�����t���O
            strSql = strSql & SS & "TekichunashiFlag" & SE & "='" & Replace(.TekichunashiFlag, "'", "''") & "',"  '' �I�����t���O
            strSql = strSql & SS & "CarryoverSyoki" & SE & "='" & Replace(.COShoki, "'", "''") & "',"             '' �L�����[�I�[�o�[���z����
            strSql = strSql & SS & "CarryoverZandaka" & SE & "='" & Replace(.COZanDaka, "'", "''") & "',"         '' �L�����[�I�[�o�[���z�c��

            strSql = Left(strSql, Len(strSql) - 1) ' Delete Last Comma

            With .KaisaiDate
                strSql = strSql & " WHERE " & SS & "Year" & SE & "='" & Replace(.Year, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "MonthDay" & SE & "='" & Replace(.Month & .Day, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "MakeDate" & SE & " <= '" & Replace(strMakeDate, "'", "''") & "'"
                System.Diagnostics.Debug.WriteLine("UPDATE JYUSYOSIKI_HEAD : " & .Year & .Month & .Day)
            End With ' KaisaiDate
        End With

        gCon.Execute(strSql)

        For i = 0 To 242
            If mBuf.WFPayInfo(i).Kumiban <> "          " Then
                strSql = "UPDATE JYUSYOSIKI SET "
                With mBuf
                    strSql = strSql & SS & "MakeDate" & SE & "='" & Replace(strMakeDate, "'", "''") & "',"          '' �N����
                    With .KaisaiDate
                        strSql = strSql & SS & "Year" & SE & "='" & Replace(.Year, "'", "''") & "',"                '' �J�ÔN
                        strSql = strSql & SS & "MonthDay" & SE & "='" & Replace(.Month & .Day, "'", "''") & "',"    '' �J�Ì���
                    End With ' KaisaiDate
                    With .WFPayInfo(i)
                        strSql = strSql & SS & "Kumi" & SE & "='" & Replace(.Kumiban, "'", "''") & "',"             '' �g��
                        strSql = strSql & SS & "PayJyushosiki" & SE & "='" & Replace(.Pay, "'", "''") & "',"        '' �d�������ߋ�
                        strSql = strSql & SS & "TekichuHyo" & SE & "='" & Replace(.Tekichu_Hyo, "'", "''") & "',"   '' �I���[��
                    End With ' WFPayInfo

                    strSql = Left(strSql, Len(strSql) - 1) ' Delete Last Comma

                    With .KaisaiDate
                        strSql = strSql & " WHERE " & SS & "Year" & SE & "='" & Replace(.Year, "'", "''") & "'"
                        strSql = strSql & " AND " & SS & "MonthDay" & SE & "='" & Replace(.Month & .Day, "'", "''") & "'"
                        strSql = strSql & " AND " & SS & "Kumi" & SE & "='" & Replace(mBuf.WFPayInfo(i).Kumiban, "'", "''") & "'"
                        strSql = strSql & " AND " & SS & "MakeDate" & SE & " <= '" & Replace(strMakeDate, "'", "''") & "'"
                        System.Diagnostics.Debug.WriteLine("UPDATE JYUSYOSIKI : " & .Year & .Month & .Day & mBuf.WFPayInfo(i).Kumiban)
                    End With ' KaisaiDate
                End With
                gCon.Execute(strSql)
            End If
        Next i

        gCon.CommitTrans()
        System.Diagnostics.Debug.WriteLine("CommitTrans")

        UpdateDB = True

ExitHandler:
        Exit Function
ErrorHandler:
        System.Diagnostics.Debug.WriteLine(Err.Description)
        System.Diagnostics.Debug.WriteLine("RollbackTrans")
        gCon.RollbackTrans()
        UpdateDB = False
        Resume ExitHandler
    End Function
End Class