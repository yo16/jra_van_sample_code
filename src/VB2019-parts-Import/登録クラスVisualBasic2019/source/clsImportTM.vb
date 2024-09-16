Option Strict Off
Option Explicit On
Option Compare Binary
Friend Class clsImportTM
	' @(h) clsReadTM.cls
	' @(s)
	' JVData "TM" �f�[�^�x�[�X�o�^�N���X
	'
	
	Private mBuf As JV_TM_INFO '' �ΐ�^�f�[�^�}�C�j���O�\�z�\����
	Private mRS As ADODB.Recordset
	
	
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
        strSql = "SELECT * FROM TAISENGATA_MINING"
        mRS = New ADODB.Recordset()
        mRS.Open(strSql, gCon, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockPessimistic)

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
        mRS.Close()

        mRS = Nothing

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
    ' ������    : lBuf - JVData ���ʎq"TM" �̂P�s
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

        mRS.AddNew()

        With mBuf
            With .head
                mRS.Fields("RecordSpec").Value = .RecordSpec             '' ���R�[�h���
                mRS.Fields("DataKubun").Value = .DataKubun               '' �f�[�^�敪
                With .MakeDate
                    mRS.Fields("MakeDate").Value = .Year & .Month & .Day '' �N����
                End With ' MakeDate
            End With ' head
            With .id
                mRS.Fields("Year").Value = .Year                          '' �J�ÔN
                mRS.Fields("MonthDay").Value = .MonthDay                  '' �J�Ì���
                mRS.Fields("JyoCD").Value = .JyoCD                        '' ��R�[�h
                mRS.Fields("Kaiji").Value = .Kaiji                        '' ��
                mRS.Fields("Nichiji").Value = .Nichiji                    '' ��
                mRS.Fields("RaceNum").Value = .RaceNum                    '' ���[�X�ԍ�
            End With ' id
            With .MakeHM
                mRS.Fields("MakeHM").Value = .Hour & .Minute              '' �f�[�^�쐬����
            End With
            For i = 0 To 17
                With .TMInfo(i)
                    mRS.Fields("Umaban" & i + 1).Value = .Umaban          '' �n��
                    mRS.Fields("TMScore" & i + 1).Value = .TMScore        '' �\���X�R�A
                End With
            Next i
        End With

        With mBuf.id
            System.Diagnostics.Debug.WriteLine("Insert TAISENGATA_MINING : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum)
        End With ' id

        mRS.Update()

        gCon.CommitTrans()
        System.Diagnostics.Debug.WriteLine("CommitTrans")

        InsertDB = True

ExitHandler:
        Exit Function
ErrorHandler:
        InsertDB = False
        mRS.CancelUpdate()
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

        strSql = "UPDATE TAISENGATA_MINING SET "
        With mBuf
            With .head
                strSql = strSql & SS & "RecordSpec" & SE & "='" & Replace(.RecordSpec, "'", "''") & "',"  '' ���R�[�h���
                strSql = strSql & SS & "DataKubun" & SE & "='" & Replace(.DataKubun, "'", "''") & "',"    '' �f�[�^�敪
                strSql = strSql & SS & "MakeDate" & SE & "='" & Replace(strMakeDate, "'", "''") & "',"    '' �N����
            End With ' head
            With .id
                strSql = strSql & SS & "Year" & SE & "='" & Replace(.Year, "'", "''") & "',"              '' �J�ÔN
                strSql = strSql & SS & "MonthDay" & SE & "='" & Replace(.MonthDay, "'", "''") & "',"      '' �J�Ì���
                strSql = strSql & SS & "JyoCD" & SE & "='" & Replace(.JyoCD, "'", "''") & "',"            '' ��R�[�h
                strSql = strSql & SS & "Kaiji" & SE & "='" & Replace(.Kaiji, "'", "''") & "',"            '' ��
                strSql = strSql & SS & "Nichiji" & SE & "='" & Replace(.Nichiji, "'", "''") & "',"        '' ��
                strSql = strSql & SS & "RaceNum" & SE & "='" & Replace(.RaceNum, "'", "''") & "',"        '' ���[�X�ԍ�
            End With ' id
            With .MakeHM
                strSql = strSql & SS & "MakeHM" & SE & "='" & Replace(.Hour & .Minute, "'", "''") & "',"  '' �f�[�^�쐬����
            End With ' MakeHM
            For i = 0 To 17
                With .TMInfo(i)
                    strSql = strSql & SS & "Umaban" & i + 1 & SE & "='" & Replace(.Umaban, " '", "''") & "',"      '' �n��
                    strSql = strSql & SS & "TMScore" & i + 1 & SE & "='" & Replace(.TMScore, " '", "''") & "',"    '' �\���X�R�A
                End With
            Next i

            strSql = Left(strSql, Len(strSql) - 1) ' Delete Last Comma

            With .id
                strSql = strSql & " WHERE " & SS & "Year" & SE & "='" & Replace(.Year, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "MonthDay" & SE & "='" & Replace(.MonthDay, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "JyoCD" & SE & "='" & Replace(.JyoCD, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "Kaiji" & SE & "='" & Replace(.Kaiji, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "Nichiji" & SE & "='" & Replace(.Nichiji, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "RaceNum" & SE & "='" & Replace(.RaceNum, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "MakeDate" & SE & " <= '" & Replace(strMakeDate, "'", "''") & "'"
            End With
        End With

        With mBuf.id
            System.Diagnostics.Debug.WriteLine("UPDATE TAISENGATA_MINING : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum)
        End With ' id

        gCon.Execute(strSql)

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