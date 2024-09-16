Option Strict Off
Option Explicit On
Option Compare Binary
Friend Class clsImportYS
	' @(h) clsReadYS.cls
	' @(s)
	' JVData "YS" �f�[�^�x�[�X�o�^�N���X
	'
	
	Private mBuf As JV_YS_SCHEDULE ''�N�ԃX�P�W���[���\����
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
        strSql = "SELECT * FROM SCHEDULE"
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
    ' ������    : lBuf - JVData ���ʎq"YS" �̂P�s
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
        Dim j As Short '' ���[�v�J�E���^
        Dim k As Short '' ���[�v�J�E���^

        gCon.BeginTrans()
        System.Diagnostics.Debug.WriteLine("BeginTrans")

        mRS.AddNew()

        With mBuf
            With .head
                mRS.Fields("RecordSpec").Value = .RecordSpec '' ���R�[�h���
                mRS.Fields("DataKubun").Value = .DataKubun '' �f�[�^�敪
                With .MakeDate
                    mRS.Fields("MakeDate").Value = .Year & .Month & .Day '' �N����
                End With ' MakeDate
            End With ' head
            With .id
                mRS.Fields("Year").Value = .Year '' �J�ÔN
                mRS.Fields("MonthDay").Value = .MonthDay '' �J�Ì���
                mRS.Fields("JyoCD").Value = .JyoCD '' ���n��R�[�h
                mRS.Fields("Kaiji").Value = .Kaiji '' �J�É��N��
                mRS.Fields("Nichiji").Value = .Nichiji '' �J�Ó���N����
            End With ' id
            mRS.Fields("YoubiCD").Value = .YoubiCD '' �j���R�[�h
            For i = 0 To 2
                With .JyusyoInfo(i)
                    mRS.Fields("Jyusyo" & i + 1 & "TokuNum").Value = .TokuNum '' ���ʋ����ԍ�
                    mRS.Fields("Jyusyo" & i + 1 & "Hondai").Value = .Hondai '' �������{��
                    mRS.Fields("Jyusyo" & i + 1 & "Ryakusyo10").Value = .Ryakusyo10 '' ����������10��
                    mRS.Fields("Jyusyo" & i + 1 & "Ryakusyo6").Value = .Ryakusyo6 '' ����������6��
                    mRS.Fields("Jyusyo" & i + 1 & "Ryakusyo3").Value = .Ryakusyo3 '' ����������3��
                    mRS.Fields("Jyusyo" & i + 1 & "Nkai").Value = .Nkai '' �d�܉񎟑�N��
                    mRS.Fields("Jyusyo" & i + 1 & "GradeCD").Value = .GradeCD '' �O���[�h�R�[�h
                    mRS.Fields("Jyusyo" & i + 1 & "SyubetuCD").Value = .SyubetuCD '' ������ʃR�[�h
                    mRS.Fields("Jyusyo" & i + 1 & "KigoCD").Value = .KigoCD '' �����L���R�[�h
                    mRS.Fields("Jyusyo" & i + 1 & "JyuryoCD").Value = .JyuryoCD '' �d�ʎ�ʃR�[�h
                    mRS.Fields("Jyusyo" & i + 1 & "Kyori").Value = .Kyori '' ����
                    mRS.Fields("Jyusyo" & i + 1 & "TrackCD").Value = .TrackCD '' �g���b�N�R�[�h
                End With ' JyusyoInfo
            Next i
        End With

        With mBuf.id
            System.Diagnostics.Debug.WriteLine("Insert SCHEDULE : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji)
        End With ' id

        mRS.Update()

        gCon.CommitTrans()
        System.Diagnostics.Debug.WriteLine("CommitTrans")

        InsertDB = True

ExitHandler:
        Exit Function
ErrorHandler:
        InsertDB = False
        gCon.RollbackTrans()
        mRS.CancelUpdate()
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
        Dim i As Short '' ���[�v�J�E���^
        Dim j As Short '' ���[�v�J�E���^
        Dim k As Short '' ���[�v�J�E���^
        Dim strSql As String '' SQL��

        System.Diagnostics.Debug.WriteLine("BeginTrans")
        gCon.BeginTrans()

        strSql = "UPDATE SCHEDULE SET "
        With mBuf
            With .head
                strSql = strSql & SS & "RecordSpec" & SE & "='" & Replace(.RecordSpec, "'", "''") & "'," '' ���R�[�h���
                strSql = strSql & SS & "DataKubun" & SE & "='" & Replace(.DataKubun, "'", "''") & "'," '' �f�[�^�敪
                strSql = strSql & SS & "MakeDate" & SE & "='" & Replace(strMakeDate, "'", "''") & "'," '' �N����
            End With ' head
            With .id
                strSql = strSql & SS & "Year" & SE & "='" & Replace(.Year, "'", "''") & "'," '' �J�ÔN
                strSql = strSql & SS & "MonthDay" & SE & "='" & Replace(.MonthDay, "'", "''") & "'," '' �J�Ì���
                strSql = strSql & SS & "JyoCD" & SE & "='" & Replace(.JyoCD, "'", "''") & "'," '' ���n��R�[�h
                strSql = strSql & SS & "Kaiji" & SE & "='" & Replace(.Kaiji, "'", "''") & "'," '' �J�É��N��
                strSql = strSql & SS & "Nichiji" & SE & "='" & Replace(.Nichiji, "'", "''") & "'," '' �J�Ó���N����
            End With ' id
            strSql = strSql & SS & "YoubiCD" & SE & "='" & Replace(.YoubiCD, "'", "''") & "'," '' �j���R�[�h
            With .JyusyoInfo(0)
                strSql = strSql & SS & "Jyusyo1TokuNum" & SE & "='" & Replace(.TokuNum, "'", "''") & "'," '' ���ʋ����ԍ�
                strSql = strSql & SS & "Jyusyo1Hondai" & SE & "='" & Replace(.Hondai, "'", "''") & "'," '' �������{��
                strSql = strSql & SS & "Jyusyo1Ryakusyo10" & SE & "='" & Replace(.Ryakusyo10, "'", "''") & "'," '' ����������10��
                strSql = strSql & SS & "Jyusyo1Ryakusyo6" & SE & "='" & Replace(.Ryakusyo6, "'", "''") & "'," '' ����������6��
                strSql = strSql & SS & "Jyusyo1Ryakusyo3" & SE & "='" & Replace(.Ryakusyo3, "'", "''") & "'," '' ����������3��
                strSql = strSql & SS & "Jyusyo1Nkai" & SE & "='" & Replace(.Nkai, "'", "''") & "'," '' �d�܉񎟑�N��
                strSql = strSql & SS & "Jyusyo1GradeCD" & SE & "='" & Replace(.GradeCD, "'", "''") & "'," '' �O���[�h�R�[�h
                strSql = strSql & SS & "Jyusyo1SyubetuCD" & SE & "='" & Replace(.SyubetuCD, "'", "''") & "'," '' ������ʃR�[�h
                strSql = strSql & SS & "Jyusyo1KigoCD" & SE & "='" & Replace(.KigoCD, "'", "''") & "'," '' �����L���R�[�h
                strSql = strSql & SS & "Jyusyo1JyuryoCD" & SE & "='" & Replace(.JyuryoCD, "'", "''") & "'," '' �d�ʎ�ʃR�[�h
                strSql = strSql & SS & "Jyusyo1Kyori" & SE & "='" & Replace(.Kyori, "'", "''") & "'," '' ����
                strSql = strSql & SS & "Jyusyo1TrackCD" & SE & "='" & Replace(.TrackCD, "'", "''") & "'," '' �g���b�N�R�[�h
            End With ' JyusyoInfo(0)
            With .JyusyoInfo(1)
                strSql = strSql & SS & "Jyusyo2TokuNum" & SE & "='" & Replace(.TokuNum, "'", "''") & "'," '' ���ʋ����ԍ�
                strSql = strSql & SS & "Jyusyo2Hondai" & SE & "='" & Replace(.Hondai, "'", "''") & "'," '' �������{��
                strSql = strSql & SS & "Jyusyo2Ryakusyo10" & SE & "='" & Replace(.Ryakusyo10, "'", "''") & "'," '' ����������10��
                strSql = strSql & SS & "Jyusyo2Ryakusyo6" & SE & "='" & Replace(.Ryakusyo6, "'", "''") & "'," '' ����������6��
                strSql = strSql & SS & "Jyusyo2Ryakusyo3" & SE & "='" & Replace(.Ryakusyo3, "'", "''") & "'," '' ����������3��
                strSql = strSql & SS & "Jyusyo2Nkai" & SE & "='" & Replace(.Nkai, "'", "''") & "'," '' �d�܉񎟑�N��
                strSql = strSql & SS & "Jyusyo2GradeCD" & SE & "='" & Replace(.GradeCD, "'", "''") & "'," '' �O���[�h�R�[�h
                strSql = strSql & SS & "Jyusyo2SyubetuCD" & SE & "='" & Replace(.SyubetuCD, "'", "''") & "'," '' ������ʃR�[�h
                strSql = strSql & SS & "Jyusyo2KigoCD" & SE & "='" & Replace(.KigoCD, "'", "''") & "'," '' �����L���R�[�h
                strSql = strSql & SS & "Jyusyo2JyuryoCD" & SE & "='" & Replace(.JyuryoCD, "'", "''") & "'," '' �d�ʎ�ʃR�[�h
                strSql = strSql & SS & "Jyusyo2Kyori" & SE & "='" & Replace(.Kyori, "'", "''") & "'," '' ����
                strSql = strSql & SS & "Jyusyo2TrackCD" & SE & "='" & Replace(.TrackCD, "'", "''") & "'," '' �g���b�N�R�[�h
            End With ' JyusyoInfo(1)
            With .JyusyoInfo(2)
                strSql = strSql & SS & "Jyusyo3TokuNum" & SE & "='" & Replace(.TokuNum, "'", "''") & "'," '' ���ʋ����ԍ�
                strSql = strSql & SS & "Jyusyo3Hondai" & SE & "='" & Replace(.Hondai, "'", "''") & "'," '' �������{��
                strSql = strSql & SS & "Jyusyo3Ryakusyo10" & SE & "='" & Replace(.Ryakusyo10, "'", "''") & "'," '' ����������10��
                strSql = strSql & SS & "Jyusyo3Ryakusyo6" & SE & "='" & Replace(.Ryakusyo6, "'", "''") & "'," '' ����������6��
                strSql = strSql & SS & "Jyusyo3Ryakusyo3" & SE & "='" & Replace(.Ryakusyo3, "'", "''") & "'," '' ����������3��
                strSql = strSql & SS & "Jyusyo3Nkai" & SE & "='" & Replace(.Nkai, "'", "''") & "'," '' �d�܉񎟑�N��
                strSql = strSql & SS & "Jyusyo3GradeCD" & SE & "='" & Replace(.GradeCD, "'", "''") & "'," '' �O���[�h�R�[�h
                strSql = strSql & SS & "Jyusyo3SyubetuCD" & SE & "='" & Replace(.SyubetuCD, "'", "''") & "'," '' ������ʃR�[�h
                strSql = strSql & SS & "Jyusyo3KigoCD" & SE & "='" & Replace(.KigoCD, "'", "''") & "'," '' �����L���R�[�h
                strSql = strSql & SS & "Jyusyo3JyuryoCD" & SE & "='" & Replace(.JyuryoCD, "'", "''") & "'," '' �d�ʎ�ʃR�[�h
                strSql = strSql & SS & "Jyusyo3Kyori" & SE & "='" & Replace(.Kyori, "'", "''") & "'," '' ����
                strSql = strSql & SS & "Jyusyo3TrackCD" & SE & "='" & Replace(.TrackCD, "'", "''") & "'" '' �g���b�N�R�[�h
            End With ' JyusyoInfo(2)
            With .id
                strSql = strSql & " WHERE " & SS & "Year" & SE & "='" & Replace(.Year, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "MonthDay" & SE & "='" & Replace(.MonthDay, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "JyoCD" & SE & "='" & Replace(.JyoCD, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "Kaiji" & SE & "='" & Replace(.Kaiji, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "Nichiji" & SE & "='" & Replace(.Nichiji, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "MakeDate" & SE & " <= '" & Replace(strMakeDate, "'", "''") & "'"
            End With
        End With
        gCon.Execute(strSql)
        With mBuf.id
            System.Diagnostics.Debug.WriteLine("UPDATE SCHEDULE : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji)
        End With ' id
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