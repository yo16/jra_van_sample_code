Option Strict Off
Option Explicit On
Option Compare Binary
Friend Class clsImportRA
	' @(h) clsReadRA.cls
	'
	' @(s)
	' JVData "RA" �f�[�^�x�[�X�o�^�N���X
	'
	
	Private mBuf As JV_RA_RACE ''���[�X�ڍ׍\����
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
        strSql = "SELECT * FROM RACE"
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
    ' ������    : lBuf - JVData ���ʎq"RA" �̂P�s
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
                mRS.Fields("RaceNum").Value = .RaceNum '' ���[�X�ԍ�
            End With ' id
            With .RaceInfo
                mRS.Fields("YoubiCD").Value = .YoubiCD '' �j���R�[�h
                mRS.Fields("TokuNum").Value = .TokuNum '' ���ʋ����ԍ�
                mRS.Fields("Hondai").Value = .Hondai '' �������{��
                mRS.Fields("Fukudai").Value = .Fukudai '' ����������
                mRS.Fields("Kakko").Value = .Kakko '' �������J�b�R��
                mRS.Fields("HondaiEng").Value = .HondaiEng '' �������{�艢��
                mRS.Fields("FukudaiEng").Value = .FukudaiEng '' ���������艢��
                mRS.Fields("KakkoEng").Value = .KakkoEng '' �������J�b�R������
                mRS.Fields("Ryakusyo10").Value = .Ryakusyo10 '' ���������̂P�O��
                mRS.Fields("Ryakusyo6").Value = .Ryakusyo6 '' ���������̂U��
                mRS.Fields("Ryakusyo3").Value = .Ryakusyo3 '' ���������̂R��
                mRS.Fields("Kubun").Value = .Kubun '' �������敪
                mRS.Fields("Nkai").Value = .Nkai '' �d�܉񎟑�N��
            End With ' RaceInfo
            mRS.Fields("GradeCD").Value = .GradeCD '' �O���[�h�R�[�h
            mRS.Fields("GradeCDBefore").Value = .GradeCDBefore '' �ύX�O�O���[�h�R�[�h
            With .JyokenInfo
                mRS.Fields("SyubetuCD").Value = .SyubetuCD '' ������ʃR�[�h
                mRS.Fields("KigoCD").Value = .KigoCD '' �����L���R�[�h
                mRS.Fields("JyuryoCD").Value = .JyuryoCD '' �d�ʎ�ʃR�[�h
                For j = 0 To 4
                    mRS.Fields("JyokenCD" & j + 1).Value = .JyokenCD(j) '' ���������R�[�h
                Next j
            End With ' JyokenInfo
            mRS.Fields("JyokenName").Value = .JyokenName '' ������������
            mRS.Fields("Kyori").Value = .Kyori '' ����
            mRS.Fields("KyoriBefore").Value = .KyoriBefore '' �ύX�O����
            mRS.Fields("TrackCD").Value = .TrackCD '' �g���b�N�R�[�h
            mRS.Fields("TrackCDBefore").Value = .TrackCDBefore '' �ύX�O�g���b�N�R�[�h
            mRS.Fields("CourseKubunCD").Value = .CourseKubunCD '' �R�[�X�敪
            mRS.Fields("CourseKubunCDBefore").Value = .CourseKubunCDBefore '' �ύX�O�R�[�X�敪
            For i = 0 To 6
                mRS.Fields("Honsyokin" & i + 1).Value = .Honsyokin(i) '' �{�܋�
            Next i
            For i = 0 To 4
                mRS.Fields("HonsyokinBefore" & i + 1).Value = .HonsyokinBefore(i) '' �ύX�O�{�܋�
            Next i
            For i = 0 To 4
                mRS.Fields("Fukasyokin" & i + 1).Value = .Fukasyokin(i) '' �t���܋�
            Next i
            For i = 0 To 2
                mRS.Fields("FukasyokinBefore" & i + 1).Value = .FukasyokinBefore(i) '' �ύX�O�t���܋�
            Next i
            mRS.Fields("HassoTime").Value = .HassoTime '' ��������
            mRS.Fields("HassoTimeBefore").Value = .HassoTimeBefore '' �ύX�O��������
            mRS.Fields("TorokuTosu").Value = .TorokuTosu '' �o�^����
            mRS.Fields("SyussoTosu").Value = .SyussoTosu '' �o������
            mRS.Fields("NyusenTosu").Value = .NyusenTosu '' ��������
            With .TenkoBaba
                mRS.Fields("TenkoCD").Value = .TenkoCD '' �V��R�[�h
                mRS.Fields("SibaBabaCD").Value = .SibaBabaCD '' �Ŕn���ԃR�[�h
                mRS.Fields("DirtBabaCD").Value = .DirtBabaCD '' �_�[�g�n���ԃR�[�h
            End With ' TenkoBaba
            For i = 0 To 24
                mRS.Fields("LapTime" & i + 1).Value = .LapTime(i) '' ���b�v�^�C��
            Next i
            mRS.Fields("SyogaiMileTime").Value = .SyogaiMileTime '' ��Q�}�C���^�C��
            mRS.Fields("HaronTimeS3").Value = .HaronTimeS3 '' �O�R�n�����^�C��
            mRS.Fields("HaronTimeS4").Value = .HaronTimeS4 '' �O�S�n�����^�C��
            mRS.Fields("HaronTimeL3").Value = .HaronTimeL3 '' ��R�n�����^�C��
            mRS.Fields("HaronTimeL4").Value = .HaronTimeL4 '' ��S�n�����^�C��
            For i = 0 To 3
                With .CornerInfo(i)
                    mRS.Fields("Corner" & i + 1).Value = .Corner '' �R�[�i�[
                    mRS.Fields("Syukaisu" & i + 1).Value = .Syukaisu '' ����
                    mRS.Fields("Jyuni" & i + 1).Value = .Jyuni '' �e�ʉߏ���
                End With ' CornerInfo
            Next i
            mRS.Fields("RecordUpKubun").Value = .RecordUpKubun '' ���R�[�h�X�V�敪
        End With

        With mBuf.id
            System.Diagnostics.Debug.WriteLine("Insert RACE : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum)
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
        Dim i As Short '' ���[�v�J�E���^
        Dim j As Short '' ���[�v�J�E���^
        Dim k As Short '' ���[�v�J�E���^
        Dim strSql As String '' SQL��

        gCon.BeginTrans()
        System.Diagnostics.Debug.WriteLine("BeginTrans")

        strSql = "UPDATE RACE SET "
        With mBuf
            With .head
                strSql = strSql & SS & "RecordSpec" & SE & "='" & Replace(.RecordSpec, "'", "''") & "'," '' ���R�[�h���
                strSql = strSql & SS & "DataKubun" & SE & "='" & Replace(.DataKubun, "'", "''") & "'," '' �f�[�^�敪
                strSql = strSql & SS & "MakeDate" & SE & "= '" & Replace(strMakeDate, "'", "''") & "'," '' �N����
            End With ' head
            With .id
                strSql = strSql & SS & "Year" & SE & "='" & Replace(.Year, "'", "''") & "'," '' �J�ÔN
                strSql = strSql & SS & "MonthDay" & SE & "='" & Replace(.MonthDay, "'", "''") & "'," '' �J�Ì���
                strSql = strSql & SS & "JyoCD" & SE & "='" & Replace(.JyoCD, "'", "''") & "'," '' ���n��R�[�h
                strSql = strSql & SS & "Kaiji" & SE & "='" & Replace(.Kaiji, "'", "''") & "'," '' �J�É��N��
                strSql = strSql & SS & "Nichiji" & SE & "='" & Replace(.Nichiji, "'", "''") & "'," '' �J�Ó���N����
                strSql = strSql & SS & "RaceNum" & SE & "='" & Replace(.RaceNum, "'", "''") & "'," '' ���[�X�ԍ�
            End With ' id
            With .RaceInfo
                strSql = strSql & SS & "YoubiCD" & SE & "='" & Replace(.YoubiCD, "'", "''") & "'," '' �j���R�[�h
                strSql = strSql & SS & "TokuNum" & SE & "='" & Replace(.TokuNum, "'", "''") & "'," '' ���ʋ����ԍ�
                strSql = strSql & SS & "Hondai" & SE & "='" & Replace(.Hondai, "'", "''") & "'," '' �������{��
                strSql = strSql & SS & "Fukudai" & SE & "='" & Replace(.Fukudai, "'", "''") & "'," '' ����������
                strSql = strSql & SS & "Kakko" & SE & "='" & Replace(.Kakko, "'", "''") & "'," '' �������J�b�R��
                strSql = strSql & SS & "HondaiEng" & SE & "='" & Replace(.HondaiEng, "'", "''") & "'," '' �������{�艢��
                strSql = strSql & SS & "FukudaiEng" & SE & "='" & Replace(.FukudaiEng, "'", "''") & "'," '' ���������艢��
                strSql = strSql & SS & "KakkoEng" & SE & "='" & Replace(.KakkoEng, "'", "''") & "'," '' �������J�b�R������
                strSql = strSql & SS & "Ryakusyo10" & SE & "='" & Replace(.Ryakusyo10, "'", "''") & "'," '' ���������̂P�O��
                strSql = strSql & SS & "Ryakusyo6" & SE & "='" & Replace(.Ryakusyo6, "'", "''") & "'," '' ���������̂U��
                strSql = strSql & SS & "Ryakusyo3" & SE & "='" & Replace(.Ryakusyo3, "'", "''") & "'," '' ���������̂R��
                strSql = strSql & SS & "Kubun" & SE & "='" & Replace(.Kubun, "'", "''") & "'," '' �������敪
                strSql = strSql & SS & "Nkai" & SE & "='" & Replace(.Nkai, "'", "''") & "'," '' �d�܉񎟑�N��
            End With ' RaceInfo
            strSql = strSql & SS & "GradeCD" & SE & "='" & Replace(.GradeCD, "'", "''") & "'," '' �O���[�h�R�[�h
            strSql = strSql & SS & "GradeCDBefore" & SE & "='" & Replace(.GradeCDBefore, "'", "''") & "'," '' �ύX�O�O���[�h�R�[�h
            With .JyokenInfo
                strSql = strSql & SS & "SyubetuCD" & SE & "='" & Replace(.SyubetuCD, "'", "''") & "'," '' ������ʃR�[�h
                strSql = strSql & SS & "KigoCD" & SE & "='" & Replace(.KigoCD, "'", "''") & "'," '' �����L���R�[�h
                strSql = strSql & SS & "JyuryoCD" & SE & "='" & Replace(.JyuryoCD, "'", "''") & "'," '' �d�ʎ�ʃR�[�h
                For j = 0 To 4
                    strSql = strSql & SS & "JyokenCD" & j + 1 & SE & "='" & Replace(.JyokenCD(j), "'", "''") & "'," '' ���������R�[�h
                Next j
            End With ' JyokenInfo
            strSql = strSql & SS & "JyokenName" & SE & "='" & Replace(.JyokenName, "'", "''") & "'," '' ������������
            strSql = strSql & SS & "Kyori" & SE & "='" & Replace(.Kyori, "'", "''") & "'," '' ����
            strSql = strSql & SS & "KyoriBefore" & SE & "='" & Replace(.KyoriBefore, "'", "''") & "'," '' �ύX�O����
            strSql = strSql & SS & "TrackCD" & SE & "='" & Replace(.TrackCD, "'", "''") & "'," '' �g���b�N�R�[�h
            strSql = strSql & SS & "TrackCDBefore" & SE & "='" & Replace(.TrackCDBefore, "'", "''") & "'," '' �ύX�O�g���b�N�R�[�h
            strSql = strSql & SS & "CourseKubunCD" & SE & "='" & Replace(.CourseKubunCD, "'", "''") & "'," '' �R�[�X�敪
            strSql = strSql & SS & "CourseKubunCDBefore" & SE & "='" & Replace(.CourseKubunCDBefore, "'", "''") & "'," '' �ύX�O�R�[�X�敪
            For i = 0 To 6
                strSql = strSql & SS & "Honsyokin" & i + 1 & SE & "='" & Replace(.Honsyokin(i), "'", "''") & "'," '' �{�܋�
            Next i
            For i = 0 To 4
                strSql = strSql & SS & "HonsyokinBefore" & i + 1 & SE & "='" & Replace(.HonsyokinBefore(i), "'", "''") & "'," '' �ύX�O�{�܋�
            Next i
            For i = 0 To 4
                strSql = strSql & SS & "Fukasyokin" & i + 1 & SE & "='" & Replace(.Fukasyokin(i), "'", "''") & "'," '' �t���܋�
            Next i
            For i = 0 To 2
                strSql = strSql & SS & "FukasyokinBefore" & i + 1 & SE & "='" & Replace(.FukasyokinBefore(i), "'", "''") & "'," '' �ύX�O�t���܋�
            Next i
            strSql = strSql & SS & "HassoTime" & SE & "='" & Replace(.HassoTime, "'", "''") & "'," '' ��������
            strSql = strSql & SS & "HassoTimeBefore" & SE & "='" & Replace(.HassoTimeBefore, "'", "''") & "'," '' �ύX�O��������
            strSql = strSql & SS & "TorokuTosu" & SE & "='" & Replace(.TorokuTosu, "'", "''") & "'," '' �o�^����
            strSql = strSql & SS & "SyussoTosu" & SE & "='" & Replace(.SyussoTosu, "'", "''") & "'," '' �o������
            strSql = strSql & SS & "NyusenTosu" & SE & "='" & Replace(.NyusenTosu, "'", "''") & "'," '' ��������
            With .TenkoBaba
                strSql = strSql & SS & "TenkoCD" & SE & "='" & Replace(.TenkoCD, "'", "''") & "'," '' �V��R�[�h
                strSql = strSql & SS & "SibaBabaCD" & SE & "='" & Replace(.SibaBabaCD, "'", "''") & "'," '' �Ŕn���ԃR�[�h
                strSql = strSql & SS & "DirtBabaCD" & SE & "='" & Replace(.DirtBabaCD, "'", "''") & "'," '' �_�[�g�n���ԃR�[�h
            End With ' TenkoBaba
            For i = 0 To 24
                strSql = strSql & SS & "LapTime" & i + 1 & SE & "='" & Replace(.LapTime(i), "'", "''") & "'," '' ���b�v�^�C��
            Next i
            strSql = strSql & SS & "SyogaiMileTime" & SE & "='" & Replace(.SyogaiMileTime, "'", "''") & "'," '' ��Q�}�C���^�C��
            strSql = strSql & SS & "HaronTimeS3" & SE & "='" & Replace(.HaronTimeS3, "'", "''") & "'," '' �O�R�n�����^�C��
            strSql = strSql & SS & "HaronTimeS4" & SE & "='" & Replace(.HaronTimeS4, "'", "''") & "'," '' �O�S�n�����^�C��
            strSql = strSql & SS & "HaronTimeL3" & SE & "='" & Replace(.HaronTimeL3, "'", "''") & "'," '' ��R�n�����^�C��
            strSql = strSql & SS & "HaronTimeL4" & SE & "='" & Replace(.HaronTimeL4, "'", "''") & "'," '' ��S�n�����^�C��
            For i = 0 To 3
                With .CornerInfo(i)
                    strSql = strSql & SS & "Corner" & i + 1 & SE & "='" & Replace(.Corner, "'", "''") & "'," '' �R�[�i�[
                    strSql = strSql & SS & "Syukaisu" & i + 1 & SE & "='" & Replace(.Syukaisu, "'", "''") & "'," '' ����
                    strSql = strSql & SS & "Jyuni" & i + 1 & SE & "='" & Replace(.Jyuni, "'", "''") & "'," '' �e�ʉߏ���
                End With ' CornerInfo
            Next i
            strSql = strSql & SS & "RecordUpKubun" & SE & "='" & Replace(.RecordUpKubun, "'", "''") & "'," '' ���R�[�h�X�V�敪
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
            System.Diagnostics.Debug.WriteLine("UPDATE RACE : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum)
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