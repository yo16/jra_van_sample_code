' @(h) clsImportRA.vb
'
' @(s)
' JVData "RA" �f�[�^�x�[�X�A�N�Z�X�N���X

Option Strict Off
Option Explicit On
Option Compare Binary

Friend Class clsImportRA

    '���[�X�ڍ׍\����
    Private mBuf As JV_RA_RACE
    Private mRS As ADODB.Recordset


    ' @(f)
    '
    ' �@�\      : ����������
    '
    ' ������    :
    '
    ' �Ԃ�l    :
    '
    ' �@�\����  :
    '
    Public Sub New()

        MyBase.New()
        Class_Initialize_Renamed()

    End Sub


    ' @(f)
    '
    ' �@�\      : �I������
    '
    ' ������    :
    '
    ' �Ԃ�l    :
    '
    ' �@�\����  :
    '
    Protected Overrides Sub Finalize()

        Class_Terminate_Renamed()
        MyBase.Finalize()

    End Sub


    ' @(f)
    '
    ' �@�\      : �������A�R�l�N�V�����A���R�[�h�Z�b�g�I�u�W�F�N�g�̃C���X�^���X����
    '
    ' ������    :
    '
    ' �Ԃ�l    :
    '
    ' �@�\����  :
    '
    '
    Private Sub Class_Initialize_Renamed()
        On Error GoTo ErrorHandler

        ' SQL��
        Dim strSql As String
        strSql = "SELECT * FROM RACE"

        ' ���R�[�h�Z�b�g�I�[�v��
        mRS = New ADODB.Recordset()
        mRS.Open(strSql, gCon, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockPessimistic)

ExitHandler:
        Exit Sub

ErrorHandler:
        System.Diagnostics.Debug.WriteLine(Err.Description)
        Resume ExitHandler

    End Sub


    ' @(f)
    '
    ' �@�\      : �I������
    '
    ' ������    :
    '
    ' �Ԃ�l    :
    '
    ' �@�\����  :
    '
    Private Sub Class_Terminate_Renamed()
        On Error GoTo ErrorHandler

ExitHandler:
        Exit Sub

ErrorHandler:
        System.Diagnostics.Debug.WriteLine(Err.Description)
        Resume ExitHandler

    End Sub


    ' @(f)
    '
    ' �@�\      : �����o�[�ϐ��̃��R�[�h�Z�b�g�̃N���[�Y����
    '
    ' ������    :
    '
    ' �Ԃ�l    :
    '
    ' �@�\����  : �K�[�x�b�W�R���N�V������Close���Ă΂���ƁA�����ŌĂ΂�邩
    '            ������Ȃ��ׁA�����I�ɌĂяo���K�v������܂��B
    '
    Public Sub Close()

        '���R�[�h�Z�b�g�N���[�Y
        mRS.Close()
        mRS = Nothing

    End Sub


    ' @(f)
    '
    ' �@�\      : ���R�[�h�̒��o(SELECT)����
    '
    ' ������    : SQL������
    '
    ' �Ԃ�l    : ���[�X�ڍ׍\���̔z��
    '
    ' �@�\����  :
    '
    Public Function SelectDB(ByVal strSQL As String) As JV_RA_RACE()
        On Error GoTo ErrorHandler

        Dim IsDBOpen As Boolean = False

        ' ADODB.Recordset
        Dim dbRS As ADODB.Recordset

        ' ADODB.Fields
        Dim dbFld As ADODB.Fields

        ' ���[�X�ڍ׍\����
        Dim structRA(0) As JV_RA_RACE

        ' ���[�v�J�E���^
        Dim iLoopCnt As Integer

        ' ���R�[�h����
        Dim lRecCount As Long
        lRecCount = 0

        ' ���R�[�h������
        Dim strBuff As String

        ' ���R�[�h�Z�b�g�̐���
        dbRS = New ADODB.Recordset()
        ' ���R�[�h�Z�b�g�̃I�[�v��
        dbRS.Open(strSQL, gCon, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockPessimistic)
        IsDBOpen = True

        While Not dbRS.EOF
            ' �t�B�[���h�̎擾
            dbFld = dbRS.Fields

            ReDim Preserve structRA(lRecCount)

            ' �\���̐ݒ�p�p�����[�^�쐬
            strBuff = dbFld("RecordSpec").Value().PadRight(2)
            strBuff = strBuff + dbFld("DataKubun").Value().PadRight(1)
            strBuff = strBuff + dbFld("MakeDate").Value().PadRight(8)
            strBuff = strBuff + dbFld("Year").Value().PadRight(4)
            strBuff = strBuff + dbFld("MonthDay").Value().PadRight(4)
            strBuff = strBuff + dbFld("JyoCD").Value().PadRight(2)
            strBuff = strBuff + dbFld("Kaiji").Value().PadRight(2)
            strBuff = strBuff + dbFld("Nichiji").Value().PadRight(2)
            strBuff = strBuff + dbFld("RaceNum").Value().PadRight(2)
            strBuff = strBuff + dbFld("YoubiCD").Value().PadRight(1)
            strBuff = strBuff + dbFld("TokuNum").Value().PadRight(4)
            strBuff = strBuff + bPadR(dbFld("Hondai").Value(), 60)
            strBuff = strBuff + bPadR(dbFld("Fukudai").Value(), 60)
            strBuff = strBuff + bPadR(dbFld("Kakko").Value(), 60)
            strBuff = strBuff + dbFld("HondaiEng").Value().PadRight(120)
            strBuff = strBuff + dbFld("FukudaiEng").Value().PadRight(120)
            strBuff = strBuff + dbFld("KakkoEng").Value().PadRight(120)
            strBuff = strBuff + bPadR(dbFld("Ryakusyo10").Value(), 20)
            strBuff = strBuff + bPadR(dbFld("Ryakusyo6").Value(), 12)
            strBuff = strBuff + bPadR(dbFld("Ryakusyo3").Value(), 6)
            strBuff = strBuff + dbFld("Kubun").Value().PadRight(1)
            strBuff = strBuff + dbFld("Nkai").Value().PadRight(3)
            strBuff = strBuff + dbFld("GradeCD").Value().PadRight(1)
            strBuff = strBuff + dbFld("GradeCDBefore").Value().PadRight(1)
            strBuff = strBuff + dbFld("SyubetuCD").Value().PadRight(2)
            strBuff = strBuff + dbFld("KigoCD").Value().PadRight(3)
            strBuff = strBuff + dbFld("JyuryoCD").Value().PadRight(1)
            For iLoopCnt = 0 To 4
                strBuff = strBuff + dbFld("JyokenCD" & iLoopCnt + 1).Value().PadRight(3)
            Next iLoopCnt
            strBuff = strBuff + bPadR(dbFld("JyokenName").Value(), 60)
            strBuff = strBuff + dbFld("Kyori").Value().PadRight(4)
            strBuff = strBuff + dbFld("KyoriBefore").Value().PadRight(4)
            strBuff = strBuff + dbFld("TrackCD").Value().PadRight(2)
            strBuff = strBuff + dbFld("TrackCDBefore").Value().PadRight(2)
            strBuff = strBuff + dbFld("CourseKubunCD").Value().PadRight(2)
            strBuff = strBuff + dbFld("CourseKubunCDBefore").Value().PadRight(2)
            For iLoopCnt = 0 To 6
                strBuff = strBuff + dbFld("Honsyokin" & iLoopCnt + 1).Value().PadRight(8)
            Next iLoopCnt
            For iLoopCnt = 0 To 4
                strBuff = strBuff + dbFld("HonsyokinBefore" & iLoopCnt + 1).Value().PadRight(8)
            Next iLoopCnt
            For iLoopCnt = 0 To 4
                strBuff = strBuff + dbFld("Fukasyokin" & iLoopCnt + 1).Value().PadRight(8)
            Next iLoopCnt
            For iLoopCnt = 0 To 2
                strBuff = strBuff + dbFld("FukasyokinBefore" & iLoopCnt + 1).Value().PadRight(8)
            Next iLoopCnt
            strBuff = strBuff + dbFld("HassoTime").Value().PadRight(4)
            strBuff = strBuff + dbFld("HassoTimeBefore").Value().PadRight(4)
            strBuff = strBuff + dbFld("TorokuTosu").Value().PadRight(2)
            strBuff = strBuff + dbFld("SyussoTosu").Value().PadRight(2)
            strBuff = strBuff + dbFld("NyusenTosu").Value().PadRight(2)
            strBuff = strBuff + dbFld("TenkoCD").Value().PadRight(1)
            strBuff = strBuff + dbFld("SibaBabaCD").Value().PadRight(1)
            strBuff = strBuff + dbFld("DirtBabaCD").Value().PadRight(1)
            For iLoopCnt = 0 To 24
                strBuff = strBuff + dbFld("LapTime" & iLoopCnt + 1).Value().PadRight(3)
            Next iLoopCnt
            strBuff = strBuff + dbFld("SyogaiMileTime").Value().PadRight(4)
            strBuff = strBuff + dbFld("HaronTimeS3").Value().PadRight(3)
            strBuff = strBuff + dbFld("HaronTimeS4").Value().PadRight(3)
            strBuff = strBuff + dbFld("HaronTimeL3").Value().PadRight(3)
            strBuff = strBuff + dbFld("HaronTimeL4").Value().PadRight(3)
            For iLoopCnt = 0 To 3
                strBuff = strBuff + dbFld("Corner" & iLoopCnt + 1).Value().PadRight(1)
                strBuff = strBuff + dbFld("Syukaisu" & iLoopCnt + 1).Value().PadRight(1)
                strBuff = strBuff + dbFld("Jyuni" & iLoopCnt + 1).Value().PadRight(70)
            Next iLoopCnt
            strBuff = strBuff + dbFld("RecordUpKubun").Value().PadRight(1) + vbCrLf

            ' �\���̂֊i�[
            structRA(lRecCount).SetData(strBuff)

            ' ���R�[�h�����J�E���g
            lRecCount = lRecCount + 1

            ' �����R�[�h��
            dbRS.MoveNext()

        End While

ExitHandler:
        ' ���R�[�h�Z�b�g�̃N���[�Y
        If dbRS Is Nothing = False And IsDBOpen = True Then
            dbRS.Close()
        End If
        dbRS = Nothing

        ' �擾�����\���̔z������^�[��
        SelectDB = structRA

        Exit Function

ErrorHandler:
        System.Diagnostics.Debug.WriteLine(Err.Description)
        Resume ExitHandler

    End Function


    ' @(f)
    '
    ' �@�\      : ���R�[�h�̍폜(DELETE)����
    '
    ' ������    : SQL������
    '
    ' �Ԃ�l    : True - ����, False - ���s
    '
    ' �@�\����  :
    '
    Public Function DeleteDB(ByVal strSQL As String) As Boolean
        On Error GoTo ErrorHandler

        Dim bRetStatus As Boolean
        bRetStatus = True

        ' �g�����U�N�V�����J�n
        gCon.BeginTrans()

        '�e�[�u���̃��R�[�h���p�����[�^��SQL�ō폜����
        gCon.Execute(strSQL)

        ' �g�����U�N�V�����I��(�R�~�b�g)
        gCon.CommitTrans()
        System.Diagnostics.Debug.WriteLine("CommitTrans")

ExitHandler:
        DeleteDB = bRetStatus
        Exit Function

ErrorHandler:
        System.Diagnostics.Debug.WriteLine(Err.Description)
        bRetStatus = False

        ' �g�����U�N�V�����I��(���[���o�b�N)
        gCon.RollbackTrans()
        System.Diagnostics.Debug.WriteLine("RollbackTrans")
        Resume ExitHandler

    End Function


    ' @(f)
    '
    ' �@�\      : JVRead�̕Ԃ��P�s���f�[�^�x�[�X�ɓo�^����
    '
    ' ������    : strBuf - JVData ���ʎq"RA" �̂P�s
    '             lngBufSize - ���g�p
    '
    ' �Ԃ�l    : True - ����, False - ���s
    '
    ' �@�\����  : clsIRead�C���^�[�t�F�C�XAdd�v���V�[�W���̎���
    '
    Public Function Add(ByRef strBuf As String, ByVal lngBufSize As Integer) As Boolean
        On Error GoTo ErrorHandler

        ' �o�^����f�[�^�̍쐬�N����
        Dim strMakeDate As String

        '�\���̂Ƀf�[�^�Z�b�g
        mBuf.SetData(strBuf)

        With mBuf.head.MakeDate
            strMakeDate = .Year & .Month & .Day
        End With

        ' INSERT����
        If Not InsertDB() Then
            'UPDATE�����iINSERT�����s�����ꍇ�j
            If Not UpdateDB(strMakeDate) Then System.Diagnostics.Debug.WriteLine("�X�V�Ɏ��s���܂����B" & Left(strBuf, 2))
        End If

        Add = True

ExitHandler:
        Exit Function

ErrorHandler:
        System.Diagnostics.Debug.WriteLine(Err.Description)
        Add = False

        Resume ExitHandler

    End Function


    ' @(f)
    '
    ' �@�\      : ���R�[�h�̑}��(INSERT)����
    '
    ' ������    : 
    '
    ' �Ԃ�l    : True - ����, False - ���s
    '
    ' �@�\����  :
    '
    Public Function InsertDB() As Boolean
        On Error GoTo ErrorHandler

        ' ���[�v�J�E���^
        Dim iLoopCnt As Integer

        ' �g�����U�N�V�����J�n
        gCon.BeginTrans()
        System.Diagnostics.Debug.WriteLine("BeginTrans")

        mRS.AddNew()

        With mBuf
            With .head
                ' ���R�[�h���
                mRS.Fields("RecordSpec").Value = .RecordSpec
                ' �f�[�^�敪
                mRS.Fields("DataKubun").Value = .DataKubun
                With .MakeDate
                    ' �N����
                    mRS.Fields("MakeDate").Value = .Year & .Month & .Day
                End With ' MakeDate
            End With ' head
            With .id
                ' �J�ÔN
                mRS.Fields("Year").Value = .Year
                ' �J�Ì���
                mRS.Fields("MonthDay").Value = .MonthDay
                ' ���n��R�[�h
                mRS.Fields("JyoCD").Value = .JyoCD
                ' �J�É��N��
                mRS.Fields("Kaiji").Value = .Kaiji
                ' �J�Ó���N����
                mRS.Fields("Nichiji").Value = .Nichiji
                ' ���[�X�ԍ�
                mRS.Fields("RaceNum").Value = .RaceNum
            End With ' id
            With .RaceInfo
                ' �j���R�[�h
                mRS.Fields("YoubiCD").Value = .YoubiCD
                ' ���ʋ����ԍ�
                mRS.Fields("TokuNum").Value = .TokuNum
                ' �������{��
                mRS.Fields("Hondai").Value = .Hondai
                ' ����������
                mRS.Fields("Fukudai").Value = .Fukudai
                ' �������J�b�R��
                mRS.Fields("Kakko").Value = .Kakko
                ' �������{�艢��
                mRS.Fields("HondaiEng").Value = .HondaiEng
                ' ���������艢��
                mRS.Fields("FukudaiEng").Value = .FukudaiEng
                ' �������J�b�R������
                mRS.Fields("KakkoEng").Value = .KakkoEng
                ' ���������̂P�O��
                mRS.Fields("Ryakusyo10").Value = .Ryakusyo10
                ' ���������̂U��
                mRS.Fields("Ryakusyo6").Value = .Ryakusyo6
                ' ���������̂R��
                mRS.Fields("Ryakusyo3").Value = .Ryakusyo3
                ' �������敪
                mRS.Fields("Kubun").Value = .Kubun
                ' �d�܉񎟑�N��
                mRS.Fields("Nkai").Value = .Nkai
            End With ' RaceInfo
            ' �O���[�h�R�[�h
            mRS.Fields("GradeCD").Value = .GradeCD
            ' �ύX�O�O���[�h�R�[�h
            mRS.Fields("GradeCDBefore").Value = .GradeCDBefore
            With .JyokenInfo
                ' ������ʃR�[�h
                mRS.Fields("SyubetuCD").Value = .SyubetuCD
                ' �����L���R�[�h
                mRS.Fields("KigoCD").Value = .KigoCD
                ' �d�ʎ�ʃR�[�h
                mRS.Fields("JyuryoCD").Value = .JyuryoCD
                For iLoopCnt = 0 To 4
                    ' ���������R�[�h
                    mRS.Fields("JyokenCD" & iLoopCnt + 1).Value = .JyokenCD(iLoopCnt)
                Next iLoopCnt
            End With ' JyokenInfo
            ' ������������
            mRS.Fields("JyokenName").Value = .JyokenName
            ' ����
            mRS.Fields("Kyori").Value = .Kyori
            ' �ύX�O����
            mRS.Fields("KyoriBefore").Value = .KyoriBefore
            ' �g���b�N�R�[�h
            mRS.Fields("TrackCD").Value = .TrackCD
            ' �ύX�O�g���b�N�R�[�h
            mRS.Fields("TrackCDBefore").Value = .TrackCDBefore
            ' �R�[�X�敪
            mRS.Fields("CourseKubunCD").Value = .CourseKubunCD
            ' �ύX�O�R�[�X�敪
            mRS.Fields("CourseKubunCDBefore").Value = .CourseKubunCDBefore
            For iLoopCnt = 0 To 6
                ' �{�܋�
                mRS.Fields("Honsyokin" & iLoopCnt + 1).Value = .Honsyokin(iLoopCnt)
            Next iLoopCnt
            For iLoopCnt = 0 To 4
                ' �ύX�O�{�܋�
                mRS.Fields("HonsyokinBefore" & iLoopCnt + 1).Value = .HonsyokinBefore(iLoopCnt)
            Next iLoopCnt
            For iLoopCnt = 0 To 4
                ' �t���܋�
                mRS.Fields("Fukasyokin" & iLoopCnt + 1).Value = .Fukasyokin(iLoopCnt)
            Next iLoopCnt
            For iLoopCnt = 0 To 2
                ' �ύX�O�t���܋�
                mRS.Fields("FukasyokinBefore" & iLoopCnt + 1).Value = .FukasyokinBefore(iLoopCnt)
            Next iLoopCnt
            ' ��������
            mRS.Fields("HassoTime").Value = .HassoTime
            ' �ύX�O��������
            mRS.Fields("HassoTimeBefore").Value = .HassoTimeBefore
            ' �o�^����
            mRS.Fields("TorokuTosu").Value = .TorokuTosu
            ' �o������
            mRS.Fields("SyussoTosu").Value = .SyussoTosu
            ' ��������
            mRS.Fields("NyusenTosu").Value = .NyusenTosu
            With .TenkoBaba
                ' �V��R�[�h
                mRS.Fields("TenkoCD").Value = .TenkoCD
                ' �Ŕn���ԃR�[�h
                mRS.Fields("SibaBabaCD").Value = .SibaBabaCD
                ' �_�[�g�n���ԃR�[�h
                mRS.Fields("DirtBabaCD").Value = .DirtBabaCD
            End With ' TenkoBaba
            For iLoopCnt = 0 To 24
                ' ���b�v�^�C��
                mRS.Fields("LapTime" & iLoopCnt + 1).Value = .LapTime(iLoopCnt)
            Next iLoopCnt
            ' ��Q�}�C���^�C��
            mRS.Fields("SyogaiMileTime").Value = .SyogaiMileTime
            ' �O�R�n�����^�C��
            mRS.Fields("HaronTimeS3").Value = .HaronTimeS3
            ' �O�S�n�����^�C��
            mRS.Fields("HaronTimeS4").Value = .HaronTimeS4
            ' ��R�n�����^�C��
            mRS.Fields("HaronTimeL3").Value = .HaronTimeL3
            ' ��S�n�����^�C��
            mRS.Fields("HaronTimeL4").Value = .HaronTimeL4
            For iLoopCnt = 0 To 3
                With .CornerInfo(iLoopCnt)
                    ' �R�[�i�[
                    mRS.Fields("Corner" & iLoopCnt + 1).Value = .Corner
                    ' ����
                    mRS.Fields("Syukaisu" & iLoopCnt + 1).Value = .Syukaisu
                    ' �e�ʉߏ���
                    mRS.Fields("Jyuni" & iLoopCnt + 1).Value = .Jyuni
                End With ' CornerInfo
            Next iLoopCnt
            ' ���R�[�h�X�V�敪
            mRS.Fields("RecordUpKubun").Value = .RecordUpKubun
        End With

        With mBuf.id
            System.Diagnostics.Debug.WriteLine("Insert RACE : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum)
        End With ' id

        mRS.Update()

        ' �g�����U�N�V�����I��(�R�~�b�g)
        gCon.CommitTrans()
        System.Diagnostics.Debug.WriteLine("CommitTrans")

        InsertDB = True

ExitHandler:
        Exit Function

ErrorHandler:
        System.Diagnostics.Debug.WriteLine(Err.Description)
        InsertDB = False

        mRS.CancelUpdate()

        ' �g�����U�N�V�����I��(���[���o�b�N)
        gCon.RollbackTrans()
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

        ' ���[�v�J�E���^
        Dim iLoopCnt As Short

        ' SQL��
        Dim strSql As String

        ' �g�����U�N�V�����J�n
        gCon.BeginTrans()
        System.Diagnostics.Debug.WriteLine("BeginTrans")

        strSql = "UPDATE RACE SET "
        With mBuf
            With .head
                ' ���R�[�h���
                strSql = strSql & SS & "RecordSpec" & SE & "='" & Replace(.RecordSpec, "'", "''") & "',"
                ' �f�[�^�敪
                strSql = strSql & SS & "DataKubun" & SE & "='" & Replace(.DataKubun, "'", "''") & "',"
                ' �N����
                strSql = strSql & SS & "MakeDate" & SE & "='" & Replace(strMakeDate, "'", "''") & "',"
            End With ' head
            With .id
                ' �J�ÔN
                strSql = strSql & SS & "Year" & SE & "='" & Replace(.Year, "'", "''") & "',"
                ' �J�Ì���
                strSql = strSql & SS & "MonthDay" & SE & "='" & Replace(.MonthDay, "'", "''") & "',"
                ' ���n��R�[�h
                strSql = strSql & SS & "JyoCD" & SE & "='" & Replace(.JyoCD, "'", "''") & "',"
                ' �J�É��N��
                strSql = strSql & SS & "Kaiji" & SE & "='" & Replace(.Kaiji, "'", "''") & "',"
                ' �J�Ó���N����
                strSql = strSql & SS & "Nichiji" & SE & "='" & Replace(.Nichiji, "'", "''") & "',"
                ' ���[�X�ԍ�
                strSql = strSql & SS & "RaceNum" & SE & "='" & Replace(.RaceNum, "'", "''") & "',"
            End With ' id
            With .RaceInfo
                ' �j���R�[�h
                strSql = strSql & SS & "YoubiCD" & SE & "='" & Replace(.YoubiCD, "'", "''") & "',"
                ' ���ʋ����ԍ�
                strSql = strSql & SS & "TokuNum" & SE & "='" & Replace(.TokuNum, "'", "''") & "',"
                ' �������{��
                strSql = strSql & SS & "Hondai" & SE & "='" & Replace(.Hondai, "'", "''") & "',"
                ' ����������
                strSql = strSql & SS & "Fukudai" & SE & "='" & Replace(.Fukudai, "'", "''") & "',"
                ' �������J�b�R��
                strSql = strSql & SS & "Kakko" & SE & "='" & Replace(.Kakko, "'", "''") & "',"
                ' �������{�艢��
                strSql = strSql & SS & "HondaiEng" & SE & "='" & Replace(.HondaiEng, "'", "''") & "',"
                ' ���������艢��
                strSql = strSql & SS & "FukudaiEng" & SE & "='" & Replace(.FukudaiEng, "'", "''") & "',"
                ' �������J�b�R������
                strSql = strSql & SS & "KakkoEng" & SE & "='" & Replace(.KakkoEng, "'", "''") & "',"
                ' ���������̂P�O��
                strSql = strSql & SS & "Ryakusyo10" & SE & "='" & Replace(.Ryakusyo10, "'", "''") & "',"
                ' ���������̂U��
                strSql = strSql & SS & "Ryakusyo6" & SE & "='" & Replace(.Ryakusyo6, "'", "''") & "',"
                ' ���������̂R��
                strSql = strSql & SS & "Ryakusyo3" & SE & "='" & Replace(.Ryakusyo3, "'", "''") & "',"
                ' �������敪
                strSql = strSql & SS & "Kubun" & SE & "='" & Replace(.Kubun, "'", "''") & "',"
                ' �d�܉񎟑�N��
                strSql = strSql & SS & "Nkai" & SE & "='" & Replace(.Nkai, "'", "''") & "',"
            End With ' RaceInfo
            ' �O���[�h�R�[�h
            strSql = strSql & SS & "GradeCD" & SE & "='" & Replace(.GradeCD, "'", "''") & "',"
            ' �ύX�O�O���[�h�R�[�h
            strSql = strSql & SS & "GradeCDBefore" & SE & "='" & Replace(.GradeCDBefore, "'", "''") & "',"
            With .JyokenInfo
                ' ������ʃR�[�h
                strSql = strSql & SS & "SyubetuCD" & SE & "='" & Replace(.SyubetuCD, "'", "''") & "',"
                ' �����L���R�[�h
                strSql = strSql & SS & "KigoCD" & SE & "='" & Replace(.KigoCD, "'", "''") & "',"
                ' �d�ʎ�ʃR�[�h
                strSql = strSql & SS & "JyuryoCD" & SE & "='" & Replace(.JyuryoCD, "'", "''") & "',"
                For iLoopCnt = 0 To 4
                    ' ���������R�[�h
                    strSql = strSql & SS & "JyokenCD" & iLoopCnt + 1 & "" & SE & "='" & Replace(.JyokenCD(iLoopCnt), "'", "''") & "',"
                Next iLoopCnt
            End With ' JyokenInfo
            ' ������������
            strSql = strSql & SS & "JyokenName" & SE & "='" & Replace(.JyokenName, "'", "''") & "',"
            ' ����
            strSql = strSql & SS & "Kyori" & SE & "='" & Replace(.Kyori, "'", "''") & "',"
            ' �ύX�O����
            strSql = strSql & SS & "KyoriBefore" & SE & "='" & Replace(.KyoriBefore, "'", "''") & "',"
            ' �g���b�N�R�[�h
            strSql = strSql & SS & "TrackCD" & SE & "='" & Replace(.TrackCD, "'", "''") & "',"
            ' �ύX�O�g���b�N�R�[�h
            strSql = strSql & SS & "TrackCDBefore" & SE & "='" & Replace(.TrackCDBefore, "'", "''") & "',"
            ' �R�[�X�敪
            strSql = strSql & SS & "CourseKubunCD" & SE & "='" & Replace(.CourseKubunCD, "'", "''") & "',"
            ' �ύX�O�R�[�X�敪
            strSql = strSql & SS & "CourseKubunCDBefore" & SE & "='" & Replace(.CourseKubunCDBefore, "'", "''") & "',"
            For iLoopCnt = 0 To 6
                ' �{�܋�
                strSql = strSql & SS & "Honsyokin" & iLoopCnt + 1 & "" & SE & "='" & Replace(.Honsyokin(iLoopCnt), "'", "''") & "',"
            Next iLoopCnt
            For iLoopCnt = 0 To 4
                ' �ύX�O�{�܋�
                strSql = strSql & SS & "HonsyokinBefore" & iLoopCnt + 1 & "" & SE & "='" & Replace(.HonsyokinBefore(iLoopCnt), "'", "''") & "',"
            Next iLoopCnt
            For iLoopCnt = 0 To 4
                ' �t���܋�
                strSql = strSql & SS & "Fukasyokin" & iLoopCnt + 1 & "" & SE & "='" & Replace(.Fukasyokin(iLoopCnt), "'", "''") & "',"
            Next iLoopCnt
            For iLoopCnt = 0 To 2
                ' �ύX�O�t���܋�
                strSql = strSql & SS & "FukasyokinBefore" & iLoopCnt + 1 & "" & SE & "='" & Replace(.FukasyokinBefore(iLoopCnt), "'", "''") & "',"
            Next iLoopCnt
            ' ��������
            strSql = strSql & SS & "HassoTime" & SE & "='" & Replace(.HassoTime, "'", "''") & "',"
            ' �ύX�O��������
            strSql = strSql & SS & "HassoTimeBefore" & SE & "='" & Replace(.HassoTimeBefore, "'", "''") & "',"
            ' �o�^����
            strSql = strSql & SS & "TorokuTosu" & SE & "='" & Replace(.TorokuTosu, "'", "''") & "',"
            ' �o������
            strSql = strSql & SS & "SyussoTosu" & SE & "='" & Replace(.SyussoTosu, "'", "''") & "',"
            ' ��������
            strSql = strSql & SS & "NyusenTosu" & SE & "='" & Replace(.NyusenTosu, "'", "''") & "',"
            With .TenkoBaba
                ' �V��R�[�h
                strSql = strSql & SS & "TenkoCD" & SE & "='" & Replace(.TenkoCD, "'", "''") & "',"
                ' �Ŕn���ԃR�[�h
                strSql = strSql & SS & "SibaBabaCD" & SE & "='" & Replace(.SibaBabaCD, "'", "''") & "',"
                ' �_�[�g�n���ԃR�[�h
                strSql = strSql & SS & "DirtBabaCD" & SE & "='" & Replace(.DirtBabaCD, "'", "''") & "',"
            End With ' TenkoBaba
            For iLoopCnt = 0 To 24
                ' ���b�v�^�C��
                strSql = strSql & SS & "LapTime" & iLoopCnt + 1 & "" & SE & "='" & Replace(.LapTime(iLoopCnt), "'", "''") & "',"
            Next iLoopCnt
            ' ��Q�}�C���^�C��
            strSql = strSql & SS & "SyogaiMileTime" & SE & "='" & Replace(.SyogaiMileTime, "'", "''") & "',"
            ' �O�R�n�����^�C��
            strSql = strSql & SS & "HaronTimeS3" & SE & "='" & Replace(.HaronTimeS3, "'", "''") & "',"
            ' �O�S�n�����^�C��
            strSql = strSql & SS & "HaronTimeS4" & SE & "='" & Replace(.HaronTimeS4, "'", "''") & "',"
            ' ��R�n�����^�C��
            strSql = strSql & SS & "HaronTimeL3" & SE & "='" & Replace(.HaronTimeL3, "'", "''") & "',"
            ' ��S�n�����^�C��
            strSql = strSql & SS & "HaronTimeL4" & SE & "='" & Replace(.HaronTimeL4, "'", "''") & "',"
            For iLoopCnt = 0 To 3
                With .CornerInfo(iLoopCnt)
                    ' �R�[�i�[
                    strSql = strSql & SS & "Corner" & iLoopCnt + 1 & "" & SE & "='" & Replace(.Corner, "'", "''") & "',"
                    ' ����
                    strSql = strSql & SS & "Syukaisu" & iLoopCnt + 1 & "" & SE & "='" & Replace(.Syukaisu, "'", "''") & "',"
                    ' �e�ʉߏ���
                    strSql = strSql & SS & "Jyuni" & iLoopCnt + 1 & "" & SE & "='" & Replace(.Jyuni, "'", "''") & "',"
                End With ' CornerInfo
            Next iLoopCnt
            ' ���R�[�h�X�V�敪
            strSql = strSql & SS & "RecordUpKubun" & SE & "='" & Replace(.RecordUpKubun, "'", "''") & "',"

            strSql = Left(strSql, Len(strSql) - 1) ' Delete Last Comma
            With .id
                strSql = strSql & " WHERE " & SS & "Year" & SE & "='" & Replace(.Year, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "MonthDay" & SE & "='" & Replace(.MonthDay, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "JyoCD" & SE & "='" & Replace(.JyoCD, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "Kaiji" & SE & "='" & Replace(.Kaiji, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "Nichiji" & SE & "='" & Replace(.Nichiji, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "RaceNum" & SE & "='" & Replace(.RaceNum, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "MakeDate" & SE & "<= '" & Replace(strMakeDate, "'", "''") & "'"
            End With
        End With

        With mBuf.id
            System.Diagnostics.Debug.WriteLine("UPDATE RACE : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum)
        End With ' id

        ' SQL���s
        gCon.Execute(strSql)

        ' �g�����U�N�V�����I��(�R�~�b�g)
        gCon.CommitTrans()
        System.Diagnostics.Debug.WriteLine("CommitTrans")

        UpdateDB = True

ExitHandler:
        Exit Function

ErrorHandler:
        System.Diagnostics.Debug.WriteLine(Err.Description)
        UpdateDB = False

        ' �g�����U�N�V�����I��(���[���o�b�N)
        gCon.RollbackTrans()
        System.Diagnostics.Debug.WriteLine("RollbackTrans")
        Resume ExitHandler

    End Function

End Class