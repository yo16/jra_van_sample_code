Option Strict Off
Option Explicit On
Option Compare Binary
Friend Class clsImportTK
	' @(h) clsReadTK
	' @(s)
	' JVData "TK" �f�[�^�x�[�X�o�^�N���X
	'
	
	Private mBuf As JV_TK_TOKUUMA '' ���ʓo�^�n�\����
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
        strSql = "SELECT * FROM TOKU_RACE"
        mRS1 = New ADODB.Recordset()
        mRS1.Open(strSql, gCon, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockPessimistic)

        '���R�[�h�Z�b�g�I�[�v��
        strSql = "SELECT * FROM TOKU"
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
        '���R�[�h�Z�b�g�N���[�Y
        mRS1.Close()
        mRS2.Close()

        mRS1 = Nothing
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
    ' ������    : lBuf - JVData ���ʎq"TK" �̂P�s
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

        mRS1.AddNew()

        ' �w�b�_����
        With mBuf
            With .head
                mRS1.Fields("RecordSpec").Value = .RecordSpec '' ���R�[�h���
                mRS1.Fields("DataKubun").Value = .DataKubun '' �f�[�^�敪
                With .MakeDate
                    mRS1.Fields("MakeDate").Value = .Year & .Month & .Day '' �N����
                End With ' MakeDate
            End With ' head
            With .id
                mRS1.Fields("Year").Value = .Year '' �J�ÔN
                mRS1.Fields("MonthDay").Value = .MonthDay '' �J�Ì���
                mRS1.Fields("JyoCD").Value = .JyoCD '' ���n��R�[�h
                mRS1.Fields("Kaiji").Value = .Kaiji '' �J�É��N��
                mRS1.Fields("Nichiji").Value = .Nichiji '' �J�Ó���N����
                mRS1.Fields("RaceNum").Value = .RaceNum '' ���[�X�ԍ�
            End With ' id
            With .RaceInfo
                mRS1.Fields("YoubiCD").Value = .YoubiCD '' �j���R�[�h
                mRS1.Fields("TokuNum").Value = .TokuNum '' ���ʋ����ԍ�
                mRS1.Fields("Hondai").Value = .Hondai '' �������{��
                mRS1.Fields("Fukudai").Value = .Fukudai '' ����������
                mRS1.Fields("Kakko").Value = .Kakko '' �������J�b�R��
                mRS1.Fields("HondaiEng").Value = .HondaiEng '' �������{�艢��
                mRS1.Fields("FukudaiEng").Value = .FukudaiEng '' ���������艢��
                mRS1.Fields("KakkoEng").Value = .KakkoEng '' �������J�b�R������
                mRS1.Fields("Ryakusyo10").Value = .Ryakusyo10 '' ���������̂P�O��
                mRS1.Fields("Ryakusyo6").Value = .Ryakusyo6 '' ���������̂U��
                mRS1.Fields("Ryakusyo3").Value = .Ryakusyo3 '' ���������̂R��
                mRS1.Fields("Kubun").Value = .Kubun '' �������敪
                mRS1.Fields("Nkai").Value = .Nkai '' �d�܉񎟑�N��
            End With ' RaceInfo
            mRS1.Fields("GradeCD").Value = .GradeCD '' �O���[�h�R�[�h
            With .JyokenInfo
                mRS1.Fields("SyubetuCD").Value = .SyubetuCD '' ������ʃR�[�h
                mRS1.Fields("KigoCD").Value = .KigoCD '' �����L���R�[�h
                mRS1.Fields("JyuryoCD").Value = .JyuryoCD '' �d�ʎ�ʃR�[�h
                For j = 0 To 4
                    mRS1.Fields("JyokenCD" & j + 1).Value = .JyokenCD(j) '' ���������R�[�h
                Next j
            End With ' JyokenInfo
            mRS1.Fields("Kyori").Value = .Kyori '' ����
            mRS1.Fields("TrackCD").Value = .TrackCD '' �g���b�N�R�[�h
            mRS1.Fields("CourseKubunCD").Value = .CourseKubunCD '' �R�[�X�敪
            With .HandiDate
                mRS1.Fields("HandiDate").Value = .Year & .Month & .Day '' �N����
            End With ' HandiDate
            mRS1.Fields("TorokuTosu").Value = .TorokuTosu '' �o�^����
        End With

        With mBuf.id
            System.Diagnostics.Debug.WriteLine("Insert TOKU_RACE : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum)
        End With ' id

        mRS1.Update()


        ' �n������
        For i = 0 To CDbl(mBuf.TorokuTosu) - 1
            mRS2.AddNew()
            With mBuf
                With .head
                    With .MakeDate
                        mRS2.Fields("MakeDate").Value = .Year & .Month & .Day '' �N����
                    End With ' MakeDate
                End With
                With .id
                    mRS2.Fields("Year").Value = .Year '' �J�ÔN
                    mRS2.Fields("MonthDay").Value = .MonthDay '' �J�Ì���
                    mRS2.Fields("JyoCD").Value = .JyoCD '' ���n��R�[�h
                    mRS2.Fields("Kaiji").Value = .Kaiji '' �J�É��N��
                    mRS2.Fields("Nichiji").Value = .Nichiji '' �J�Ó���N����
                    mRS2.Fields("RaceNum").Value = .RaceNum '' ���[�X�ԍ�
                End With ' id
                With .TokuUmaInfo(i)
                    mRS2.Fields("Num").Value = .Num '' �A��
                    mRS2.Fields("KettoNum").Value = .KettoNum '' �����o�^�ԍ�
                    mRS2.Fields("Bamei").Value = .Bamei '' �n��
                    mRS2.Fields("UmaKigoCD").Value = .UmaKigoCD '' �n�L���R�[�h
                    mRS2.Fields("SexCD").Value = .SexCD '' ���ʃR�[�h
                    mRS2.Fields("TozaiCD").Value = .TozaiCD '' �����t���������R�[�h
                    mRS2.Fields("ChokyosiCode").Value = .ChokyosiCode '' �����t�R�[�h
                    mRS2.Fields("ChokyosiRyakusyo").Value = .ChokyosiRyakusyo '' �����t������
                    mRS2.Fields("Futan").Value = .Futan '' ���S�d��
                    mRS2.Fields("Koryu").Value = .Koryu '' �𗬋敪
                End With ' TokuUmaInfo
            End With

            With mBuf.id
                System.Diagnostics.Debug.WriteLine("Insert TOKU : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum & mBuf.TokuUmaInfo(i).Num)
            End With ' id
            mRS2.Update()
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
    Public Function UpdateDB(ByVal strMakeDate As String) As Boolean
        On Error GoTo ErrorHandler
        Dim i As Short '' ���[�v�J�E���^
        Dim j As Short '' ���[�v�J�E���^
        Dim k As Short '' ���[�v�J�E���^
        Dim strSql As String '' SQL��

        gCon.BeginTrans()
        System.Diagnostics.Debug.WriteLine("BeginTrans")

        ' �w�b�_����
        strSql = "UPDATE TOKU_RACE SET "
        With mBuf
            With .head
                strSql = strSql & SS & "RecordSpec" & SE & " = '" & Replace(.RecordSpec, "'", "''") & "'," '' ���R�[�h���
                strSql = strSql & SS & "DataKubun" & SE & " = '" & Replace(.DataKubun, "'", "''") & "'," '' �f�[�^�敪
                With .MakeDate
                    strSql = strSql & SS & "MakeDate" & SE & " = '" & Replace(.Year & .Month & .Day, "'", "''") & "'," '' �N
                End With ' MakeDate
            End With ' head
            With .RaceInfo
                strSql = strSql & SS & "YoubiCD" & SE & " = '" & Replace(.YoubiCD, "'", "''") & "'," '' �j���R�[�h
                strSql = strSql & SS & "TokuNum" & SE & " = '" & Replace(.TokuNum, "'", "''") & "'," '' ���ʋ����ԍ�
                strSql = strSql & SS & "Hondai" & SE & " = '" & Replace(.Hondai, "'", "''") & "'," '' �������{��
                strSql = strSql & SS & "Fukudai" & SE & " = '" & Replace(.Fukudai, "'", "''") & "'," '' ����������
                strSql = strSql & SS & "Kakko" & SE & " = '" & Replace(.Kakko, "'", "''") & "'," '' �������J�b�R��
                strSql = strSql & SS & "HondaiEng" & SE & " = '" & Replace(.HondaiEng, "'", "''") & "'," '' �������{�艢��
                strSql = strSql & SS & "FukudaiEng" & SE & " = '" & Replace(.FukudaiEng, "'", "''") & "'," '' ���������艢��
                strSql = strSql & SS & "KakkoEng" & SE & " = '" & Replace(.KakkoEng, "'", "''") & "'," '' �������J�b�R������
                strSql = strSql & SS & "Ryakusyo10" & SE & " = '" & Replace(.Ryakusyo10, "'", "''") & "'," '' ���������̂P�O��
                strSql = strSql & SS & "Ryakusyo6" & SE & " = '" & Replace(.Ryakusyo6, "'", "''") & "'," '' ���������̂U��
                strSql = strSql & SS & "Ryakusyo3" & SE & " = '" & Replace(.Ryakusyo3, "'", "''") & "'," '' ���������̂R��
                strSql = strSql & SS & "Kubun" & SE & " = '" & Replace(.Kubun, "'", "''") & "'," '' �������敪
                strSql = strSql & SS & "Nkai" & SE & " = '" & Replace(.Nkai, "'", "''") & "'," '' �d�܉񎟑�N��
            End With ' RaceInfo
            strSql = strSql & SS & "GradeCD" & SE & " = '" & Replace(.GradeCD, "'", "''") & "'," '' �O���[�h�R�[�h
            With .JyokenInfo
                strSql = strSql & SS & "SyubetuCD" & SE & " = '" & Replace(.SyubetuCD, "'", "''") & "'," '' ������ʃR�[�h
                strSql = strSql & SS & "KigoCD" & SE & " = '" & Replace(.KigoCD, "'", "''") & "'," '' �����L���R�[�h
                strSql = strSql & SS & "JyuryoCD" & SE & " = '" & Replace(.JyuryoCD, "'", "''") & "'," '' �d�ʎ�ʃR�[�h
                For j = 0 To 4
                    strSql = strSql & SS & "JyokenCD" & j + 1 & SE & " = '" & Replace(.JyokenCD(j), "'", "''") & "'," '' ���������R�[�h
                Next j
            End With ' JyokenInfo
            strSql = strSql & SS & "Kyori" & SE & " = '" & Replace(.Kyori, "'", "''") & "'," '' ����
            strSql = strSql & SS & "TrackCD" & SE & " = '" & Replace(.TrackCD, "'", "''") & "'," '' �g���b�N�R�[�h
            strSql = strSql & SS & "CourseKubunCD" & SE & " = '" & Replace(.CourseKubunCD, "'", "''") & "'," '' �R�[�X�敪
            With .HandiDate
                strSql = strSql & SS & "HandiDate" & SE & " = '" & Replace(.Year & .Month & .Day, "'", "''") & "'," '' HandiDate
            End With ' HandiDate
            strSql = strSql & SS & "TorokuTosu" & SE & " = '" & Replace(.TorokuTosu, "'", "''") & "'," '' �o�^����
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
            System.Diagnostics.Debug.WriteLine("UPDATE TOKU_RACE : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum)
        End With ' id

        Dim tmpRS As ADODB.Recordset
        tmpRS = gCon.Execute(strSql)
        If tmpRS Is Nothing Then
        Else

            ' �n������
            ' �n�����̓��R�[�h�����ς��邽�߁A�����̃f�[�^��S�č폜���Ă���o�^���Ȃ����܂��B

            With mBuf.id
                strSql = "DELETE FROM TOKU"
                strSql = strSql & " WHERE " & SS & "Year" & SE & "='" & Replace(.Year, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "MonthDay" & SE & "='" & Replace(.MonthDay, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "JyoCD" & SE & "='" & Replace(.JyoCD, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "Kaiji" & SE & "='" & Replace(.Kaiji, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "Nichiji" & SE & "='" & Replace(.Nichiji, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "RaceNum" & SE & "='" & Replace(.RaceNum, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "MakeDate" & SE & " <= '" & Replace(strMakeDate, "'", "''") & "'"

                With mBuf.id
                    System.Diagnostics.Debug.WriteLine("DELETE TOKU : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum)
                End With ' id

                gCon.Execute(strSql)
            End With

            ' �S���ēo�^����
            For i = 0 To CDbl(mBuf.TorokuTosu) - 1
                strSql = "INSERT INTO TOKU VALUES( "
                With mBuf
                    With .head
                        strSql = strSql & "'" & Replace(strMakeDate, "'", "''") & "'," ''�񋟔N����
                    End With
                    With .id
                        strSql = strSql & "'" & Replace(.Year, "'", "''") & "'," '' �J�ÔN
                        strSql = strSql & "'" & Replace(.MonthDay, "'", "''") & "'," '' �J�Ì���
                        strSql = strSql & "'" & Replace(.JyoCD, "'", "''") & "'," '' ���n��R�[�h
                        strSql = strSql & "'" & Replace(.Kaiji, "'", "''") & "'," '' �J�É��N��
                        strSql = strSql & "'" & Replace(.Nichiji, "'", "''") & "'," '' �J�Ó���N����
                        strSql = strSql & "'" & Replace(.RaceNum, "'", "''") & "'," '' ���[�X�ԍ�
                    End With ' id
                    With .TokuUmaInfo(i)
                        strSql = strSql & "'" & Replace(.Num, "'", "''") & "'," '' �A��
                        strSql = strSql & "'" & Replace(.KettoNum, "'", "''") & "'," '' �����o�^�ԍ�
                        strSql = strSql & "'" & Replace(.Bamei, "'", "''") & "'," '' �n��
                        strSql = strSql & "'" & Replace(.UmaKigoCD, "'", "''") & "'," '' �n�L���R�[�h
                        strSql = strSql & "'" & Replace(.SexCD, "'", "''") & "'," '' ���ʃR�[�h
                        strSql = strSql & "'" & Replace(.TozaiCD, "'", "''") & "'," '' �����t���������R�[�h
                        strSql = strSql & "'" & Replace(.ChokyosiCode, "'", "''") & "'," '' �����t�R�[�h
                        strSql = strSql & "'" & Replace(.ChokyosiRyakusyo, "'", "''") & "'," '' �����t������
                        strSql = strSql & "'" & Replace(.Futan, "'", "''") & "'," '' ���S�d��
                        strSql = strSql & "'" & Replace(.Koryu, "'", "''") & "'," '' �𗬋敪
                    End With ' TokuUmaInfo
                End With

                strSql = Left(strSql, Len(strSql) - 1) ' Delete Last Comma
                strSql = strSql & ")"

                With mBuf.id
                    System.Diagnostics.Debug.WriteLine("Insert TOKU : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum & mBuf.TokuUmaInfo(i).Num)
                End With ' id
                gCon.Execute(strSql)
            Next i
            tmpRS = Nothing
        End If

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