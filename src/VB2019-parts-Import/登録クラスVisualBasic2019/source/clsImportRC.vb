Option Strict Off
Option Explicit On
Option Compare Binary
Friend Class clsImportRC
	' @(h) clsReadRC.cls
	' @(s)
	' JVData "RC" �f�[�^�x�[�X�o�^�N���X
	'
	
	Private mBuf As JV_RC_RECORD ''���R�[�h�}�X�^�\����
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
		strSql = "SELECT * FROM RECORD"
		mRS = New ADODB.Recordset
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
    ' ������    : lBuf - JVData ���ʎq"RC" �̂P�s
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
            mRS.Fields("RecInfoKubun").Value = .RecInfoKubun '' ���R�[�h���ʋ敪
            With .id
                mRS.Fields("Year").Value = .Year '' �J�ÔN
                mRS.Fields("MonthDay").Value = .MonthDay '' �J�Ì���
                mRS.Fields("JyoCD").Value = .JyoCD '' ���n��R�[�h
                mRS.Fields("Kaiji").Value = .Kaiji '' �J�É��N��
                mRS.Fields("Nichiji").Value = .Nichiji '' �J�Ó���N����
                mRS.Fields("RaceNum").Value = .RaceNum '' ���[�X�ԍ�
            End With ' id
            mRS.Fields("TokuNum").Value = .TokuNum '' ���ʋ����ԍ�
            mRS.Fields("Hondai").Value = .Hondai '' �������{��
            mRS.Fields("GradeCD").Value = .GradeCD '' �O���[�h�R�[�h
            mRS.Fields("SyubetuCD_TrackCD").Value = .SyubetuCD & .TrackCD '' ������ʃR�[�h
            mRS.Fields("Kyori").Value = .Kyori '' ����
            mRS.Fields("RecKubun").Value = .RecKubun '' ���R�[�h�敪
            mRS.Fields("RecTime").Value = .RecTime '' ���R�[�h�^�C��
            With .TenkoBaba
                mRS.Fields("TenkoCD").Value = .TenkoCD '' �V��R�[�h
                mRS.Fields("SibaBabaCD").Value = .SibaBabaCD '' �Ŕn���ԃR�[�h
                mRS.Fields("DirtBabaCD").Value = .DirtBabaCD '' �_�[�g�n���ԃR�[�h
            End With ' TenkoBaba
            For i = 0 To 2
                With .RecUmaInfo(i)
                    mRS.Fields("RecUmaKettoNum" & i + 1).Value = .KettoNum '' �����o�^�ԍ�
                    mRS.Fields("RecUmaBamei" & i + 1).Value = .Bamei '' �n��
                    mRS.Fields("RecUmaUmaKigoCD" & i + 1).Value = .UmaKigoCD '' �n�L���R�[�h
                    mRS.Fields("RecUmaSexCD" & i + 1).Value = .SexCD '' ���ʃR�[�h
                    mRS.Fields("RecUmaChokyosiCode" & i + 1).Value = .ChokyosiCode '' �����t�R�[�h
                    mRS.Fields("RecUmaChokyosiName" & i + 1).Value = .ChokyosiName '' �����t��
                    mRS.Fields("RecUmaFutan" & i + 1).Value = .Futan '' ���S�d��
                    mRS.Fields("RecUmaKisyuCode" & i + 1).Value = .KisyuCode '' �R��R�[�h
                    mRS.Fields("RecUmaKisyuName" & i + 1).Value = .KisyuName '' �R�薼
                End With ' RecUmaInfo
            Next i
        End With

        With mBuf.id
            System.Diagnostics.Debug.WriteLine("Insert RECORD : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum & mBuf.SyubetuCD & mBuf.TrackCD & mBuf.Kyori)
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

        strSql = "UPDATE RECORD SET "

        With mBuf
            With .head
                strSql = strSql & SS & "RecordSpec" & SE & "='" & Replace(.RecordSpec, "'", "''") & "'," '' ���R�[�h���
                strSql = strSql & SS & "DataKubun" & SE & "='" & Replace(.DataKubun, "'", "''") & "'," '' �f�[�^�敪
                strSql = strSql & SS & "MakeDate" & SE & "='" & Replace(strMakeDate, "'", "''") & "'," '' �N����
            End With ' head
            strSql = strSql & SS & "RecInfoKubun" & SE & "='" & Replace(.RecInfoKubun, "'", "''") & "'," '' ���R�[�h���ʋ敪
            With .id
                strSql = strSql & SS & "Year" & SE & "='" & Replace(.Year, "'", "''") & "'," '' �J�ÔN
                strSql = strSql & SS & "MonthDay" & SE & "='" & Replace(.MonthDay, "'", "''") & "'," '' �J�Ì���
                strSql = strSql & SS & "JyoCD" & SE & "='" & Replace(.JyoCD, "'", "''") & "'," '' ���n��R�[�h
                strSql = strSql & SS & "Kaiji" & SE & "='" & Replace(.Kaiji, "'", "''") & "'," '' �J�É��N��
                strSql = strSql & SS & "Nichiji" & SE & "='" & Replace(.Nichiji, "'", "''") & "'," '' �J�Ó���N����
                strSql = strSql & SS & "RaceNum" & SE & "='" & Replace(.RaceNum, "'", "''") & "'," '' ���[�X�ԍ�
            End With ' id
            strSql = strSql & SS & "TokuNum" & SE & "='" & Replace(.TokuNum, "'", "''") & "'," '' ���ʋ����ԍ�
            strSql = strSql & SS & "Hondai" & SE & "='" & Replace(.Hondai, "'", "''") & "'," '' �������{��
            strSql = strSql & SS & "GradeCD" & SE & "='" & Replace(.GradeCD, "'", "''") & "'," '' �O���[�h�R�[�h
            strSql = strSql & SS & "SyubetuCD_TrackCD" & SE & "='" & Replace(.SyubetuCD & .TrackCD, "'", "''") & "'," '' ������ʃR�[�h
            strSql = strSql & SS & "Kyori" & SE & "='" & Replace(.Kyori, "'", "''") & "'," '' ����
            strSql = strSql & SS & "RecKubun" & SE & "='" & Replace(.RecKubun, "'", "''") & "'," '' ���R�[�h�敪
            strSql = strSql & SS & "RecTime" & SE & "='" & Replace(.RecTime, "'", "''") & "'," '' ���R�[�h�^�C��
            With .TenkoBaba
                strSql = strSql & SS & "TenkoCD" & SE & "='" & Replace(.TenkoCD, "'", "''") & "'," '' �V��R�[�h
                strSql = strSql & SS & "SibaBabaCD" & SE & "='" & Replace(.SibaBabaCD, "'", "''") & "'," '' �Ŕn���ԃR�[�h
                strSql = strSql & SS & "DirtBabaCD" & SE & "='" & Replace(.DirtBabaCD, "'", "''") & "'," '' �_�[�g�n���ԃR�[�h
            End With ' TenkoBaba
            With .RecUmaInfo(0)
                strSql = strSql & SS & "RecUmaKettoNum1" & SE & "='" & Replace(.KettoNum, "'", "''") & "'," '' �����o�^�ԍ�
                strSql = strSql & SS & "RecUmaBamei1" & SE & "='" & Replace(.Bamei, "'", "''") & "'," '' �n��
                strSql = strSql & SS & "RecUmaUmaKigoCD1" & SE & "='" & Replace(.UmaKigoCD, "'", "''") & "'," '' �n�L���R�[�h
                strSql = strSql & SS & "RecUmaSexCD1" & SE & "='" & Replace(.SexCD, "'", "''") & "'," '' ���ʃR�[�h
                strSql = strSql & SS & "RecUmaChokyosiCode1" & SE & "='" & Replace(.ChokyosiCode, "'", "''") & "'," '' �����t�R�[�h
                strSql = strSql & SS & "RecUmaChokyosiName1" & SE & "='" & Replace(.ChokyosiName, "'", "''") & "'," '' �����t��
                strSql = strSql & SS & "RecUmaFutan1" & SE & "='" & Replace(.Futan, "'", "''") & "'," '' ���S�d��
                strSql = strSql & SS & "RecUmaKisyuCode1" & SE & "='" & Replace(.KisyuCode, "'", "''") & "'," '' �R��R�[�h
                strSql = strSql & SS & "RecUmaKisyuName1" & SE & "='" & Replace(.KisyuName, "'", "''") & "'," '' �R�薼
            End With ' RecUmaInfo
            With .RecUmaInfo(1)
                strSql = strSql & SS & "RecUmaKettoNum2" & SE & "='" & Replace(.KettoNum, "'", "''") & "'," '' �����o�^�ԍ�
                strSql = strSql & SS & "RecUmaBamei2" & SE & "='" & Replace(.Bamei, "'", "''") & "'," '' �n��
                strSql = strSql & SS & "RecUmaUmaKigoCD2" & SE & "='" & Replace(.UmaKigoCD, "'", "''") & "'," '' �n�L���R�[�h
                strSql = strSql & SS & "RecUmaSexCD2" & SE & "='" & Replace(.SexCD, "'", "''") & "'," '' ���ʃR�[�h
                strSql = strSql & SS & "RecUmaChokyosiCode2" & SE & "='" & Replace(.ChokyosiCode, "'", "''") & "'," '' �����t�R�[�h
                strSql = strSql & SS & "RecUmaChokyosiName2" & SE & "='" & Replace(.ChokyosiName, "'", "''") & "'," '' �����t��
                strSql = strSql & SS & "RecUmaFutan2" & SE & "='" & Replace(.Futan, "'", "''") & "'," '' ���S�d��
                strSql = strSql & SS & "RecUmaKisyuCode2" & SE & "='" & Replace(.KisyuCode, "'", "''") & "'," '' �R��R�[�h
                strSql = strSql & SS & "RecUmaKisyuName2" & SE & "='" & Replace(.KisyuName, "'", "''") & "'," '' �R�薼
            End With ' RecUmaInfo
            With .RecUmaInfo(2)
                strSql = strSql & SS & "RecUmaKettoNum3" & SE & "='" & Replace(.KettoNum, "'", "''") & "'," '' �����o�^�ԍ�
                strSql = strSql & SS & "RecUmaBamei3" & SE & "='" & Replace(.Bamei, "'", "''") & "'," '' �n��
                strSql = strSql & SS & "RecUmaUmaKigoCD3" & SE & "='" & Replace(.UmaKigoCD, "'", "''") & "'," '' �n�L���R�[�h
                strSql = strSql & SS & "RecUmaSexCD3" & SE & "='" & Replace(.SexCD, "'", "''") & "'," '' ���ʃR�[�h
                strSql = strSql & SS & "RecUmaChokyosiCode3" & SE & "='" & Replace(.ChokyosiCode, "'", "''") & "'," '' �����t�R�[�h
                strSql = strSql & SS & "RecUmaChokyosiName3" & SE & "='" & Replace(.ChokyosiName, "'", "''") & "'," '' �����t��
                strSql = strSql & SS & "RecUmaFutan3" & SE & "='" & Replace(.Futan, "'", "''") & "'," '' ���S�d��
                strSql = strSql & SS & "RecUmaKisyuCode3" & SE & "='" & Replace(.KisyuCode, "'", "''") & "'," '' �R��R�[�h
                strSql = strSql & SS & "RecUmaKisyuName3" & SE & "='" & Replace(.KisyuName, "'", "''") & "'," '' �R�薼
            End With ' RecUmaInfo
            strSql = Left(strSql, Len(strSql) - 1) ' Delete Last Comma

            strSql = strSql & " WHERE " & SS & "RecInfoKubun" & SE & "='" & Replace(.RecInfoKubun, "'", "''") & "'"
            With .id
                strSql = strSql & " AND " & SS & "Year" & SE & "='" & Replace(.Year, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "MonthDay" & SE & "='" & Replace(.MonthDay, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "JyoCD" & SE & "='" & Replace(.JyoCD, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "Kaiji" & SE & "='" & Replace(.Kaiji, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "Nichiji" & SE & "='" & Replace(.Nichiji, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "RaceNum" & SE & "='" & Replace(.RaceNum, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "TokuNum" & SE & "='" & Replace(mBuf.TokuNum, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "SyubetuCD_TrackCD" & SE & "='" & Replace(mBuf.SyubetuCD & mBuf.TrackCD, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "Kyori" & SE & "='" & Replace(mBuf.Kyori, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "MakeDate" & SE & " <= '" & Replace(strMakeDate, "'", "''") & "'"
            End With
        End With
        With mBuf.id
            System.Diagnostics.Debug.WriteLine("UPDATE RECORD : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum & mBuf.SyubetuCD & mBuf.TrackCD & mBuf.Kyori)
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