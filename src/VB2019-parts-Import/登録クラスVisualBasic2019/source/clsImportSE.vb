Option Strict Off
Option Explicit On
Option Compare Binary
Friend Class clsImportSE
	' @(h) clsReadSE.cls
	' @(s)
	' JVData "SE" �f�[�^�x�[�X�o�^�N���X
	'
	
	Private mBuf As JV_SE_RACE_UMA ''�n�����[�X���\����
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
        strSql = "SELECT * FROM UMA_RACE"
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
    ' ������    : lBuf - JVData ���ʎq"SE" �̂P�s
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
            mRS.Fields("Wakuban").Value = .Wakuban '' �g��
            mRS.Fields("Umaban").Value = .Umaban '' �n��
            mRS.Fields("KettoNum").Value = .KettoNum '' �����o�^�ԍ�
            mRS.Fields("Bamei").Value = .Bamei '' �n��
            mRS.Fields("UmaKigoCD").Value = .UmaKigoCD '' �n�L���R�[�h
            mRS.Fields("SexCD").Value = .SexCD '' ���ʃR�[�h
            mRS.Fields("HinsyuCD").Value = .HinsyuCD '' �i��R�[�h
            mRS.Fields("KeiroCD").Value = .KeiroCD '' �ѐF�R�[�h
            mRS.Fields("Barei").Value = .Barei '' �n��
            mRS.Fields("TozaiCD").Value = .TozaiCD '' ���������R�[�h
            mRS.Fields("ChokyosiCode").Value = .ChokyosiCode '' �����t�R�[�h
            mRS.Fields("ChokyosiRyakusyo").Value = .ChokyosiRyakusyo '' �����t������
            mRS.Fields("BanusiCode").Value = .BanusiCode '' �n��R�[�h
            mRS.Fields("BanusiName").Value = .BanusiName '' �n�喼
            mRS.Fields("Fukusyoku").Value = .Fukusyoku '' ���F�W��
            mRS.Fields("reserved1").Value = .reserved1 '' �\��
            mRS.Fields("Futan").Value = .Futan '' ���S�d��
            mRS.Fields("FutanBefore").Value = .FutanBefore '' �ύX�O���S�d��
            mRS.Fields("Blinker").Value = .Blinker '' �u�����J�[�g�p�敪
            mRS.Fields("reserved2").Value = .reserved2 '' �\��
            mRS.Fields("KisyuCode").Value = .KisyuCode '' �R��R�[�h
            mRS.Fields("KisyuCodeBefore").Value = .KisyuCodeBefore '' �ύX�O�R��R�[�h
            mRS.Fields("KisyuRyakusyo").Value = .KisyuRyakusyo '' �R�薼����
            mRS.Fields("KisyuRyakusyoBefore").Value = .KisyuRyakusyoBefore '' �ύX�O�R�薼����
            mRS.Fields("MinaraiCD").Value = .MinaraiCD '' �R�茩�K�R�[�h
            mRS.Fields("MinaraiCDBefore").Value = .MinaraiCDBefore '' �ύX�O�R�茩�K�R�[�h
            mRS.Fields("BaTaijyu").Value = .BaTaijyu '' �n�̏d
            mRS.Fields("ZogenFugo").Value = .ZogenFugo '' ��������
            mRS.Fields("ZogenSa").Value = .ZogenSa '' ������
            mRS.Fields("IJyoCD").Value = .IJyoCD '' �ُ�敪�R�[�h
            mRS.Fields("NyusenJyuni").Value = .NyusenJyuni '' ��������
            mRS.Fields("KakuteiJyuni").Value = .KakuteiJyuni '' �m�蒅��
            mRS.Fields("DochakuKubun").Value = .DochakuKubun '' �����敪
            mRS.Fields("DochakuTosu").Value = .DochakuTosu '' ��������
            mRS.Fields("Time").Value = .Time '' ���j�^�C��
            mRS.Fields("ChakusaCD").Value = .ChakusaCD '' �����R�[�h
            mRS.Fields("ChakusaCDP").Value = .ChakusaCDP '' +�����R�[�h
            mRS.Fields("ChakusaCDPP").Value = .ChakusaCDPP '' ++�����R�[�h
            mRS.Fields("Jyuni1c").Value = .Jyuni1c '' 1�R�[�i�[�ł̏���
            mRS.Fields("Jyuni2c").Value = .Jyuni2c '' 2�R�[�i�[�ł̏���
            mRS.Fields("Jyuni3c").Value = .Jyuni3c '' 3�R�[�i�[�ł̏���
            mRS.Fields("Jyuni4c").Value = .Jyuni4c '' 4�R�[�i�[�ł̏���
            mRS.Fields("Odds").Value = .Odds '' �P���I�b�Y
            mRS.Fields("Ninki").Value = .Ninki '' �P���l�C��
            mRS.Fields("Honsyokin").Value = .Honsyokin '' �l���{�܋�
            mRS.Fields("Fukasyokin").Value = .Fukasyokin '' �l���t���܋�
            mRS.Fields("reserved3").Value = .reserved3 '' �\��
            mRS.Fields("reserved4").Value = .reserved4 '' �\��
            mRS.Fields("HaronTimeL4").Value = .HaronTimeL4 '' ��S�n�����^�C��
            mRS.Fields("HaronTimeL3").Value = .HaronTimeL3 '' ��R�n�����^�C��
            For i = 0 To 2
                With .ChakuUmaInfo(i)
                    mRS.Fields("KettoNum" & i + 1).Value = .KettoNum '' �����o�^�ԍ�
                    mRS.Fields("Bamei" & i + 1).Value = .Bamei '' �n��
                End With ' ChakuUmaInfo
            Next i
            mRS.Fields("TimeDiff").Value = .TimeDiff '' �^�C����
            mRS.Fields("RecordUpKubun").Value = .RecordUpKubun '' ���R�[�h�X�V�敪
            mRS.Fields("DMKubun").Value = .DMKubun '' �}�C�j���O�敪
            mRS.Fields("DMTime").Value = .DMTime '' �}�C�j���O�\�z���j�^�C��
            mRS.Fields("DMGosaP").Value = .DMGosaP '' �\���덷(�M���x)�{
            mRS.Fields("DMGosaM").Value = .DMGosaM '' �\���덷(�M���x)�|
            mRS.Fields("DMJyuni").Value = .DMJyuni '' �}�C�j���O�\�z����
            mRS.Fields("KyakusituKubun").Value = .KyakusituKubun '' ���񃌁[�X�r������
        End With

        With mBuf.id
            System.Diagnostics.Debug.WriteLine("Insert UMA_RACE : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum)
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

        strSql = "UPDATE UMA_RACE SET "

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
                strSql = strSql & SS & "RaceNum" & SE & "='" & Replace(.RaceNum, "'", "''") & "'," '' ���[�X�ԍ�
            End With ' id
            strSql = strSql & SS & "Wakuban" & SE & "='" & Replace(.Wakuban, "'", "''") & "'," '' �g��
            strSql = strSql & SS & "Umaban" & SE & "='" & Replace(.Umaban, "'", "''") & "'," '' �n��
            strSql = strSql & SS & "KettoNum" & SE & "='" & Replace(.KettoNum, "'", "''") & "'," '' �����o�^�ԍ�
            strSql = strSql & SS & "Bamei" & SE & "='" & Replace(.Bamei, "'", "''") & "'," '' �n��
            strSql = strSql & SS & "UmaKigoCD" & SE & "='" & Replace(.UmaKigoCD, "'", "''") & "'," '' �n�L���R�[�h
            strSql = strSql & SS & "SexCD" & SE & "='" & Replace(.SexCD, "'", "''") & "'," '' ���ʃR�[�h
            strSql = strSql & SS & "HinsyuCD" & SE & "='" & Replace(.HinsyuCD, "'", "''") & "'," '' �i��R�[�h
            strSql = strSql & SS & "KeiroCD" & SE & "='" & Replace(.KeiroCD, "'", "''") & "'," '' �ѐF�R�[�h
            strSql = strSql & SS & "Barei" & SE & "='" & Replace(.Barei, "'", "''") & "'," '' �n��
            strSql = strSql & SS & "TozaiCD" & SE & "='" & Replace(.TozaiCD, "'", "''") & "'," '' ���������R�[�h
            strSql = strSql & SS & "ChokyosiCode" & SE & "='" & Replace(.ChokyosiCode, "'", "''") & "'," '' �����t�R�[�h
            strSql = strSql & SS & "ChokyosiRyakusyo" & SE & "='" & Replace(.ChokyosiRyakusyo, "'", "''") & "'," '' �����t������
            strSql = strSql & SS & "BanusiCode" & SE & "='" & Replace(.BanusiCode, "'", "''") & "'," '' �n��R�[�h
            strSql = strSql & SS & "BanusiName" & SE & "='" & Replace(.BanusiName, "'", "''") & "'," '' �n�喼
            strSql = strSql & SS & "Fukusyoku" & SE & "='" & Replace(.Fukusyoku, "'", "''") & "'," '' ���F�W��
            strSql = strSql & SS & "reserved1" & SE & "='" & Replace(.reserved1, "'", "''") & "'," '' �\��
            strSql = strSql & SS & "Futan" & SE & "='" & Replace(.Futan, "'", "''") & "'," '' ���S�d��
            strSql = strSql & SS & "FutanBefore" & SE & "='" & Replace(.FutanBefore, "'", "''") & "'," '' �ύX�O���S�d��
            strSql = strSql & SS & "Blinker" & SE & "='" & Replace(.Blinker, "'", "''") & "'," '' �u�����J�[�g�p�敪
            strSql = strSql & SS & "reserved2" & SE & "='" & Replace(.reserved2, "'", "''") & "'," '' �\��
            strSql = strSql & SS & "KisyuCode" & SE & "='" & Replace(.KisyuCode, "'", "''") & "'," '' �R��R�[�h
            strSql = strSql & SS & "KisyuCodeBefore" & SE & "='" & Replace(.KisyuCodeBefore, "'", "''") & "'," '' �ύX�O�R��R�[�h
            strSql = strSql & SS & "KisyuRyakusyo" & SE & "='" & Replace(.KisyuRyakusyo, "'", "''") & "'," '' �R�薼����
            strSql = strSql & SS & "KisyuRyakusyoBefore" & SE & "='" & Replace(.KisyuRyakusyoBefore, "'", "''") & "'," '' �ύX�O�R�薼����
            strSql = strSql & SS & "MinaraiCD" & SE & "='" & Replace(.MinaraiCD, "'", "''") & "'," '' �R�茩�K�R�[�h
            strSql = strSql & SS & "MinaraiCDBefore" & SE & "='" & Replace(.MinaraiCDBefore, "'", "''") & "'," '' �ύX�O�R�茩�K�R�[�h
            strSql = strSql & SS & "BaTaijyu" & SE & "='" & Replace(.BaTaijyu, "'", "''") & "'," '' �n�̏d
            strSql = strSql & SS & "ZogenFugo" & SE & "='" & Replace(.ZogenFugo, "'", "''") & "'," '' ��������
            strSql = strSql & SS & "ZogenSa" & SE & "='" & Replace(.ZogenSa, "'", "''") & "'," '' ������
            strSql = strSql & SS & "IJyoCD" & SE & "='" & Replace(.IJyoCD, "'", "''") & "'," '' �ُ�敪�R�[�h
            strSql = strSql & SS & "NyusenJyuni" & SE & "='" & Replace(.NyusenJyuni, "'", "''") & "'," '' ��������
            strSql = strSql & SS & "KakuteiJyuni" & SE & "='" & Replace(.KakuteiJyuni, "'", "''") & "'," '' �m�蒅��
            strSql = strSql & SS & "DochakuKubun" & SE & "='" & Replace(.DochakuKubun, "'", "''") & "'," '' �����敪
            strSql = strSql & SS & "DochakuTosu" & SE & "='" & Replace(.DochakuTosu, "'", "''") & "'," '' ��������
            strSql = strSql & SS & "Time" & SE & "='" & Replace(.Time, "'", "''") & "'," '' ���j�^�C��
            strSql = strSql & SS & "ChakusaCD" & SE & "='" & Replace(.ChakusaCD, "'", "''") & "'," '' �����R�[�h
            strSql = strSql & SS & "ChakusaCDP" & SE & "='" & Replace(.ChakusaCDP, "'", "''") & "'," '' +�����R�[�h
            strSql = strSql & SS & "ChakusaCDPP" & SE & "='" & Replace(.ChakusaCDPP, "'", "''") & "'," '' ++�����R�[�h
            strSql = strSql & SS & "Jyuni1c" & SE & "='" & Replace(.Jyuni1c, "'", "''") & "'," '' 1�R�[�i�[�ł̏���
            strSql = strSql & SS & "Jyuni2c" & SE & "='" & Replace(.Jyuni2c, "'", "''") & "'," '' 2�R�[�i�[�ł̏���
            strSql = strSql & SS & "Jyuni3c" & SE & "='" & Replace(.Jyuni3c, "'", "''") & "'," '' 3�R�[�i�[�ł̏���
            strSql = strSql & SS & "Jyuni4c" & SE & "='" & Replace(.Jyuni4c, "'", "''") & "'," '' 4�R�[�i�[�ł̏���
            strSql = strSql & SS & "Odds" & SE & "='" & Replace(.Odds, "'", "''") & "'," '' �P���I�b�Y
            strSql = strSql & SS & "Ninki" & SE & "='" & Replace(.Ninki, "'", "''") & "'," '' �P���l�C��
            strSql = strSql & SS & "Honsyokin" & SE & "='" & Replace(.Honsyokin, "'", "''") & "'," '' �l���{�܋�
            strSql = strSql & SS & "Fukasyokin" & SE & "='" & Replace(.Fukasyokin, "'", "''") & "'," '' �l���t���܋�
            strSql = strSql & SS & "reserved3" & SE & "='" & Replace(.reserved3, "'", "''") & "'," '' �\��
            strSql = strSql & SS & "reserved4" & SE & "='" & Replace(.reserved4, "'", "''") & "'," '' �\��
            strSql = strSql & SS & "HaronTimeL4" & SE & "='" & Replace(.HaronTimeL4, "'", "''") & "'," '' ��S�n�����^�C��
            strSql = strSql & SS & "HaronTimeL3" & SE & "='" & Replace(.HaronTimeL3, "'", "''") & "'," '' ��R�n�����^�C��
            For i = 0 To 2
                With .ChakuUmaInfo(i)
                    strSql = strSql & SS & "KettoNum" & i + 1 & SE & "='" & Replace(.KettoNum, "'", "''") & "'," '' �����o�^�ԍ��i����n1�`3�j
                    strSql = strSql & SS & "Bamei" & i + 1 & SE & "='" & Replace(.Bamei, "'", "''") & "'," '' �n���i����n1�`3�j
                End With ' ChakuUmaInfo
            Next i
            strSql = strSql & SS & "TimeDiff" & SE & "='" & Replace(.TimeDiff, "'", "''") & "'," '' �^�C����
            strSql = strSql & SS & "RecordUpKubun" & SE & "='" & Replace(.RecordUpKubun, "'", "''") & "'," '' ���R�[�h�X�V�敪
            strSql = strSql & SS & "DMKubun" & SE & "='" & Replace(.DMKubun, "'", "''") & "'," '' �}�C�j���O�敪
            strSql = strSql & SS & "DMTime" & SE & "='" & Replace(.DMTime, "'", "''") & "'," '' �}�C�j���O�\�z���j�^�C��
            strSql = strSql & SS & "DMGosaP" & SE & "='" & Replace(.DMGosaP, "'", "''") & "'," '' �\���덷(�M���x)�{
            strSql = strSql & SS & "DMGosaM" & SE & "='" & Replace(.DMGosaM, "'", "''") & "'," '' �\���덷(�M���x)�|
            strSql = strSql & SS & "DMJyuni" & SE & "='" & Replace(.DMJyuni, "'", "''") & "'," '' �}�C�j���O�\�z����
            strSql = strSql & SS & "KyakusituKubun" & SE & "='" & Replace(.KyakusituKubun, "'", "''") & "'," '' ���񃌁[�X�r������

            strSql = Left(strSql, Len(strSql) - 1) ' Delete Last Comma

            With .id
                strSql = strSql & " WHERE " & SS & "Year" & SE & "='" & Replace(.Year, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "MonthDay" & SE & "='" & Replace(.MonthDay, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "JyoCD" & SE & "='" & Replace(.JyoCD, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "Kaiji" & SE & "='" & Replace(.Kaiji, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "Nichiji" & SE & "='" & Replace(.Nichiji, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "RaceNum" & SE & "='" & Replace(.RaceNum, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "Umaban" & SE & "='" & Replace(mBuf.Umaban, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "KettoNum" & SE & "='" & Replace(mBuf.KettoNum, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "MakeDate" & SE & " <= '" & Replace(strMakeDate, "'", "''") & "'"
            End With
        End With

        With mBuf.id
            System.Diagnostics.Debug.WriteLine("UPDATE UMA_RACE : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum & mBuf.Umaban & mBuf.KettoNum)
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