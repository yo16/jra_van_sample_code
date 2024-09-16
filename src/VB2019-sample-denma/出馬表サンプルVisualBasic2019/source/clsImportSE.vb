' @(h) clsImportSE.vb
'
' @(s)
' JVData "SE" �f�[�^�x�[�X�A�N�Z�X�N���X

Option Strict Off
Option Explicit On
Option Compare Binary

Friend Class clsImportSE

    '�n�����[�X���\����
    Private mBuf As JV_SE_RACE_UMA
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
        strSql = "SELECT * FROM UMA_RACE"

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
    Public Function SelectDB(ByVal strSQL As String) As JV_SE_RACE_UMA()
        On Error GoTo ErrorHandler

        Dim IsDBOpen As Boolean = False

        ' ADODB.Recordset
        Dim dbRS As ADODB.Recordset

        ' ADODB.Fields
        Dim dbFld As ADODB.Fields

        ' ���[�X�ڍ׍\����
        Dim structSE(0) As JV_SE_RACE_UMA

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

            ReDim Preserve structSE(lRecCount)

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
            strBuff = strBuff + dbFld("Wakuban").Value().PadRight(1)
            strBuff = strBuff + dbFld("Umaban").Value().PadRight(2)
            strBuff = strBuff + dbFld("KettoNum").Value().PadRight(10)
            strBuff = strBuff + bPadR(dbFld("Bamei").Value(), 36)
            strBuff = strBuff + dbFld("UmaKigoCD").Value().PadRight(2)
            strBuff = strBuff + dbFld("SexCD").Value().PadRight(1)
            strBuff = strBuff + dbFld("HinsyuCD").Value().PadRight(1)
            strBuff = strBuff + dbFld("KeiroCD").Value().PadRight(2)
            strBuff = strBuff + dbFld("Barei").Value().PadRight(2)
            strBuff = strBuff + dbFld("TozaiCD").Value().PadRight(1)
            strBuff = strBuff + dbFld("ChokyosiCode").Value().PadRight(5)
            strBuff = strBuff + bPadR(dbFld("ChokyosiRyakusyo").Value(), 8)
            strBuff = strBuff + dbFld("BanusiCode").Value().PadRight(6)
            strBuff = strBuff + bPadR(dbFld("BanusiName").Value(), 64)
            strBuff = strBuff + bPadR(dbFld("Fukusyoku").Value(), 60)
            strBuff = strBuff + bPadR(dbFld("reserved1").Value(), 60)
            strBuff = strBuff + dbFld("Futan").Value().PadRight(3)
            strBuff = strBuff + dbFld("FutanBefore").Value().PadRight(3)
            strBuff = strBuff + dbFld("Blinker").Value().PadRight(1)
            strBuff = strBuff + dbFld("reserved2").Value().PadRight(1)
            strBuff = strBuff + dbFld("KisyuCode").Value().PadRight(5)
            strBuff = strBuff + dbFld("KisyuCodeBefore").Value().PadRight(5)
            strBuff = strBuff + bPadR(dbFld("KisyuRyakusyo").Value(), 8)
            strBuff = strBuff + bPadR(dbFld("KisyuRyakusyoBefore").Value(), 8)
            strBuff = strBuff + dbFld("MinaraiCD").Value().PadRight(1)
            strBuff = strBuff + dbFld("MinaraiCDBefore").Value().PadRight(1)
            strBuff = strBuff + dbFld("BaTaijyu").Value().PadRight(3)
            strBuff = strBuff + dbFld("ZogenFugo").Value().PadRight(1)
            strBuff = strBuff + dbFld("ZogenSa").Value().PadRight(3)
            strBuff = strBuff + dbFld("IJyoCD").Value().PadRight(1)
            strBuff = strBuff + dbFld("NyusenJyuni").Value().PadRight(2)
            strBuff = strBuff + dbFld("KakuteiJyuni").Value().PadRight(2)
            strBuff = strBuff + dbFld("DochakuKubun").Value().PadRight(1)
            strBuff = strBuff + dbFld("DochakuTosu").Value().PadRight(1)
            strBuff = strBuff + dbFld("Time").Value().PadRight(4)
            strBuff = strBuff + dbFld("ChakusaCD").Value().PadRight(3)
            strBuff = strBuff + dbFld("ChakusaCDP").Value().PadRight(3)
            strBuff = strBuff + dbFld("ChakusaCDPP").Value().PadRight(3)
            strBuff = strBuff + dbFld("Jyuni1c").Value().PadRight(2)
            strBuff = strBuff + dbFld("Jyuni2c").Value().PadRight(2)
            strBuff = strBuff + dbFld("Jyuni3c").Value().PadRight(2)
            strBuff = strBuff + dbFld("Jyuni4c").Value().PadRight(2)
            strBuff = strBuff + dbFld("Odds").Value().PadRight(4)
            strBuff = strBuff + dbFld("Ninki").Value().PadRight(2)
            strBuff = strBuff + dbFld("Honsyokin").Value().PadRight(8)
            strBuff = strBuff + dbFld("Fukasyokin").Value().PadRight(8)
            strBuff = strBuff + dbFld("reserved3").Value().PadRight(3)
            strBuff = strBuff + dbFld("reserved4").Value().PadRight(3)
            strBuff = strBuff + dbFld("HaronTimeL4").Value().PadRight(3)
            strBuff = strBuff + dbFld("HaronTimeL3").Value().PadRight(3)
            For iLoopCnt = 0 To 2
                strBuff = strBuff + dbFld("KettoNum" & iLoopCnt + 1).Value().PadRight(10)
                strBuff = strBuff + bPadR(dbFld("Bamei" & iLoopCnt + 1).Value(), 36)
            Next iLoopCnt
            strBuff = strBuff + dbFld("TimeDiff").Value().PadRight(4)
            strBuff = strBuff + dbFld("RecordUpKubun").Value().PadRight(1)
            strBuff = strBuff + dbFld("DMKubun").Value().PadRight(1)
            strBuff = strBuff + dbFld("DMTime").Value().PadRight(5)
            strBuff = strBuff + dbFld("DMGosaP").Value().PadRight(4)
            strBuff = strBuff + dbFld("DMGosaM").Value().PadRight(4)
            strBuff = strBuff + dbFld("DMJyuni").Value().PadRight(2)
            strBuff = strBuff + dbFld("KyakusituKubun").Value().PadRight(1) + vbCrLf

            ' �\���̂֊i�[
            structSE(lRecCount).SetData(strBuff)

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
        SelectDB = structSE

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
    ' ������    : strBuf - JVData ���ʎq"SE" �̂P�s
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
            ' �g��
            mRS.Fields("Wakuban").Value = .Wakuban
            ' �n��
            mRS.Fields("Umaban").Value = .Umaban
            ' �����o�^�ԍ�
            mRS.Fields("KettoNum").Value = .KettoNum
            ' �n��
            mRS.Fields("Bamei").Value = .Bamei
            ' �n�L���R�[�h
            mRS.Fields("UmaKigoCD").Value = .UmaKigoCD
            ' ���ʃR�[�h
            mRS.Fields("SexCD").Value = .SexCD
            ' �i��R�[�h
            mRS.Fields("HinsyuCD").Value = .HinsyuCD
            ' �ѐF�R�[�h
            mRS.Fields("KeiroCD").Value = .KeiroCD
            ' �n��
            mRS.Fields("Barei").Value = .Barei
            ' ���������R�[�h
            mRS.Fields("TozaiCD").Value = .TozaiCD
            ' �����t�R�[�h
            mRS.Fields("ChokyosiCode").Value = .ChokyosiCode
            ' �����t������
            mRS.Fields("ChokyosiRyakusyo").Value = .ChokyosiRyakusyo
            ' �n��R�[�h
            mRS.Fields("BanusiCode").Value = .BanusiCode
            ' �n�喼
            mRS.Fields("BanusiName").Value = .BanusiName
            ' ���F�W��
            mRS.Fields("Fukusyoku").Value = .Fukusyoku
            ' �\��
            mRS.Fields("reserved1").Value = .reserved1
            ' ���S�d��
            mRS.Fields("Futan").Value = .Futan
            ' �ύX�O���S�d��
            mRS.Fields("FutanBefore").Value = .FutanBefore
            ' �u�����J�[�g�p�敪
            mRS.Fields("Blinker").Value = .Blinker
            ' �\��
            mRS.Fields("reserved2").Value = .reserved2
            ' �R��R�[�h
            mRS.Fields("KisyuCode").Value = .KisyuCode
            ' �ύX�O�R��R�[�h
            mRS.Fields("KisyuCodeBefore").Value = .KisyuCodeBefore
            ' �R�薼����
            mRS.Fields("KisyuRyakusyo").Value = .KisyuRyakusyo
            ' �ύX�O�R�薼����
            mRS.Fields("KisyuRyakusyoBefore").Value = .KisyuRyakusyoBefore
            ' �R�茩�K�R�[�h
            mRS.Fields("MinaraiCD").Value = .MinaraiCD
            ' �ύX�O�R�茩�K�R�[�h
            mRS.Fields("MinaraiCDBefore").Value = .MinaraiCDBefore
            ' �n�̏d
            mRS.Fields("BaTaijyu").Value = .BaTaijyu
            ' ��������
            mRS.Fields("ZogenFugo").Value = .ZogenFugo
            ' ������
            mRS.Fields("ZogenSa").Value = .ZogenSa
            ' �ُ�敪�R�[�h
            mRS.Fields("IJyoCD").Value = .IJyoCD
            ' ��������
            mRS.Fields("NyusenJyuni").Value = .NyusenJyuni
            ' �m�蒅��
            mRS.Fields("KakuteiJyuni").Value = .KakuteiJyuni
            ' �����敪
            mRS.Fields("DochakuKubun").Value = .DochakuKubun
            ' ��������
            mRS.Fields("DochakuTosu").Value = .DochakuTosu
            ' ���j�^�C��
            mRS.Fields("Time").Value = .Time
            ' �����R�[�h
            mRS.Fields("ChakusaCD").Value = .ChakusaCD
            ' +�����R�[�h
            mRS.Fields("ChakusaCDP").Value = .ChakusaCDP
            ' ++�����R�[�h
            mRS.Fields("ChakusaCDPP").Value = .ChakusaCDPP
            ' 1�R�[�i�[�ł̏���
            mRS.Fields("Jyuni1c").Value = .Jyuni1c
            ' 2�R�[�i�[�ł̏���
            mRS.Fields("Jyuni2c").Value = .Jyuni2c
            ' 3�R�[�i�[�ł̏���
            mRS.Fields("Jyuni3c").Value = .Jyuni3c
            ' 4�R�[�i�[�ł̏���
            mRS.Fields("Jyuni4c").Value = .Jyuni4c
            ' �P���I�b�Y
            mRS.Fields("Odds").Value = .Odds
            ' �P���l�C��
            mRS.Fields("Ninki").Value = .Ninki
            ' �l���{�܋�
            mRS.Fields("Honsyokin").Value = .Honsyokin
            ' �l���t���܋�
            mRS.Fields("Fukasyokin").Value = .Fukasyokin
            ' �\��
            mRS.Fields("reserved3").Value = .reserved3
            ' �\��
            mRS.Fields("reserved4").Value = .reserved4
            ' ��S�n�����^�C��
            mRS.Fields("HaronTimeL4").Value = .HaronTimeL4
            ' ��R�n�����^�C��
            mRS.Fields("HaronTimeL3").Value = .HaronTimeL3
            For iLoopCnt = 0 To 2
                With .ChakuUmaInfo(iLoopCnt)
                    ' �����o�^�ԍ�
                    mRS.Fields("KettoNum" & iLoopCnt + 1).Value = .KettoNum
                    ' �n��
                    mRS.Fields("Bamei" & iLoopCnt + 1).Value = .Bamei
                End With ' ChakuUmaInfo
            Next iLoopCnt
            ' �^�C����
            mRS.Fields("TimeDiff").Value = .TimeDiff
            ' ���R�[�h�X�V�敪
            mRS.Fields("RecordUpKubun").Value = .RecordUpKubun
            ' �}�C�j���O�敪
            mRS.Fields("DMKubun").Value = .DMKubun
            ' �}�C�j���O�\�z���j�^�C��
            mRS.Fields("DMTime").Value = .DMTime
            ' �\���덷(�M���x)�{
            mRS.Fields("DMGosaP").Value = .DMGosaP
            ' �\���덷(�M���x)�|
            mRS.Fields("DMGosaM").Value = .DMGosaM
            ' �}�C�j���O�\�z����
            mRS.Fields("DMJyuni").Value = .DMJyuni
            ' ���񃌁[�X�r������
            mRS.Fields("KyakusituKubun").Value = .KyakusituKubun
        End With

        With mBuf.id
            System.Diagnostics.Debug.WriteLine("Insert UMA_RACE : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum)
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

        strSql = "UPDATE UMA_RACE SET "
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
            ' �g��
            strSql = strSql & SS & "Wakuban" & SE & "='" & Replace(.Wakuban, "'", "''") & "',"
            ' �n��
            strSql = strSql & SS & "Umaban" & SE & "='" & Replace(.Umaban, "'", "''") & "',"
            ' �����o�^�ԍ�
            strSql = strSql & SS & "KettoNum" & SE & "='" & Replace(.KettoNum, "'", "''") & "',"
            ' �n��
            strSql = strSql & SS & "Bamei" & SE & "='" & Replace(.Bamei, "'", "''") & "',"
            ' �n�L���R�[�h
            strSql = strSql & SS & "UmaKigoCD" & SE & "='" & Replace(.UmaKigoCD, "'", "''") & "',"
            ' ���ʃR�[�h
            strSql = strSql & SS & "SexCD" & SE & "='" & Replace(.SexCD, "'", "''") & "',"
            ' �i��R�[�h
            strSql = strSql & SS & "HinsyuCD" & SE & "='" & Replace(.HinsyuCD, "'", "''") & "',"
            ' �ѐF�R�[�h
            strSql = strSql & SS & "KeiroCD" & SE & "='" & Replace(.KeiroCD, "'", "''") & "',"
            ' �n��
            strSql = strSql & SS & "Barei" & SE & "='" & Replace(.Barei, "'", "''") & "',"
            ' ���������R�[�h
            strSql = strSql & SS & "TozaiCD" & SE & "='" & Replace(.TozaiCD, "'", "''") & "',"
            ' �����t�R�[�h
            strSql = strSql & SS & "ChokyosiCode" & SE & "='" & Replace(.ChokyosiCode, "'", "''") & "',"
            ' �����t������
            strSql = strSql & SS & "ChokyosiRyakusyo" & SE & "='" & Replace(.ChokyosiRyakusyo, "'", "''") & "',"
            ' �n��R�[�h
            strSql = strSql & SS & "BanusiCode" & SE & "='" & Replace(.BanusiCode, "'", "''") & "',"
            ' �n�喼
            strSql = strSql & SS & "BanusiName" & SE & "='" & Replace(.BanusiName, "'", "''") & "',"
            ' ���F�W��
            strSql = strSql & SS & "Fukusyoku" & SE & "='" & Replace(.Fukusyoku, "'", "''") & "',"
            ' �\��
            strSql = strSql & SS & "reserved1" & SE & "='" & Replace(.reserved1, "'", "''") & "',"
            ' ���S�d��
            strSql = strSql & SS & "Futan" & SE & "='" & Replace(.Futan, "'", "''") & "',"
            ' �ύX�O���S�d��
            strSql = strSql & SS & "FutanBefore" & SE & "='" & Replace(.FutanBefore, "'", "''") & "',"
            ' �u�����J�[�g�p�敪
            strSql = strSql & SS & "Blinker" & SE & "='" & Replace(.Blinker, "'", "''") & "',"
            ' �\��
            strSql = strSql & SS & "reserved2" & SE & "='" & Replace(.reserved2, "'", "''") & "',"
            ' �R��R�[�h
            strSql = strSql & SS & "KisyuCode" & SE & "='" & Replace(.KisyuCode, "'", "''") & "',"
            ' �ύX�O�R��R�[�h
            strSql = strSql & SS & "KisyuCodeBefore" & SE & "='" & Replace(.KisyuCodeBefore, "'", "''") & "',"
            ' �R�薼����
            strSql = strSql & SS & "KisyuRyakusyo" & SE & "='" & Replace(.KisyuRyakusyo, "'", "''") & "',"
            ' �ύX�O�R�薼����
            strSql = strSql & SS & "KisyuRyakusyoBefore" & SE & "='" & Replace(.KisyuRyakusyoBefore, "'", "''") & "',"
            ' �R�茩�K�R�[�h
            strSql = strSql & SS & "MinaraiCD" & SE & "='" & Replace(.MinaraiCD, "'", "''") & "',"
            ' �ύX�O�R�茩�K�R�[�h
            strSql = strSql & SS & "MinaraiCDBefore" & SE & "='" & Replace(.MinaraiCDBefore, "'", "''") & "',"
            ' �n�̏d
            strSql = strSql & SS & "BaTaijyu" & SE & "='" & Replace(.BaTaijyu, "'", "''") & "',"
            ' ��������
            strSql = strSql & SS & "ZogenFugo" & SE & "='" & Replace(.ZogenFugo, "'", "''") & "',"
            ' ������
            strSql = strSql & SS & "ZogenSa" & SE & "='" & Replace(.ZogenSa, "'", "''") & "',"
            ' �ُ�敪�R�[�h
            strSql = strSql & SS & "IJyoCD" & SE & "='" & Replace(.IJyoCD, "'", "''") & "',"
            ' ��������
            strSql = strSql & SS & "NyusenJyuni" & SE & "='" & Replace(.NyusenJyuni, "'", "''") & "',"
            ' �m�蒅��
            strSql = strSql & SS & "KakuteiJyuni" & SE & "='" & Replace(.KakuteiJyuni, "'", "''") & "',"
            ' �����敪
            strSql = strSql & SS & "DochakuKubun" & SE & "='" & Replace(.DochakuKubun, "'", "''") & "',"
            ' ��������
            strSql = strSql & SS & "DochakuTosu" & SE & "='" & Replace(.DochakuTosu, "'", "''") & "',"
            ' ���j�^�C��
            strSql = strSql & SS & "Time" & SE & "='" & Replace(.Time, "'", "''") & "',"
            ' �����R�[�h
            strSql = strSql & SS & "ChakusaCD" & SE & "='" & Replace(.ChakusaCD, "'", "''") & "',"
            ' +�����R�[�h
            strSql = strSql & SS & "ChakusaCDP" & SE & "='" & Replace(.ChakusaCDP, "'", "''") & "',"
            ' ++�����R�[�h
            strSql = strSql & SS & "ChakusaCDPP" & SE & "='" & Replace(.ChakusaCDPP, "'", "''") & "',"
            ' 1�R�[�i�[�ł̏���
            strSql = strSql & SS & "Jyuni1c" & SE & "='" & Replace(.Jyuni1c, "'", "''") & "',"
            ' 2�R�[�i�[�ł̏���
            strSql = strSql & SS & "Jyuni2c" & SE & "='" & Replace(.Jyuni2c, "'", "''") & "',"
            ' 3�R�[�i�[�ł̏���
            strSql = strSql & SS & "Jyuni3c" & SE & "='" & Replace(.Jyuni3c, "'", "''") & "',"
            ' 4�R�[�i�[�ł̏���
            strSql = strSql & SS & "Jyuni4c" & SE & "='" & Replace(.Jyuni4c, "'", "''") & "',"
            ' �P���I�b�Y
            strSql = strSql & SS & "Odds" & SE & "='" & Replace(.Odds, "'", "''") & "',"
            ' �P���l�C��
            strSql = strSql & SS & "Ninki" & SE & "='" & Replace(.Ninki, "'", "''") & "',"
            ' �l���{�܋�
            strSql = strSql & SS & "Honsyokin" & SE & "='" & Replace(.Honsyokin, "'", "''") & "',"
            ' �l���t���܋�
            strSql = strSql & SS & "Fukasyokin" & SE & "='" & Replace(.Fukasyokin, "'", "''") & "',"
            ' �\��
            strSql = strSql & SS & "reserved3" & SE & "='" & Replace(.reserved3, "'", "''") & "',"
            ' �\��
            strSql = strSql & SS & "reserved4" & SE & "='" & Replace(.reserved4, "'", "''") & "',"
            ' ��S�n�����^�C��
            strSql = strSql & SS & "HaronTimeL4" & SE & "='" & Replace(.HaronTimeL4, "'", "''") & "',"
            ' ��R�n�����^�C��
            strSql = strSql & SS & "HaronTimeL3" & SE & "='" & Replace(.HaronTimeL3, "'", "''") & "',"
            For iLoopCnt = 0 To 2
                With .ChakuUmaInfo(iLoopCnt)
                    ' �����o�^�ԍ�
                    strSql = strSql & SS & "KettoNum" & iLoopCnt + 1 & "" & SE & "='" & Replace(.KettoNum, "'", "''") & "',"
                    ' �n��
                    strSql = strSql & SS & "Bamei" & iLoopCnt + 1 & "" & SE & "='" & Replace(.Bamei, "'", "''") & "',"
                End With ' ChakuUmaInfo
            Next iLoopCnt

            ' �^�C����
            strSql = strSql & SS & "TimeDiff" & SE & "='" & Replace(.TimeDiff, "'", "''") & "',"
            ' ���R�[�h�X�V�敪
            strSql = strSql & SS & "RecordUpKubun" & SE & "='" & Replace(.RecordUpKubun, "'", "''") & "',"
            ' �}�C�j���O�敪
            strSql = strSql & SS & "DMKubun" & SE & "='" & Replace(.DMKubun, "'", "''") & "',"
            ' �}�C�j���O�\�z���j�^�C��
            strSql = strSql & SS & "DMTime" & SE & "='" & Replace(.DMTime, "'", "''") & "',"
            ' �\���덷(�M���x)�{
            strSql = strSql & SS & "DMGosaP" & SE & "='" & Replace(.DMGosaP, "'", "''") & "',"
            ' �\���덷(�M���x)�|
            strSql = strSql & SS & "DMGosaM" & SE & "='" & Replace(.DMGosaM, "'", "''") & "',"
            ' �}�C�j���O�\�z����
            strSql = strSql & SS & "DMJyuni" & SE & "='" & Replace(.DMJyuni, "'", "''") & "',"
            ' ���񃌁[�X�r������
            strSql = strSql & SS & "KyakusituKubun" & SE & "='" & Replace(.KyakusituKubun, "'", "''") & "',"

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
                strSql = strSql & " AND " & SS & "MakeDate" & SE & "<= '" & Replace(strMakeDate, "'", "''") & "'"
            End With
        End With

        With mBuf.id
            System.Diagnostics.Debug.WriteLine("UPDATE UMA_RACE : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum & mBuf.Umaban & mBuf.KettoNum)
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