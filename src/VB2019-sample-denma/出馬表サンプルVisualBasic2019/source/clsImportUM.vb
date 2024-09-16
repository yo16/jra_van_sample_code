' @(h) clsImportUM.vb
'
' @(s)
' JVData "UM" �f�[�^�x�[�X�A�N�Z�X�N���X

Option Strict Off
Option Explicit On
Option Compare Binary

Friend Class clsImportUM

    '�����n�}�X�^�\����
    Private mBuf As JV_UM_UMA
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
        strSql = "SELECT * FROM UMA"

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
    Public Function SelectDB(ByVal strSQL As String) As JV_UM_UMA()
        On Error GoTo ErrorHandler

        Dim IsDBOpen As Boolean = False

        ' ADODB.Recordset
        Dim dbRS As ADODB.Recordset

        ' ADODB.Fields
        Dim dbFld As ADODB.Fields

        ' ���[�X�ڍ׍\����
        Dim structUM(0) As JV_UM_UMA

        ' ���[�v�J�E���^
        Dim iLoopCnt1 As Integer
        Dim iLoopCnt2 As Integer

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

            ReDim Preserve structUM(lRecCount)

            ' �\���̐ݒ�p�p�����[�^�쐬
            strBuff = dbFld("RecordSpec").Value().PadRight(2)
            strBuff = strBuff + dbFld("DataKubun").Value().PadRight(1)
            strBuff = strBuff + dbFld("MakeDate").Value().PadRight(8)
            strBuff = strBuff + dbFld("KettoNum").Value().PadRight(10)
            strBuff = strBuff + dbFld("DelKubun").Value().PadRight(1)
            strBuff = strBuff + dbFld("RegDate").Value().PadRight(8)
            strBuff = strBuff + dbFld("DelDate").Value().PadRight(8)
            strBuff = strBuff + dbFld("BirthDate").Value().PadRight(8)
            strBuff = strBuff + bPadR(dbFld("Bamei").Value(), 36)
            strBuff = strBuff + dbFld("BameiKana").Value().PadRight(36)
            strBuff = strBuff + dbFld("BameiEng").Value().PadRight(60)
            strBuff = strBuff + dbFld("ZaikyuFlag").Value().PadRight(1)
            strBuff = strBuff + dbFld("Reserved").Value().PadRight(19)
            strBuff = strBuff + dbFld("UmaKigoCD").Value().PadRight(2)
            strBuff = strBuff + dbFld("SexCD").Value().PadRight(1)
            strBuff = strBuff + dbFld("HinsyuCD").Value().PadRight(1)
            strBuff = strBuff + dbFld("KeiroCD").Value().PadRight(2)
            For iLoopCnt1 = 0 To 13
                strBuff = strBuff + dbFld("Ketto3InfoHansyokuNum" & iLoopCnt1 + 1).Value().PadRight(10)
                strBuff = strBuff + bPadR(dbFld("Ketto3InfoBamei" & iLoopCnt1 + 1).Value(), 36)
            Next iLoopCnt1
            strBuff = strBuff + dbFld("TozaiCD").Value().PadRight(1)
            strBuff = strBuff + dbFld("ChokyosiCode").Value().PadRight(5)
            strBuff = strBuff + bPadR(dbFld("ChokyosiRyakusyo").Value(), 8)
            strBuff = strBuff + bPadR(dbFld("Syotai").Value(), 20)
            strBuff = strBuff + dbFld("BreederCode").Value().PadRight(8)
            strBuff = strBuff + bPadR(dbFld("BreederName").Value(), 72)
            strBuff = strBuff + bPadR(dbFld("SanchiName").Value(), 20)
            strBuff = strBuff + dbFld("BanusiCode").Value().PadRight(6)
            strBuff = strBuff + bPadR(dbFld("BanusiName").Value(), 64)
            strBuff = strBuff + dbFld("RuikeiHonsyoHeiti").Value().PadRight(9)
            strBuff = strBuff + dbFld("RuikeiHonsyoSyogai").Value().PadRight(9)
            strBuff = strBuff + dbFld("RuikeiFukaHeichi").Value().PadRight(9)
            strBuff = strBuff + dbFld("RuikeiFukaSyogai").Value().PadRight(9)
            strBuff = strBuff + dbFld("RuikeiSyutokuHeichi").Value().PadRight(9)
            strBuff = strBuff + dbFld("RuikeiSyutokuSyogai").Value().PadRight(9)
            For iLoopCnt1 = 0 To 5
                strBuff = strBuff + dbFld("SogoChakukaisu" & iLoopCnt1 + 1).Value().PadRight(3)
            Next iLoopCnt1
            For iLoopCnt1 = 0 To 5
                strBuff = strBuff + dbFld("ChuoChakukaisu" & iLoopCnt1 + 1).Value().PadRight(3)
            Next iLoopCnt1
            For iLoopCnt1 = 0 To 6
                For iLoopCnt2 = 0 To 5
                    strBuff = strBuff + dbFld("Ba" & iLoopCnt1 + 1 & "Chakukaisu" & iLoopCnt2 + 1).Value().PadRight(3)
                Next iLoopCnt2
            Next iLoopCnt1
            For iLoopCnt1 = 0 To 11
                For iLoopCnt2 = 0 To 5
                    strBuff = strBuff + dbFld("Jyotai" & iLoopCnt1 + 1 & "Chakukaisu" & iLoopCnt2 + 1).Value().PadRight(3)
                Next iLoopCnt2
            Next iLoopCnt1
            For iLoopCnt1 = 0 To 5
                For iLoopCnt2 = 0 To 5
                    strBuff = strBuff + dbFld("Kyori" & iLoopCnt1 + 1 & "Chakukaisu" & iLoopCnt2 + 1).Value().PadRight(3)
                Next iLoopCnt2
            Next iLoopCnt1
            For iLoopCnt1 = 0 To 3
                strBuff = strBuff + dbFld("Kyakusitu" & iLoopCnt1 + 1).Value().PadRight(3)
            Next iLoopCnt1
            strBuff = strBuff + dbFld("RaceCount").Value().PadRight(3) & vbCrLf

            ' �\���̂֊i�[
            structUM(lRecCount).SetData(strBuff)

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
        SelectDB = structUM

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
    ' ������    : strBuf - JVData ���ʎq"UM" �̂P�s
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
        Dim iLoopCnt1 As Integer
        Dim iLoopCnt2 As Integer

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
            ' �����o�^�ԍ�
            mRS.Fields("KettoNum").Value = .KettoNum
            ' �����n�����敪
            mRS.Fields("DelKubun").Value = .DelKubun
            With .RegDate
                ' �N����
                mRS.Fields("RegDate").Value = .Year & .Month & .Day
            End With ' RegDate
            With .DelDate
                ' �N����
                mRS.Fields("DelDate").Value = .Year & .Month & .Day
            End With ' DelDate
            With .BirthDate
                ' �N����
                mRS.Fields("BirthDate").Value = .Year & .Month & .Day
            End With ' BirthDate
            ' �n��
            mRS.Fields("Bamei").Value = .Bamei
            ' �n�����p�J�i
            mRS.Fields("BameiKana").Value = .BameiKana
            ' �n������
            mRS.Fields("BameiEng").Value = .BameiEng
            ' JRA�{�ݍ݂��イ�t���O
            mRS.Fields("ZaikyuFlag").Value = .ZaikyuFlag
            ' �\��
            mRS.Fields("Reserved").Value = .Reserved
            ' �n�L���R�[�h
            mRS.Fields("UmaKigoCD").Value = .UmaKigoCD
            ' ���ʃR�[�h
            mRS.Fields("SexCD").Value = .SexCD
            ' �i��R�[�h
            mRS.Fields("HinsyuCD").Value = .HinsyuCD
            ' �ѐF�R�[�h
            mRS.Fields("KeiroCD").Value = .KeiroCD
            For iLoopCnt1 = 0 To 13
                With .Ketto3Info(iLoopCnt1)
                    ' �ɐB�o�^�ԍ�
                    mRS.Fields("Ketto3InfoHansyokuNum" & iLoopCnt1 + 1).Value = .HansyokuNum
                    ' �n��
                    mRS.Fields("Ketto3InfoBamei" & iLoopCnt1 + 1).Value = .Bamei
                End With ' Ketto3Info
            Next iLoopCnt1
            ' ���������R�[�h
            mRS.Fields("TozaiCD").Value = .TozaiCD
            ' �����t�R�[�h
            mRS.Fields("ChokyosiCode").Value = .ChokyosiCode
            ' �����t������
            mRS.Fields("ChokyosiRyakusyo").Value = .ChokyosiRyakusyo
            ' ���Ғn�於
            mRS.Fields("Syotai").Value = .Syotai
            ' ���Y�҃R�[�h
            mRS.Fields("BreederCode").Value = .BreederCode
            ' ���Y�Җ�
            mRS.Fields("BreederName").Value = .BreederName
            ' �Y�n��
            mRS.Fields("SanchiName").Value = .SanchiName
            ' �n��R�[�h
            mRS.Fields("BanusiCode").Value = .BanusiCode
            ' �n�喼
            mRS.Fields("BanusiName").Value = .BanusiName
            ' ���n�{�܋��݌v
            mRS.Fields("RuikeiHonsyoHeiti").Value = .RuikeiHonsyoHeiti
            ' ��Q�{�܋��݌v
            mRS.Fields("RuikeiHonsyoSyogai").Value = .RuikeiHonsyoSyogai
            ' ���n�t���܋��݌v
            mRS.Fields("RuikeiFukaHeichi").Value = .RuikeiFukaHeichi
            ' ��Q�t���܋��݌v
            mRS.Fields("RuikeiFukaSyogai").Value = .RuikeiFukaSyogai
            ' ���n�����܋��݌v
            mRS.Fields("RuikeiSyutokuHeichi").Value = .RuikeiSyutokuHeichi
            ' ��Q�����܋��݌v
            mRS.Fields("RuikeiSyutokuSyogai").Value = .RuikeiSyutokuSyogai
            With .ChakuSogo
                For iLoopCnt1 = 0 To 5
                    mRS.Fields("SogoChakukaisu" & iLoopCnt1 + 1).Value = .Chakukaisu(iLoopCnt1)
                Next iLoopCnt1
            End With ' ChakuSogo
            With .ChakuChuo
                For iLoopCnt1 = 0 To 5
                    mRS.Fields("ChuoChakukaisu" & iLoopCnt1 + 1).Value = .Chakukaisu(iLoopCnt1)
                Next iLoopCnt1
            End With ' ChakuChuo
            For iLoopCnt1 = 0 To 6
                With .ChakuKaisuBa(iLoopCnt1)
                    For iLoopCnt2 = 0 To 5
                        mRS.Fields("Ba" & iLoopCnt1 + 1 & "Chakukaisu" & iLoopCnt2 + 1).Value = .Chakukaisu(iLoopCnt2)
                    Next iLoopCnt2
                End With ' ChakuKaisuBa
            Next iLoopCnt1
            For iLoopCnt1 = 0 To 11
                With .ChakuKaisuJyotai(iLoopCnt1)
                    For iLoopCnt2 = 0 To 5
                        mRS.Fields("Jyotai" & iLoopCnt1 + 1 & "Chakukaisu" & iLoopCnt2 + 1).Value = .Chakukaisu(iLoopCnt2)
                    Next iLoopCnt2
                End With ' ChakuKaisuJyotai
            Next iLoopCnt1
            For iLoopCnt1 = 0 To 5
                With .ChakuKaisuKyori(iLoopCnt1)
                    For iLoopCnt2 = 0 To 5
                        mRS.Fields("Kyori" & iLoopCnt1 + 1 & "Chakukaisu" & iLoopCnt2 + 1).Value = .Chakukaisu(iLoopCnt2)
                    Next iLoopCnt2
                End With ' ChakuKaisuKyori
            Next iLoopCnt1
            For iLoopCnt1 = 0 To 3
                ' �r���X��
                mRS.Fields("Kyakusitu" & iLoopCnt1 + 1).Value = .Kyakusitu(iLoopCnt1)
            Next iLoopCnt1
            ' �o�^���[�X��
            mRS.Fields("RaceCount").Value = .RaceCount
        End With

        With mBuf
            System.Diagnostics.Debug.WriteLine("Insert UMA : " & .KettoNum)
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
        Dim iLoopCnt1 As Short
        Dim iLoopCnt2 As Short

        ' SQL��
        Dim strSql As String

        ' �g�����U�N�V�����J�n
        gCon.BeginTrans()
        System.Diagnostics.Debug.WriteLine("BeginTrans")

        strSql = "UPDATE UMA SET "
        With mBuf
            ' �����o�^�ԍ�
            strSql = strSql & SS & "KettoNum" & SE & "='" & Replace(.KettoNum, "'", "''") & "',"
            ' �����n�����敪
            strSql = strSql & SS & "DelKubun" & SE & "='" & Replace(.DelKubun, "'", "''") & "',"
            With .RegDate
                ' �N����
                strSql = strSql & SS & "RegDate" & SE & "='" & Replace(.Year & .Month & .Day, "'", "''") & "',"
            End With ' RegDate
            With .DelDate
                ' �N����
                strSql = strSql & SS & "DelDate" & SE & "='" & Replace(.Year & .Month & .Day, "'", "''") & "',"
            End With ' DelDate
            With .BirthDate
                ' �N����
                strSql = strSql & SS & "BirthDate" & SE & "='" & Replace(.Year & .Month & .Day, "'", "''") & "',"
            End With ' BirthDate
            ' �n��
            strSql = strSql & SS & "Bamei" & SE & "='" & Replace(.Bamei, "'", "''") & "',"
            ' �n�����p�J�i
            strSql = strSql & SS & "BameiKana" & SE & "='" & Replace(.BameiKana, "'", "''") & "',"
            ' �n������
            strSql = strSql & SS & "BameiEng" & SE & "='" & Replace(.BameiEng, "'", "''") & "',"
            ' JRA�{�ݍ݂��イ�t���O
            strSql = strSql & SS & "ZaikyuFlag" & SE & "='" & Replace(.ZaikyuFlag, "'", "''") & "',"
            ' �\��
            strSql = strSql & SS & "Reserved" & SE & "='" & Replace(.Reserved, "'", "''") & "',"
            ' �n�L���R�[�h
            strSql = strSql & SS & "UmaKigoCD" & SE & "='" & Replace(.UmaKigoCD, "'", "''") & "',"
            ' ���ʃR�[�h
            strSql = strSql & SS & "SexCD" & SE & "='" & Replace(.SexCD, "'", "''") & "',"
            ' �i��R�[�h
            strSql = strSql & SS & "HinsyuCD" & SE & "='" & Replace(.HinsyuCD, "'", "''") & "',"
            ' �ѐF�R�[�h
            strSql = strSql & SS & "KeiroCD" & SE & "='" & Replace(.KeiroCD, "'", "''") & "',"
            For iLoopCnt1 = 0 To 13
                With .Ketto3Info(iLoopCnt1)
                    ' �ɐB�o�^�ԍ�
                    strSql = strSql & SS & "Ketto3InfoHansyokuNum" & iLoopCnt1 + 1 & "" & SE & "='" & Replace(.HansyokuNum, "'", "''") & "',"
                    ' �n��
                    strSql = strSql & SS & "Ketto3InfoBamei" & iLoopCnt1 + 1 & "" & SE & "='" & Replace(.Bamei, "'", "''") & "',"
                End With ' Ketto3Info
            Next iLoopCnt1
            ' ���������R�[�h
            strSql = strSql & SS & "TozaiCD" & SE & "='" & Replace(.TozaiCD, "'", "''") & "',"
            ' �����t�R�[�h
            strSql = strSql & SS & "ChokyosiCode" & SE & "='" & Replace(.ChokyosiCode, "'", "''") & "',"
            ' �����t������
            strSql = strSql & SS & "ChokyosiRyakusyo" & SE & "='" & Replace(.ChokyosiRyakusyo, "'", "''") & "',"
            ' ���Ғn�於
            strSql = strSql & SS & "Syotai" & SE & "='" & Replace(.Syotai, "'", "''") & "',"
            ' ���Y�҃R�[�h
            strSql = strSql & SS & "BreederCode" & SE & "='" & Replace(.BreederCode, "'", "''") & "',"
            ' ���Y�Җ�
            strSql = strSql & SS & "BreederName" & SE & "='" & Replace(.BreederName, "'", "''") & "',"
            ' �Y�n��
            strSql = strSql & SS & "SanchiName" & SE & "='" & Replace(.SanchiName, "'", "''") & "',"
            ' �n��R�[�h
            strSql = strSql & SS & "BanusiCode" & SE & "='" & Replace(.BanusiCode, "'", "''") & "',"
            ' �n�喼
            strSql = strSql & SS & "BanusiName" & SE & "='" & Replace(.BanusiName, "'", "''") & "',"
            ' ���n�{�܋��݌v
            strSql = strSql & SS & "RuikeiHonsyoHeiti" & SE & "='" & Replace(.RuikeiHonsyoHeiti, "'", "''") & "',"
            ' ��Q�{�܋��݌v
            strSql = strSql & SS & "RuikeiHonsyoSyogai" & SE & "='" & Replace(.RuikeiHonsyoSyogai, "'", "''") & "',"
            ' ���n�t���܋��݌v
            strSql = strSql & SS & "RuikeiFukaHeichi" & SE & "='" & Replace(.RuikeiFukaHeichi, "'", "''") & "',"
            ' ��Q�t���܋��݌v
            strSql = strSql & SS & "RuikeiFukaSyogai" & SE & "='" & Replace(.RuikeiFukaSyogai, "'", "''") & "',"
            ' ���n�����܋��݌v
            strSql = strSql & SS & "RuikeiSyutokuHeichi" & SE & "='" & Replace(.RuikeiSyutokuHeichi, "'", "''") & "',"
            ' ��Q�����܋��݌v
            strSql = strSql & SS & "RuikeiSyutokuSyogai" & SE & "='" & Replace(.RuikeiSyutokuSyogai, "'", "''") & "',"
            With .ChakuSogo
                For iLoopCnt1 = 0 To 5
                    strSql = strSql & SS & "SogoChakukaisu" & iLoopCnt1 + 1 & "" & SE & "='" & Replace(.Chakukaisu(iLoopCnt1), "'", "''") & "',"
                Next iLoopCnt1
            End With ' ChakuSogo
            With .ChakuChuo
                For iLoopCnt1 = 0 To 5
                    strSql = strSql & SS & "ChuoChakukaisu" & iLoopCnt1 + 1 & "" & SE & "='" & Replace(.Chakukaisu(iLoopCnt1), "'", "''") & "',"
                Next iLoopCnt1
            End With ' ChakuChuo
            For iLoopCnt1 = 0 To 6
                With .ChakuKaisuBa(iLoopCnt1)
                    For iLoopCnt2 = 0 To 5
                        strSql = strSql & SS & "Ba" & iLoopCnt1 + 1 & "Chakukaisu" & iLoopCnt2 + 1 & "" & SE & "='" & Replace(.Chakukaisu(iLoopCnt2), "'", "''") & "'"
                        If iLoopCnt1 <> 6 Or iLoopCnt2 <> 5 Then
                            strSql = strSql & ","
                        End If
                    Next iLoopCnt2
                End With ' ChakuKaisuBa
            Next iLoopCnt1

            'strSql = Left(strSql, Len(strSql) - 1) ' Delete Last Comma

            strSql = strSql & " WHERE " & SS & "KettoNum" & SE & "='" & Replace(.KettoNum, "'", "''") & "'"
            strSql = strSql & " AND " & SS & "MakeDate" & SE & "<='" & Replace(strMakeDate, "'", "''") & "'"


            gCon.Execute(strSql)

            '��x�ɍX�V�ł���t�B�[���h������127�܂ł̈� �����X�V�iJET�d�l�j 

            strSql = "UPDATE UMA SET "
            '�w�b�_�̍X�V�͌㔼�̍X�V�ōs��
            With .head
                ' ���R�[�h���
                strSql = strSql & SS & "RecordSpec" & SE & "='" & Replace(.RecordSpec, "'", "''") & "',"
                ' �f�[�^�敪
                strSql = strSql & SS & "DataKubun" & SE & "='" & Replace(.DataKubun, "'", "''") & "',"
                ' �N����
                strSql = strSql & SS & "MakeDate" & SE & "='" & Replace(strMakeDate, "'", "''") & "',"
            End With ' head
            For iLoopCnt1 = 0 To 11
                With .ChakuKaisuJyotai(iLoopCnt1)
                    For iLoopCnt2 = 0 To 5
                        strSql = strSql & SS & "Jyotai" & iLoopCnt1 + 1 & "Chakukaisu" & iLoopCnt2 + 1 & "" & SE & "='" & Replace(.Chakukaisu(iLoopCnt2), "'", "''") & "',"
                    Next iLoopCnt2
                End With ' ChakuKaisuJyotai
            Next iLoopCnt1
            For iLoopCnt1 = 0 To 5
                With .ChakuKaisuKyori(iLoopCnt1)
                    For iLoopCnt2 = 0 To 5
                        strSql = strSql & SS & "Kyori" & iLoopCnt1 + 1 & "Chakukaisu" & iLoopCnt2 + 1 & "" & SE & "='" & Replace(.Chakukaisu(iLoopCnt2), "'", "''") & "',"
                    Next iLoopCnt2
                End With ' ChakuKaisuKyori
            Next iLoopCnt1
            For iLoopCnt1 = 0 To 3
                ' �r���X��
                strSql = strSql & SS & "Kyakusitu" & iLoopCnt1 + 1 & "" & SE & "='" & Replace(.Kyakusitu(iLoopCnt1), "'", "''") & "',"
            Next iLoopCnt1
            ' �o�^���[�X��
            strSql = strSql & SS & "RaceCount" & SE & "='" & Replace(.RaceCount, "'", "''") & "'"
            strSql = strSql & " WHERE " & SS & "KettoNum" & SE & "='" & Replace(.KettoNum, "'", "''") & "'"
            strSql = strSql & " AND " & SS & "MakeDate" & SE & "<= '" & Replace(strMakeDate, "'", "''") & "'"
        End With

        With mBuf
            System.Diagnostics.Debug.WriteLine("UPDATE UMA : " & .KettoNum)
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