Option Strict Off
Option Explicit On
Option Compare Binary
Friend Class clsImportUM
	' @(h) clsReadUM.cls
	' @(s)
	' JVData "UM" �f�[�^�x�[�X�o�^�N���X
	'
	
	Private mBuf As JV_UM_UMA ''�����n�}�X�^�\����
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
        strSql = "SELECT * FROM UMA"
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
    ' ������    : lBuf - JVData ���ʎq"UM" �̂P�s
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


    '
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
    Private Function InsertDB() As Boolean
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
            mRS.Fields("KettoNum").Value = .KettoNum '' �����o�^�ԍ�
            mRS.Fields("DelKubun").Value = .DelKubun '' �����n�����敪
            With .RegDate
                mRS.Fields("RegDate").Value = .Year & .Month & .Day '' �N����
            End With ' RegDate
            With .DelDate
                mRS.Fields("DelDate").Value = .Year & .Month & .Day '' �N����
            End With ' DelDate
            With .BirthDate
                mRS.Fields("BirthDate").Value = .Year & .Month & .Day '' �N����
            End With ' BirthDate
            mRS.Fields("Bamei").Value = .Bamei '' �n��
            mRS.Fields("BameiKana").Value = .BameiKana '' �n�����p�J�i
            mRS.Fields("BameiEng").Value = .BameiEng '' �n������
            mRS.Fields("ZaikyuFlag").Value = .ZaikyuFlag '' JRA�{�ݍ݂��イ�t���O
            mRS.Fields("Reserved").Value = .Reserved '' �\��
            mRS.Fields("UmaKigoCD").Value = .UmaKigoCD '' �n�L���R�[�h
            mRS.Fields("SexCD").Value = .SexCD '' ���ʃR�[�h
            mRS.Fields("HinsyuCD").Value = .HinsyuCD '' �i��R�[�h
            mRS.Fields("KeiroCD").Value = .KeiroCD '' �ѐF�R�[�h
            For i = 0 To 13
                With .Ketto3Info(i)
                    mRS.Fields("Ketto3InfoHansyokuNum" & i + 1).Value = .HansyokuNum '' �ɐB�o�^�ԍ�
                    mRS.Fields("Ketto3InfoBamei" & i + 1).Value = .Bamei '' �n��
                End With ' Ketto3Info
            Next i
            mRS.Fields("TozaiCD").Value = .TozaiCD '' ���������R�[�h
            mRS.Fields("ChokyosiCode").Value = .ChokyosiCode '' �����t�R�[�h
            mRS.Fields("ChokyosiRyakusyo").Value = .ChokyosiRyakusyo '' �����t������
            mRS.Fields("Syotai").Value = .Syotai '' ���Ғn�於
            mRS.Fields("BreederCode").Value = .BreederCode '' ���Y�҃R�[�h
            mRS.Fields("BreederName").Value = .BreederName '' ���Y�Җ�
            mRS.Fields("SanchiName").Value = .SanchiName '' �Y�n��
            mRS.Fields("BanusiCode").Value = .BanusiCode '' �n��R�[�h
            mRS.Fields("BanusiName").Value = .BanusiName '' �n�喼
            mRS.Fields("RuikeiHonsyoHeiti").Value = .RuikeiHonsyoHeiti '' ���n�{�܋��݌v
            mRS.Fields("RuikeiHonsyoSyogai").Value = .RuikeiHonsyoSyogai '' ��Q�{�܋��݌v
            mRS.Fields("RuikeiFukaHeichi").Value = .RuikeiFukaHeichi '' ���n�t���܋��݌v
            mRS.Fields("RuikeiFukaSyogai").Value = .RuikeiFukaSyogai '' ��Q�t���܋��݌v
            mRS.Fields("RuikeiSyutokuHeichi").Value = .RuikeiSyutokuHeichi '' ���n�����܋��݌v
            mRS.Fields("RuikeiSyutokuSyogai").Value = .RuikeiSyutokuSyogai '' ��Q�����܋��݌v
            With .ChakuSogo
                For j = 0 To 5
                    mRS.Fields("SogoChakukaisu" & j + 1).Value = .Chakukaisu(j)
                Next j
            End With ' ChakuSogo
            With .ChakuChuo
                For j = 0 To 5
                    mRS.Fields("ChuoChakukaisu" & j + 1).Value = .Chakukaisu(j)
                Next j
            End With ' ChakuChuo
            For i = 0 To 6
                With .ChakuKaisuBa(i)
                    For j = 0 To 5
                        mRS.Fields("Ba" & i + 1 & "Chakukaisu" & j + 1).Value = .Chakukaisu(j)
                    Next j
                End With ' ChakuKaisuBa
            Next i
            For i = 0 To 11
                With .ChakuKaisuJyotai(i)
                    For j = 0 To 5
                        mRS.Fields("Jyotai" & i + 1 & "Chakukaisu" & j + 1).Value = .Chakukaisu(j)
                    Next j
                End With ' ChakuKaisuJyotai
            Next i
            For i = 0 To 5
                With .ChakuKaisuKyori(i)
                    For j = 0 To 5
                        mRS.Fields("Kyori" & i + 1 & "Chakukaisu" & j + 1).Value = .Chakukaisu(j)
                    Next j
                End With ' ChakuKaisuKyori
            Next i
            For i = 0 To 3
                mRS.Fields("Kyakusitu" & i + 1).Value = .Kyakusitu(i) '' �r���X��
            Next i
            mRS.Fields("RaceCount").Value = .RaceCount '' �o�^���[�X��
        End With

        With mBuf
            System.Diagnostics.Debug.WriteLine("Insert UMA : " & .KettoNum)
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

        With mBuf
            strSql = "UPDATE UMA SET "
            strSql = strSql & SS & "KettoNum" & SE & "='" & Replace(.KettoNum, "'", "''") & "'," '' �����o�^�ԍ�
            strSql = strSql & SS & "DelKubun" & SE & "='" & Replace(.DelKubun, "'", "''") & "'," '' �����n�����敪
            With .RegDate
                strSql = strSql & SS & "RegDate" & SE & "='" & Replace(.Year & .Month & .Day, "'", "''") & "'," '' �N����
            End With ' RegDate
            With .DelDate
                strSql = strSql & SS & "DelDate" & SE & "='" & Replace(.Year & .Month & .Day, "'", "''") & "'," '' �N����
            End With ' DelDate
            With .BirthDate
                strSql = strSql & SS & "BirthDate" & SE & "='" & Replace(.Year & .Month & .Day, "'", "''") & "'," '' �N����
            End With ' BirthDate
            strSql = strSql & SS & "Bamei" & SE & "='" & Replace(.Bamei, "'", "''") & "'," '' �n��
            strSql = strSql & SS & "BameiKana" & SE & "='" & Replace(.BameiKana, "'", "''") & "'," '' �n�����p�J�i
            strSql = strSql & SS & "BameiEng" & SE & "='" & Replace(.BameiEng, "'", "''") & "'," '' �n������
            strSql = strSql & SS & "ZaikyuFlag" & SE & "='" & Replace(.ZaikyuFlag, "'", "''") & "'," '' JRA�{�ݍ݂��イ�t���O
            strSql = strSql & SS & "Reserved" & SE & "='" & Replace(.Reserved, "'", "''") & "'," '' �\��
            strSql = strSql & SS & "UmaKigoCD" & SE & "='" & Replace(.UmaKigoCD, "'", "''") & "'," '' �n�L���R�[�h
            strSql = strSql & SS & "SexCD" & SE & "='" & Replace(.SexCD, "'", "''") & "'," '' ���ʃR�[�h
            strSql = strSql & SS & "HinsyuCD" & SE & "='" & Replace(.HinsyuCD, "'", "''") & "'," '' �i��R�[�h
            strSql = strSql & SS & "KeiroCD" & SE & "='" & Replace(.KeiroCD, "'", "''") & "'," '' �ѐF�R�[�h
            For i = 0 To 13
                With .Ketto3Info(i)
                    strSql = strSql & SS & "Ketto3InfoHansyokuNum" & i + 1 & SE & "='" & Replace(.HansyokuNum, "'", "''") & "'," '' �ɐB�o�^�ԍ�
                    strSql = strSql & SS & "Ketto3InfoBamei" & i + 1 & SE & "='" & Replace(.Bamei, "'", "''") & "'," '' �n��
                End With ' Ketto3Info
            Next i
            strSql = strSql & SS & "TozaiCD" & SE & "='" & Replace(.TozaiCD, "'", "''") & "'," '' ���������R�[�h
            strSql = strSql & SS & "ChokyosiCode" & SE & "='" & Replace(.ChokyosiCode, "'", "''") & "'," '' �����t�R�[�h
            strSql = strSql & SS & "ChokyosiRyakusyo" & SE & "='" & Replace(.ChokyosiRyakusyo, "'", "''") & "'," '' �����t������
            strSql = strSql & SS & "Syotai" & SE & "='" & Replace(.Syotai, "'", "''") & "'," '' ���Ғn�於
            strSql = strSql & SS & "BreederCode" & SE & "='" & Replace(.BreederCode, "'", "''") & "'," '' ���Y�҃R�[�h
            strSql = strSql & SS & "BreederName" & SE & "='" & Replace(.BreederName, "'", "''") & "'," '' ���Y�Җ�
            strSql = strSql & SS & "SanchiName" & SE & "='" & Replace(.SanchiName, "'", "''") & "'," '' �Y�n��
            strSql = strSql & SS & "BanusiCode" & SE & "='" & Replace(.BanusiCode, "'", "''") & "'," '' �n��R�[�h
            strSql = strSql & SS & "BanusiName" & SE & "='" & Replace(.BanusiName, "'", "''") & "'," '' �n�喼
            strSql = strSql & SS & "RuikeiHonsyoHeiti" & SE & "='" & Replace(.RuikeiHonsyoHeiti, "'", "''") & "'," '' ���n�{�܋��݌v
            strSql = strSql & SS & "RuikeiHonsyoSyogai" & SE & "='" & Replace(.RuikeiHonsyoSyogai, "'", "''") & "'," '' ��Q�{�܋��݌v
            strSql = strSql & SS & "RuikeiFukaHeichi" & SE & "='" & Replace(.RuikeiFukaHeichi, "'", "''") & "'," '' ���n�t���܋��݌v
            strSql = strSql & SS & "RuikeiFukaSyogai" & SE & "='" & Replace(.RuikeiFukaSyogai, "'", "''") & "'," '' ��Q�t���܋��݌v
            strSql = strSql & SS & "RuikeiSyutokuHeichi" & SE & "='" & Replace(.RuikeiSyutokuHeichi, "'", "''") & "'," '' ���n�����܋��݌v
            strSql = strSql & SS & "RuikeiSyutokuSyogai" & SE & "='" & Replace(.RuikeiSyutokuSyogai, "'", "''") & "'," '' ��Q�����܋��݌v
            With .ChakuSogo
                For j = 0 To 5
                    strSql = strSql & SS & "SogoChakukaisu" & j + 1 & SE & "='" & Replace(.Chakukaisu(j), "'", "''") & "',"
                Next j
            End With ' ChakuSogo
            With .ChakuChuo
                For j = 0 To 5
                    strSql = strSql & SS & "ChuoChakukaisu" & j + 1 & SE & "='" & Replace(.Chakukaisu(j), "'", "''") & "',"
                Next j
            End With ' ChakuChuo
            For i = 0 To 6
                With .ChakuKaisuBa(i)
                    For j = 0 To 5
                        strSql = strSql & SS & "Ba" & i + 1 & "Chakukaisu" & j + 1 & SE & "='" & Replace(.Chakukaisu(j), "'", "''") & "'"
                        If i <> 6 Or j <> 5 Then
                            strSql = strSql & ","
                        End If
                    Next j
                End With ' ChakuKaisuBa
            Next i

            'strSql = Left(strSql, Len(strSql) - 1) ' Delete Last Comma

            strSql = strSql & " WHERE " & SS & "KettoNum" & SE & " ='" & Replace(.KettoNum, "'", "''") & "'"
            strSql = strSql & " AND " & SS & "MakeDate" & SE & " <='" & Replace(strMakeDate, "'", "''") & "'"


            gCon.Execute(strSql)

            ''��x�ɍX�V�ł���t�B�[���h������127�܂ł̈� �����X�V�iJET�d�l�j
            strSql = "UPDATE UMA SET "
            '�w�b�_�̍X�V�͌㔼�̍X�V�ōs��
            With .head
                strSql = strSql & SS & "RecordSpec" & SE & "='" & Replace(.RecordSpec, "'", "''") & "'," '' ���R�[�h���
                strSql = strSql & SS & "DataKubun" & SE & "='" & Replace(.DataKubun, "'", "''") & "'," '' �f�[�^�敪
                strSql = strSql & SS & "MakeDate" & SE & "='" & Replace(strMakeDate, "'", "''") & "'," '' �N����
            End With ' head
            For i = 0 To 11
                With .ChakuKaisuJyotai(i)
                    For j = 0 To 5
                        strSql = strSql & SS & "Jyotai" & i + 1 & "Chakukaisu" & j + 1 & SE & "='" & Replace(.Chakukaisu(j), "'", "''") & "',"
                    Next j
                End With ' ChakuKaisuJyotai
            Next i
            For i = 0 To 5
                With .ChakuKaisuKyori(i)
                    For j = 0 To 5
                        strSql = strSql & SS & "Kyori" & i + 1 & "Chakukaisu" & j + 1 & SE & "='" & Replace(.Chakukaisu(j), "'", "''") & "',"
                    Next j
                End With ' ChakuKaisuKyori
            Next i
            For i = 0 To 3
                strSql = strSql & SS & "Kyakusitu" & i + 1 & SE & "='" & Replace(.Kyakusitu(i), "'", "''") & "'," '' �r���X��
            Next i
            strSql = strSql & SS & "RaceCount" & SE & "='" & Replace(.RaceCount, "'", "''") & "'" '' �o�^���[�X��
            strSql = strSql & " WHERE " & SS & "KettoNum" & SE & "='" & Replace(.KettoNum, "'", "''") & "'"
            strSql = strSql & " AND " & SS & "MakeDate" & SE & " <= '" & Replace(strMakeDate, "'", "''") & "'"


            gCon.Execute(strSql)

        End With

        With mBuf
            System.Diagnostics.Debug.WriteLine("UPDATE UMA : " & .KettoNum)
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