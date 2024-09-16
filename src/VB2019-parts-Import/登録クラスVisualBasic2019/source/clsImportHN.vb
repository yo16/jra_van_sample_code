Option Strict Off
Option Explicit On
Option Compare Binary
Friend Class clsImportHN
	' @(h) clsReadHN.cls
	' @(s)
	' JVData "HN" �f�[�^�x�[�X�o�^�N���X
	'
	
	Private mBuf As JV_HN_HANSYOKU '' �ɐB�n�}�X�^�\����
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
        strSql = "SELECT * FROM HANSYOKU"
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
        System.Diagnostics.Debug.WriteLine("mRS.Close")
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
    ' ������    : lBuf - JVData ���ʎq"HN" �̂P�s
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
            mRS.Fields("HansyokuNum").Value = .HansyokuNum '' �ɐB�o�^�ԍ�
            mRS.Fields("reserved").Value = .reserved '' �\��
            mRS.Fields("KettoNum").Value = .KettoNum '' �����o�^�ԍ�
            mRS.Fields("DelKubun").Value = .DelKubun '' �ɐB�n�����敪(���݂͗\���Ƃ��Ďg�p)
            mRS.Fields("Bamei").Value = .Bamei '' �n��
            mRS.Fields("BameiKana").Value = .BameiKana '' �n�����p�J�i
            mRS.Fields("BameiEng").Value = .BameiEng '' �n������
            mRS.Fields("BirthYear").Value = .BirthYear '' ���N
            mRS.Fields("SexCD").Value = .SexCD '' ���ʃR�[�h
            mRS.Fields("HinsyuCD").Value = .HinsyuCD '' �i��R�[�h
            mRS.Fields("KeiroCD").Value = .KeiroCD '' �ѐF�R�[�h
            mRS.Fields("HansyokuMochiKubun").Value = .HansyokuMochiKubun '' �ɐB�n�����敪
            mRS.Fields("ImportYear").Value = .ImportYear '' �A���N
            mRS.Fields("SanchiName").Value = .SanchiName '' �Y�n��
            mRS.Fields("HansyokuFNum").Value = .HansyokuFNum '' ���n�ɐB�o�^�ԍ�
            mRS.Fields("HansyokuMNum").Value = .HansyokuMNum '' ��n�ɐB�o�^�ԍ�
        End With

        With mBuf
            System.Diagnostics.Debug.WriteLine("Insert HANSYOKU : " & .HansyokuNum)
        End With ' mBuf

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
    Public Function UpdateDB(ByVal strMakeDate As String) As Boolean
        On Error GoTo ErrorHandler
        Dim i As Short '' ���[�v�J�E���^
        Dim j As Short '' ���[�v�J�E���^
        Dim k As Short '' ���[�v�J�E���^
        Dim strSql As String '' SQL��

        gCon.BeginTrans()
        System.Diagnostics.Debug.WriteLine("BeginTrans")

        strSql = "UPDATE HANSYOKU SET "
        With mBuf
            With .head
                strSql = strSql & SS & "RecordSpec" & SE & "='" & Replace(.RecordSpec, "'", "''") & "'," '' ���R�[�h���
                strSql = strSql & SS & "DataKubun" & SE & "='" & Replace(.DataKubun, "'", "''") & "'," '' �f�[�^�敪
                strSql = strSql & SS & "MakeDate" & SE & "= '" & Replace(strMakeDate, "'", "''") & "'," '' �N����
            End With ' head
            strSql = strSql & SS & "HansyokuNum" & SE & "='" & Replace(.HansyokuNum, "'", "''") & "'," '' �ɐB�o�^�ԍ�
            strSql = strSql & SS & "reserved" & SE & "='" & Replace(.reserved, "'", "''") & "'," '' �\��
            strSql = strSql & SS & "KettoNum" & SE & "='" & Replace(.KettoNum, "'", "''") & "'," '' �����o�^�ԍ�
            strSql = strSql & SS & "DelKubun" & SE & "='" & Replace(.DelKubun, "'", "''") & "'," '' �ɐB�n�����敪(���݂͗\���Ƃ��Ďg�p)
            strSql = strSql & SS & "Bamei" & SE & "='" & Replace(.Bamei, "'", "''") & "'," '' �n��
            strSql = strSql & SS & "BameiKana" & SE & "='" & Replace(.BameiKana, "'", "''") & "'," '' �n�����p�J�i
            strSql = strSql & SS & "BameiEng" & SE & "='" & Replace(.BameiEng, "'", "''") & "'," '' �n������
            strSql = strSql & SS & "BirthYear" & SE & "='" & Replace(.BirthYear, "'", "''") & "'," '' ���N
            strSql = strSql & SS & "SexCD" & SE & "='" & Replace(.SexCD, "'", "''") & "'," '' ���ʃR�[�h
            strSql = strSql & SS & "HinsyuCD" & SE & "='" & Replace(.HinsyuCD, "'", "''") & "'," '' �i��R�[�h
            strSql = strSql & SS & "KeiroCD" & SE & "='" & Replace(.KeiroCD, "'", "''") & "'," '' �ѐF�R�[�h
            strSql = strSql & SS & "HansyokuMochiKubun" & SE & "='" & Replace(.HansyokuMochiKubun, "'", "''") & "'," '' �ɐB�n�����敪
            strSql = strSql & SS & "ImportYear" & SE & "='" & Replace(.ImportYear, "'", "''") & "'," '' �A���N
            strSql = strSql & SS & "SanchiName" & SE & "='" & Replace(.SanchiName, "'", "''") & "'," '' �Y�n��
            strSql = strSql & SS & "HansyokuFNum" & SE & "='" & Replace(.HansyokuFNum, "'", "''") & "'," '' ���n�ɐB�o�^�ԍ�
            strSql = strSql & SS & "HansyokuMNum" & SE & "='" & Replace(.HansyokuMNum, "'", "''") & "'," '' ��n�ɐB�o�^�ԍ�

            strSql = Left(strSql, Len(strSql) - 1) ' Delete Last Comma

            strSql = strSql & " WHERE " & SS & "HansyokuNum" & SE & "='" & Replace(.HansyokuNum, "'", "''") & "'"
            strSql = strSql & " AND " & SS & "MakeDate" & SE & " <= '" & Replace(strMakeDate, "'", "''") & "'"
        End With

        With mBuf
            System.Diagnostics.Debug.WriteLine("UPDATE HANSYOKU : " & .HansyokuNum)
        End With ' mBuf

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