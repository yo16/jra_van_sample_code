Option Strict Off
Option Explicit On
Option Compare Binary
Friend Class clsImportSK
	' @(h) clsReadSK.cls
	' @(s)
	' JVData "SK" �f�[�^�x�[�X�o�^�N���X
	'
	
	Private mBuf As JV_SK_SANKU '' �Y��}�X�^�\����
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
        strSql = "SELECT * FROM SANKU"
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
    ' ������    : lBuf - JVData ���ʎq"SK" �̂P�s
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
            mRS.Fields("KettoNum").Value = .KettoNum '' �����o�^�ԍ�
            With .BirthDate
                mRS.Fields("BirthDate").Value = .Year & .Month & .Day '' �N����
            End With ' BirthDate
            mRS.Fields("SexCD").Value = .SexCD '' ���ʃR�[�h
            mRS.Fields("HinsyuCD").Value = .HinsyuCD '' �i��R�[�h
            mRS.Fields("KeiroCD").Value = .KeiroCD '' �ѐF�R�[�h
            mRS.Fields("SankuMochiKubun").Value = .SankuMochiKubun '' �Y����敪
            mRS.Fields("ImportYear").Value = .ImportYear '' �A���N
            mRS.Fields("BreederCode").Value = .BreederCode '' ���Y�҃R�[�h
            mRS.Fields("SanchiName").Value = .SanchiName '' �Y�n��
            mRS.Fields("FNum").Value = .HansyokuNum(0)
            mRS.Fields("MNum").Value = .HansyokuNum(1)
            mRS.Fields("FFNum").Value = .HansyokuNum(2)
            mRS.Fields("FMNum").Value = .HansyokuNum(3)
            mRS.Fields("MFNum").Value = .HansyokuNum(4)
            mRS.Fields("MMNum").Value = .HansyokuNum(5)
            mRS.Fields("FFFNum").Value = .HansyokuNum(6)
            mRS.Fields("FFMNum").Value = .HansyokuNum(7)
            mRS.Fields("FMFNum").Value = .HansyokuNum(8)
            mRS.Fields("FMMNum").Value = .HansyokuNum(9)
            mRS.Fields("MFFNum").Value = .HansyokuNum(10)
            mRS.Fields("MFMNum").Value = .HansyokuNum(11)
            mRS.Fields("MMFNum").Value = .HansyokuNum(12)
            mRS.Fields("MMMNum").Value = .HansyokuNum(13)
        End With

        With mBuf
            System.Diagnostics.Debug.WriteLine("Insert SANKU : " & .KettoNum)
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

        strSql = "UPDATE SANKU SET "
        With mBuf
            With .head
                strSql = strSql & SS & "RecordSpec" & SE & "='" & Replace(.RecordSpec, "'", "''") & "'," '' ���R�[�h���
                strSql = strSql & SS & "DataKubun" & SE & "='" & Replace(.DataKubun, "'", "''") & "'," '' �f�[�^�敪
                strSql = strSql & SS & "MakeDate" & SE & "='" & Replace(strMakeDate, "'", "''") & "'," '' �N����
            End With ' head
            strSql = strSql & SS & "KettoNum" & SE & "='" & Replace(.KettoNum, "'", "''") & "'," '' �����o�^�ԍ�
            With .BirthDate
                strSql = strSql & SS & "BirthDate" & SE & "='" & Replace(.Year & .Month & .Day, "'", "''") & "'," '' �N����
            End With ' BirthDate
            strSql = strSql & SS & "SexCD" & SE & "='" & Replace(.SexCD, "'", "''") & "'," '' ���ʃR�[�h
            strSql = strSql & SS & "HinsyuCD" & SE & "='" & Replace(.HinsyuCD, "'", "''") & "'," '' �i��R�[�h
            strSql = strSql & SS & "KeiroCD" & SE & "='" & Replace(.KeiroCD, "'", "''") & "'," '' �ѐF�R�[�h
            strSql = strSql & SS & "SankuMochiKubun" & SE & "='" & Replace(.SankuMochiKubun, "'", "''") & "'," '' �Y����敪
            strSql = strSql & SS & "ImportYear" & SE & "='" & Replace(.ImportYear, "'", "''") & "'," '' �A���N
            strSql = strSql & SS & "BreederCode" & SE & "='" & Replace(.BreederCode, "'", "''") & "'," '' ���Y�҃R�[�h
            strSql = strSql & SS & "SanchiName" & SE & "='" & Replace(.SanchiName, "'", "''") & "'," '' �Y�n��
            strSql = strSql & SS & "FNum" & SE & "='" & Replace(.HansyokuNum(0), "'", "''") & "'," '' ���ɐB�o�^�ԍ�
            strSql = strSql & SS & "MNum" & SE & "='" & Replace(.HansyokuNum(1), "'", "''") & "'," '' ��ɐB�o�^�ԍ�
            strSql = strSql & SS & "FFNum" & SE & "='" & Replace(.HansyokuNum(2), "'", "''") & "'," '' �����ɐB�o�^�ԍ�
            strSql = strSql & SS & "FMNum" & SE & "='" & Replace(.HansyokuNum(3), "'", "''") & "'," '' ����ɐB�o�^�ԍ�
            strSql = strSql & SS & "MFNum" & SE & "='" & Replace(.HansyokuNum(4), "'", "''") & "'," '' �ꕃ�ɐB�o�^�ԍ�
            strSql = strSql & SS & "MMNum" & SE & "='" & Replace(.HansyokuNum(5), "'", "''") & "'," '' ���ɐB�o�^�ԍ�
            strSql = strSql & SS & "FFFNum" & SE & "='" & Replace(.HansyokuNum(6), "'", "''") & "'," '' �������ɐB�o�^�ԍ�
            strSql = strSql & SS & "FFMNum" & SE & "='" & Replace(.HansyokuNum(7), "'", "''") & "'," '' ������ɐB�o�^�ԍ�
            strSql = strSql & SS & "FMFNum" & SE & "='" & Replace(.HansyokuNum(8), "'", "''") & "'," '' ���ꕃ�ɐB�o�^�ԍ�
            strSql = strSql & SS & "FMMNum" & SE & "='" & Replace(.HansyokuNum(9), "'", "''") & "'," '' �����ɐB�o�^�ԍ�
            strSql = strSql & SS & "MFFNum" & SE & "='" & Replace(.HansyokuNum(10), "'", "''") & "'," '' �ꕃ���ɐB�o�^�ԍ�
            strSql = strSql & SS & "MFMNum" & SE & "='" & Replace(.HansyokuNum(11), "'", "''") & "'," '' �ꕃ��ɐB�o�^�ԍ�
            strSql = strSql & SS & "MMFNum" & SE & "='" & Replace(.HansyokuNum(12), "'", "''") & "'," '' ��ꕃ�ɐB�o�^�ԍ�
            strSql = strSql & SS & "MMMNum" & SE & "='" & Replace(.HansyokuNum(13), "'", "''") & "'," '' ����ɐB�o�^�ԍ�

            strSql = Left(strSql, Len(strSql) - 1) ' Delete Last Comma

            strSql = strSql & " WHERE " & SS & "KettoNum" & SE & "='" & Replace(.KettoNum, "'", "''") & "'"
            strSql = strSql & " AND " & SS & "MakeDate" & SE & " <= '" & Replace(strMakeDate, "'", "''") & "'"
        End With
        With mBuf
            System.Diagnostics.Debug.WriteLine("UPDATE SANKU : " & .KettoNum)
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