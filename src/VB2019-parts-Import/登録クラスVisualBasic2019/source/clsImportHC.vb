Option Strict Off
Option Explicit On
Option Compare Binary
Friend Class clsImportHC
	' @(h) clsReadHC.cls
	' @(s)
	' JVData "HC" �f�[�^�x�[�X�o�^�N���X
	'
	
	Private mBuf As JV_HC_HANRO '' �⓹�����\����
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
        strSql = "SELECT * FROM HANRO"
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
    ' ������    : lBuf - JVData ���ʎq"HC" �̂P�s
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
            mRS.Fields("TresenKubun").Value = .TresenKubun '' �g���Z���敪
            With .ChokyoDate
                mRS.Fields("ChokyoDate").Value = .Year & .Month & .Day '' �N����
            End With ' ChokyoDate
            mRS.Fields("ChokyoTime").Value = .ChokyoTime '' ��������
            mRS.Fields("KettoNum").Value = .KettoNum '' �����o�^�ԍ�
            mRS.Fields("HaronTime4").Value = .HaronTime4 '' 4�n�����^�C�����v(800M-0M)
            mRS.Fields("LapTime4").Value = .LapTime4 '' ���b�v�^�C��(800M-600M)
            mRS.Fields("HaronTime3").Value = .HaronTime3 '' 3�n�����^�C�����v(600M-0M)
            mRS.Fields("LapTime3").Value = .LapTime3 '' ���b�v�^�C��(600M-400M)
            mRS.Fields("HaronTime2").Value = .HaronTime2 '' 2�n�����^�C�����v(400M-0M)
            mRS.Fields("LapTime2").Value = .LapTime2 '' ���b�v�^�C��(400M-200M)
            mRS.Fields("LapTime1").Value = .LapTime1 '' ���b�v�^�C��(200M-0M)
        End With

        With mBuf
            System.Diagnostics.Debug.WriteLine("INSERT HANRO : " & .TresenKubun & .ChokyoDate.Year & .ChokyoDate.Month & .ChokyoDate.Day & .ChokyoTime & .KettoNum)
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

        strSql = "UPDATE HANRO SET "
        With mBuf
            With .head
                strSql = strSql & SS & "RecordSpec" & SE & "='" & Replace(.RecordSpec, "'", "''") & "'," '' ���R�[�h���
                strSql = strSql & SS & "DataKubun" & SE & "='" & Replace(.DataKubun, "'", "''") & "'," '' �f�[�^�敪
                strSql = strSql & SS & "MakeDate" & SE & "= '" & Replace(strMakeDate, "'", "''") & "'," '' �N����
            End With ' head
            strSql = strSql & SS & "TresenKubun" & SE & "='" & Replace(.TresenKubun, "'", "''") & "'," '' �g���Z���敪
            With .ChokyoDate
                strSql = strSql & SS & "ChokyoDate" & SE & "='" & Replace(.Year & .Month & .Day, "'", "''") & "'," '' �N����
            End With ' ChokyoDate
            strSql = strSql & SS & "ChokyoTime" & SE & "='" & Replace(.ChokyoTime, "'", "''") & "'," '' ��������
            strSql = strSql & SS & "KettoNum" & SE & "='" & Replace(.KettoNum, "'", "''") & "'," '' �����o�^�ԍ�
            strSql = strSql & SS & "HaronTime4" & SE & "='" & Replace(.HaronTime4, "'", "''") & "'," '' 4�n�����^�C�����v(800M-0M)
            strSql = strSql & SS & "LapTime4" & SE & "='" & Replace(.LapTime4, "'", "''") & "'," '' ���b�v�^�C��(800M-600M)
            strSql = strSql & SS & "HaronTime3" & SE & "='" & Replace(.HaronTime3, "'", "''") & "'," '' 3�n�����^�C�����v(600M-0M)
            strSql = strSql & SS & "LapTime3" & SE & "='" & Replace(.LapTime3, "'", "''") & "'," '' ���b�v�^�C��(600M-400M)
            strSql = strSql & SS & "HaronTime2" & SE & "='" & Replace(.HaronTime2, "'", "''") & "'," '' 2�n�����^�C�����v(400M-0M)
            strSql = strSql & SS & "LapTime2" & SE & "='" & Replace(.LapTime2, "'", "''") & "'," '' ���b�v�^�C��(400M-200M)
            strSql = strSql & SS & "LapTime1" & SE & "='" & Replace(.LapTime1, "'", "''") & "'," '' ���b�v�^�C��(200M-0M)

            strSql = Left(strSql, Len(strSql) - 1) ' Delete Last Comma

            strSql = strSql & " WHERE " & SS & "TresenKubun" & SE & "='" & Replace(.TresenKubun, "'", "''") & "'"
            With .ChokyoDate
                strSql = strSql & " AND " & SS & "ChokyoDate" & SE & "='" & Replace(.Year & .Month & .Day, "'", "''") & "'"
            End With ' ChokyoDate
            strSql = strSql & " AND " & SS & "ChokyoTime" & SE & "='" & Replace(.ChokyoTime, "'", "''") & "'"
            strSql = strSql & " AND " & SS & "KettoNum" & SE & "='" & Replace(.KettoNum, "'", "''") & "'"
            strSql = strSql & " AND " & SS & "MakeDate" & SE & " <= '" & Replace(strMakeDate, "'", "''") & "'"
        End With ' mBuf

        With mBuf
            System.Diagnostics.Debug.WriteLine("UPDATE HANRO : " & .TresenKubun & .ChokyoDate.Year & .ChokyoDate.Month & .ChokyoDate.Day & .ChokyoTime & .KettoNum)
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