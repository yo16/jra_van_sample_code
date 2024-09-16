Option Strict Off
Option Explicit On
Option Compare Binary
Friend Class clsImportWE
	' @(h) clsReadWE.cls
	' @(s)
	' JVData "WE" �f�[�^�x�[�X�o�^�N���X
	'
	
	Private mBuf As JV_WE_WEATHER ''�V��n����
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
        strSql = "SELECT * FROM TENKO_BABA"
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
    ' ������    : lBuf - JVData ���ʎq"WE" �̂P�s
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
            End With ' id
            With .HappyoTime
                mRS.Fields("HappyoTime").Value = .Month & .Day & .Hour & .Minute '' ��������
            End With ' HappyoTime
            mRS.Fields("HenkoID").Value = .HenkoID '' �ύX����
            With .TenkoBaba
                mRS.Fields("AtoTenkoCD").Value = .TenkoCD '' �V��R�[�h
                mRS.Fields("AtoSibaBabaCD").Value = .SibaBabaCD '' �Ŕn���ԃR�[�h
                mRS.Fields("AtoDirtBabaCD").Value = .DirtBabaCD '' �_�[�g�n���ԃR�[�h
            End With ' TenkoBaba
            With .TenkoBabaBefore
                mRS.Fields("MaeTenkoCD").Value = .TenkoCD '' �V��R�[�h
                mRS.Fields("MaeSibaBabaCD").Value = .SibaBabaCD '' �Ŕn���ԃR�[�h
                mRS.Fields("MaeDirtBabaCD").Value = .DirtBabaCD '' �_�[�g�n���ԃR�[�h
            End With ' TenkoBabaBefore
        End With

        With mBuf.id
            System.Diagnostics.Debug.WriteLine("Insert JV_WE_WEATHER : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & mBuf.HappyoTime.Month & mBuf.HappyoTime.Day & mBuf.HappyoTime.Hour & mBuf.HappyoTime.Minute & mBuf.HenkoID)
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
    Public Function UpdateDB(ByVal strMakeDate As String) As Boolean
        On Error GoTo ErrorHandler
        Dim i As Short '' ���[�v�J�E���^
        Dim j As Short '' ���[�v�J�E���^
        Dim k As Short '' ���[�v�J�E���^
        Dim strSql As String '' SQL��

        gCon.BeginTrans()
        System.Diagnostics.Debug.WriteLine("BeginTrans")

        strSql = "UPDATE TENKO_BABA SET "
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
            End With ' id
            With .HappyoTime
                strSql = strSql & SS & "HappyoTime" & SE & "='" & Replace(.Month & .Day & .Hour & .Minute, "'", "''") & "',"
            End With ' HappyoTime
            strSql = strSql & SS & "HenkoID" & SE & "='" & Replace(.HenkoID, "'", "''") & "'," '' �ύX����
            With .TenkoBaba
                strSql = strSql & SS & "AtoTenkoCD" & SE & "='" & Replace(.TenkoCD, "'", "''") & "'," '' �V��R�[�h
                strSql = strSql & SS & "AtoSibaBabaCD" & SE & "='" & Replace(.SibaBabaCD, "'", "''") & "'," '' �Ŕn���ԃR�[�h
                strSql = strSql & SS & "AtoDirtBabaCD" & SE & "='" & Replace(.DirtBabaCD, "'", "''") & "'," '' �_�[�g�n���ԃR�[�h
            End With ' TenkoBaba
            With .TenkoBabaBefore
                strSql = strSql & SS & "MaeTenkoCD" & SE & "='" & Replace(.TenkoCD, "'", "''") & "'," '' �V��R�[�h
                strSql = strSql & SS & "MaeSibaBabaCD" & SE & "='" & Replace(.SibaBabaCD, "'", "''") & "'," '' �Ŕn���ԃR�[�h
                strSql = strSql & SS & "MaeDirtBabaCD" & SE & "='" & Replace(.DirtBabaCD, "'", "''") & "'" '' �_�[�g�n���ԃR�[�h
            End With ' TenkoBabaBefore
            With .id
                strSql = strSql & " WHERE " & SS & "Year" & SE & "='" & Replace(.Year, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "MonthDay" & SE & "='" & Replace(.MonthDay, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "JyoCD" & SE & "='" & Replace(.JyoCD, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "Kaiji" & SE & "='" & Replace(.Kaiji, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "Nichiji" & SE & "='" & Replace(.Nichiji, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "HappyoTime" & SE & "='" & Replace(mBuf.HappyoTime.Month & mBuf.HappyoTime.Day & mBuf.HappyoTime.Hour & mBuf.HappyoTime.Minute, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "HenkoID" & SE & "='" & Replace(mBuf.HenkoID, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "MakeDate" & SE & " <= '" & Replace(strMakeDate, "'", "''") & "'"
            End With
        End With

        gCon.Execute(strSql)

        With mBuf.id
            System.Diagnostics.Debug.WriteLine("UPDATE JV_WE_WEATHER : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & mBuf.HappyoTime.Month & mBuf.HappyoTime.Day & mBuf.HappyoTime.Hour & mBuf.HappyoTime.Minute & mBuf.HenkoID)
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