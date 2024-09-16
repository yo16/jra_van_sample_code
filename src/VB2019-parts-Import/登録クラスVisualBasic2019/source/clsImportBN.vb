Option Strict Off
Option Explicit On
Option Compare Binary
Friend Class clsImportBN
	' @(h) clsReadBN.cls
	' @(s)
	' JVData "BN" �f�[�^�x�[�X�o�^�N���X
	'
	
	Private mBuf As JV_BN_BANUSI '' �n��}�X�^�\����
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
        strSql = "SELECT * FROM BANUSI"
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
    ' ������    : lBuf - JVData ���ʎq"BN" �̂P�s
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
        Dim s1 As String = "" '' �擪������

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
            mRS.Fields("BanusiCode").Value = .BanusiCode '' �n��R�[�h
            mRS.Fields("BanusiName_Co").Value = .BanusiName_Co '' �n�喼�i�@�l�i�L�j
            mRS.Fields("BanusiName").Value = .BanusiName '' �n�喼�i�@�l�i���j
            mRS.Fields("BanusiNameKana").Value = .BanusiNameKana '' �n�喼���p�J�i
            mRS.Fields("BanusiNameEng").Value = .BanusiNameEng '' �n�喼����
            mRS.Fields("Fukusyoku").Value = .Fukusyoku '' ���F�W��
            For i = 0 To 1
                With .HonRuikei(i)
                    If i = 0 Then s1 = "H"
                    If i = 1 Then s1 = "R"

                    mRS.Fields(s1 & "_SetYear").Value = .SetYear '' �ݒ�N

                    mRS.Fields(s1 & "_HonSyokinTotal").Value = .HonSyokinTotal '' �{�܋����v

                    mRS.Fields(s1 & "_FukaSyokin").Value = .FukaSyokin '' �t���܋����v
                    For j = 0 To 5

                        mRS.Fields(s1 & "_Chakukaisu" & j + 1).Value = .ChakuKaisu(j) '' ����
                    Next j
                End With ' HonRuikei
            Next i
        End With

        mRS.Update()

        With mBuf
            System.Diagnostics.Debug.WriteLine("INSERT BANUSI : " & .BanusiCode)
        End With ' mBuf

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

        strSql = "UPDATE BANUSI SET "
        With mBuf
            With .head

                strSql = strSql & SS & "RecordSpec" & SE & "='" & Replace(.RecordSpec, "'", "''") & "'," '' ���R�[�h���

                strSql = strSql & SS & "DataKubun" & SE & "='" & Replace(.DataKubun, "'", "''") & "'," '' �f�[�^�敪
                strSql = strSql & SS & "MakeDate" & SE & "= '" & Replace(strMakeDate, "'", "''") & "'," '' �N����
            End With ' head
            strSql = strSql & SS & "BanusiCode" & SE & "='" & Replace(.BanusiCode, "'", "''") & "'," '' �n��R�[�h
            strSql = strSql & SS & "BanusiName_Co" & SE & "='" & Replace(.BanusiName_Co, "'", "''") & "'," '' �n�喼(�@�l�i�L)
            strSql = strSql & SS & "BanusiName" & SE & "='" & Replace(.BanusiName, "'", "''") & "'," '' �n�喼(�@�l�i��)
            strSql = strSql & SS & "BanusiNameKana" & SE & "='" & Replace(.BanusiNameKana, "'", "''") & "'," '' �n�喼���p�J�i
            strSql = strSql & SS & "BanusiNameEng" & SE & "='" & Replace(.BanusiNameEng, "'", "''") & "'," '' �n�喼����
            strSql = strSql & SS & "Fukusyoku" & SE & "='" & Replace(.Fukusyoku, "'", "''") & "'," '' ���F�W��
            With .HonRuikei(0)
                strSql = strSql & SS & "H_SetYear" & SE & "='" & Replace(.SetYear, "'", "''") & "'," '' �ݒ�N
                strSql = strSql & SS & "H_HonSyokinTotal" & SE & "='" & Replace(.HonSyokinTotal, "'", "''") & "'," '' �{�܋����v
                strSql = strSql & SS & "H_Fukasyokin" & SE & "='" & Replace(.FukaSyokin, "'", "''") & "'," '' �t���܋����v
                For j = 0 To 5
                    strSql = strSql & SS & "H_Chakukaisu" & j + 1 & SE & "='" & Replace(.ChakuKaisu(j), "'", "''") & "'," '' ����
                Next j
            End With ' HonRuikei(0)
            With .HonRuikei(1)
                strSql = strSql & SS & "R_SetYear" & SE & "='" & Replace(.SetYear, "'", "''") & "'," '' �ݒ�N
                strSql = strSql & SS & "R_HonSyokinTotal" & SE & "='" & Replace(.HonSyokinTotal, "'", "''") & "'," '' �{�܋����v
                strSql = strSql & SS & "R_Fukasyokin" & SE & "='" & Replace(.FukaSyokin, "'", "''") & "'," '' �t���܋����v
                For j = 0 To 5
                    strSql = strSql & SS & "R_Chakukaisu" & j + 1 & SE & "='" & Replace(.ChakuKaisu(j), "'", "''") & "'," '' ����
                Next j
            End With ' HonRuikei(1)
            strSql = Left(strSql, Len(strSql) - 1) ' Delete Last Comma
            strSql = strSql & " WHERE " & SS & "BanusiCode" & SE & "='" & Replace(.BanusiCode, "'", "''") & "'"
            strSql = strSql & " AND " & SS & "MakeDate" & SE & " <= '" & Replace(strMakeDate, "'", "''") & "'"
        End With ' mBuf

        With mBuf
            System.Diagnostics.Debug.WriteLine("UPDATE BANUSI : " & .BanusiCode)
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