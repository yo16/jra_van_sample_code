Option Strict Off
Option Explicit On
Option Compare Binary
Friend Class clsImportHS
	' @(h) clsReadHS.cls
	' @(s)
	' JVData "HS" �f�[�^�x�[�X�o�^�N���X
	'
	
	Private mBuf As JV_HS_SALE '' �����n�s�������i�\����
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
        strSql = "SELECT * FROM SALE"
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
    ' ������    : lBuf - JVData ���ʎq"HS" �̂P�s
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

        gCon.BeginTrans()
        System.Diagnostics.Debug.WriteLine("BeginTrans")

        mRS.AddNew()

        With mBuf
            With .head
                mRS.Fields("RecordSpec").Value = .RecordSpec             '' ���R�[�h���
                mRS.Fields("DataKubun").Value = .DataKubun               '' �f�[�^�敪
                With .MakeDate
                    mRS.Fields("MakeDate").Value = .Year & .Month & .Day '' �N����
                End With ' MakeDate
            End With ' head
            mRS.Fields("KettoNum").Value = .KettoNum                     ''�����o�^�ԍ�
            mRS.Fields("HansyokuFNum").Value = .HansyokuFNum             ''���n�ɐB�o�^�ԍ�
            mRS.Fields("HansyokuMNum").Value = .HansyokuMNum             ''��n�ɐB�o�^�ԍ�
            mRS.Fields("BirthYear").Value = .BirthYear                   ''���N
            mRS.Fields("SaleCode").Value = .SaleCode                     ''��ÎҁE�s��R�[�h
            mRS.Fields("SaleHostName").Value = .SaleHostName             ''��ÎҖ���
            mRS.Fields("SaleName").Value = .SaleName                     ''�s��̖���
            With .FromDate
                mRS.Fields("FromDate").Value = .Year & .Month & .Day     ''�s��̊J�Ê���(�J�n��)
            End With ' FromDate
            With .ToDate
                mRS.Fields("ToDate").Value = .Year & .Month & .Day       ''�s��̊J�Ê���(�I����)
            End With ' ToDate
            mRS.Fields("Barei").Value = .Barei                           ''������̋����n�̔N��
            mRS.Fields("Price").Value = .Price                           ''������i
        End With

        With mBuf
            System.Diagnostics.Debug.WriteLine("Insert SALE : " & .KettoNum & .SaleCode & .FromDate.Year &  .FromDate.Month &  .FromDate.Day)
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
        Dim strSql As String '' SQL��

        gCon.BeginTrans()
        System.Diagnostics.Debug.WriteLine("BeginTrans")

        strSql = "UPDATE SALE SET "
        With mBuf
            With .head
                strSql = strSql & SS & "RecordSpec" & SE & "='" & Replace(.RecordSpec, "'", "''") & "',"          '' ���R�[�h���
                strSql = strSql & SS & "DataKubun" & SE & "='" & Replace(.DataKubun, "'", "''") & "',"            '' �f�[�^�敪
                strSql = strSql & SS & "MakeDate" & SE & "='" & Replace(strMakeDate, "'", "''") & "',"            '' �N����
            End With ' head
            strSql = strSql & SS & "KettoNum" & SE & "='" & Replace(.KettoNum, "'", "''") & "',"                  ''�����o�^�ԍ�
            strSql = strSql & SS & "HansyokuFNum" & SE & "='" & Replace(.HansyokuFNum, "'", "''") & "',"          ''���n�ɐB�o�^�ԍ�
            strSql = strSql & SS & "HansyokuMNum" & SE & "='" & Replace(.HansyokuMNum, "'", "''") & "',"          ''��n�ɐB�o�^�ԍ�
            strSql = strSql & SS & "BirthYear" & SE & "='" & Replace(.BirthYear, "'", "''") & "',"                ''���N
            strSql = strSql & SS & "SaleCode" & SE & "='" & Replace(.SaleCode, "'", "''") & "',"                  ''��ÎҁE�s��R�[�h
            strSql = strSql & SS & "SaleHostName" & SE & "='" & Replace(.SaleHostName, "'", "''") & "',"          ''��ÎҖ���
            strSql = strSql & SS & "SaleName" & SE & "='" & Replace(.SaleName, "'", "''") & "',"                  ''�s��̖���
            With .FromDate
                strSql = strSql & SS & "FromDate" & SE & "='" & Replace(.Year & .Month & .Day, "'", "''") & "',"  ''�s��̊J�Ê���(�J�n��)
            End With ' FromDate
            With .ToDate
                strSql = strSql & SS & "ToDate" & SE & "='" & Replace(.Year & .Month & .Day, "'", "''") & "',"    ''�s��̊J�Ê���(�I����)
            End With ' ToDate
            strSql = strSql & SS & "Barei" & SE & "='" & Replace(.Barei, "'", "''") & "',"                        '' ������̋����n�̔N��
            strSql = strSql & SS & "Price" & SE & "='" & Replace(.Price, "'", "''") & "',"                        '' ������i

            strSql = Left(strSql, Len(strSql) - 1) ' Delete Last Comma

            strSql = strSql & " WHERE " & SS & "KettoNum" & SE & "='" & Replace(.KettoNum, "'", "''") & "'"                 ''�����o�^�ԍ�
            strSql = strSql & " AND " & SS & "SaleCode" & SE & "='" & Replace(.SaleCode, "'", "''") & "'"                   ''��ÎҁE�s��R�[�h
            With .FromDate
                strSql = strSql & "AND " & SS & "FromDate" & SE & "='" & Replace(.Year & .Month & .Day, "'", "''") & "'"   ''�s��̊J�Ê���(�J�n��)
            End With ' FromDate
            strSql = strSql & " AND " & SS & "MakeDate" & SE & " <= '" & Replace(strMakeDate, "'", "''") & "'"              '' �N����
        End With
        With mBuf
            System.Diagnostics.Debug.WriteLine("UPDATE SALE : " & .KettoNum & .SaleCode & .FromDate.Year &  .FromDate.Month &  .FromDate.Day)
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