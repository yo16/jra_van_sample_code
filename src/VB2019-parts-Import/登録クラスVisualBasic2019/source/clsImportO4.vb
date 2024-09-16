Option Strict Off
Option Explicit On
Option Compare Binary
Friend Class clsImportO4
	' @(h) clsReadO4.cls
	' @(s)
	' JVData "O4" �f�[�^�x�[�X�o�^�N���X
	'
	
	Private mBuf As JV_O4_ODDS_UMATAN '' �I�b�Y�i�n�P�j�\����
	Private mRS1 As ADODB.Recordset
	Private mRS2 As ADODB.Recordset
	
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
        strSql = "SELECT * FROM ODDS_UMATAN_HEAD"
        mRS1 = New ADODB.Recordset()
        mRS1.Open(strSql, gCon, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockPessimistic)

        strSql = "SELECT * FROM ODDS_UMATAN"
        mRS2 = New ADODB.Recordset()
        mRS2.Open(strSql, gCon, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockPessimistic)

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
        mRS1.Close()

        mRS1 = Nothing
        mRS2.Close()

        mRS2 = Nothing

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
    ' ������    : lBuf - JVData ���ʎq"O4" �̂P�s
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

        mRS1.AddNew()

        ' ODDS_UMATAN_HEAD (�I�b�Y_�n�P_�w�b�_)
        With mBuf
            With .head
                mRS1.Fields("RecordSpec").Value = .RecordSpec '' ���R�[�h���
                mRS1.Fields("DataKubun").Value = .DataKubun '' �f�[�^�敪
                With .MakeDate
                    mRS1.Fields("MakeDate").Value = .Year & .Month & .Day '' �N����
                End With ' MakeDate
            End With ' head
            With .id
                mRS1.Fields("Year").Value = .Year '' �J�ÔN
                mRS1.Fields("MonthDay").Value = .MonthDay '' �J�Ì���
                mRS1.Fields("JyoCD").Value = .JyoCD '' ���n��R�[�h
                mRS1.Fields("Kaiji").Value = .Kaiji '' �J�É��N��
                mRS1.Fields("Nichiji").Value = .Nichiji '' �J�Ó���N����
                mRS1.Fields("RaceNum").Value = .RaceNum '' ���[�X�ԍ�
            End With ' id
            With .HappyoTime
                mRS1.Fields("HappyoTime").Value = .Month & .Day & .Hour & .Minute '' ��������
            End With ' HappyoTime
            mRS1.Fields("TorokuTosu").Value = .TorokuTosu '' �o�^����
            mRS1.Fields("SyussoTosu").Value = .SyussoTosu '' �o������
            mRS1.Fields("UmatanFlag").Value = .UmatanFlag '' �����t���O�@�n�P
            mRS1.Fields("TotalHyosuUmatan").Value = .TotalHyosuUmatan '' �n�P�[�����v
        End With

        With mBuf.id
            System.Diagnostics.Debug.WriteLine("Insert ODDS_UMATAN_HEAD : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum)
        End With ' id

        mRS1.Update()


        ' ODDS_UMATAN (�I�b�Y_�n�P)
        If mBuf.UmatanFlag <> "0" Then
            For i = 0 To 305
                If mBuf.OddsUmatanInfo(i).Kumi <> "    " Then
                    mRS2.AddNew()
                    With mBuf
                        With .head
                            With .MakeDate
                                mRS2.Fields("MakeDate").Value = .Year & .Month & .Day '' �N����
                            End With ' MakeDate
                        End With ' head
                        With .id
                            mRS2.Fields("Year").Value = .Year '' �J�ÔN
                            mRS2.Fields("MonthDay").Value = .MonthDay '' �J�Ì���
                            mRS2.Fields("JyoCD").Value = .JyoCD '' ���n��R�[�h
                            mRS2.Fields("Kaiji").Value = .Kaiji '' �J�É��N��
                            mRS2.Fields("Nichiji").Value = .Nichiji '' �J�Ó���N����
                            mRS2.Fields("RaceNum").Value = .RaceNum '' ���[�X�ԍ�
                        End With ' id
                        With .OddsUmatanInfo(i)
                            mRS2.Fields("Kumi").Value = .Kumi '' �g��
                            mRS2.Fields("Odds").Value = .Odds '' �I�b�Y
                            mRS2.Fields("Ninki").Value = .Ninki '' �l�C��
                        End With ' OddsUmatanInfo
                    End With

                    With mBuf.id
                        System.Diagnostics.Debug.WriteLine("Insert ODDS_UMATAN : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum & mBuf.OddsUmatanInfo(i).Kumi)
                    End With ' id

                    mRS2.Update()

                End If
            Next i
        End If

        gCon.CommitTrans()
        System.Diagnostics.Debug.WriteLine("CommitTrans")

        InsertDB = True

ExitHandler:
        Exit Function
ErrorHandler:
        InsertDB = False
        mRS1.CancelUpdate()
        mRS2.CancelUpdate()
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

        ' ODDS_UMATAN_HEAD (�I�b�Y_�n�P_�w�b�_)
        strSql = "UPDATE ODDS_UMATAN_HEAD SET "
        With mBuf
            With .head
                strSql = strSql & SS & "RecordSpec" & SE & "='" & Replace(.RecordSpec, "'", "''") & "'," '' ���R�[�h���
                strSql = strSql & SS & "DataKubun" & SE & "='" & Replace(.DataKubun, "'", "''") & "'," '' �f�[�^�敪
                strSql = strSql & SS & "MakeDate" & SE & "= '" & Replace(strMakeDate, "'", "''") & "'," '' �N����
            End With ' head
            With .id
                strSql = strSql & SS & "Year" & SE & "='" & Replace(.Year, "'", "''") & "'," '' �J�ÔN
                strSql = strSql & SS & "MonthDay" & SE & "='" & Replace(.MonthDay, "'", "''") & "'," '' �J�Ì���
                strSql = strSql & SS & "JyoCD" & SE & "='" & Replace(.JyoCD, "'", "''") & "'," '' ���n��R�[�h
                strSql = strSql & SS & "Kaiji" & SE & "='" & Replace(.Kaiji, "'", "''") & "'," '' �J�É��N��
                strSql = strSql & SS & "Nichiji" & SE & "='" & Replace(.Nichiji, "'", "''") & "'," '' �J�Ó���N����
                strSql = strSql & SS & "RaceNum" & SE & "='" & Replace(.RaceNum, "'", "''") & "'," '' ���[�X�ԍ�
            End With ' id
            With .HappyoTime
                strSql = strSql & SS & "HappyoTime" & SE & "='" & Replace(.Month & .Day & .Hour & .Minute, "'", "''") & "',"
            End With ' HappyoTime
            strSql = strSql & SS & "TorokuTosu" & SE & "='" & Replace(.TorokuTosu, "'", "''") & "'," '' �o�^����
            strSql = strSql & SS & "SyussoTosu" & SE & "='" & Replace(.SyussoTosu, "'", "''") & "'," '' �o������
            strSql = strSql & SS & "UmatanFlag" & SE & "='" & Replace(.UmatanFlag, "'", "''") & "'," '' �����t���O�@�n�P
            strSql = strSql & SS & "TotalHyosuUmatan" & SE & "='" & Replace(.TotalHyosuUmatan, "'", "''") & "'," '' �n�P�[�����v

            strSql = Left(strSql, Len(strSql) - 1) ' Delete Last Comma

            With .id
                strSql = strSql & " WHERE " & SS & "Year" & SE & "='" & Replace(.Year, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "MonthDay" & SE & "='" & Replace(.MonthDay, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "JyoCD" & SE & "='" & Replace(.JyoCD, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "Kaiji" & SE & "='" & Replace(.Kaiji, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "Nichiji" & SE & "='" & Replace(.Nichiji, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "RaceNum" & SE & "='" & Replace(.RaceNum, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "MakeDate" & SE & " <= '" & Replace(strMakeDate, "'", "''") & "'"
            End With
        End With

        With mBuf.id
            System.Diagnostics.Debug.WriteLine("UPDATE ODDS_UMATAN_HEAD : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum)
        End With ' id

        gCon.Execute(strSql)


        ' ODDS_UMATAN (�I�b�Y_�n�P)
        For i = 0 To 305
            strSql = "UPDATE ODDS_UMATAN SET "
            With mBuf
                strSql = strSql & SS & "MakeDate" & SE & "='" & Replace(strMakeDate, "'", "''") & "'," '' �N����
                With .id
                    strSql = strSql & SS & "Year" & SE & "='" & Replace(.Year, "'", "''") & "'," '' �J�ÔN
                    strSql = strSql & SS & "MonthDay" & SE & "='" & Replace(.MonthDay, "'", "''") & "'," '' �J�Ì���
                    strSql = strSql & SS & "JyoCD" & SE & "='" & Replace(.JyoCD, "'", "''") & "'," '' ���n��R�[�h
                    strSql = strSql & SS & "Kaiji" & SE & "='" & Replace(.Kaiji, "'", "''") & "'," '' �J�É��N��
                    strSql = strSql & SS & "Nichiji" & SE & "='" & Replace(.Nichiji, "'", "''") & "'," '' �J�Ó���N����
                    strSql = strSql & SS & "RaceNum" & SE & "='" & Replace(.RaceNum, "'", "''") & "'," '' ���[�X�ԍ�
                End With ' id
                With .OddsUmatanInfo(i)
                    strSql = strSql & SS & "Kumi" & SE & "='" & Replace(.Kumi, "'", "''") & "'," '' �g��
                    strSql = strSql & SS & "Odds" & SE & "='" & Replace(.Odds, "'", "''") & "'," '' �I�b�Y
                    strSql = strSql & SS & "Ninki" & SE & "='" & Replace(.Ninki, "'", "''") & "'," '' �l�C��
                End With ' OddsUmatanInfo

                strSql = Left(strSql, Len(strSql) - 1) ' Delete Last Comma

                With .id
                    strSql = strSql & " WHERE " & SS & "Year" & SE & "='" & Replace(.Year, "'", "''") & "'"
                    strSql = strSql & " AND " & SS & "MonthDay" & SE & "='" & Replace(.MonthDay, "'", "''") & "'"
                    strSql = strSql & " AND " & SS & "JyoCD" & SE & "='" & Replace(.JyoCD, "'", "''") & "'"
                    strSql = strSql & " AND " & SS & "Kaiji" & SE & "='" & Replace(.Kaiji, "'", "''") & "'"
                    strSql = strSql & " AND " & SS & "Nichiji" & SE & "='" & Replace(.Nichiji, "'", "''") & "'"
                    strSql = strSql & " AND " & SS & "RaceNum" & SE & "='" & Replace(.RaceNum, "'", "''") & "'"
                    strSql = strSql & " AND " & SS & "Kumi" & SE & "='" & Replace(mBuf.OddsUmatanInfo(i).Kumi, "'", "''") & "'"
                    strSql = strSql & " AND " & SS & "MakeDate" & SE & " <= '" & Replace(strMakeDate, "'", "''") & "'"
                End With
            End With

            With mBuf.id
                System.Diagnostics.Debug.WriteLine("UPDATE ODDS_UMATAN : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum & mBuf.OddsUmatanInfo(i).Kumi)
            End With ' id

            gCon.Execute(strSql)

        Next i

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