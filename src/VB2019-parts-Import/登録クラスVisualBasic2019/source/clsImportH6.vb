Option Strict Off
Option Explicit On
Option Compare Binary
Friend Class clsImportH6
    ' @(h) clsReadH1.cls
    ' @(s)
    ' JVData "H6" �f�[�^�x�[�X�o�^�N���X
    '

    Private mBuf As JV_H6_HYOSU_SANRENTAN '' �[���i�R�A�P�j�\����
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
        strSql = "SELECT * FROM HYOSU2"
        mRS1 = New ADODB.Recordset()
        mRS1.Open(strSql, gCon, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockPessimistic)

        strSql = "SELECT * FROM HYOSU_SANRENTAN"
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
    ' ������    : lBuf - JVData ���ʎq"H6" �̂P�s
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

        ' HYOSU2 (�[��2)
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
            mRS1.Fields("TorokuTosu").Value = .TorokuTosu '' �o�^����
            mRS1.Fields("SyussoTosu").Value = .SyussoTosu '' �o������
            mRS1.Fields("HatubaiFlag1").Value = .HatubaiFlag '' �����t���O
            For i = 0 To 17
                mRS1.Fields("HenkanUma" & i + 1).Value = .HenkanUma(i) '' �ԊҔn�ԏ��(�n��01�`18)
            Next i
            For i = 0 To 1
                mRS1.Fields("HyoTotal" & i + 1).Value = .HyoTotal(i) '' �[�����v
            Next i
        End With ' mBuf

        With mBuf.id
            System.Diagnostics.Debug.WriteLine("Insert HYOSU2 : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum)
        End With ' id

        mRS1.Update()


        ' HYOSU_SANRENTAN (�[��_�O�A�P)
            For i = 0 To 4895
                If mBuf.HyoSanrentan(i).Kumi <> "      " Then
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
                        With .HyoSanrentan(i)
                            mRS2.Fields("Kumi").Value = .Kumi '' �g��
                            mRS2.Fields("Hyo").Value = .Hyo '' �[��
                            mRS2.Fields("Ninki").Value = .Ninki '' �l�C
                        End With ' HyoSanrenpuku
                    End With ' mBuf

                    With mBuf.id
                        System.Diagnostics.Debug.WriteLine("Insert HYOSU_SANRENTAN : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum & mBuf.HyoSanrentan(i).Kumi)
                    End With ' id
                    mRS2.Update()
                End If
            Next i

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
    Public Function UpdateDB(ByVal strMakeDate As String) As Boolean
        On Error GoTo ErrorHandler
        Dim i As Short '' ���[�v�J�E���^
        Dim j As Short '' ���[�v�J�E���^
        Dim k As Short '' ���[�v�J�E���^
        Dim strSql As String '' SQL��

        gCon.BeginTrans()
        System.Diagnostics.Debug.WriteLine("BeginTrans")

        ' HYOSU2 (�[��2)
        strSql = "UPDATE HYOSU2 SET "
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
            strSql = strSql & SS & "TorokuTosu" & SE & "='" & Replace(.TorokuTosu, "'", "''") & "'," '' �o�^����
            strSql = strSql & SS & "SyussoTosu" & SE & "='" & Replace(.SyussoTosu, "'", "''") & "'," '' �o������
            strSql = strSql & SS & "HatubaiFlag1" & SE & "='" & Replace(.HatubaiFlag, "'", "''") & "'," '' �����t���O
            For i = 0 To 17
                strSql = strSql & SS & "HenkanUma" & i + 1 & SE & "='" & Replace(.HenkanUma(i), "'", "''") & "'," '' �ԊҔn�ԏ��(�n��01�`18)
            Next i
            For i = 0 To 1
                strSql = strSql & SS & "HyoTotal" & i + 1 & SE & "='" & Replace(.HyoTotal(i), "'", "''") & "'," '' �[�����v
            Next i

            strSql = Left(strSql, Len(strSql) - 1) ' Delete Last Comma

            With .id
                strSql = strSql & " WHERE " & SS & "Year" & SE & "='" & Replace(.Year, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "MonthDay" & SE & "='" & Replace(.MonthDay, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "JyoCD" & SE & "='" & Replace(.JyoCD, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "Kaiji" & SE & "='" & Replace(.Kaiji, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "Nichiji" & SE & "='" & Replace(.Nichiji, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "RaceNum" & SE & "='" & Replace(.RaceNum, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "MakeDate" & SE & " <= '" & Replace(strMakeDate, "'", "''") & "'"
            End With ' id
        End With ' mBuf

        With mBuf.id
            System.Diagnostics.Debug.WriteLine("UPDATE HYOSU2 : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum)
        End With ' id

        gCon.Execute(strSql)

        ' HYOSU_SANRENTAN (�[��_�O�A�P)
        For i = 0 To 4895
            strSql = "UPDATE HYOSU_SANRENTAN SET "
            With mBuf
                strSql = strSql & SS & "MakeDate" & SE & "= '" & Replace(strMakeDate, "'", "''") & "'," '' �N����
                With .id
                    strSql = strSql & SS & "Year" & SE & "='" & Replace(.Year, "'", "''") & "'," '' �J�ÔN
                    strSql = strSql & SS & "MonthDay" & SE & "='" & Replace(.MonthDay, "'", "''") & "'," '' �J�Ì���
                    strSql = strSql & SS & "JyoCD" & SE & "='" & Replace(.JyoCD, "'", "''") & "'," '' ���n��R�[�h
                    strSql = strSql & SS & "Kaiji" & SE & "='" & Replace(.Kaiji, "'", "''") & "'," '' �J�É��N��
                    strSql = strSql & SS & "Nichiji" & SE & "='" & Replace(.Nichiji, "'", "''") & "'," '' �J�Ó���N����
                    strSql = strSql & SS & "RaceNum" & SE & "='" & Replace(.RaceNum, "'", "''") & "'," '' ���[�X�ԍ�
                End With ' id
                With .HyoSanrentan(i)
                    strSql = strSql & SS & "Kumi" & SE & "='" & Replace(.Kumi, "'", "''") & "'," '' �g��
                    strSql = strSql & SS & "Hyo" & SE & "='" & Replace(.Hyo, "'", "''") & "'," '' �[��
                    strSql = strSql & SS & "Ninki" & SE & "='" & Replace(.Ninki, "'", "''") & "'," '' �l�C
                End With ' HyoSanrenpuku

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
                strSql = strSql & " AND " & SS & "Kumi" & SE & "='" & Replace(.HyoSanrentan(i).Kumi, "'", "''") & "'"
            End With ' mBuf

            With mBuf.id
                System.Diagnostics.Debug.WriteLine("UPDATE HYOSU_SANRENTAN : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum & mBuf.HyoSanrentan(i).Kumi)
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