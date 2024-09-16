Option Strict Off
Option Explicit On
Option Compare Binary
Friend Class clsImportH1
	' @(h) clsReadH1.cls
	' @(s)
	' JVData "H1" �f�[�^�x�[�X�o�^�N���X
	'
	
	Private mBuf As JV_H1_HYOSU_ZENKAKE '' �[���i�S�q���j�\����
	Private mRS1 As ADODB.Recordset
	Private mRS2 As ADODB.Recordset
	Private mRS3 As ADODB.Recordset
	Private mRS4 As ADODB.Recordset
	Private mRS5 As ADODB.Recordset
	Private mRS6 As ADODB.Recordset
	
	
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
        strSql = "SELECT * FROM HYOSU"
        mRS1 = New ADODB.Recordset()
        mRS1.Open(strSql, gCon, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockPessimistic)

        strSql = "SELECT * FROM HYOSU_TANPUKU"
        mRS2 = New ADODB.Recordset()
        mRS2.Open(strSql, gCon, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockPessimistic)

        strSql = "SELECT * FROM HYOSU_WAKU"
        mRS3 = New ADODB.Recordset()
        mRS3.Open(strSql, gCon, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockPessimistic)

        strSql = "SELECT * FROM HYOSU_UMARENWIDE"
        mRS4 = New ADODB.Recordset()
        mRS4.Open(strSql, gCon, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockPessimistic)

        strSql = "SELECT * FROM HYOSU_UMATAN"
        mRS5 = New ADODB.Recordset()
        mRS5.Open(strSql, gCon, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockPessimistic)

        strSql = "SELECT * FROM HYOSU_SANREN"
        mRS6 = New ADODB.Recordset()
        mRS6.Open(strSql, gCon, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockPessimistic)

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
        mRS3.Close()

        mRS3 = Nothing
        mRS4.Close()

        mRS4 = Nothing
        mRS5.Close()

        mRS5 = Nothing
        mRS6.Close()

        mRS6 = Nothing


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
    ' ������    : lBuf - JVData ���ʎq"H1" �̂P�s
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

        ' HYOSU (�[��)
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
            For i = 0 To 6
                mRS1.Fields("HatubaiFlag" & i + 1).Value = .HatubaiFlag(i) '' �����t���O
            Next i
            mRS1.Fields("FukuChakuBaraiKey").Value = .FukuChakuBaraiKey '' ���������L�[
            For i = 0 To 27
                mRS1.Fields("HenkanUma" & i + 1).Value = .HenkanUma(i) '' �ԊҔn�ԏ��(�n��01�`28)
            Next i
            For i = 0 To 7
                mRS1.Fields("HenkanWaku" & i + 1).Value = .HenkanWaku(i) '' �ԊҘg�ԏ��(�g��1�`8)
            Next i
            For i = 0 To 7
                mRS1.Fields("HenkanDoWaku" & i + 1).Value = .HenkanDoWaku(i) '' �Ԋғ��g���(�g��1�`8)
            Next i
            For i = 0 To 13
                mRS1.Fields("HyoTotal" & i + 1).Value = .HyoTotal(i) '' �[�����v
            Next i
        End With ' mBuf

        With mBuf.id
            System.Diagnostics.Debug.WriteLine("Insert HYOSU : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum)
        End With ' id

        mRS1.Update()


        ' HYOSU_TANPUKU (�[��_�P��)
            For i = 0 To 27
                If mBuf.HyoTansyo(i).Umaban <> "  " Then
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
                        With .HyoTansyo(i)
                            mRS2.Fields("Umaban").Value = .Umaban '' �n��
                            mRS2.Fields("TanHyo").Value = .Hyo '' �[��
                            mRS2.Fields("TanNinki").Value = .Ninki '' �l�C
                        End With ' HyoTansyo
                        With .HyoFukusyo(i)
                            mRS2.Fields("FukuHyo").Value = .Hyo '' �[��
                            mRS2.Fields("FukuNinki").Value = .Ninki '' �l�C
                        End With ' HyoFukusyo
                    End With ' mBuf

                    mRS2.Update()

                    With mBuf.id
                        System.Diagnostics.Debug.WriteLine("Insert HYOSU_TANPUKU : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum & mBuf.HyoTansyo(i).Umaban)
                    End With ' id
                End If
            Next i

        ' HYOSU_WAKU (�[��_�g�A)
            For i = 0 To 35
                If mBuf.HyoWakuren(i).Umaban <> "  " Then
                    mRS3.AddNew()
                    With mBuf
                        With .head
                            With .MakeDate
                                mRS3.Fields("MakeDate").Value = .Year & .Month & .Day '' �N����
                            End With ' MakeDate
                        End With ' head
                        With .id
                            mRS3.Fields("Year").Value = .Year '' �J�ÔN
                            mRS3.Fields("MonthDay").Value = .MonthDay '' �J�Ì���
                            mRS3.Fields("JyoCD").Value = .JyoCD '' ���n��R�[�h
                            mRS3.Fields("Kaiji").Value = .Kaiji '' �J�É��N��
                            mRS3.Fields("Nichiji").Value = .Nichiji '' �J�Ó���N����
                            mRS3.Fields("RaceNum").Value = .RaceNum '' ���[�X�ԍ�
                        End With ' id
                        With .HyoWakuren(i)
                            mRS3.Fields("Kumi").Value = .Umaban '' �g��
                            mRS3.Fields("Hyo").Value = .Hyo '' �[��
                            mRS3.Fields("Ninki").Value = .Ninki '' �l�C
                        End With ' HyoWakuren
                    End With ' mBuf

                    With mBuf.id
                        System.Diagnostics.Debug.WriteLine("Insert HYOSU_WAKU : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum & mBuf.HyoWakuren(i).Umaban)
                    End With ' id
                    mRS3.Update()
                End If
            Next i

        ' HYOSU_UMARENWIDE (�[��_�n�A�E���C�h)
            For i = 0 To 152
                If mBuf.HyoUmaren(i).Kumi <> "    " Then
                    mRS4.AddNew()
                    With mBuf
                        With .head
                            With .MakeDate
                                mRS4.Fields("MakeDate").Value = .Year & .Month & .Day '' �N����
                            End With ' MakeDate
                        End With ' head
                        With .id
                            mRS4.Fields("Year").Value = .Year '' �J�ÔN
                            mRS4.Fields("MonthDay").Value = .MonthDay '' �J�Ì���
                            mRS4.Fields("JyoCD").Value = .JyoCD '' ���n��R�[�h
                            mRS4.Fields("Kaiji").Value = .Kaiji '' �J�É��N��
                            mRS4.Fields("Nichiji").Value = .Nichiji '' �J�Ó���N����
                            mRS4.Fields("RaceNum").Value = .RaceNum '' ���[�X�ԍ�
                        End With ' id
                        With .HyoUmaren(i)
                            mRS4.Fields("Kumi").Value = .Kumi '' �g��
                            mRS4.Fields("UmarenHyo").Value = .Hyo '' �[��
                            mRS4.Fields("UmarenNinki").Value = .Ninki '' �l�C
                        End With ' HyoUmaren
                        With .HyoWide(i)
                            mRS4.Fields("WideHyo").Value = .Hyo '' �[��
                            mRS4.Fields("WideNinki").Value = .Ninki '' �l�C
                        End With ' HyoWide
                    End With ' mBuf
                    With mBuf.id
                        System.Diagnostics.Debug.WriteLine("Insert HYOSU_UMARENWIDE : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum & mBuf.HyoUmaren(i).Kumi)
                    End With ' id
                    mRS4.Update()
                End If
            Next i

        ' HYOSU_UMATAN (�[��_�n�P)
            For i = 0 To 305
                If mBuf.HyoUmatan(i).Kumi <> "    " Then
                    mRS5.AddNew()
                    With mBuf
                        With .head
                            With .MakeDate
                                mRS5.Fields("MakeDate").Value = .Year & .Month & .Day '' �N����
                            End With ' MakeDate
                        End With ' head
                        With .id
                            mRS5.Fields("Year").Value = .Year '' �J�ÔN
                            mRS5.Fields("MonthDay").Value = .MonthDay '' �J�Ì���
                            mRS5.Fields("JyoCD").Value = .JyoCD '' ���n��R�[�h
                            mRS5.Fields("Kaiji").Value = .Kaiji '' �J�É��N��
                            mRS5.Fields("Nichiji").Value = .Nichiji '' �J�Ó���N����
                            mRS5.Fields("RaceNum").Value = .RaceNum '' ���[�X�ԍ�
                        End With ' id
                        With .HyoUmatan(i)
                            mRS5.Fields("Kumi").Value = .Kumi '' �g��
                            mRS5.Fields("Hyo").Value = .Hyo '' �[��
                            mRS5.Fields("Ninki").Value = .Ninki '' �l�C
                        End With ' HyoUmatan
                    End With ' mBuf
                    With mBuf.id
                        System.Diagnostics.Debug.WriteLine("Insert HYOSU_UMATAN : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum & mBuf.HyoUmatan(i).Kumi)
                    End With ' id
                    mRS5.Update()
                End If
            Next i

        ' HYOSU_SANREN (�[��_�O�A)
            For i = 0 To 815
                If mBuf.HyoSanrenpuku(i).Kumi <> "      " Then
                    mRS6.AddNew()
                    With mBuf
                        With .head
                            With .MakeDate
                                mRS6.Fields("MakeDate").Value = .Year & .Month & .Day '' �N����
                            End With ' MakeDate
                        End With ' head
                        With .id
                            mRS6.Fields("Year").Value = .Year '' �J�ÔN
                            mRS6.Fields("MonthDay").Value = .MonthDay '' �J�Ì���
                            mRS6.Fields("JyoCD").Value = .JyoCD '' ���n��R�[�h
                            mRS6.Fields("Kaiji").Value = .Kaiji '' �J�É��N��
                            mRS6.Fields("Nichiji").Value = .Nichiji '' �J�Ó���N����
                            mRS6.Fields("RaceNum").Value = .RaceNum '' ���[�X�ԍ�
                        End With ' id
                        With .HyoSanrenpuku(i)
                            mRS6.Fields("Kumi").Value = .Kumi '' �g��
                            mRS6.Fields("Hyo").Value = .Hyo '' �[��
                            mRS6.Fields("Ninki").Value = .Ninki '' �l�C
                        End With ' HyoSanrenpuku
                    End With ' mBuf

                    With mBuf.id
                        System.Diagnostics.Debug.WriteLine("Insert HYOSU_SANREN : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum & mBuf.HyoSanrenpuku(i).Kumi)
                    End With ' id
                    mRS6.Update()
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
        mRS3.CancelUpdate()
        mRS4.CancelUpdate()
        mRS5.CancelUpdate()
        mRS6.CancelUpdate()
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

        ' HYOSU (�[��)
        strSql = "UPDATE HYOSU SET "
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
            For i = 0 To 6
                strSql = strSql & SS & "HatubaiFlag" & i + 1 & SE & "='" & Replace(.HatubaiFlag(i), "'", "''") & "'," '' �����t���O
            Next i
            strSql = strSql & SS & "FukuChakuBaraiKey" & SE & "='" & Replace(.FukuChakuBaraiKey, "'", "''") & "'," '' ���������L�[
            For i = 0 To 27
                strSql = strSql & SS & "HenkanUma" & i + 1 & SE & "='" & Replace(.HenkanUma(i), "'", "''") & "'," '' �ԊҔn�ԏ��(�n��01�`28)
            Next i
            For i = 0 To 7
                strSql = strSql & SS & "HenkanWaku" & i + 1 & SE & "='" & Replace(.HenkanWaku(i), "'", "''") & "'," '' �ԊҘg�ԏ��(�g��1�`8)
            Next i
            For i = 0 To 7
                strSql = strSql & SS & "HenkanDoWaku" & i + 1 & SE & "='" & Replace(.HenkanDoWaku(i), "'", "''") & "'," '' �Ԋғ��g���(�g��1�`8)
            Next i
            For i = 0 To 13
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
            System.Diagnostics.Debug.WriteLine("UPDATE HYOSU : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum)
        End With ' id

        gCon.Execute(strSql)

        ' HYOSU_TANPUKU (�[��_�P��)
        For i = 0 To 27
            strSql = "UPDATE HYOSU_TANPUKU SET "
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
                With .HyoTansyo(i)
                    strSql = strSql & SS & "Umaban" & SE & "='" & Replace(.Umaban, "'", "''") & "'," '' �n��
                    strSql = strSql & SS & "TanHyo" & SE & "='" & Replace(.Hyo, "'", "''") & "'," '' �[��
                    strSql = strSql & SS & "TanNinki" & SE & "='" & Replace(.Ninki, "'", "''") & "'," '' �l�C
                End With ' HyoTansyo
                With .HyoFukusyo(i)
                    strSql = strSql & SS & "FukuHyo" & SE & "='" & Replace(.Hyo, "'", "''") & "'," '' �[��
                    strSql = strSql & SS & "FukuNinki" & SE & "='" & Replace(.Ninki, "'", "''") & "'," '' �l�C
                End With ' HyoFukusyo

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
                strSql = strSql & " AND " & SS & "Umaban" & SE & "='" & Replace(.HyoTansyo(i).Umaban, "'", "''") & "'"
            End With ' mBuf

            With mBuf.id
                System.Diagnostics.Debug.WriteLine("UPDATE HYOSU_TANPUKU : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum & mBuf.HyoTansyo(i).Umaban)
            End With ' id

            gCon.Execute(strSql)
        Next i

        ' HYOSU_WAKU (�[��_�g�A)
        For i = 0 To 35
            strSql = "UPDATE HYOSU_WAKU SET "
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
                With .HyoWakuren(i)
                    strSql = strSql & SS & "Kumi" & SE & "='" & Replace(.Umaban, "'", "''") & "'," '' �n��
                    strSql = strSql & SS & "Hyo" & SE & "='" & Replace(.Hyo, "'", "''") & "'," '' �[��
                    strSql = strSql & SS & "Ninki" & SE & "='" & Replace(.Ninki, "'", "''") & "'," '' �l�C
                End With ' HyoWakuren

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
                strSql = strSql & " AND " & SS & "Kumi" & SE & "='" & Replace(.HyoWakuren(i).Umaban, "'", "''") & "'"
            End With ' mBuf

            With mBuf.id
                System.Diagnostics.Debug.WriteLine("UPDATE HYOSU_WAKU : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum & mBuf.HyoWakuren(i).Umaban)
            End With ' id

            gCon.Execute(strSql)
        Next i

        ' HYOSU_UMARENWIDE (�[��_�n�A�E���C�h)
        For i = 0 To 152
            strSql = "UPDATE HYOSU_UMARENWIDE SET "
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
                With .HyoUmaren(i)
                    strSql = strSql & SS & "Kumi" & SE & "='" & Replace(.Kumi, "'", "''") & "'," '' �g��
                    strSql = strSql & SS & "UmarenHyo" & SE & "='" & Replace(.Hyo, "'", "''") & "'," '' �[��
                    strSql = strSql & SS & "UmarenNinki" & SE & "='" & Replace(.Ninki, "'", "''") & "'," '' �l�C
                End With ' HyoUmaren
                With .HyoWide(i)
                    strSql = strSql & SS & "WideHyo" & SE & "='" & Replace(.Hyo, "'", "''") & "'," '' �[��
                    strSql = strSql & SS & "WideNinki" & SE & "='" & Replace(.Ninki, "'", "''") & "'," '' �l�C
                End With ' HyoWide

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
                strSql = strSql & " AND " & SS & "Kumi" & SE & "='" & Replace(.HyoUmaren(i).Kumi, "'", "''") & "'"
            End With ' mBuf

            With mBuf.id
                System.Diagnostics.Debug.WriteLine("UPDATE HYOSU_UMARENWIDE : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum & mBuf.HyoUmaren(i).Kumi)
            End With ' id

            gCon.Execute(strSql)

        Next i

        ' HYOSU_UMATAN (�[��_�n�P)
        For i = 0 To 305
            strSql = "UPDATE HYOSU_UMATAN SET "
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
                With .HyoUmatan(i)
                    strSql = strSql & SS & "Kumi" & SE & "='" & Replace(.Kumi, "'", "''") & "'," '' �g��
                    strSql = strSql & SS & "Hyo" & SE & "='" & Replace(.Hyo, "'", "''") & "'," '' �[��
                    strSql = strSql & SS & "Ninki" & SE & "='" & Replace(.Ninki, "'", "''") & "'," '' �l�C
                End With ' HyoUmatan

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
                strSql = strSql & " AND " & SS & "Kumi" & SE & "='" & Replace(.HyoUmatan(i).Kumi, "'", "''") & "'"
            End With ' mBuf

            With mBuf.id
                System.Diagnostics.Debug.WriteLine("UPDATE HYOSU_UMATAN : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum & mBuf.HyoUmatan(i).Kumi)
            End With ' id

            gCon.Execute(strSql)
        Next i

        ' HYOSU_SANREN (�[��_�O�A)
        For i = 0 To 815
            strSql = "UPDATE HYOSU_SANREN SET "
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
                With .HyoSanrenpuku(i)
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
                strSql = strSql & " AND " & SS & "Kumi" & SE & "='" & Replace(.HyoSanrenpuku(i).Kumi, "'", "''") & "'"
            End With ' mBuf

            With mBuf.id
                System.Diagnostics.Debug.WriteLine("UPDATE HYOSU_SANREN : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum & mBuf.HyoSanrenpuku(i).Kumi)
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