Option Strict Off
Option Explicit On
Option Compare Binary
Friend Class clsImportCH
	' @(h) clsReadCH.cls
	' @(s)
	' JVData "CH" �f�[�^�x�[�X�o�^�N���X
	'
	
	Private mBuf As JV_CH_CHOKYOSI '' �����t�}�X�^�\����
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
        strSql = "SELECT * FROM CHOKYO"
        mRS1 = New ADODB.Recordset()
        mRS1.Open(strSql, gCon, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockPessimistic)


        '���R�[�h�Z�b�g�I�[�v��
        strSql = "SELECT * FROM CHOKYO_SEISEKI"
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
        '���R�[�h�Z�b�g�N���[�Y
        mRS1.Close()

        mRS1 = Nothing
        mRS2.Close()

        mRS2 = Nothing

        System.Diagnostics.Debug.WriteLine("mRS.Close")

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
    ' ������    : lBuf - JVData ���ʎq"CH" �̂P�s
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

        If Not InsertDB() Then
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

        System.Diagnostics.Debug.WriteLine("BeginTrans")
        gCon.BeginTrans()

        mRS1.AddNew()

        With mBuf
            With .head
                mRS1.Fields("RecordSpec").Value = .RecordSpec '' ���R�[�h���
                mRS1.Fields("DataKubun").Value = .DataKubun '' �f�[�^�敪
                With .MakeDate
                    mRS1.Fields("MakeDate").Value = .Year & .Month & .Day '' �N����
                End With ' MakeDate
            End With ' head
            mRS1.Fields("ChokyosiCode").Value = .ChokyosiCode '' �����t�R�[�h
            mRS1.Fields("DelKubun").Value = .DelKubun '' �����t�����敪
            With .IssueDate
                mRS1.Fields("IssueDate").Value = .Year & .Month & .Day '' �N����
            End With ' IssueDate
            With .DelDate
                mRS1.Fields("DelDate").Value = .Year & .Month & .Day '' �N����
            End With ' DelDate
            With .BirthDate
                mRS1.Fields("BirthDate").Value = .Year & .Month & .Day '' �N����
            End With ' BirthDate
            mRS1.Fields("ChokyosiName").Value = .ChokyosiName '' �����t������
            mRS1.Fields("ChokyosiNameKana").Value = .ChokyosiNameKana '' �����t�����p�J�i
            mRS1.Fields("ChokyosiRyakusyo").Value = .ChokyosiRyakusyo '' �����t������
            mRS1.Fields("ChokyosiNameEng").Value = .ChokyosiNameEng '' �����t������
            mRS1.Fields("SexCD").Value = .SexCD '' ���ʋ敪
            mRS1.Fields("TozaiCD").Value = .TozaiCD '' �����t���������R�[�h
            mRS1.Fields("Syotai").Value = .Syotai '' ���Ғn�於

            For i = 0 To 2
                With .SaikinJyusyo(i)
                    With .SaikinJyusyoid '' �ŋߏd�܏��
                        mRS1.Fields("SaikinJyusyo" & i + 1 & "SaikinJyusyoid").Value = .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum
                    End With ' SaikinJyusyoid
                    mRS1.Fields("SaikinJyusyo" & i + 1 & "Hondai").Value = .Hondai '' �������{��
                    mRS1.Fields("SaikinJyusyo" & i + 1 & "Ryakusyo10").Value = .Ryakusyo10 '' ����������10��
                    mRS1.Fields("SaikinJyusyo" & i + 1 & "Ryakusyo6").Value = .Ryakusyo6 '' ����������6��
                    mRS1.Fields("SaikinJyusyo" & i + 1 & "Ryakusyo3").Value = .Ryakusyo3 '' ����������3��
                    mRS1.Fields("SaikinJyusyo" & i + 1 & "GradeCD").Value = .GradeCD '' �O���[�h�R�[�h
                    mRS1.Fields("SaikinJyusyo" & i + 1 & "SyussoTosu").Value = .SyussoTosu '' �o������
                    mRS1.Fields("SaikinJyusyo" & i + 1 & "KettoNum").Value = .KettoNum '' �����o�^�ԍ�
                    mRS1.Fields("SaikinJyusyo" & i + 1 & "Bamei").Value = .Bamei '' �n��
                End With ' SaikinJyusyo
            Next i
        End With ' mBuf

        mRS1.Update()

        With mBuf
            System.Diagnostics.Debug.WriteLine("Insert CHOKYO : " & .ChokyosiCode)
        End With ' id

        ' ���ѕ���
        For i = 0 To 2
            With mBuf
                mRS2.AddNew()
                With .head
                    With .MakeDate
                        mRS2.Fields("MakeDate").Value = .Year & .Month & .Day '' �N����
                    End With ' MakeDate
                End With ' head
                mRS2.Fields("ChokyosiCode").Value = .ChokyosiCode '' �����t�R�[�h
                mRS2.Fields("Num").Value = i '' �A��
                With .HonZenRuikei(i)
                    mRS2.Fields("SetYear").Value = .SetYear '' �ݒ�N
                    mRS2.Fields("HonSyokinHeichi").Value = .HonSyokinHeichi '' ���n�{�܋����v
                    mRS2.Fields("HonSyokinSyogai").Value = .HonSyokinSyogai '' ��Q�{�܋����v
                    mRS2.Fields("FukaSyokinHeichi").Value = .FukaSyokinHeichi '' ���n�t���܋����v
                    mRS2.Fields("FukaSyokinSyogai").Value = .FukaSyokinSyogai '' ��Q�t���܋����v
                    With .ChakuKaisuHeichi

                        For k = 0 To 5
                            mRS2.Fields("HeichiChakuKaisu" & k + 1).Value = .Chakukaisu(k)
                        Next k

                    End With ' ChakuKaisuHeichi

                    With .ChakuKaisuSyogai
                        For k = 0 To 5
                            mRS2.Fields("SyogaiChakuKaisu" & k + 1).Value = .Chakukaisu(k)
                        Next k
                    End With ' ChakuKaisuSyogai

                    For j = 0 To 19
                        With .ChakuKaisuJyo(j)
                            For k = 0 To 5
                                mRS2.Fields("Jyo" & j + 1 & "ChakuKaisu" & k + 1).Value = .Chakukaisu(k)
                            Next k
                        End With ' ChakuKaisuJyo
                    Next j

                    For j = 0 To 5
                        With .ChakuKaisuKyori(j)
                            For k = 0 To 5
                                mRS2.Fields("Kyori" & j + 1 & "ChakuKaisu" & k + 1).Value = .Chakukaisu(k)
                            Next k
                        End With ' ChakuKaisuKyori
                    Next j
                End With ' HonZenRuikei
            End With

            With mBuf
                System.Diagnostics.Debug.WriteLine("Insert CHOKYO : " & .ChokyosiCode & CStr(i))
            End With ' mBuf

            mRS2.Update()

        Next i

        gCon.CommitTrans()
        System.Diagnostics.Debug.WriteLine("CommitTrans")

        InsertDB = True

ExitHandler:
        Exit Function
ErrorHandler:
        InsertDB = False
        gCon.RollbackTrans()
        mRS1.CancelUpdate()
        mRS2.CancelUpdate()
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

        ' �w�b�_����
        strSql = "UPDATE CHOKYO SET "
        With mBuf
            With .head
                strSql = strSql & SS & "RecordSpec" & SE & "='" & Replace(.RecordSpec, "'", "''") & "'," '' ���R�[�h���
                strSql = strSql & SS & "DataKubun" & SE & "='" & Replace(.DataKubun, "'", "''") & "'," '' �f�[�^�敪
                strSql = strSql & SS & "MakeDate" & SE & "= '" & Replace(strMakeDate, "'", "''") & "'," '' �N����
            End With ' head
            strSql = strSql & SS & "ChokyosiCode" & SE & "='" & Replace(.ChokyosiCode, "'", "''") & "'," '' �����t�R�[�h
            strSql = strSql & SS & "DelKubun" & SE & "='" & Replace(.DelKubun, "'", "''") & "'," '' �����t�����敪
            With .IssueDate
                strSql = strSql & SS & "IssueDate" & SE & "= '" & Replace(.Year & .Month & .Day, "'", "''") & "'," '' �N����
            End With ' IssueDate
            With .DelDate
                strSql = strSql & SS & "DelDate" & SE & "= '" & Replace(.Year & .Month & .Day, "'", "''") & "'," '' �N����
            End With ' DelDate
            With .BirthDate
                strSql = strSql & SS & "BirthDate" & SE & "= '" & Replace(.Year & .Month & .Day, "'", "''") & "'," '' �N����
            End With ' BirthDate
            strSql = strSql & SS & "ChokyosiName" & SE & "='" & Replace(.ChokyosiName, "'", "''") & "'," '' �����t������
            strSql = strSql & SS & "ChokyosiNameKana" & SE & "='" & Replace(.ChokyosiNameKana, "'", "''") & "'," '' �����t�����p�J�i
            strSql = strSql & SS & "ChokyosiRyakusyo" & SE & "='" & Replace(.ChokyosiRyakusyo, "'", "''") & "'," '' �����t������
            strSql = strSql & SS & "ChokyosiNameEng" & SE & "='" & Replace(.ChokyosiNameEng, "'", "''") & "'," '' �����t������
            strSql = strSql & SS & "SexCD" & SE & "='" & Replace(.SexCD, "'", "''") & "'," '' ���ʋ敪
            strSql = strSql & SS & "TozaiCD" & SE & "='" & Replace(.TozaiCD, "'", "''") & "'," '' �����t���������R�[�h
            strSql = strSql & SS & "Syotai" & SE & "='" & Replace(.Syotai, "'", "''") & "'," '' ���Ғn�於
            For i = 0 To 2
                With .SaikinJyusyo(i)
                    With .SaikinJyusyoid
                        strSql = strSql & SS & "SaikinJyusyo" & i + 1 & "SaikinJyusyoid" & SE & "='" & Replace(.Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum, "'", "''") & "',"
                        '' �J�ÔN �J�Ì��� ���n��R�[�h �J�É�[��N��] �J�Ó���[N����] ���[�X�ԍ�
                    End With ' SaikinJyusyoid
                    strSql = strSql & SS & "SaikinJyusyo" & i + 1 & "Hondai" & SE & "='" & Replace(.Hondai, "'", "''") & "'," '' �������{��
                    strSql = strSql & SS & "SaikinJyusyo" & i + 1 & "Ryakusyo10" & SE & "='" & Replace(.Ryakusyo10, "'", "''") & "'," '' ����������10��
                    strSql = strSql & SS & "SaikinJyusyo" & i + 1 & "Ryakusyo6" & SE & "='" & Replace(.Ryakusyo6, "'", "''") & "'," '' ����������6��
                    strSql = strSql & SS & "SaikinJyusyo" & i + 1 & "Ryakusyo3" & SE & "='" & Replace(.Ryakusyo3, "'", "''") & "'," '' ����������3��
                    strSql = strSql & SS & "SaikinJyusyo" & i + 1 & "GradeCD" & SE & "='" & Replace(.GradeCD, "'", "''") & "'," '' �O���[�h�R�[�h
                    strSql = strSql & SS & "SaikinJyusyo" & i + 1 & "SyussoTosu" & SE & "='" & Replace(.SyussoTosu, "'", "''") & "'," '' �o������
                    strSql = strSql & SS & "SaikinJyusyo" & i + 1 & "KettoNum" & SE & "='" & Replace(.KettoNum, "'", "''") & "'," '' �����o�^�ԍ�
                    strSql = strSql & SS & "SaikinJyusyo" & i + 1 & "Bamei" & SE & "='" & Replace(.Bamei, "'", "''") & "'," '' �n��
                End With ' SaikinJyusyo
            Next i

            strSql = Left(strSql, Len(strSql) - 1) ' Delete Last Comma
            strSql = strSql & " WHERE " & SS & "ChokyosiCode" & SE & "='" & Replace(.ChokyosiCode, "'", "''") & "'"
            strSql = strSql & " AND " & SS & "MakeDate" & SE & " <= '" & Replace(strMakeDate, "'", "''") & "'"
        End With ' mBuf

        With mBuf
            System.Diagnostics.Debug.WriteLine("UPDATE CHOKYO : " & .ChokyosiCode)
        End With ' mBuf

        gCon.Execute(strSql)

        ' ���ѕ���
        For i = 0 To 2
            With mBuf
                strSql = "UPDATE CHOKYO_SEISEKI SET "
                strSql = strSql & SS & "MakeDate" & SE & "= '" & Replace(strMakeDate, "'", "''") & "'," '' �N����
                strSql = strSql & SS & "ChokyosiCode" & SE & "='" & Replace(.ChokyosiCode, "'", "''") & "'," '' �����t�R�[�h
                strSql = strSql & SS & "Num" & SE & "='" & Replace(CStr(i), "'", "''") & "'," '' �A��
                With .HonZenRuikei(i)
                    strSql = strSql & SS & "SetYear" & SE & "='" & Replace(.SetYear, "'", "''") & "'," '' �ݒ�N
                    strSql = strSql & SS & "HonSyokinHeichi" & SE & "='" & Replace(.HonSyokinHeichi, "'", "''") & "'," '' ���n�{�܋����v
                    strSql = strSql & SS & "HonSyokinSyogai" & SE & "='" & Replace(.HonSyokinSyogai, "'", "''") & "'," '' ��Q�{�܋����v
                    strSql = strSql & SS & "FukaSyokinHeichi" & SE & "='" & Replace(.FukaSyokinHeichi, "'", "''") & "'," '' ���n�t���܋����v
                    strSql = strSql & SS & "FukaSyokinSyogai" & SE & "='" & Replace(.FukaSyokinSyogai, "'", "''") & "'," '' ��Q�t���܋����v
                    With .ChakuKaisuHeichi
                        For k = 0 To 5
                            strSql = strSql & SS & "HeichiChakukaisu" & k + 1 & SE & "='" & Replace(.Chakukaisu(k), "'", "''") & "',"
                        Next k
                    End With ' ChakuKaisuHeichi
                    With .ChakuKaisuSyogai
                        For k = 0 To 5
                            strSql = strSql & SS & "SyogaiChakukaisu" & k + 1 & SE & "='" & Replace(.Chakukaisu(k), "'", "''") & "',"
                        Next k
                    End With ' ChakuKaisuSyogai
                    For j = 0 To 5
                        With .ChakuKaisuKyori(j)
                            For k = 0 To 5
                                strSql = strSql & SS & "Kyori" & j + 1 & "Chakukaisu" & k + 1 & SE & "='" & Replace(.Chakukaisu(k), "'", "''") & "',"
                            Next k
                        End With ' ChakuKaisuKyori
                    Next j
                End With ' HonZenRuikei

                strSql = Left(strSql, Len(strSql) - 1) ' Delete Last Comma

                strSql = strSql & " WHERE " & SS & "ChokyosiCode" & SE & "='" & Replace(.ChokyosiCode, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "Num" & SE & "='" & Replace(CStr(i), "'", "''") & "'"
                strSql = strSql & " AND " & SS & "MakeDate" & SE & " <= '" & Replace(strMakeDate, "'", "''") & "'"

                gCon.Execute(strSql)

                ''��x�ɍX�V�ł���t�B�[���h������127�܂ł̈� �����X�V�iJET�d�l�j
                strSql = "UPDATE CHOKYO_SEISEKI SET "
                strSql = strSql & SS & "MakeDate" & SE & "= '" & Replace(strMakeDate, "'", "''") & "'," '' �N����
                With .HonZenRuikei(i)
                    For j = 0 To 19

                        With .ChakuKaisuJyo(j)
                            For k = 0 To 5

                                strSql = strSql & SS & "Jyo" & j + 1 & "Chakukaisu" & k + 1 & SE & "='" & Replace(.Chakukaisu(k), "'", "''") & "',"
                            Next k
                        End With ' ChakuKaisuJyo
                    Next j
                End With ' HonZenRuikei

                strSql = Left(strSql, Len(strSql) - 1) ' Delete Last Comma

                strSql = strSql & " WHERE " & SS & "ChokyosiCode" & SE & "='" & Replace(.ChokyosiCode, "'", "''") & "'"
                strSql = strSql & " AND " & SS & "Num" & SE & "='" & Replace(CStr(i), "'", "''") & "'"
                strSql = strSql & " AND " & SS & "MakeDate" & SE & " <= '" & Replace(strMakeDate, "'", "''") & "'"

            End With ' mBuf

            With mBuf
                System.Diagnostics.Debug.WriteLine("UPDATE CHOKYO_SEISEKI : " & .ChokyosiCode & CStr(i))
            End With ' mBuf

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