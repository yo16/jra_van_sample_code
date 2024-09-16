Option Strict Off
Option Explicit On
Option Compare Binary
Friend Class clsImportHR
	' @(h) clsReadHR.cls
	' @(s)
	' JVData "HR" �f�[�^�x�[�X�o�^�N���X
	'
	
	Private mBuf As JV_HR_PAY '' ���ߍ\����
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
        strSql = "SELECT * FROM HARAI"
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
    ' �@�\      : Add�v���V�[�W�����Ă�
    '
    ' ������    : lBuf - JVData ���ʎq"HR" �̂P�s
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
        System.Diagnostics.Debug.WriteLine("mRs.AddNew")
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
                mRS.Fields("RaceNum").Value = .RaceNum '' ���[�X�ԍ�
            End With ' id
            mRS.Fields("TorokuTosu").Value = .TorokuTosu '' �o�^����
            mRS.Fields("SyussoTosu").Value = .SyussoTosu '' �o������
            For i = 0 To 8
                mRS.Fields("FuseirituFlag" & i + 1).Value = .FuseirituFlag(i) '' �s�����t���O
            Next i
            For i = 0 To 8
                mRS.Fields("TokubaraiFlag" & i + 1).Value = .TokubaraiFlag(i) '' �����t���O
            Next i
            For i = 0 To 8
                mRS.Fields("HenkanFlag" & i + 1).Value = .HenkanFlag(i) '' �Ԋ҃t���O
            Next i
            For i = 0 To 27
                mRS.Fields("HenkanUma" & i + 1).Value = .HenkanUma(i) '' �ԊҔn�ԏ��(�n��01�`28)
            Next i
            For i = 0 To 7
                mRS.Fields("HenkanWaku" & i + 1).Value = .HenkanWaku(i) '' �ԊҘg�ԏ��(�g��1�`8)
            Next i
            For i = 0 To 7
                mRS.Fields("HenkanDoWaku" & i + 1).Value = .HenkanDoWaku(i) '' �Ԋғ��g���(�g��1�`8)
            Next i
            For i = 0 To 2
                With .PayTansyo(i)
                    mRS.Fields("PayTansyoUmaban" & i + 1).Value = .Umaban '' �n��
                    mRS.Fields("PayTansyoPay" & i + 1).Value = .Pay '' ���ߋ�
                    mRS.Fields("PayTansyoNinki" & i + 1).Value = .Ninki '' �l�C��
                End With ' PayTansyo
            Next i
            For i = 0 To 4
                With .PayFukusyo(i)
                    mRS.Fields("PayFukusyoUmaban" & i + 1).Value = .Umaban '' �n��
                    mRS.Fields("PayFukusyoPay" & i + 1).Value = .Pay '' ���ߋ�
                    mRS.Fields("PayFukusyoNinki" & i + 1).Value = .Ninki '' �l�C��
                End With ' PayFukusyo
            Next i
            For i = 0 To 2
                With .PayWakuren(i)
                    mRS.Fields("PayWakurenKumi" & i + 1).Value = .Umaban '' �g��
                    mRS.Fields("PayWakurenPay" & i + 1).Value = .Pay '' ���ߋ�
                    mRS.Fields("PayWakurenNinki" & i + 1).Value = .Ninki '' �l�C��
                End With ' PayWakuren
            Next i
            For i = 0 To 2
                With .PayUmaren(i)
                    mRS.Fields("PayUmarenKumi" & i + 1).Value = .Kumi '' �g��
                    mRS.Fields("PayUmarenPay" & i + 1).Value = .Pay '' ���ߋ�
                    mRS.Fields("PayUmarenNinki" & i + 1).Value = .Ninki '' �l�C��
                End With ' PayUmaren
            Next i
            For i = 0 To 6
                With .PayWide(i)
                    mRS.Fields("PayWideKumi" & i + 1).Value = .Kumi '' �g��
                    mRS.Fields("PayWidePay" & i + 1).Value = .Pay '' ���ߋ�
                    mRS.Fields("PayWideNinki" & i + 1).Value = .Ninki '' �l�C��
                End With ' PayWide
            Next i
            For i = 0 To 2
                With .PayReserved1(i)
                    mRS.Fields("PayReserved1Kumi" & i + 1).Value = .Kumi '' �g��
                    mRS.Fields("PayReserved1Pay" & i + 1).Value = .Pay '' ���ߋ�
                    mRS.Fields("PayReserved1Ninki" & i + 1).Value = .Ninki '' �l�C��
                End With ' PayReserved1
            Next i
            For i = 0 To 5
                With .PayUmatan(i)
                    mRS.Fields("PayUmatanKumi" & i + 1).Value = .Kumi '' �g��
                    mRS.Fields("PayUmatanPay" & i + 1).Value = .Pay '' ���ߋ�
                    mRS.Fields("PayUmatanNinki" & i + 1).Value = .Ninki '' �l�C��
                End With ' PayUmatan
            Next i
            For i = 0 To 2
                With .PaySanrenpuku(i)
                    mRS.Fields("PaySanrenpukuKumi" & i + 1).Value = .Kumi '' �g��
                    mRS.Fields("PaySanrenpukuPay" & i + 1).Value = .Pay '' ���ߋ�
                    mRS.Fields("PaySanrenpukuNinki" & i + 1).Value = .Ninki '' �l�C��
                End With ' PaySanrenpuku
            Next i
            For i = 0 To 5
                With .PaySanrentan(i)
                    mRS.Fields("PaySanrentanKumi" & i + 1).Value = .Kumi '' �g��
                    mRS.Fields("PaySanrentanPay" & i + 1).Value = .Pay '' ���ߋ�
                    mRS.Fields("PaySanrentanNinki" & i + 1).Value = .Ninki '' �l�C��
                End With ' PayReserved2
            Next i
        End With

        With mBuf.id
            System.Diagnostics.Debug.WriteLine("Insert HARAI : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum)
        End With ' id

        mRS.Update()
        System.Diagnostics.Debug.WriteLine("mRS.update")

        gCon.CommitTrans()
        System.Diagnostics.Debug.WriteLine("CommitTrans")

        InsertDB = True

ExitHandler:
        Exit Function
ErrorHandler:
        InsertDB = False
        mRS.CancelUpdate()
        System.Diagnostics.Debug.WriteLine("mRS.CancelUpdate")
        gCon.RollbackTrans()
        System.Diagnostics.Debug.WriteLine(Err.Description)
        System.Diagnostics.Debug.WriteLine("Insert RollbackTrans")
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

        strSql = "UPDATE HARAI SET "
        With mBuf

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
            For i = 0 To 8
                strSql = strSql & SS & "FuseirituFlag" & i + 1 & SE & "='" & Replace(.FuseirituFlag(i), "'", "''") & "'," '' �s�����t���O
            Next i
            For i = 0 To 8
                strSql = strSql & SS & "TokubaraiFlag" & i + 1 & SE & "='" & Replace(.TokubaraiFlag(i), "'", "''") & "'," '' �����t���O
            Next i
            For i = 0 To 8
                strSql = strSql & SS & "HenkanFlag" & i + 1 & SE & "='" & Replace(.HenkanFlag(i), "'", "''") & "'," '' �Ԋ҃t���O
            Next i
            For i = 0 To 27
                strSql = strSql & SS & "HenkanUma" & i + 1 & SE & "='" & Replace(.HenkanUma(i), "'", "''") & "'," '' �ԊҔn�ԏ��(�n��01�`28)
            Next i
            For i = 0 To 7
                strSql = strSql & SS & "HenkanWaku" & i + 1 & SE & "='" & Replace(.HenkanWaku(i), "'", "''") & "'," '' �ԊҘg�ԏ��(�g��1�`8)
            Next i
            For i = 0 To 7
                strSql = strSql & SS & "HenkanDoWaku" & i + 1 & SE & "='" & Replace(.HenkanDoWaku(i), "'", "''") & "'," '' �Ԋғ��g���(�g��1�`8)
            Next i
            For i = 0 To 2
                With .PayTansyo(i)
                    strSql = strSql & SS & "PayTansyoUmaban" & i + 1 & SE & "='" & Replace(.Umaban, "'", "''") & "'," '' �n��
                    strSql = strSql & SS & "PayTansyoPay" & i + 1 & SE & "='" & Replace(.Pay, "'", "''") & "'," '' ���ߋ�
                    strSql = strSql & SS & "PayTansyoNinki" & i + 1 & SE & "='" & Replace(.Ninki, "'", "''") & "'," '' �l�C��
                End With ' PayTansyo
            Next i
            For i = 0 To 4
                With .PayFukusyo(i)
                    strSql = strSql & SS & "PayFukusyoUmaban" & i + 1 & SE & "='" & Replace(.Umaban, "'", "''") & "'," '' �n��
                    strSql = strSql & SS & "PayFukusyoPay" & i + 1 & SE & "='" & Replace(.Pay, "'", "''") & "'," '' ���ߋ�
                    strSql = strSql & SS & "PayFukusyoNinki" & i + 1 & SE & "='" & Replace(.Ninki, "'", "''") & "'," '' �l�C��
                End With ' PayFukusyo
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
            End With

            gCon.Execute(strSql)

            ''��x�ɍX�V�ł���t�B�[���h������127�܂ł̈� �����X�V�iJET�d�l�j
            strSql = "UPDATE HARAI SET "
            With .head

                strSql = strSql & SS & "RecordSpec" & SE & "='" & Replace(.RecordSpec, "'", "''") & "'," '' ���R�[�h���

                strSql = strSql & SS & "DataKubun" & SE & "='" & Replace(.DataKubun, "'", "''") & "'," '' �f�[�^�敪
                strSql = strSql & SS & "MakeDate" & SE & "= '" & Replace(strMakeDate, "'", "''") & "'," '' �N����
            End With ' head
            For i = 0 To 2
                With .PayWakuren(i)
                    strSql = strSql & SS & "PayWakurenKumi" & i + 1 & SE & "='" & Replace(.Umaban, "'", "''") & "'," '' �n��
                    strSql = strSql & SS & "PayWakurenPay" & i + 1 & SE & "='" & Replace(.Pay, "'", "''") & "'," '' ���ߋ�
                    strSql = strSql & SS & "PayWakurenNinki" & i + 1 & SE & "='" & Replace(.Ninki, "'", "''") & "'," '' �l�C��
                End With ' PayWakuren
            Next i
            For i = 0 To 2
                With .PayUmaren(i)
                    strSql = strSql & SS & "PayUmarenKumi" & i + 1 & SE & "='" & Replace(.Kumi, "'", "''") & "'," '' �g��
                    strSql = strSql & SS & "PayUmarenPay" & i + 1 & SE & "='" & Replace(.Pay, "'", "''") & "'," '' ���ߋ�
                    strSql = strSql & SS & "PayUmarenNinki" & i + 1 & SE & "='" & Replace(.Ninki, "'", "''") & "'," '' �l�C��
                End With ' PayUmaren
            Next i
            For i = 0 To 6
                With .PayWide(i)
                    strSql = strSql & SS & "PayWideKumi" & i + 1 & SE & "='" & Replace(.Kumi, "'", "''") & "'," '' �g��
                    strSql = strSql & SS & "PayWidePay" & i + 1 & SE & "='" & Replace(.Pay, "'", "''") & "'," '' ���ߋ�
                    strSql = strSql & SS & "PayWideNinki" & i + 1 & SE & "='" & Replace(.Ninki, "'", "''") & "'," '' �l�C��
                End With ' PayWide
            Next i
            For i = 0 To 2
                With .PayReserved1(i)
                    strSql = strSql & SS & "PayReserved1Kumi" & i + 1 & SE & "='" & Replace(.Kumi, "'", "''") & "'," '' �g��
                    strSql = strSql & SS & "PayReserved1Pay" & i + 1 & SE & "='" & Replace(.Pay, "'", "''") & "'," '' ���ߋ�
                    strSql = strSql & SS & "PayReserved1Ninki" & i + 1 & SE & "='" & Replace(.Ninki, "'", "''") & "'," '' �l�C��
                End With ' PayReserved1
            Next i
            For i = 0 To 5
                With .PayUmatan(i)
                    strSql = strSql & SS & "PayUmatanKumi" & i + 1 & SE & "='" & Replace(.Kumi, "'", "''") & "'," '' �g��
                    strSql = strSql & SS & "PayUmatanPay" & i + 1 & SE & "='" & Replace(.Pay, "'", "''") & "'," '' ���ߋ�
                    strSql = strSql & SS & "PayUmatanNinki" & i + 1 & SE & "='" & Replace(.Ninki, "'", "''") & "'," '' �l�C��
                End With ' PayUmatan
            Next i
            For i = 0 To 2
                With .PaySanrenpuku(i)
                    strSql = strSql & SS & "PaySanrenpukuKumi" & i + 1 & SE & "='" & Replace(.Kumi, "'", "''") & "'," '' �g��
                    strSql = strSql & SS & "PaySanrenpukuPay" & i + 1 & SE & "='" & Replace(.Pay, "'", "''") & "'," '' ���ߋ�
                    strSql = strSql & SS & "PaySanrenpukuNinki" & i + 1 & SE & "='" & Replace(.Ninki, "'", "''") & "'," '' �l�C��
                End With ' PaySanrenpuku
            Next i
            For i = 0 To 5
                With .PaySanrentan(i)
                    strSql = strSql & SS & "PaySanrentanKumi" & i + 1 & SE & "='" & Replace(.Kumi, "'", "''") & "'," '' �g��
                    strSql = strSql & SS & "PaySanrentanPay" & i + 1 & SE & "='" & Replace(.Pay, "'", "''") & "'," '' ���ߋ�
                    strSql = strSql & SS & "PaySanrentanNinki" & i + 1 & SE & "='" & Replace(.Ninki, "'", "''") & "'," '' �l�C��
                End With ' PayReserved2
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
            End With
        End With

        With mBuf.id
            System.Diagnostics.Debug.WriteLine("UPDATE HARAI : " & .Year & .MonthDay & .JyoCD & .Kaiji & .Nichiji & .RaceNum)
        End With ' id

        gCon.Execute(strSql)

        System.Diagnostics.Debug.WriteLine("CommitTrans")
        gCon.CommitTrans()

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