Option Strict Off
Option Explicit On
Friend Class clsCodeConv
    ' �R�[�h���̎擾���W���[��
    ' �R�[�h���疼�̂��擾����
    '
    Private Structure mudtCodeLine ''�R�[�h���̍\��
        Dim strCodeNo As String ''�R�[�hNo.
        Dim StrCode As String ''�R�[�h
        Dim strNames As String ''���̗�i��������ꍇ�C�A������������̂܂܊i�[�j
    End Structure

    Private mFileName As String ''���̓t�@�C����
    Private mArrData() As clsCodeConv.mudtCodeLine ''�R�[�h���̃f�[�^
    Private blnFlag As Boolean ''�f�[�^�Ǎ��m�F�t���O


    ' @(f)
    '
    ' �@�\�@�@ : �f�[�^�̊i�[
    '
    ' �������@ : ARG1 - �t�@�C����
    '
    ' �Ԃ�l�@ : �Ȃ�
    '
    ' �@�\���� : �w�肳�ꂽ�t�@�C���̃f�[�^����������Ɋi�[����
    '
    Public WriteOnly Property FileName() As String
        Set(ByVal Value As String)
            On Error GoTo err_Renamed
            mFileName = Value
            Call SetData()
ext:
            Exit Property
err_Renamed:
            MsgBox(Err.Description)
            Resume ext
        End Set
    End Property

    ' @(f)
    '
    ' �@�\�@�@ : ���̂̎擾
    '
    ' �������@ : ARG1 - �R�[�hNo.
    ' �@�@�@�@   ARG2 - �R�[�h
    '
    ' �Ԃ�l�@ : ����
    '
    ' �@�\���� : ��������Ɋi�[�����f�[�^���R�[�h�ɂ�茟�������̂��擾����
    '
    Public Function GetCodeName(ByVal strCodeNo As String, ByVal StrCode As String, Optional ByVal intNo As Short = 1) As String
        On Error GoTo err_Renamed
        Dim i As Short '���[�v�J�E���^
        Dim j As Short '���[�v�J�E���^
        Dim ct As Short '���̎擾�p�J�E���^
        Dim strName As String = String.Empty '����

        '�f�[�^���ǂݍ��߂Ă��Ȃ��ꍇ

        If Not blnFlag Then
            GetCodeName = ""
            GoTo ext
        End If

        '���̕����񂩂�w��Ԗڂ̖��̂�Ԃ�
        For i = 0 To UBound(mArrData, 1)
            If mArrData(i).strCodeNo = strCodeNo And mArrData(i).StrCode = StrCode Then
                ct = 1
                For j = 1 To Len(mArrData(i).strNames)
                    If Mid(mArrData(i).strNames, j, 1) = "," Then
                        ct = ct + 1
                        If ct > intNo Then Exit For
                    ElseIf ct = intNo Then
                        strName = strName & Mid(mArrData(i).strNames, j, 1)
                    End If
                Next j
                Exit For
            End If
        Next i
        GetCodeName = strName

ext:
        Exit Function
err_Renamed:
        MsgBox(Err.Description)
        Resume ext
    End Function

    ' @(f)
    '
    ' �@�\�@�@ : �f�[�^�̊J��
    '
    ' �������@ : �Ȃ�
    '
    ' �Ԃ�l�@ : �Ȃ�
    '
    ' �@�\���� : ��������Ɋi�[�����f�[�^���J������
    '
    Private Sub Class_Terminate_Renamed()
        On Error GoTo err_Renamed
        Erase mArrData
ext:
        Exit Sub
err_Renamed:
        MsgBox(Err.Description)
        Resume ext
    End Sub
    Protected Overrides Sub Finalize()
        Class_Terminate_Renamed()
        MyBase.Finalize()
    End Sub

    ' @(f)
    '
    ' �@�\�@�@ : �f�[�^��1�s������
    '
    ' �������@ : �Ȃ�
    '
    ' �Ԃ�l�@ : �Ȃ�
    '
    ' �@�\���� : CSV�f�[�^��1�s������؂��ď�������
    '
    Private Function SetData() As Object
        On Error GoTo err_Renamed
        Dim strRt As String '���s����
        Dim lngLnRt As Integer '���s�����̕�����
        Dim strData As String 'CSV�t�@�C�����󂯂镶����
        Dim lnglenData As Integer 'strData�̕�����
        Dim lngRt As Integer 'strData����strRt�̈ʒu
        Dim lngCt As Integer 'Rt�̃J�E���^�CmArrData�̍s��
        Dim lngBeforeRt As Integer '�ЂƂO��lngRt
        Dim strLine As String 'CSV�t�@�C����s��
        Dim intFileNo As Short '�g�p����t�@�C��No.
        Dim bytData() As Byte '�t�@�C���̃f�[�^�i�[��

        blnFlag = True

        '���s�����̌���
        strRt = vbCrLf
        lngLnRt = Len(strRt)

        '�t�@�C���̒��g�𕶎���Ƃ��Ď擾
        intFileNo = FreeFile()
        FileOpen(intFileNo, mFileName, OpenMode.Binary, OpenAccess.Read)
        ReDim bytData(LOF(intFileNo) - 1)
        FileGet(intFileNo, bytData)
        FileClose(intFileNo)

        '�G���R�[�h
        strData = System.Text.Encoding.GetEncoding(932).GetString(bytData)

        '�z��N���A
        Erase bytData

        lnglenData = Len(strData)

        '���g����C�������̓t�@�C�������݂��Ȃ��ꍇ
        If Len(strData) = 0 Then
            blnFlag = False
            GoTo ext
        End If

        '��s������
        lngBeforeRt = 1 - lngLnRt '��s�ڂ̑O�ɉ��s������Ɖ���
        Do While lngRt < lnglenData
            lngRt = InStr(lngRt + 1, strData, strRt, CompareMethod.Binary)
            If lngRt = 0 Then Exit Do
            ReDim Preserve mArrData(lngCt)
            strLine = Mid(strData, lngBeforeRt + lngLnRt, lngRt - lngBeforeRt - lngLnRt)
            SetLine(strLine, lngCt)
            lngCt = lngCt + 1
            lngBeforeRt = lngRt
        Loop

ext:
        Exit Function
err_Renamed:
        blnFlag = False
        MsgBox(Err.Description)
        Resume ext
    End Function

    ' @(f)
    '
    ' �@�\�@�@ : �z��Ɋi�[
    '
    ' �������@ : ARG1 - ��s���̕�����
    ' �@�@�@�@ : ARG2 - ���݂̍s�ԍ�
    '
    ' �Ԃ�l�@ : �Ȃ�
    '
    ' �@�\���� : 1�s�����\���̂ɕϊ����Ĕz��Ɋi�[����
    '
    Private Function SetLine(ByRef strLine As String, ByRef lngCt As Integer) As Object
        On Error GoTo err_Renamed
        Dim bytFieldCt As Byte '�t�B�[���h�i��j��
        Dim strDelimiter As String '��؂�q
        Dim lngDelimiter As Integer '��؂�q�̈ʒu
        Dim lngBeforeDel As Integer '�O�̋�؂�q�̈ʒu
        Dim strWord As String '�t�B�[���h1���̕�����
        Dim udtWords As clsCodeConv.mudtCodeLine = New mudtCodeLine() '��s����strWord���i�[

        '��؂�q�̌���
        strDelimiter = ","

        '���[�U��`�^mudtCodeLine�ɕϊ�
        Do While bytFieldCt <= 2
            If bytFieldCt < 2 Then
                lngDelimiter = InStr(lngDelimiter + 1, strLine, strDelimiter, CompareMethod.Binary)
            Else
                lngDelimiter = Len(strLine) + 1
            End If

            '�t�B�[���h��2�ȉ��̏ꍇ
            If lngDelimiter = 0 Then MsgBox("CSV�t�@�C�����s���ł�") : blnFlag = False : GoTo ext

            strWord = Mid(strLine, lngBeforeDel + 1, lngDelimiter - lngBeforeDel - 1)

            Select Case bytFieldCt
                Case 0
                    udtWords.strCodeNo = strWord
                Case 1
                    udtWords.StrCode = strWord
                Case 2
                    udtWords.strNames = strWord
                Case Else
                    GoTo ext
            End Select

            bytFieldCt = bytFieldCt + 1
            lngBeforeDel = lngDelimiter
        Loop

        '���[�U��`�^mudtCodeLine��z��ɑ��
        mArrData(lngCt) = udtWords

ext:
        Exit Function
err_Renamed:
        MsgBox(Err.Description)
        Resume ext
    End Function
End Class