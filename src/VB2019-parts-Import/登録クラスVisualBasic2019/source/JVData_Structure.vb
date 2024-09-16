Module JVLink_Stluct
    '========================================================================
    '  JRA-VAN Data Lab. JV-Data�\����
    '
    '
    '	�쐬: JRA-VAN �\�t�g�E�F�A�H�[
    '	�X�V:                           2009�N 9�� 8��
    '
    '========================================================================
    '	(C) Copyright JRA SYSTEM SERVICE CO.,LTD. 2007 All rights reserved
    '========================================================================


    '''''''''''''''''''' �Z�b�g�f�[�^�̃v���O���~���O�p�[�c '''''''''''''''''''''''''''''''''''''

    '------------------------------------------------------------------------
    '�@�@��������o�C�g���Ő؏o��
    '------------------------------------------------------------------------
    '�@ [����]
    '		myByte			= ������
    '		strStart		= �J�n�ʒu
    '		strLength		= �o�C�g��
    '	[�߂�l]
    '		String			= ������
    '------------------------------------------------------------------------
    Public Function MidB2S(ByRef myByte As Byte(), _
          ByVal bSt As Long, _
          ByVal bLen As Long) As String
        '������C�ӂɐ؏o��
        MidB2S = System.Text.Encoding.GetEncoding(932).GetString(myByte, bSt - 1, bLen)
    End Function

    '------------------------------------------------------------------------
    '�@�@�o�C�g�z����o�C�g���Ő؏o��
    '------------------------------------------------------------------------
    '�@ [����]
    '		myByte			= ������
    '		strStart		= �J�n�ʒu
    '		strLength		= �o�C�g��
    '	[�߂�l]
    '		String			= ������
    '------------------------------------------------------------------------
    Public Function MidB2B(ByRef myByte As Byte(), _
           ByVal bSt As Long, _
           ByVal bLen As Long) As Byte()
        Dim cBt As Byte()
        ReDim cBt(bLen - 1)
        ReDim MidB2B(bLen - 1)

        '������o�C�g�C�Ӑ؂�o��
        Dim i, j As Integer
        j = 0
        i = 0
        For i = bSt - 1 To bSt - 1 + bLen - 1
            cBt(j) = myByte(i)
            j = j + 1
        Next
        MidB2B = cBt
    End Function

    '------------------------------------------------------------------------
    '�@�@��������o�C�g�z��ɕϊ�
    '------------------------------------------------------------------------
    '�@ [����]
    '		myString		= ������
    '	[�߂�l]
    '		Byte()			= �o�C�g�z��
    '------------------------------------------------------------------------
    Public Function Str2Byte(ByRef myString As String) As Byte()
        'Shift JIS�ɕϊ�����
        Str2Byte = System.Text.Encoding.GetEncoding(932).GetBytes(myString)
    End Function


    '''''''''''''''''''' ���ʍ\���� ''''''''''''''''''''''''''''''''''''''''

    '<�N����>
    Public Structure YMD
        Public Year As String     ''�N
        Public Month As String     ''��
        Public Day As String     ''��
        '�f�[�^�Z�b�g
        Public Sub SetDataB(ByVal bBuff As Byte())
            Year = MidB2S(bBuff, 1, 4)
            Month = MidB2S(bBuff, 5, 2)
            Day = MidB2S(bBuff, 7, 2)
        End Sub
    End Structure

    '<�����b>
    Public Structure HMS
        Public Hour As String     ''��
        Public Minute As String     ''��
        Public Second As String     ''�b
        '�f�[�^�Z�b�g
        Public Sub SetDataB(ByVal bBuff As Byte())
            Hour = MidB2S(bBuff, 1, 2)
            Minute = MidB2S(bBuff, 3, 2)
            Second = MidB2S(bBuff, 5, 2)
        End Sub
    End Structure

    '<����>
    Public Structure HM
        Public Hour As String     ''��
        Public Minute As String     ''��
        '�f�[�^�Z�b�g
        Public Sub SetDataB(ByVal bBuff As Byte())
            Hour = MidB2S(bBuff, 1, 2)
            Minute = MidB2S(bBuff, 3, 2)
        End Sub
    End Structure

    '<��������>
    Public Structure MDHM
        Public Month As String     ''��
        Public Day As String     ''��
        Public Hour As String     ''��
        Public Minute As String     ''��
        '�f�[�^�Z�b�g
        Public Sub SetDataB(ByVal bBuff As Byte())
            Month = MidB2S(bBuff, 1, 2)
            Day = MidB2S(bBuff, 3, 2)
            Hour = MidB2S(bBuff, 5, 2)
            Minute = MidB2S(bBuff, 7, 2)
        End Sub
    End Structure

    '<���R�[�h�w�b�_>
    Public Structure RECORD_ID
        Public RecordSpec As String    ''���R�[�h���
        Public DataKubun As String    ''�f�[�^�敪
        Public MakeDate As YMD     ''�f�[�^�쐬�N����
        '�f�[�^�Z�b�g
        Public Sub SetDataB(ByVal bBuff As Byte())
            RecordSpec = MidB2S(bBuff, 1, 2)
            DataKubun = MidB2S(bBuff, 3, 1)
            MakeDate.SetDataB(MidB2B(bBuff, 4, 8))
        End Sub
    End Structure

    '<�������ʏ��>
    Public Structure RACE_ID
        Public Year As String     ''�J�ÔN
        Public MonthDay As String    ''�J�Ì���
        Public JyoCD As String     ''���n��R�[�h
        Public Kaiji As String     ''�J�É�[��N��]
        Public Nichiji As String    ''�J�Ó���[N����]
        Public RaceNum As String    ''���[�X�ԍ�
        '�f�[�^�Z�b�g
        Public Sub SetDataB(ByVal bBuff As Byte())
            Year = MidB2S(bBuff, 1, 4)
            MonthDay = MidB2S(bBuff, 5, 4)
            JyoCD = MidB2S(bBuff, 9, 2)
            Kaiji = MidB2S(bBuff, 11, 2)
            Nichiji = MidB2S(bBuff, 13, 2)
            RaceNum = MidB2S(bBuff, 15, 2)
        End Sub
    End Structure

    '<�������ʏ��Q>
    Public Structure RACE_ID2
        Public Year As String     ''�J�ÔN
        Public MonthDay As String    ''�J�Ì���
        Public JyoCD As String     ''���n��R�[�h
        Public Kaiji As String     ''�J�É�[��N��]
        Public Nichiji As String    ''�J�Ó���[N����]
        '�f�[�^�Z�b�g
        Public Sub SetDataB(ByVal bBuff As Byte())
            Year = MidB2S(bBuff, 1, 4)
            MonthDay = MidB2S(bBuff, 5, 4)
            JyoCD = MidB2S(bBuff, 9, 2)
            Kaiji = MidB2S(bBuff, 11, 2)
            Nichiji = MidB2S(bBuff, 13, 2)
        End Sub
    End Structure

    '<�{�N�E�݌v���я��>
    Public Structure SEI_RUIKEI_INFO
        Public SetYear As String    ''�ݒ�N
        Public HonSyokinTotal As String   ''�{�܋����v
        Public FukaSyokin As String    ''�t���܋����v
        Public ChakuKaisu() As String   ''����
        '�z��̏�����
        Public Sub Initialize()
            ReDim ChakuKaisu(5)
        End Sub
        '�f�[�^�Z�b�g
        Public Sub SetDataB(ByVal bBuff As Byte())
            Initialize()      ''�z��̏�����
            SetYear = MidB2S(bBuff, 1, 4)
            HonSyokinTotal = MidB2S(bBuff, 5, 10)
            FukaSyokin = MidB2S(bBuff, 15, 10)
            Dim i As Integer = 0
            For i = 0 To 5
                ChakuKaisu(i) = MidB2S(bBuff, 25 + 6 * i, 6)
            Next i
        End Sub
    End Structure

    '<�ŋߏd�܏������>
    Public Structure SAIKIN_JYUSYO_INFO
        Public SaikinJyusyoid As RACE_ID  ''<�N��������R>
        Public Hondai As String     ''�������{��
        Public Ryakusyo10 As String    ''����������10��
        Public Ryakusyo6 As String    ''����������6��
        Public Ryakusyo3 As String    ''����������3��
        Public GradeCD As String    ''�O���[�h�R�[�h
        Public SyussoTosu As String    ''�o������
        Public KettoNum As String    ''�����o�^�ԍ�
        Public Bamei As String     ''�n��
        '�f�[�^�Z�b�g
        Public Sub SetDataB(ByVal bBuff As Byte())
            SaikinJyusyoid.SetDataB(MidB2B(bBuff, 1, 16))
            Hondai = MidB2S(bBuff, 17, 60)
            Ryakusyo10 = MidB2S(bBuff, 77, 20)
            Ryakusyo6 = MidB2S(bBuff, 97, 12)
            Ryakusyo3 = MidB2S(bBuff, 109, 6)
            GradeCD = MidB2S(bBuff, 115, 1)
            SyussoTosu = MidB2S(bBuff, 116, 2)
            KettoNum = MidB2S(bBuff, 118, 10)
            Bamei = MidB2S(bBuff, 128, 36)
        End Sub
    End Structure

    '<�{�N�E�O�N�E�݌v���я��>
    Public Structure HON_ZEN_RUIKEISEI_INFO
        Public SetYear As String    ''�ݒ�N
        Public HonSyokinHeichi As String  ''���n�{�܋����v
        Public HonSyokinSyogai As String  ''��Q�{�܋����v
        Public FukaSyokinHeichi As String  ''���n�t���܋����v
        Public FukaSyokinSyogai As String  ''��Q�t���܋����v
        Public ChakuKaisuHeichi As CHAKUKAISU6_INFO  ''���n����
        Public ChakuKaisuSyogai As CHAKUKAISU6_INFO  ''��Q����
        Public ChakuKaisuJyo() As CHAKUKAISU6_INFO  ''���n��ʒ���
        Public ChakuKaisuKyori() As CHAKUKAISU6_INFO ''�����ʒ���
        '�z��̏�����
        Public Sub Initialize()
            ReDim ChakuKaisuJyo(19)
            ReDim ChakuKaisuKyori(5)
        End Sub
        '�f�[�^�Z�b�g
        Public Sub SetDataB(ByVal bBuff As Byte())
            Initialize()      ''�z��̏�����
            SetYear = MidB2S(bBuff, 1, 4)
            HonSyokinHeichi = MidB2S(bBuff, 5, 10)
            HonSyokinSyogai = MidB2S(bBuff, 15, 10)
            FukaSyokinHeichi = MidB2S(bBuff, 25, 10)
            FukaSyokinSyogai = MidB2S(bBuff, 35, 10)
            ChakuKaisuHeichi.SetDataB(MidB2B(bBuff, 45, 36))
            ChakuKaisuSyogai.SetDataB(MidB2B(bBuff, 81, 36))
            Dim i As Integer = 0
            For i = 0 To 19
                ChakuKaisuJyo(i).SetDataB(MidB2B(bBuff, 117 + 36 * i, 36))
            Next i
            For i = 0 To 5
                ChakuKaisuKyori(i).SetDataB(MidB2B(bBuff, 837 + 36 * i, 36))
            Next i
        End Sub
    End Structure

    '<���[�X���>
    Public Structure RACE_INFO
        Public YoubiCD As String    ''�j���R�[�h
        Public TokuNum As String    ''���ʋ����ԍ�
        Public Hondai As String     ''�������{��
        Public Fukudai As String    ''����������
        Public Kakko As String     ''�������J�b�R��
        Public HondaiEng As String    ''�������{�艢��
        Public FukudaiEng As String    ''���������艢��
        Public KakkoEng As String    ''�������J�b�R������
        Public Ryakusyo10 As String    ''���������̂P�O��
        Public Ryakusyo6 As String    ''���������̂U��
        Public Ryakusyo3 As String    ''���������̂R��
        Public Kubun As String     ''�������敪
        Public Nkai As String     ''�d�܉�[��N��]
        '�f�[�^�Z�b�g
        Public Sub SetDataB(ByVal bBuff As Byte())
            YoubiCD = MidB2S(bBuff, 1, 1)
            TokuNum = MidB2S(bBuff, 2, 4)
            Hondai = MidB2S(bBuff, 6, 60)
            Fukudai = MidB2S(bBuff, 66, 60)
            Kakko = MidB2S(bBuff, 126, 60)
            HondaiEng = MidB2S(bBuff, 186, 120)
            FukudaiEng = MidB2S(bBuff, 306, 120)
            KakkoEng = MidB2S(bBuff, 426, 120)
            Ryakusyo10 = MidB2S(bBuff, 546, 20)
            Ryakusyo6 = MidB2S(bBuff, 566, 12)
            Ryakusyo3 = MidB2S(bBuff, 578, 6)
            Kubun = MidB2S(bBuff, 584, 1)
            Nkai = MidB2S(bBuff, 585, 3)
        End Sub
    End Structure

    '<�V��E�n����>
    Public Structure TENKO_BABA_INFO
        Public TenkoCD As String    ''�V��R�[�h
        Public SibaBabaCD As String    ''�Ŕn���ԃR�[�h
        Public DirtBabaCD As String    ''�_�[�g�n���ԃR�[�h
        '�f�[�^�Z�b�g
        Public Sub SetDataB(ByVal bBuff As Byte())
            TenkoCD = MidB2S(bBuff, 1, 1)
            SibaBabaCD = MidB2S(bBuff, 2, 1)
            DirtBabaCD = MidB2S(bBuff, 3, 1)
        End Sub
    End Structure

    '<���񐔁i�T�C�Y3byte�j>
    Public Structure CHAKUKAISU3_INFO
        Public Chakukaisu() As String
        '�z��̏�����
        Public Sub Initialize()
            ReDim Chakukaisu(5)
        End Sub
        '�f�[�^�Z�b�g
        Public Sub SetDataB(ByVal bBuff As Byte())
            Initialize()      ''�z��̏�����
            Dim i As Integer = 0
            For i = 0 To 5
                Chakukaisu(i) = MidB2S(bBuff, 1 + 3 * i, 3)
            Next i
        End Sub
    End Structure

    '<���񐔁i�T�C�Y4byte�j>
    Public Structure CHAKUKAISU4_INFO
        Public Chakukaisu() As String
        '�z��̏�����
        Public Sub Initialize()
            ReDim Chakukaisu(5)
        End Sub
        '�f�[�^�Z�b�g
        Public Sub SetDataB(ByVal bBuff As Byte())
            Initialize()      ''�z��̏�����
            Dim i As Integer = 0
            For i = 0 To 5
                Chakukaisu(i) = MidB2S(bBuff, 1 + 4 * i, 4)
            Next i
        End Sub
    End Structure

    '<���񐔁i�T�C�Y5byte�j>
    Public Structure CHAKUKAISU5_INFO
        Public Chakukaisu() As String
        '�z��̏�����
        Public Sub Initialize()
            ReDim Chakukaisu(5)
        End Sub
        '�f�[�^�Z�b�g
        Public Sub SetDataB(ByVal bBuff As Byte())
            Initialize()      ''�z��̏�����
            Dim i As Integer = 0
            For i = 0 To 5
                Chakukaisu(i) = MidB2S(bBuff, 1 + 5 * i, 5)
            Next i
        End Sub
    End Structure

    '<���񐔁i�T�C�Y6byte�j>
    Public Structure CHAKUKAISU6_INFO
        Public Chakukaisu() As String
        '�z��̏�����
        Public Sub Initialize()
            ReDim Chakukaisu(5)
        End Sub
        '�f�[�^�Z�b�g
        Public Sub SetDataB(ByVal bBuff As Byte())
            Initialize()      ''�z��̏�����
            Dim i As Integer = 0
            For i = 0 To 5
                Chakukaisu(i) = MidB2S(bBuff, 1 + (6 * i), 6)
            Next i
        End Sub
    End Structure

    '<���������R�[�h>
    Public Structure RACE_JYOKEN
        Public SyubetuCD As String      ''������ʃR�[�h
        Public KigoCD As String       ''�����L���R�[�h
        Public JyuryoCD As String      ''�d�ʎ�ʃR�[�h
        Public JyokenCD() As String      ''���������R�[�h
        '�z��̏�����
        Public Sub Initialize()
            ReDim JyokenCD(4)
        End Sub
        '�f�[�^�Z�b�g
        Public Sub SetDataB(ByVal bBuff As Byte())
            Initialize()      ''�z��̏�����
            SyubetuCD = MidB2S(bBuff, 1, 2)
            KigoCD = MidB2S(bBuff, 3, 3)
            JyuryoCD = MidB2S(bBuff, 6, 1)
            Dim i As Integer = 0
            For i = 0 To 4
                JyokenCD(i) = MidB2S(bBuff, 7 + 3 * i, 3)
            Next i
        End Sub
    End Structure

    '''''''''''''''''''' �f�[�^�\���� ''''''''''''''''''''''''''''''

    '****** �P�D���ʓo�^�n ****************************************
    '<�o�^�n�����>
    Public Structure TOKUUMA_INFO
        Public Num As String     ''�A��
        Public KettoNum As String    ''�����o�^�ԍ�
        Public Bamei As String     ''�n��
        Public UmaKigoCD As String    ''�n�L���R�[�h
        Public SexCD As String     ''���ʃR�[�h
        Public TozaiCD As String    ''�����t���������R�[�h
        Public ChokyosiCode As String   ''�����t�R�[�h
        Public ChokyosiRyakusyo As String  ''�����t������
        Public Futan As String     ''���S�d��
        Public Koryu As String     ''�𗬋敪
        '�f�[�^�Z�b�g
        Public Sub SetDataB(ByVal bBuff As Byte())
            Num = MidB2S(bBuff, 1, 3)
            KettoNum = MidB2S(bBuff, 4, 10)
            Bamei = MidB2S(bBuff, 14, 36)
            UmaKigoCD = MidB2S(bBuff, 50, 2)
            SexCD = MidB2S(bBuff, 52, 1)
            TozaiCD = MidB2S(bBuff, 53, 1)
            ChokyosiCode = MidB2S(bBuff, 54, 5)
            ChokyosiRyakusyo = MidB2S(bBuff, 59, 8)
            Futan = MidB2S(bBuff, 67, 3)
            Koryu = MidB2S(bBuff, 70, 1)
        End Sub
    End Structure
    Public Structure JV_TK_TOKUUMA
        Public head As RECORD_ID    ''<���R�[�h�w�b�_�[>
        Public id As RACE_ID     ''<�������ʏ��>
        Public RaceInfo As RACE_INFO   ''<���[�X���>
        Public GradeCD As String    ''�O���[�h�R�[�h
        Public JyokenInfo As RACE_JYOKEN  ''<���������R�[�h>
        Public Kyori As String     ''����
        Public TrackCD As String    ''�g���b�N�R�[�h
        Public CourseKubunCD As String   ''�R�[�X�敪
        Public HandiDate As YMD     ''�n���f���\��
        Public TorokuTosu As String    ''�o�^����
        Public TokuUmaInfo() As TOKUUMA_INFO ''<�o�^�n�����>
        Public crlf As String     ''���R�[�h���
        '�z��̏�����
        Public Sub Initialize()
            ReDim TokuUmaInfo(299)
        End Sub
        '�f�[�^�Z�b�g
        Public Sub SetData(ByRef strBuff As String)
            Initialize()      ''�z��̏�����
            Dim bBuff As Byte()
            Dim bSize As Long
            bSize = 21657
            bBuff = New Byte(bSize) {}
            bBuff = Str2Byte(strBuff)

            head.SetDataB(MidB2B(bBuff, 1, 11))
            id.SetDataB(MidB2B(bBuff, 12, 16))
            RaceInfo.SetDataB(MidB2B(bBuff, 28, 587))
            GradeCD = MidB2S(bBuff, 615, 1)
            JyokenInfo.SetDataB(MidB2B(bBuff, 616, 21))
            Kyori = MidB2S(bBuff, 637, 4)
            TrackCD = MidB2S(bBuff, 641, 2)
            CourseKubunCD = MidB2S(bBuff, 643, 2)
            HandiDate.SetDataB(MidB2B(bBuff, 645, 8))
            TorokuTosu = MidB2S(bBuff, 653, 3)
            Dim i As Integer
            For i = 0 To 299
                TokuUmaInfo(i).SetDataB(MidB2B(bBuff, 656 + 70 * i, 70))

            Next i
            crlf = MidB2S(bBuff, 21656, 2)
            bBuff = Nothing
        End Sub
    End Structure

    '****** �Q�D���[�X�ڍ� ****************************************
    '<�R�[�i�[�ʉߏ���>
    Public Structure CORNER_INFO
        Public Corner As String     ''�R�[�i�[
        Public Syukaisu As String    ''����
        Public Jyuni As String     ''�e�ʉߏ���
        '�f�[�^�Z�b�g
        Public Sub SetDataB(ByVal bBuff As Byte())
            Corner = MidB2S(bBuff, 1, 1)
            Syukaisu = MidB2S(bBuff, 2, 1)
            Jyuni = MidB2S(bBuff, 3, 70)
        End Sub
    End Structure
    Public Structure JV_RA_RACE
        Public head As RECORD_ID    ''<���R�[�h�w�b�_�[>
        Public id As RACE_ID     ''<�������ʏ��>
        Public RaceInfo As RACE_INFO   ''<���[�X���>
        Public GradeCD As String    ''�O���[�h�R�[�h
        Public GradeCDBefore As String   ''�ύX�O�O���[�h�R�[�h
        Public JyokenInfo As RACE_JYOKEN  ''<���������R�[�h>
        Public JyokenName As String    ''������������
        Public Kyori As String     ''����
        Public KyoriBefore As String   ''�ύX�O����
        Public TrackCD As String    ''�g���b�N�R�[�h
        Public TrackCDBefore As String   ''�ύX�O�g���b�N�R�[�h
        Public CourseKubunCD As String   ''�R�[�X�敪
        Public CourseKubunCDBefore As String ''�ύX�O�R�[�X�敪
        Public Honsyokin() As String   ''�{�܋�
        Public HonsyokinBefore() As String  ''�ύX�O�{�܋�
        Public Fukasyokin() As String   ''�t���܋�
        Public FukasyokinBefore() As String  ''�ύX�O�t���܋�
        Public HassoTime As String    ''��������
        Public HassoTimeBefore As String  ''�ύX�O��������
        Public TorokuTosu As String    ''�o�^����
        Public SyussoTosu As String    ''�o������
        Public NyusenTosu As String    ''��������
        Public TenkoBaba As TENKO_BABA_INFO  ''�V��E�n���ԃR�[�h
        Public LapTime() As String    ''���b�v�^�C��
        Public SyogaiMileTime As String   ''��Q�}�C���^�C��
        Public HaronTimeS3 As String   ''�O�R�n�����^�C��
        Public HaronTimeS4 As String   ''�O�S�n�����^�C��
        Public HaronTimeL3 As String   ''��R�n�����^�C��
        Public HaronTimeL4 As String   ''��S�n�����^�C��
        Public CornerInfo() As CORNER_INFO  ''<�R�[�i�[�ʉߏ���>
        Public RecordUpKubun As String   ''���R�[�h�X�V�敪
        Public crlf As String     ''���R�[�h��؂�
        '�z��̏�����
        Public Sub Initialize()
            ReDim Honsyokin(6)
            ReDim HonsyokinBefore(4)
            ReDim Fukasyokin(4)
            ReDim FukasyokinBefore(2)
            ReDim LapTime(24)
            ReDim CornerInfo(3)
        End Sub
        '�f�[�^�Z�b�g
        Public Sub SetData(ByRef strBuff As String)
            Initialize()      ''�z��̏�����
            Dim i As Integer
            Dim bBuff As Byte()
            Dim bSize As Long
            bSize = 1272
            bBuff = New Byte(bSize) {}
            bBuff = Str2Byte(strBuff)

            head.SetDataB(MidB2B(bBuff, 1, 11))
            id.SetDataB(MidB2B(bBuff, 12, 16))
            RaceInfo.SetDataB(MidB2B(bBuff, 28, 587))
            GradeCD = MidB2S(bBuff, 615, 1)
            GradeCDBefore = MidB2S(bBuff, 616, 1)
            JyokenInfo.SetDataB(MidB2B(bBuff, 617, 21))
            JyokenName = MidB2S(bBuff, 638, 60)
            Kyori = MidB2S(bBuff, 698, 4)
            KyoriBefore = MidB2S(bBuff, 702, 4)
            TrackCD = MidB2S(bBuff, 706, 2)
            TrackCDBefore = MidB2S(bBuff, 708, 2)
            CourseKubunCD = MidB2S(bBuff, 710, 2)
            CourseKubunCDBefore = MidB2S(bBuff, 712, 2)
            For i = 0 To 6
                Honsyokin(i) = MidB2S(bBuff, 714 + 8 * i, 8)
            Next i
            For i = 0 To 4
                HonsyokinBefore(i) = MidB2S(bBuff, 770 + 8 * i, 8)
            Next i
            For i = 0 To 4
                Fukasyokin(i) = MidB2S(bBuff, 810 + 8 * i, 8)
            Next i
            For i = 0 To 2
                FukasyokinBefore(i) = MidB2S(bBuff, 850 + 8 * i, 8)
            Next i
            HassoTime = MidB2S(bBuff, 874, 4)
            HassoTimeBefore = MidB2S(bBuff, 878, 4)
            TorokuTosu = MidB2S(bBuff, 882, 2)
            SyussoTosu = MidB2S(bBuff, 884, 2)
            NyusenTosu = MidB2S(bBuff, 886, 2)
            TenkoBaba.SetDataB(MidB2B(bBuff, 888, 3))
            For i = 0 To 24
                LapTime(i) = MidB2S(bBuff, 891 + 3 * i, 3)
            Next i
            SyogaiMileTime = MidB2S(bBuff, 966, 4)
            HaronTimeS3 = MidB2S(bBuff, 970, 3)
            HaronTimeS4 = MidB2S(bBuff, 973, 3)
            HaronTimeL3 = MidB2S(bBuff, 976, 3)
            HaronTimeL4 = MidB2S(bBuff, 979, 3)
            For i = 0 To 3
                CornerInfo(i).SetDataB(MidB2B(bBuff, 982 + 72 * i, 72))
            Next i
            RecordUpKubun = MidB2S(bBuff, 1270, 1)
            crlf = MidB2S(bBuff, 1271, 2)
            bBuff = Nothing
        End Sub
    End Structure

    '****** �R�D�n�����[�X��� ****************************************
    '<1���n(����n)���>
    Public Structure CHAKUUMA_INFO
        Public KettoNum As String    ''�����o�^�ԍ�
        Public Bamei As String     ''�n��
        '�f�[�^�Z�b�g
        Public Sub SetDataB(ByVal bBuff As Byte())
            KettoNum = MidB2S(bBuff, 1, 10)
            Bamei = MidB2S(bBuff, 11, 36)
        End Sub
    End Structure
    Public Structure JV_SE_RACE_UMA
        Public head As RECORD_ID    ''<���R�[�h�w�b�_�[>
        Public id As RACE_ID     ''<�������ʏ��>
        Public Wakuban As String       ''�g��
        Public Umaban As String     ''�n��
        Public KettoNum As String    ''�����o�^�ԍ�
        Public Bamei As String     ''�n��
        Public UmaKigoCD As String    ''�n�L���R�[�h
        Public SexCD As String     ''���ʃR�[�h
        Public HinsyuCD As String    ''�i��R�[�h
        Public KeiroCD As String    ''�ѐF�R�[�h
        Public Barei As String     ''�n��
        Public TozaiCD As String    ''���������R�[�h
        Public ChokyosiCode As String   ''�����t�R�[�h
        Public ChokyosiRyakusyo As String  ''�����t������
        Public BanusiCode As String    ''�n��R�[�h
        Public BanusiName As String    ''�n�喼
        Public Fukusyoku As String    ''���F�W��
        Public reserved1 As String    ''�\��
        Public Futan As String     ''���S�d��
        Public FutanBefore As String   ''�ύX�O���S�d��
        Public Blinker As String    ''�u�����J�[�g�p�敪
        Public reserved2 As String    ''�\��
        Public KisyuCode As String    ''�R��R�[�h
        Public KisyuCodeBefore As String  ''�ύX�O�R��R�[�h
        Public KisyuRyakusyo As String   ''�R�薼����
        Public KisyuRyakusyoBefore As String ''�ύX�O�R�薼����
        Public MinaraiCD As String    ''�R�茩�K�R�[�h
        Public MinaraiCDBefore As String  ''�ύX�O�R�茩�K�R�[�h
        Public BaTaijyu As String    ''�n�̏d
        Public ZogenFugo As String    ''��������
        Public ZogenSa As String    ''������
        Public IJyoCD As String     ''�ُ�敪�R�[�h
        Public NyusenJyuni As String   ''��������
        Public KakuteiJyuni As String   ''�m�蒅��
        Public DochakuKubun As String   ''�����敪
        Public DochakuTosu As String   ''��������
        Public Time As String     ''���j�^�C��
        Public ChakusaCD As String    ''�����R�[�h
        Public ChakusaCDP As String    ''+�����R�[�h
        Public ChakusaCDPP As String   ''++�����R�[�h
        Public Jyuni1c As String    ''1�R�[�i�[�ł̏���
        Public Jyuni2c As String    ''2�R�[�i�[�ł̏���
        Public Jyuni3c As String    ''3�R�[�i�[�ł̏���
        Public Jyuni4c As String    ''4�R�[�i�[�ł̏���
        Public Odds As String     ''�P���I�b�Y
        Public Ninki As String     ''�P���l�C��
        Public Honsyokin As String    ''�l���{�܋�
        Public Fukasyokin As String    ''�l���t���܋�
        Public reserved3 As String    ''�\��
        Public reserved4 As String    ''�\��
        Public HaronTimeL4 As String   ''��S�n�����^�C��
        Public HaronTimeL3 As String   ''��R�n�����^�C��
        Public ChakuUmaInfo() As CHAKUUMA_INFO ''<1���n(����n)���>
        Public TimeDiff As String    ''�^�C����
        Public RecordUpKubun As String   ''���R�[�h�X�V�敪
        Public DMKubun As String    ''�}�C�j���O�敪
        Public DMTime As String     ''�}�C�j���O�\�z���j�^�C��
        Public DMGosaP As String    ''�\���덷(�M���x)�{
        Public DMGosaM As String    ''�\���덷(�M���x)�|
        Public DMJyuni As String    ''�}�C�j���O�\�z����
        Public KyakusituKubun As String   ''���񃌁[�X�r������
        Public crlf As String     ''���R�[�h��؂�
        '�z��̏�����
        Public Sub Initialize()
            ReDim ChakuUmaInfo(2)
        End Sub
        '�f�[�^�Z�b�g
        Public Sub SetData(ByRef strBuff As String)
            Initialize()      ''�z��̏�����
            Dim i As Integer
            Dim bBuff As Byte()
            Dim bSize As Long
            bSize = 555
            bBuff = New Byte(bSize) {}
            bBuff = Str2Byte(strBuff)

            head.SetDataB(MidB2B(bBuff, 1, 11))
            id.SetDataB(MidB2B(bBuff, 12, 16))
            Wakuban = MidB2S(bBuff, 28, 1)
            Umaban = MidB2S(bBuff, 29, 2)
            KettoNum = MidB2S(bBuff, 31, 10)
            Bamei = MidB2S(bBuff, 41, 36)
            UmaKigoCD = MidB2S(bBuff, 77, 2)
            SexCD = MidB2S(bBuff, 79, 1)
            HinsyuCD = MidB2S(bBuff, 80, 1)
            KeiroCD = MidB2S(bBuff, 81, 2)
            Barei = MidB2S(bBuff, 83, 2)
            TozaiCD = MidB2S(bBuff, 85, 1)
            ChokyosiCode = MidB2S(bBuff, 86, 5)
            ChokyosiRyakusyo = MidB2S(bBuff, 91, 8)
            BanusiCode = MidB2S(bBuff, 99, 6)
            BanusiName = MidB2S(bBuff, 105, 64)
            Fukusyoku = MidB2S(bBuff, 169, 60)
            reserved1 = MidB2S(bBuff, 229, 60)
            Futan = MidB2S(bBuff, 289, 3)
            FutanBefore = MidB2S(bBuff, 292, 3)
            Blinker = MidB2S(bBuff, 295, 1)
            reserved2 = MidB2S(bBuff, 296, 1)
            KisyuCode = MidB2S(bBuff, 297, 5)
            KisyuCodeBefore = MidB2S(bBuff, 302, 5)
            KisyuRyakusyo = MidB2S(bBuff, 307, 8)
            KisyuRyakusyoBefore = MidB2S(bBuff, 315, 8)
            MinaraiCD = MidB2S(bBuff, 323, 1)
            MinaraiCDBefore = MidB2S(bBuff, 324, 1)
            BaTaijyu = MidB2S(bBuff, 325, 3)
            ZogenFugo = MidB2S(bBuff, 328, 1)
            ZogenSa = MidB2S(bBuff, 329, 3)
            IJyoCD = MidB2S(bBuff, 332, 1)
            NyusenJyuni = MidB2S(bBuff, 333, 2)
            KakuteiJyuni = MidB2S(bBuff, 335, 2)
            DochakuKubun = MidB2S(bBuff, 337, 1)
            DochakuTosu = MidB2S(bBuff, 338, 1)
            Time = MidB2S(bBuff, 339, 4)
            ChakusaCD = MidB2S(bBuff, 343, 3)
            ChakusaCDP = MidB2S(bBuff, 346, 3)
            ChakusaCDPP = MidB2S(bBuff, 349, 3)
            Jyuni1c = MidB2S(bBuff, 352, 2)
            Jyuni2c = MidB2S(bBuff, 354, 2)
            Jyuni3c = MidB2S(bBuff, 356, 2)
            Jyuni4c = MidB2S(bBuff, 358, 2)
            Odds = MidB2S(bBuff, 360, 4)
            Ninki = MidB2S(bBuff, 364, 2)
            Honsyokin = MidB2S(bBuff, 366, 8)
            Fukasyokin = MidB2S(bBuff, 374, 8)
            reserved3 = MidB2S(bBuff, 382, 3)
            reserved4 = MidB2S(bBuff, 385, 3)
            HaronTimeL4 = MidB2S(bBuff, 388, 3)
            HaronTimeL3 = MidB2S(bBuff, 391, 3)
            For i = 0 To 2
                ChakuUmaInfo(i).SetDataB(MidB2B(bBuff, 394 + 46 * i, 46))
            Next i
            TimeDiff = MidB2S(bBuff, 532, 4)
            RecordUpKubun = MidB2S(bBuff, 536, 1)
            DMKubun = MidB2S(bBuff, 537, 1)
            DMTime = MidB2S(bBuff, 538, 5)
            DMGosaP = MidB2S(bBuff, 543, 4)
            DMGosaM = MidB2S(bBuff, 547, 4)
            DMJyuni = MidB2S(bBuff, 551, 2)
            KyakusituKubun = MidB2S(bBuff, 553, 1)
            crlf = MidB2S(bBuff, 554, 2)
            bBuff = Nothing
        End Sub
    End Structure

    '****** �S�D���� ****************************************

    ''<���ߏ��P �P�E���E�g>
    Public Structure PAY_INFO1
        Public Umaban As String     ''�n��
        Public Pay As String     ''���ߋ�
        Public Ninki As String     ''�l�C��	
        '�f�[�^�Z�b�g
        Public Sub SetDataB(ByVal bBuff As Byte())
            Umaban = MidB2S(bBuff, 1, 2)
            Pay = MidB2S(bBuff, 3, 9)
            Ninki = MidB2S(bBuff, 12, 2)
        End Sub
    End Structure

    ''<���ߏ��Q �n�A�E���C�h�E�\���E�n�P>
    Public Structure PAY_INFO2
        Public Kumi As String     ''�g��
        Public Pay As String     ''���ߋ�
        Public Ninki As String     ''�l�C��	
        '�f�[�^�Z�b�g
        Public Sub SetDataB(ByVal bBuff As Byte())
            Kumi = MidB2S(bBuff, 1, 4)
            Pay = MidB2S(bBuff, 5, 9)
            Ninki = MidB2S(bBuff, 14, 3)
        End Sub
    End Structure

    ''<���ߏ��R �R�A��>
    Public Structure PAY_INFO3
        Public Kumi As String     ''�g��
        Public Pay As String     ''���ߋ�
        Public Ninki As String     ''�l�C��	
        '�f�[�^�Z�b�g
        Public Sub SetDataB(ByVal bBuff As Byte())
            Kumi = MidB2S(bBuff, 1, 6)
            Pay = MidB2S(bBuff, 7, 9)
            Ninki = MidB2S(bBuff, 16, 3)
        End Sub
    End Structure

    ''<���ߏ��S �R�A�P>
    Public Structure PAY_INFO4
        Public Kumi As String     ''�g��
        Public Pay As String     ''���ߋ�
        Public Ninki As String     ''�l�C��
        '�f�[�^�Z�b�g
        Public Sub SetDataB(ByVal bBuff As Byte())
            Kumi = MidB2S(bBuff, 1, 6)
            Pay = MidB2S(bBuff, 7, 9)
            Ninki = MidB2S(bBuff, 16, 4)
        End Sub
    End Structure

    Public Structure JV_HR_PAY
        Public head As RECORD_ID            ''<���R�[�h�w�b�_�[>
        Public id As RACE_ID                ''<�������ʏ��>
        Public TorokuTosu As String         ''�o�^����
        Public SyussoTosu As String         ''�o������
        Public FuseirituFlag() As String    ''�s�����t���O
        Public TokubaraiFlag() As String    ''�����t���O
        Public HenkanFlag() As String       ''�Ԋ҃t���O
        Public HenkanUma() As String        ''�ԊҔn�ԏ��(�n��01�`28)
        Public HenkanWaku() As String       ''�ԊҘg�ԏ��(�g��1�`8)
        Public HenkanDoWaku() As String     ''�Ԋғ��g���(�g��1�`8)
        Public PayTansyo() As PAY_INFO1     ''<�P������>
        Public PayFukusyo() As PAY_INFO1    ''<��������>
        Public PayWakuren() As PAY_INFO1    ''<�g�A����>
        Public PayUmaren() As PAY_INFO2     ''<�n�A����>
        Public PayWide() As PAY_INFO2       ''<���C�h����>
        Public PayReserved1() As PAY_INFO2  ''<�\��>
        Public PayUmatan() As PAY_INFO2     ''<�n�P����>
        Public PaySanrenpuku() As PAY_INFO3 ''<3�A������>
        Public PaySanrentan() As PAY_INFO4  ''<3�A�P����>
        Public crlf As String     ''���R�[�h��؂�
        '�z��̏�����
        Public Sub Initialize()
            ReDim FuseirituFlag(8)
            ReDim TokubaraiFlag(8)
            ReDim HenkanFlag(8)
            ReDim HenkanUma(27)
            ReDim HenkanWaku(7)
            ReDim HenkanDoWaku(7)
            ReDim PayTansyo(2)
            ReDim PayFukusyo(4)
            ReDim PayWakuren(2)
            ReDim PayUmaren(2)
            ReDim PayWide(6)
            ReDim PayReserved1(2)
            ReDim PayUmatan(5)
            ReDim PaySanrenpuku(2)
            ReDim PaySanrentan(5)
        End Sub
        '�f�[�^�Z�b�g
        Public Sub SetData(ByRef strBuff As String)
            Initialize()      ''�z��̏�����
            Dim i As Integer
            Dim bBuff As Byte()
            Dim bSize As Long
            bSize = 719
            bBuff = New Byte(bSize) {}
            bBuff = Str2Byte(strBuff)

            head.SetDataB(MidB2B(bBuff, 1, 11))
            id.SetDataB(MidB2B(bBuff, 12, 16))
            TorokuTosu = MidB2S(bBuff, 28, 2)
            SyussoTosu = MidB2S(bBuff, 30, 2)
            For i = 0 To 8
                FuseirituFlag(i) = MidB2S(bBuff, 32 + (1 * i), 1)
            Next i
            For i = 0 To 8
                TokubaraiFlag(i) = MidB2S(bBuff, 41 + (1 * i), 1)
            Next i
            For i = 0 To 8
                HenkanFlag(i) = MidB2S(bBuff, 50 + (1 * i), 1)
            Next i
            For i = 0 To 27
                HenkanUma(i) = MidB2S(bBuff, 59 + (1 * i), 1)
            Next i
            For i = 0 To 7
                HenkanWaku(i) = MidB2S(bBuff, 87 + (1 * i), 1)
            Next i
            For i = 0 To 7
                HenkanDoWaku(i) = MidB2S(bBuff, 95 + (1 * i), 1)
            Next i
            For i = 0 To 2
                PayTansyo(i).SetDataB(MidB2B(bBuff, 103 + (13 * i), 13))
            Next i
            For i = 0 To 4
                PayFukusyo(i).SetDataB(MidB2B(bBuff, 142 + (13 * i), 13))
            Next i
            For i = 0 To 2
                PayWakuren(i).SetDataB(MidB2B(bBuff, 207 + (13 * i), 13))
            Next i
            For i = 0 To 2
                PayUmaren(i).SetDataB(MidB2B(bBuff, 246 + (16 * i), 16))
            Next i
            For i = 0 To 6
                PayWide(i).SetDataB(MidB2B(bBuff, 294 + (16 * i), 16))
            Next i
            For i = 0 To 2
                PayReserved1(i).SetDataB(MidB2B(bBuff, 406 + (16 * i), 16))
            Next i
            For i = 0 To 5
                PayUmatan(i).SetDataB(MidB2B(bBuff, 454 + (16 * i), 16))
            Next i
            For i = 0 To 2
                PaySanrenpuku(i).SetDataB(MidB2B(bBuff, 550 + (18 * i), 18))
            Next i
            For i = 0 To 5
                PaySanrentan(i).SetDataB(MidB2B(bBuff, 604 + (19 * i), 19))
            Next i
            crlf = MidB2S(bBuff, 718, 2)
            bBuff = Nothing
        End Sub
    End Structure

    '****** �T�D�[���i�S�|���j****************************************
    '<�[�����P �P�E���E�g>
    Public Structure HYO_INFO1
        Public Umaban As String     ''�n��		
        Public Hyo As String     ''�[��
        Public Ninki As String     ''�l�C
        '�f�[�^�Z�b�g
        Public Sub SetDataB(ByVal bBuff As Byte())
            Umaban = MidB2S(bBuff, 1, 2)
            Hyo = MidB2S(bBuff, 3, 11)
            Ninki = MidB2S(bBuff, 14, 2)
        End Sub
    End Structure
    '<�[�����Q �n�A�E���C�h�E�n�P>
    Public Structure HYO_INFO2
        Public Kumi As String     ''�g��		
        Public Hyo As String     ''�[��
        Public Ninki As String     ''�l�C
        '�f�[�^�Z�b�g
        Public Sub SetDataB(ByVal bBuff As Byte())
            Kumi = MidB2S(bBuff, 1, 4)
            Hyo = MidB2S(bBuff, 5, 11)
            Ninki = MidB2S(bBuff, 16, 3)
        End Sub
    End Structure
    '<�[�����R �R�A���[��>
    Public Structure HYO_INFO3
        Public Kumi As String     ''�g��		
        Public Hyo As String     ''�[��
        Public Ninki As String     ''�l�C
        '�f�[�^�Z�b�g
        Public Sub SetDataB(ByVal bBuff As Byte())
            Kumi = MidB2S(bBuff, 1, 6)
            Hyo = MidB2S(bBuff, 7, 11)
            Ninki = MidB2S(bBuff, 18, 3)
        End Sub
    End Structure
    '<�[�����S �R�A�P�[��>
    Public Structure HYO_INFO4
        Public Kumi As String     ''�g��		
        Public Hyo As String     ''�[��
        Public Ninki As String     ''�l�C
        '�f�[�^�Z�b�g
        Public Sub SetDataB(ByVal bBuff As Byte())
            Kumi = MidB2S(bBuff, 1, 6)
            Hyo = MidB2S(bBuff, 7, 11)
            Ninki = MidB2S(bBuff, 18, 4)
        End Sub
    End Structure

    Public Structure JV_H1_HYOSU_ZENKAKE
        Public head As RECORD_ID            ''<���R�[�h�w�b�_�[>
        Public id As RACE_ID                ''<�������ʏ��>
        Public TorokuTosu As String         ''�o�^����
        Public SyussoTosu As String         ''�o������
        Public HatubaiFlag() As String      ''�����t���O�@
        Public FukuChakuBaraiKey As String  ''���������L�[
        Public HenkanUma() As String        ''�ԊҔn�ԏ��(�n��01�`28)
        Public HenkanWaku() As String       ''�ԊҘg�ԏ��(�g��1�`8)
        Public HenkanDoWaku() As String     ''�Ԋғ��g���(�g��1�`8)
        Public HyoTansyo() As HYO_INFO1     ''<�P���[��>
        Public HyoFukusyo() As HYO_INFO1    ''<�����[��>
        Public HyoWakuren() As HYO_INFO1    ''<�g�A�[��>
        Public HyoUmaren() As HYO_INFO2     ''<�n�A�[��>
        Public HyoWide() As HYO_INFO2       ''<���C�h�[��>
        Public HyoUmatan() As HYO_INFO2     ''<�n�P�[��>
        Public HyoSanrenpuku() As HYO_INFO3 ''<3�A���[��>
        Public HyoTotal() As String         ''�[�����v
        Public crlf As String               ''���R�[�h��؂�
        '�z��̏�����
        Public Sub Initialize()
            ReDim HatubaiFlag(6)
            ReDim HenkanUma(27)
            ReDim HenkanWaku(7)
            ReDim HenkanDoWaku(7)
            ReDim HyoTansyo(27)
            ReDim HyoFukusyo(27)
            ReDim HyoWakuren(35)
            ReDim HyoUmaren(152)
            ReDim HyoWide(152)
            ReDim HyoUmatan(305)
            ReDim HyoSanrenpuku(815)
            ReDim HyoTotal(13)
        End Sub
        '�f�[�^�Z�b�g
        Public Sub SetData(ByRef strBuff As String)
            Initialize()      ''�z��̏�����
            Dim i As Integer
            Dim bBuff As Byte()
            Dim bSize As Long
            bSize = 28955
            bBuff = New Byte(bSize) {}
            bBuff = Str2Byte(strBuff)

            head.SetDataB(MidB2B(bBuff, 1, 11))
            id.SetDataB(MidB2B(bBuff, 12, 16))
            TorokuTosu = MidB2S(bBuff, 28, 2)
            SyussoTosu = MidB2S(bBuff, 30, 2)
            For i = 0 To 6
                HatubaiFlag(i) = MidB2S(bBuff, 32 + (1 * i), 1)
            Next i
            FukuChakuBaraiKey = MidB2S(bBuff, 39, 1)
            For i = 0 To 27
                HenkanUma(i) = MidB2S(bBuff, 40 + (1 * i), 1)
            Next i
            For i = 0 To 7
                HenkanWaku(i) = MidB2S(bBuff, 68 + (1 * i), 1)
            Next i
            For i = 0 To 7
                HenkanDoWaku(i) = MidB2S(bBuff, 76 + (1 * i), 1)
            Next i
            For i = 0 To 27
                HyoTansyo(i).SetDataB(MidB2B(bBuff, 84 + (15 * i), 15))
            Next i
            For i = 0 To 27
                HyoFukusyo(i).SetDataB(MidB2B(bBuff, 504 + (15 * i), 15))
            Next i
            For i = 0 To 35
                HyoWakuren(i).SetDataB(MidB2B(bBuff, 924 + (15 * i), 15))
            Next i
            For i = 0 To 152
                HyoUmaren(i).SetDataB(MidB2B(bBuff, 1464 + (18 * i), 18))
            Next i
            For i = 0 To 152
                HyoWide(i).SetDataB(MidB2B(bBuff, 4218 + (18 * i), 18))
            Next i
            For i = 0 To 305
                HyoUmatan(i).SetDataB(MidB2B(bBuff, 6972 + (18 * i), 18))
            Next i
            For i = 0 To 815
                HyoSanrenpuku(i).SetDataB(MidB2B(bBuff, 12480 + (20 * i), 20))
            Next i
            For i = 0 To 13
                HyoTotal(i) = MidB2S(bBuff, 28800 + (11 * i), 11)
            Next i
            crlf = MidB2S(bBuff, 28954, 2)
            bBuff = Nothing
        End Sub
    End Structure

    Public Structure JV_H6_HYOSU_SANRENTAN
        Public head As RECORD_ID            ''<���R�[�h�w�b�_�[>
        Public id As RACE_ID                ''<�������ʏ��>
        Public TorokuTosu As String         ''�o�^����
        Public SyussoTosu As String         ''�o������
        Public HatubaiFlag As String        ''�����t���O�@
        Public HenkanUma() As String        ''�ԊҔn�ԏ��(�n��01�`18)
        Public HyoSanrentan() As HYO_INFO4 ''<3�A�P�[��>
        Public HyoTotal() As String         ''�[�����v
        Public crlf As String               ''���R�[�h��؂�
        '�z��̏�����
        Public Sub Initialize()
            ReDim HenkanUma(17)
            ReDim HyoSanrentan(4895)
            ReDim HyoTotal(1)
        End Sub

        '�f�[�^�Z�b�g
        Public Sub SetData(ByRef strBuff As String)
            Initialize()      ''�z��̏�����
            Dim i As Integer
            Dim bBuff As Byte()
            Dim bSize As Long
            bSize = 102900
            bBuff = New Byte(bSize) {}
            bBuff = Str2Byte(strBuff)

            head.SetDataB(MidB2B(bBuff, 1, 11))
            id.SetDataB(MidB2B(bBuff, 12, 16))
            TorokuTosu = MidB2S(bBuff, 28, 2)
            SyussoTosu = MidB2S(bBuff, 30, 2)
            HatubaiFlag = MidB2S(bBuff, 32, 1)
            For i = 0 To 17
                HenkanUma(i) = MidB2S(bBuff, 33 + (1 * i), 1)
            Next i
            For i = 0 To 4895
                HyoSanrentan(i).SetDataB(MidB2B(bBuff, 51 + (21 * i), 21))
            Next i
            For i = 0 To 1
                HyoTotal(i) = MidB2S(bBuff, 102867 + (11 * i), 11)
            Next i
            crlf = MidB2S(bBuff, 102889, 2)
            bBuff = Nothing
        End Sub
    End Structure

    '****** �U�D�I�b�Y�i�P���g�j****************************************
    '<�P���I�b�Y>
    Public Structure ODDS_TANSYO_INFO
        Public Umaban As String     ''�n��
        Public Odds As String     ''�I�b�Y
        Public Ninki As String     ''�l�C��
        '�f�[�^�Z�b�g
        Public Sub SetDataB(ByVal bBuff As Byte())
            Umaban = MidB2S(bBuff, 1, 2)
            Odds = MidB2S(bBuff, 3, 4)
            Ninki = MidB2S(bBuff, 7, 2)
        End Sub
    End Structure
    '<�����I�b�Y>
    Public Structure ODDS_FUKUSYO_INFO
        Public Umaban As String     ''�n��
        Public OddsLow As String    ''�Œ�I�b�Y
        Public OddsHigh As String    ''�ō��I�b�Y
        Public Ninki As String     ''�l�C��
        '�f�[�^�Z�b�g
        Public Sub SetDataB(ByVal bBuff As Byte())
            Umaban = MidB2S(bBuff, 1, 2)
            OddsLow = MidB2S(bBuff, 3, 4)
            OddsHigh = MidB2S(bBuff, 7, 4)
            Ninki = MidB2S(bBuff, 11, 2)
        End Sub
    End Structure
    '<�g�A�I�b�Y>
    Public Structure ODDS_WAKUREN_INFO
        Public Kumi As String     ''�g
        Public Odds As String     ''�I�b�Y
        Public Ninki As String     ''�l�C��
        '�f�[�^�Z�b�g
        Public Sub SetDataB(ByVal bBuff As Byte())
            Kumi = MidB2S(bBuff, 1, 2)
            Odds = MidB2S(bBuff, 3, 5)
            Ninki = MidB2S(bBuff, 8, 2)
        End Sub
    End Structure
    Public Structure JV_O1_ODDS_TANFUKUWAKU
        Public head As RECORD_ID    ''<���R�[�h�w�b�_�[>
        Public id As RACE_ID     ''<�������ʏ��>
        Public HappyoTime As MDHM    ''���\��������
        Public TorokuTosu As String    ''�o�^����
        Public SyussoTosu As String    ''�o������
        Public TansyoFlag As String    ''�����t���O�@�P��
        Public FukusyoFlag As String   ''�����t���O�@����
        Public WakurenFlag As String   ''�����t���O�@�g�A
        Public FukuChakuBaraiKey As String  ''���������L�[
        Public OddsTansyoInfo() As ODDS_TANSYO_INFO  ''<�P���I�b�Y>
        Public OddsFukusyoInfo() As ODDS_FUKUSYO_INFO ''<�����[���I�b�Y>
        Public OddsWakurenInfo() As ODDS_WAKUREN_INFO ''<�g�A�[���I�b�Y>
        Public TotalHyosuTansyo As String  ''�P���[�����v
        Public TotalHyosuFukusyo As String  ''�����[�����v
        Public TotalHyosuWakuren As String  ''�g�A�[�����v
        Public crlf As String   ''���R�[�h��؂�
        '�z��̏�����
        Public Sub Initialize()
            ReDim OddsTansyoInfo(27)
            ReDim OddsFukusyoInfo(27)
            ReDim OddsWakurenInfo(35)
        End Sub
        '�f�[�^�Z�b�g
        Public Sub SetData(ByRef strBuff As String)
            Initialize()      ''�z��̏�����
            Dim i As Integer
            Dim bBuff As Byte()
            Dim bSize As Long
            bSize = 962
            bBuff = New Byte(bSize) {}
            bBuff = Str2Byte(strBuff)

            head.SetDataB(MidB2B(bBuff, 1, 11))
            id.SetDataB(MidB2B(bBuff, 12, 16))
            HappyoTime.SetDataB(MidB2B(bBuff, 28, 8))
            TorokuTosu = MidB2S(bBuff, 36, 2)
            SyussoTosu = MidB2S(bBuff, 38, 2)
            TansyoFlag = MidB2S(bBuff, 40, 1)
            FukusyoFlag = MidB2S(bBuff, 41, 1)
            WakurenFlag = MidB2S(bBuff, 42, 1)
            FukuChakuBaraiKey = MidB2S(bBuff, 43, 1)
            For i = 0 To 27
                OddsTansyoInfo(i).SetDataB(MidB2B(bBuff, 44 + (8 * i), 8))
            Next i
            For i = 0 To 27
                OddsFukusyoInfo(i).SetDataB(MidB2B(bBuff, 268 + (12 * i), 12))
            Next i
            For i = 0 To 35
                OddsWakurenInfo(i).SetDataB(MidB2B(bBuff, 604 + (9 * i), 9))
            Next i
            TotalHyosuTansyo = MidB2S(bBuff, 928, 11)
            TotalHyosuFukusyo = MidB2S(bBuff, 939, 11)
            TotalHyosuWakuren = MidB2S(bBuff, 950, 11)
            crlf = MidB2S(bBuff, 961, 2)
            bBuff = Nothing
        End Sub
    End Structure

    '****** �V�D�I�b�Y�i�n�A�j****************************************
    '<�n�A�I�b�Y>
    Public Structure ODDS_UMAREN_INFO
        Public Kumi As String     ''�g��
        Public Odds As String     ''�I�b�Y
        Public Ninki As String     ''�l�C��
        '�f�[�^�Z�b�g
        Public Sub SetDataB(ByVal bBuff As Byte())
            Kumi = MidB2S(bBuff, 1, 4)
            Odds = MidB2S(bBuff, 5, 6)
            Ninki = MidB2S(bBuff, 11, 3)
        End Sub
    End Structure
    Public Structure JV_O2_ODDS_UMAREN
        Public head As RECORD_ID    ''<���R�[�h�w�b�_�[>
        Public id As RACE_ID     ''<�������ʏ��>
        Public HappyoTime As MDHM    ''���\��������
        Public TorokuTosu As String    ''�o�^����
        Public SyussoTosu As String    ''�o������
        Public UmarenFlag As String    ''�����t���O�@�n�A
        Public OddsUmarenInfo() As ODDS_UMAREN_INFO  ''<�n�A�I�b�Y>
        Public TotalHyosuUmaren As String  ''�n�A�[�����v
        Public crlf As String     ''���R�[�h��؂�
        '�z��̏�����
        Public Sub Initialize()
            ReDim OddsUmarenInfo(152)
        End Sub
        '�f�[�^�Z�b�g
        Public Sub SetData(ByRef strBuff As String)
            Initialize()      ''�z��̏�����
            Dim i As Integer
            Dim bBuff As Byte()
            Dim bSize As Long
            bSize = 2042
            bBuff = New Byte(bSize) {}
            bBuff = Str2Byte(strBuff)

            head.SetDataB(MidB2B(bBuff, 1, 11))
            id.SetDataB(MidB2B(bBuff, 12, 16))
            HappyoTime.SetDataB(MidB2B(bBuff, 28, 8))
            TorokuTosu = MidB2S(bBuff, 36, 2)
            SyussoTosu = MidB2S(bBuff, 38, 2)
            UmarenFlag = MidB2S(bBuff, 40, 1)
            For i = 0 To 152
                OddsUmarenInfo(i).SetDataB(MidB2B(bBuff, 41 + (13 * i), 13))
            Next i
            TotalHyosuUmaren = MidB2S(bBuff, 2030, 11)
            crlf = MidB2S(bBuff, 2041, 2)
            bBuff = Nothing
        End Sub
    End Structure

    '****** �W�D�I�b�Y�i���C�h�j****************************************
    '<���C�h�I�b�Y>
    Public Structure ODDS_WIDE_INFO
        Public Kumi As String     ''�g��
        Public OddsLow As String    ''�Œ�I�b�Y
        Public OddsHigh As String    ''�ō��I�b�Y
        Public Ninki As String     ''�l�C��
        '�f�[�^�Z�b�g
        Public Sub SetDataB(ByVal bBuff As Byte())
            Kumi = MidB2S(bBuff, 1, 4)
            OddsLow = MidB2S(bBuff, 5, 5)
            OddsHigh = MidB2S(bBuff, 10, 5)
            Ninki = MidB2S(bBuff, 15, 3)
        End Sub
    End Structure
    Public Structure JV_O3_ODDS_WIDE
        Public head As RECORD_ID    ''<���R�[�h�w�b�_�[>
        Public id As RACE_ID     ''<�������ʏ��>
        Public HappyoTime As MDHM    ''���\��������
        Public TorokuTosu As String    ''�o�^����
        Public SyussoTosu As String    ''�o������
        Public WideFlag As String    ''�����t���O�@���C�h
        Public OddsWideInfo() As ODDS_WIDE_INFO ''<���C�h�I�b�Y>
        Public TotalHyosuWide As String   ''���C�h�[�����v
        Public crlf As String     ''���R�[�h��؂�
        '�z��̏�����
        Public Sub Initialize()
            ReDim OddsWideInfo(152)
        End Sub
        '�f�[�^�Z�b�g
        Public Sub SetData(ByRef strBuff As String)
            Initialize()      ''�z��̏�����
            Dim i As Integer
            Dim bBuff As Byte()
            Dim bSize As Long
            bSize = 2654
            bBuff = New Byte(bSize) {}
            bBuff = Str2Byte(strBuff)

            head.SetDataB(MidB2B(bBuff, 1, 11))
            id.SetDataB(MidB2B(bBuff, 12, 16))
            HappyoTime.SetDataB(MidB2B(bBuff, 28, 8))
            TorokuTosu = MidB2S(bBuff, 36, 2)
            SyussoTosu = MidB2S(bBuff, 38, 2)
            WideFlag = MidB2S(bBuff, 40, 1)
            For i = 0 To 152
                OddsWideInfo(i).SetDataB(MidB2B(bBuff, 41 + (17 * i), 17))
            Next i
            TotalHyosuWide = MidB2S(bBuff, 2642, 11)
            crlf = MidB2S(bBuff, 2653, 2)
            bBuff = Nothing
        End Sub
    End Structure

    '****** �X�D�I�b�Y�i�n�P�j ****************************************
    '<�n�P�I�b�Y>
    Public Structure ODDS_UMATAN_INFO
        Public Kumi As String     ''�g��
        Public Odds As String     ''�I�b�Y
        Public Ninki As String     ''�l�C��
        '�f�[�^�Z�b�g
        Public Sub SetDataB(ByVal bBuff As Byte())
            Kumi = MidB2S(bBuff, 1, 4)
            Odds = MidB2S(bBuff, 5, 6)
            Ninki = MidB2S(bBuff, 11, 3)
        End Sub
    End Structure
    Public Structure JV_O4_ODDS_UMATAN
        Public head As RECORD_ID    ''<���R�[�h�w�b�_�[>
        Public id As RACE_ID     ''<�������ʏ��>
        Public HappyoTime As MDHM    ''���\��������
        Public TorokuTosu As String    ''�o�^����
        Public SyussoTosu As String    ''�o������
        Public UmatanFlag As String    ''�����t���O�@�n�P
        Public OddsUmatanInfo() As ODDS_UMATAN_INFO ''<�n�P�I�b�Y>
        Public TotalHyosuUmatan As String  ''�n�P�[�����v
        Public crlf As String     ''���R�[�h��؂�
        '�z��̏�����
        Public Sub Initialize()
            ReDim OddsUmatanInfo(305)
        End Sub
        '�f�[�^�Z�b�g
        Public Sub SetData(ByRef strBuff As String)
            Initialize()      ''�z��̏�����
            Dim i As Integer
            Dim bBuff As Byte()
            Dim bSize As Long
            bSize = 4031
            bBuff = New Byte(bSize) {}
            bBuff = Str2Byte(strBuff)

            head.SetDataB(MidB2B(bBuff, 1, 11))
            id.SetDataB(MidB2B(bBuff, 12, 16))
            HappyoTime.SetDataB(MidB2B(bBuff, 28, 8))
            TorokuTosu = MidB2S(bBuff, 36, 2)
            SyussoTosu = MidB2S(bBuff, 38, 2)
            UmatanFlag = MidB2S(bBuff, 40, 1)
            For i = 0 To 305
                OddsUmatanInfo(i).SetDataB(MidB2B(bBuff, 41 + (13 * i), 13))
            Next i
            TotalHyosuUmatan = MidB2S(bBuff, 4019, 11)
            crlf = MidB2S(bBuff, 4030, 2)
            bBuff = Nothing
        End Sub
    End Structure

    '****** �P�O�D�I�b�Y�i�R�A���j****************************************
    '<3�A���I�b�Y>
    Public Structure ODDS_SANREN_INFO
        Public Kumi As String     ''�g��
        Public Odds As String     ''�I�b�Y
        Public Ninki As String     ''�l�C��
        '�f�[�^�Z�b�g
        Public Sub SetDataB(ByVal bBuff As Byte())
            Kumi = MidB2S(bBuff, 1, 6)
            Odds = MidB2S(bBuff, 7, 6)
            Ninki = MidB2S(bBuff, 13, 3)
        End Sub
    End Structure
    Public Structure JV_O5_ODDS_SANREN
        Public head As RECORD_ID    ''<���R�[�h�w�b�_�[>
        Public id As RACE_ID     ''<�������ʏ��>
        Public HappyoTime As MDHM    ''���\��������
        Public TorokuTosu As String    ''�o�^����
        Public SyussoTosu As String    ''�o������
        Public SanrenpukuFlag As String   ''�����t���O�@3�A��
        Public OddsSanrenInfo() As ODDS_SANREN_INFO ''<3�A���I�b�Y>
        Public TotalHyosuSanrenpuku As String ''3�A���[�����v
        Public crlf As String     ''���R�[�h��؂�
        '�z��̏�����
        Public Sub Initialize()
            ReDim OddsSanrenInfo(815)
        End Sub
        '�f�[�^�Z�b�g
        Public Sub SetData(ByRef strBuff As String)
            Initialize()      ''�z��̏�����
            Dim i As Integer
            Dim bBuff As Byte()
            Dim bSize As Long
            bSize = 12293
            bBuff = New Byte(bSize) {}
            bBuff = Str2Byte(strBuff)

            head.SetDataB(MidB2B(bBuff, 1, 11))
            id.SetDataB(MidB2B(bBuff, 12, 16))
            HappyoTime.SetDataB(MidB2B(bBuff, 28, 8))
            TorokuTosu = MidB2S(bBuff, 36, 2)
            SyussoTosu = MidB2S(bBuff, 38, 2)
            SanrenpukuFlag = MidB2S(bBuff, 40, 1)
            For i = 0 To 815
                OddsSanrenInfo(i).SetDataB(MidB2B(bBuff, 41 + (15 * i), 15))
            Next i
            TotalHyosuSanrenpuku = MidB2S(bBuff, 12281, 11)
            crlf = MidB2S(bBuff, 12292, 2)
            bBuff = Nothing
        End Sub
    End Structure

    '****** �P�O�|�P�D�I�b�Y�i�R�A�P�j****************************************
    '<3�A�P�I�b�Y>
    Public Structure ODDS_SANRENTAN_INFO
        Public Kumi As String       ''�g��
        Public Odds As String       ''�I�b�Y
        Public Ninki As String      ''�l�C��
        '�f�[�^�Z�b�g
        Public Sub SetDataB(ByVal bBuff As Byte())
            Kumi = MidB2S(bBuff, 1, 6)
            Odds = MidB2S(bBuff, 7, 7)
            Ninki = MidB2S(bBuff, 14, 4)
        End Sub
    End Structure

    Public Structure JV_O6_ODDS_SANRENTAN
        Public head As RECORD_ID                            ''<���R�[�h�w�b�_�[>
        Public id As RACE_ID                                ''<�������ʏ��>
        Public HappyoTime As MDHM                           ''���\��������
        Public TorokuTosu As String                         ''�o�^����
        Public SyussoTosu As String                         ''�o������
        Public SanrentanFlag As String                      ''�����t���O�@3�A�P
        Public OddsSanrentanInfo() As ODDS_SANRENTAN_INFO   ''<3�A�P�I�b�Y>
        Public TotalHyosuSanrentan As String                ''3�A�P�[�����v
        Public crlf As String                               ''���R�[�h��؂�

        '�z��̏�����
        Public Sub Initialize()
            ReDim OddsSanrentanInfo(4895)
        End Sub
        '�f�[�^�Z�b�g
        Public Sub SetData(ByRef strBuff As String)
            Initialize()      ''�z��̏�����
            Dim i As Integer
            Dim bBuff As Byte()
            Dim bSize As Long
            bSize = 83285
            bBuff = New Byte(bSize) {}
            bBuff = Str2Byte(strBuff)

            head.SetDataB(MidB2B(bBuff, 1, 11))
            id.SetDataB(MidB2B(bBuff, 12, 16))
            HappyoTime.SetDataB(MidB2B(bBuff, 28, 8))
            TorokuTosu = MidB2S(bBuff, 36, 2)
            SyussoTosu = MidB2S(bBuff, 38, 2)
            SanrentanFlag = MidB2S(bBuff, 40, 1)
            For i = 0 To 4895
                OddsSanrentanInfo(i).SetDataB(MidB2B(bBuff, 41 + (17 * i), 17))
            Next i
            TotalHyosuSanrentan = MidB2S(bBuff, 83273, 11)
            crlf = MidB2S(bBuff, 83284, 2)
            bBuff = Nothing
        End Sub
    End Structure


    '****** �P�P�D�����n�}�X�^ ****************************************
    '<�R�㌌�����>
    Public Structure KETTO3_INFO
        Public HansyokuNum As String   ''�ɐB�o�^�ԍ�
        Public Bamei As String     ''�n��
        '�f�[�^�Z�b�g
        Public Sub SetDataB(ByVal bBuff As Byte())
            HansyokuNum = MidB2S(bBuff, 1, 10)
            Bamei = MidB2S(bBuff, 11, 36)
        End Sub
    End Structure
    Public Structure JV_UM_UMA
        Public head As RECORD_ID    ''<���R�[�h�w�b�_�[>
        Public KettoNum As String    ''�����o�^�ԍ�
        Public DelKubun As String    ''�����n�����敪
        Public RegDate As YMD     ''�����n�o�^�N����
        Public DelDate As YMD     ''�����n�����N����
        Public BirthDate As YMD     ''���N����
        Public Bamei As String     ''�n��
        Public BameiKana As String    ''�n�����p�J�i
        Public BameiEng As String    ''�n������
        Public ZaikyuFlag As String    ''JRA�{�ݍ݂��イ�t���O
        Public Reserved As String    ''�\��
        Public UmaKigoCD As String    ''�n�L���R�[�h
        Public SexCD As String     ''���ʃR�[�h
        Public HinsyuCD As String    ''�i��R�[�h
        Public KeiroCD As String    ''�ѐF�R�[�h
        Public Ketto3Info() As KETTO3_INFO  ''<3�㌌�����>
        Public TozaiCD As String    ''���������R�[�h
        Public ChokyosiCode As String   ''�����t�R�[�h
        Public ChokyosiRyakusyo As String  ''�����t������
        Public Syotai As String     ''���Ғn�於
        Public BreederCode As String   ''���Y�҃R�[�h
        Public BreederName As String   ''���Y�Җ�
        Public SanchiName As String    ''�Y�n��
        Public BanusiCode As String    ''�n��R�[�h
        Public BanusiName As String    ''�n�喼
        Public RuikeiHonsyoHeiti As String  ''���n�{�܋��݌v
        Public RuikeiHonsyoSyogai As String  ''��Q�{�܋��݌v
        Public RuikeiFukaHeichi As String  ''���n�t���܋��݌v
        Public RuikeiFukaSyogai As String  ''��Q�t���܋��݌v
        Public RuikeiSyutokuHeichi As String ''���n�����܋��݌v
        Public RuikeiSyutokuSyogai As String ''��Q�����܋��݌v
        Public ChakuSogo As CHAKUKAISU3_INFO     ''��������
        Public ChakuChuo As CHAKUKAISU3_INFO     ''�������v����
        Public ChakuKaisuBa() As CHAKUKAISU3_INFO   ''�n��ʒ���
        Public ChakuKaisuJyotai() As CHAKUKAISU3_INFO  ''�n���ԕʒ���
        Public ChakuKaisuKyori() As CHAKUKAISU3_INFO  ''�����ʒ���
        Public Kyakusitu() As String   ''�r���X��
        Public RaceCount As String    ''�o�^���[�X��
        Public crlf As String     ''���R�[�h��؂�
        '�z��̏�����
        Public Sub Initialize()
            ReDim Ketto3Info(13)
            ReDim ChakuKaisuBa(6)
            ReDim ChakuKaisuJyotai(11)
            ReDim Kyakusitu(3)
            ReDim ChakuKaisuKyori(5)
        End Sub
        '�f�[�^�Z�b�g
        Public Sub SetData(ByRef strBuff As String)
            Initialize()      ''�z��̏�����
            Dim i As Integer
            Dim bBuff As Byte()
            Dim bSize As Long
            bSize = 1609
            bBuff = New Byte(bSize) {}
            bBuff = Str2Byte(strBuff)

            head.SetDataB(MidB2B(bBuff, 1, 11))
            KettoNum = MidB2S(bBuff, 12, 10)
            DelKubun = MidB2S(bBuff, 22, 1)
            RegDate.SetDataB(MidB2B(bBuff, 23, 8))
            DelDate.SetDataB(MidB2B(bBuff, 31, 8))
            BirthDate.SetDataB(MidB2B(bBuff, 39, 8))
            Bamei = MidB2S(bBuff, 47, 36)
            BameiKana = MidB2S(bBuff, 83, 36)
            BameiEng = MidB2S(bBuff, 119, 60)
            ZaikyuFlag = MidB2S(bBuff, 179, 1)
            Reserved = MidB2S(bBuff, 180, 19)
            UmaKigoCD = MidB2S(bBuff, 199, 2)
            SexCD = MidB2S(bBuff, 201, 1)
            HinsyuCD = MidB2S(bBuff, 202, 1)
            KeiroCD = MidB2S(bBuff, 203, 2)
            For i = 0 To 13
                Ketto3Info(i).SetDataB(MidB2B(bBuff, 205 + (46 * i), 46))
            Next i
            TozaiCD = MidB2S(bBuff, 849, 1)
            ChokyosiCode = MidB2S(bBuff, 850, 5)
            ChokyosiRyakusyo = MidB2S(bBuff, 855, 8)
            Syotai = MidB2S(bBuff, 863, 20)
            BreederCode = MidB2S(bBuff, 883, 8)
            BreederName = MidB2S(bBuff, 891, 72)
            SanchiName = MidB2S(bBuff, 963, 20)
            BanusiCode = MidB2S(bBuff, 983, 6)
            BanusiName = MidB2S(bBuff, 989, 64)
            RuikeiHonsyoHeiti = MidB2S(bBuff, 1053, 9)
            RuikeiHonsyoSyogai = MidB2S(bBuff, 1062, 9)
            RuikeiFukaHeichi = MidB2S(bBuff, 1071, 9)
            RuikeiFukaSyogai = MidB2S(bBuff, 1080, 9)
            RuikeiSyutokuHeichi = MidB2S(bBuff, 1089, 9)
            RuikeiSyutokuSyogai = MidB2S(bBuff, 1098, 9)
            ChakuSogo.SetDataB(MidB2B(bBuff, 1107, 18))
            ChakuChuo.SetDataB(MidB2B(bBuff, 1125, 18))
            For i = 0 To 6
                ChakuKaisuBa(i).SetDataB(MidB2B(bBuff, 1143 + (18 * i), 18))
            Next i
            For i = 0 To 11
                ChakuKaisuJyotai(i).SetDataB(MidB2B(bBuff, 1269 + (18 * i), 18))
            Next i
            For i = 0 To 5
                ChakuKaisuKyori(i).SetDataB(MidB2B(bBuff, 1485 + (18 * i), 18))
            Next i
            For i = 0 To 3
                Kyakusitu(i) = MidB2S(bBuff, 1593 + (3 * i), 3)
            Next i
            RaceCount = MidB2S(bBuff, 1605, 3)
            crlf = MidB2S(bBuff, 1608, 2)
            bBuff = Nothing
        End Sub
    End Structure

    '****** �P�Q�D�R��}�X�^ ****************************************
    '<���R����>
    Public Structure HATUKIJYO_INFO
        Public Hatukijyoid As RACE_ID   ''�N��������R
        Public SyussoTosu As String    ''�o������
        Public KettoNum As String    ''�����o�^�ԍ�
        Public Bamei As String     ''�n��
        Public KakuteiJyuni As String   ''�m�蒅��
        Public IJyoCD As String     ''�ُ�敪�R�[�h
        '�f�[�^�Z�b�g
        Public Sub SetDataB(ByVal bBuff As Byte())
            Hatukijyoid.SetDataB(MidB2B(bBuff, 1, 16))
            SyussoTosu = MidB2S(bBuff, 17, 2)
            KettoNum = MidB2S(bBuff, 19, 10)
            Bamei = MidB2S(bBuff, 29, 36)
            KakuteiJyuni = MidB2S(bBuff, 65, 2)
            IJyoCD = MidB2S(bBuff, 67, 1)
        End Sub
    End Structure
    '<���������>
    Public Structure HATUSYORI_INFO
        Public Hatusyoriid As RACE_ID   ''�N��������R
        Public SyussoTosu As String    ''�o������
        Public KettoNum As String    ''�����o�^�ԍ�
        Public Bamei As String     ''�n��
        '�f�[�^�Z�b�g
        Public Sub SetDataB(ByVal bBuff As Byte())
            Hatusyoriid.SetDataB(MidB2B(bBuff, 1, 16))
            SyussoTosu = MidB2S(bBuff, 17, 2)
            KettoNum = MidB2S(bBuff, 19, 10)
            Bamei = MidB2S(bBuff, 29, 36)
        End Sub
    End Structure
    Public Structure JV_KS_KISYU
        Public head As RECORD_ID    ''<���R�[�h�w�b�_�[>
        Public KisyuCode As String    ''�R��R�[�h
        Public DelKubun As String    ''�R�薕���敪
        Public IssueDate As YMD     ''�R��Ƌ���t�N����
        Public DelDate As YMD     ''�R��Ƌ������N����
        Public BirthDate As YMD     ''���N����
        Public KisyuName As String    ''�R�薼����
        Public reserved As String    ''�\��
        Public KisyuNameKana As String   ''�R�薼���p�J�i
        Public KisyuRyakusyo As String   ''�R�薼����
        Public KisyuNameEng As String   ''�R�薼����
        Public SexCD As String     ''���ʋ敪
        Public SikakuCD As String    ''�R�掑�i�R�[�h
        Public MinaraiCD As String    ''�R�茩�K�R�[�h
        Public TozaiCD As String    ''�R�蓌�������R�[�h
        Public Syotai As String     ''���Ғn�於
        Public ChokyosiCode As String   ''���������t�R�[�h
        Public ChokyosiRyakusyo As String  ''���������t������
        Public HatuKiJyo() As HATUKIJYO_INFO   ''<���R����>
        Public HatuSyori() As HATUSYORI_INFO   ''<���������>
        Public SaikinJyusyo() As SAIKIN_JYUSYO_INFO  ''<�ŋߏd�܏������>
        Public HonZenRuikei() As HON_ZEN_RUIKEISEI_INFO ''<�{�N�E�O�N�E�݌v���я��>
        Public crlf As String   ''���R�[�h��؂�
        '�z��̏�����
        Public Sub Initialize()
            ReDim HatuKiJyo(1)
            ReDim HatuSyori(1)
            ReDim SaikinJyusyo(2)
            ReDim HonZenRuikei(2)
        End Sub
        '�f�[�^�Z�b�g	
        Public Sub SetData(ByRef strBuff As String)
            Initialize()      ''�z��̏�����
            Dim i As Integer
            Dim bBuff As Byte()
            Dim bSize As Long
            bSize = 4173
            bBuff = New Byte(bSize) {}
            bBuff = Str2Byte(strBuff)

            head.SetDataB(MidB2B(bBuff, 1, 11))
            KisyuCode = MidB2S(bBuff, 12, 5)
            DelKubun = MidB2S(bBuff, 17, 1)
            IssueDate.SetDataB(MidB2B(bBuff, 18, 8))
            DelDate.SetDataB(MidB2B(bBuff, 26, 8))
            BirthDate.SetDataB(MidB2B(bBuff, 34, 8))
            KisyuName = MidB2S(bBuff, 42, 34)
            reserved = MidB2S(bBuff, 76, 34)
            KisyuNameKana = MidB2S(bBuff, 110, 30)
            KisyuRyakusyo = MidB2S(bBuff, 140, 8)
            KisyuNameEng = MidB2S(bBuff, 148, 80)
            SexCD = MidB2S(bBuff, 228, 1)
            SikakuCD = MidB2S(bBuff, 229, 1)
            MinaraiCD = MidB2S(bBuff, 230, 1)
            TozaiCD = MidB2S(bBuff, 231, 1)
            Syotai = MidB2S(bBuff, 232, 20)
            ChokyosiCode = MidB2S(bBuff, 252, 5)
            ChokyosiRyakusyo = MidB2S(bBuff, 257, 8)
            For i = 0 To 1
                HatuKiJyo(i).SetDataB(MidB2B(bBuff, 265 + (67 * i), 67))
            Next i
            For i = 0 To 1
                HatuSyori(i).SetDataB(MidB2B(bBuff, 399 + (64 * i), 64))
            Next i
            For i = 0 To 2
                SaikinJyusyo(i).SetDataB(MidB2B(bBuff, 527 + (163 * i), 163))
            Next i
            For i = 0 To 2
                HonZenRuikei(i).SetDataB(MidB2B(bBuff, 1016 + (1052 * i), 1052))
            Next i
            crlf = MidB2S(bBuff, 4172, 2)
            bBuff = Nothing
        End Sub
    End Structure

    '****** �P�R�D�����t�}�X�^ ****************************************
    Public Structure JV_CH_CHOKYOSI
        Public head As RECORD_ID    ''<���R�[�h�w�b�_�[>
        Public ChokyosiCode As String   ''�����t�R�[�h
        Public DelKubun As String    ''�����t�����敪
        Public IssueDate As YMD      ''�����t�Ƌ���t�N����
        Public DelDate As YMD     ''�����t�Ƌ������N����
        Public BirthDate As YMD     ''���N����
        Public ChokyosiName As String   ''�����t������
        Public ChokyosiNameKana As String  ''�����t�����p�J�i
        Public ChokyosiRyakusyo As String  ''�����t������
        Public ChokyosiNameEng As String  ''�����t������
        Public SexCD As String     ''���ʋ敪
        Public TozaiCD As String    ''�����t���������R�[�h
        Public Syotai As String     ''���Ғn�於
        Public SaikinJyusyo() As SAIKIN_JYUSYO_INFO  ''<�ŋߏd�܏������>
        Public HonZenRuikei() As HON_ZEN_RUIKEISEI_INFO ''<�{�N�E�O�N�E�݌v���я��>
        Public crlf As String     ''���R�[�h��؂�
        '�z��̏�����
        Public Sub Initialize()
            ReDim SaikinJyusyo(2)
            ReDim HonZenRuikei(2)
        End Sub
        '�f�[�^�Z�b�g
        Public Sub SetData(ByRef strBuff As String)
            Initialize()      ''�z��̏�����
            Dim i As Integer
            Dim bBuff As Byte()
            Dim bSize As Long
            bSize = 3862
            bBuff = New Byte(bSize) {}
            bBuff = Str2Byte(strBuff)

            head.SetDataB(MidB2B(bBuff, 1, 11))
            ChokyosiCode = MidB2S(bBuff, 12, 5)
            DelKubun = MidB2S(bBuff, 17, 1)
            IssueDate.SetDataB(MidB2B(bBuff, 18, 8))
            DelDate.SetDataB(MidB2B(bBuff, 26, 8))
            BirthDate.SetDataB(MidB2B(bBuff, 34, 8))
            ChokyosiName = MidB2S(bBuff, 42, 34)
            ChokyosiNameKana = MidB2S(bBuff, 76, 30)
            ChokyosiRyakusyo = MidB2S(bBuff, 106, 8)
            ChokyosiNameEng = MidB2S(bBuff, 114, 80)
            SexCD = MidB2S(bBuff, 194, 1)
            TozaiCD = MidB2S(bBuff, 195, 1)
            Syotai = MidB2S(bBuff, 196, 20)
            For i = 0 To 2
                SaikinJyusyo(i).SetDataB(MidB2B(bBuff, 216 + (163 * i), 163))
            Next i
            For i = 0 To 2
                HonZenRuikei(i).SetDataB(MidB2B(bBuff, 705 + (1052 * i), 1052))
            Next i
            crlf = MidB2S(bBuff, 3861, 2)
            bBuff = Nothing
        End Sub
    End Structure

    ''******�P�S�D���Y�҃}�X�^ ****************************************
    Public Structure JV_BR_BREEDER
        Public head As RECORD_ID    ''<���R�[�h�w�b�_�[>
        Public BreederCode As String   ''���Y�҃R�[�h
        Public BreederName_Co As String   ''���Y�Җ�(�@�l�i�L)
        Public BreederName As String   ''���Y�Җ�(�@�l�i��)
        Public BreederNameKana As String  ''���Y�Җ����p�J�i
        Public BreederNameEng As String   ''���Y�Җ�����
        Public Address As String    ''���Y�ҏZ�������Ȗ�
        Public HonRuikei() As SEI_RUIKEI_INFO ''<�{�N�E�݌v���я��>
        Public crlf As String     ''���R�[�h��؂�
        '�z��̏�����
        Public Sub Initialize()
            ReDim HonRuikei(1)
        End Sub
        '�f�[�^�Z�b�g
        Public Sub SetData(ByRef strBuff As String)
            Initialize()      ''�z��̏�����
            Dim i As Integer
            Dim bBuff As Byte()
            Dim bSize As Long
            bSize = 545
            bBuff = New Byte(bSize) {}
            bBuff = Str2Byte(strBuff)

            head.SetDataB(MidB2B(bBuff, 1, 11))
            BreederCode = MidB2S(bBuff, 12, 8)
            BreederName_Co = MidB2S(bBuff, 20, 72)
            BreederName = MidB2S(bBuff, 92, 72)
            BreederNameKana = MidB2S(bBuff, 164, 72)
            BreederNameEng = MidB2S(bBuff, 236, 168)
            Address = MidB2S(bBuff, 404, 20)
            For i = 0 To 1
                HonRuikei(i).SetDataB(MidB2B(bBuff, 424 + (60 * i), 60))
            Next i
            crlf = MidB2S(bBuff, 544, 2)
            bBuff = Nothing
        End Sub
    End Structure

    '****** �P�T�D�n��}�X�^ ****************************************
    Public Structure JV_BN_BANUSI
        Public head As RECORD_ID    ''<���R�[�h�w�b�_�[>
        Public BanusiCode As String    ''�n��R�[�h
        Public BanusiName_Co As String    ''�n�喼(�@�l�i�L)
        Public BanusiName As String    ''�n�喼(�@�l�i��)
        Public BanusiNameKana As String   ''�n�喼���p�J�i
        Public BanusiNameEng As String   ''�n�喼����
        Public Fukusyoku As String    ''���F�W��
        Public HonRuikei() As SEI_RUIKEI_INFO ''<�{�N�E�݌v���я��>
        Public crlf As String     ''���R�[�h��؂�
        '�z��̏�����
        Public Sub Initialize()
            ReDim HonRuikei(1)
        End Sub
        '�f�[�^�Z�b�g
        Public Sub SetData(ByRef strBuff As String)
            Initialize()      ''�z��̏�����
            Dim i As Integer
            Dim bBuff As Byte()
            Dim bSize As Long
            bSize = 477
            bBuff = New Byte(bSize) {}
            bBuff = Str2Byte(strBuff)

            head.SetDataB(MidB2B(bBuff, 1, 11))
            BanusiCode = MidB2S(bBuff, 12, 6)
            BanusiName_Co = MidB2S(bBuff, 18, 64)
            BanusiName = MidB2S(bBuff, 82, 64)
            BanusiNameKana = MidB2S(bBuff, 146, 50)
            BanusiNameEng = MidB2S(bBuff, 196, 100)
            Fukusyoku = MidB2S(bBuff, 296, 60)
            For i = 0 To 1
                HonRuikei(i).SetDataB(MidB2B(bBuff, 356 + (60 * i), 60))
            Next i
            crlf = MidB2S(bBuff, 476, 2)
            bBuff = Nothing
        End Sub
    End Structure

    ''****** �P�U�D�ɐB�n�}�X�^ ****************************************
    Public Structure JV_HN_HANSYOKU
        Public head As RECORD_ID    ''<���R�[�h�w�b�_�[>
        Public HansyokuNum As String   ''�ɐB�o�^�ԍ�
        Public reserved As String    ''�\��
        Public KettoNum As String    ''�����o�^�ԍ�
        Public DelKubun As String    ''�ɐB�n�����敪(���݂͗\���Ƃ��Ďg�p)
        Public Bamei As String     ''�n��
        Public BameiKana As String    ''�n�����p�J�i
        Public BameiEng As String    ''�n������
        Public BirthYear As String    ''���N
        Public SexCD As String     ''���ʃR�[�h
        Public HinsyuCD As String    ''�i��R�[�h
        Public KeiroCD As String    ''�ѐF�R�[�h
        Public HansyokuMochiKubun As String  ''�ɐB�n�����敪
        Public ImportYear As String    ''�A���N
        Public SanchiName As String    ''�Y�n��
        Public HansyokuFNum As String   ''���n�ɐB�o�^�ԍ�
        Public HansyokuMNum As String   ''��n�ɐB�o�^�ԍ�
        Public crlf As String     ''���R�[�h��؂�
        '�f�[�^�Z�b�g
        Public Sub SetData(ByRef strBuff As String)
            Dim bBuff As Byte()
            Dim bSize As Long
            bSize = 251
            bBuff = New Byte(bSize) {}
            bBuff = Str2Byte(strBuff)

            head.SetDataB(MidB2B(bBuff, 1, 11))
            HansyokuNum = MidB2S(bBuff, 12, 10)
            reserved = MidB2S(bBuff, 22, 8)
            KettoNum = MidB2S(bBuff, 30, 10)
            DelKubun = MidB2S(bBuff, 40, 1)
            Bamei = MidB2S(bBuff, 41, 36)
            BameiKana = MidB2S(bBuff, 77, 40)
            BameiEng = MidB2S(bBuff, 117, 80)
            BirthYear = MidB2S(bBuff, 197, 4)
            SexCD = MidB2S(bBuff, 201, 1)
            HinsyuCD = MidB2S(bBuff, 202, 1)
            KeiroCD = MidB2S(bBuff, 203, 2)
            HansyokuMochiKubun = MidB2S(bBuff, 205, 1)
            ImportYear = MidB2S(bBuff, 206, 4)
            SanchiName = MidB2S(bBuff, 210, 20)
            HansyokuFNum = MidB2S(bBuff, 230, 10)
            HansyokuMNum = MidB2S(bBuff, 240, 10)
            crlf = MidB2S(bBuff, 250, 2)
            bBuff = Nothing
        End Sub
    End Structure

    '****** �P�V�D�Y��}�X�^ ****************************************
    Public Structure JV_SK_SANKU
        Public head As RECORD_ID    ''<���R�[�h�w�b�_�[>
        Public KettoNum As String    ''�����o�^�ԍ�
        Public BirthDate As YMD     ''���N����
        Public SexCD As String     ''���ʃR�[�h
        Public HinsyuCD As String    ''�i��R�[�h
        Public KeiroCD As String    ''�ѐF�R�[�h
        Public SankuMochiKubun As String  ''�Y����敪
        Public ImportYear As String    ''�A���N
        Public BreederCode As String   ''���Y�҃R�[�h
        Public SanchiName As String    ''�Y�n��
        Public HansyokuNum() As String   ''3�㌌�� �ɐB�o�^�ԍ�
        Public crlf As String     ''���R�[�h��؂�
        '�z��̏�����
        Public Sub Initialize()
            ReDim HansyokuNum(13)
        End Sub
        '�f�[�^�Z�b�g
        Public Sub SetData(ByRef strBuff As String)
            Initialize()      ''�z��̏�����
            Dim i As Integer
            Dim bBuff As Byte()
            Dim bSize As Long
            bSize = 208
            bBuff = New Byte(bSize) {}
            bBuff = Str2Byte(strBuff)

            head.SetDataB(MidB2B(bBuff, 1, 11))
            KettoNum = MidB2S(bBuff, 12, 10)
            BirthDate.SetDataB(MidB2B(bBuff, 22, 8))
            SexCD = MidB2S(bBuff, 30, 1)
            HinsyuCD = MidB2S(bBuff, 31, 1)
            KeiroCD = MidB2S(bBuff, 32, 2)
            SankuMochiKubun = MidB2S(bBuff, 34, 1)
            ImportYear = MidB2S(bBuff, 35, 4)
            BreederCode = MidB2S(bBuff, 39, 8)
            SanchiName = MidB2S(bBuff, 47, 20)
            For i = 0 To 13
                HansyokuNum(i) = MidB2S(bBuff, 67 + (10 * i), 10)
            Next i
            crlf = MidB2S(bBuff, 207, 2)
            bBuff = Nothing
        End Sub
    End Structure

    '****** �P�W�D���R�[�h�}�X�^ ****************************************
    '<���R�[�h�ێ��n���>
    Public Structure RECUMA_INFO
        Public KettoNum As String    ''�����o�^�ԍ�
        Public Bamei As String     ''�n��
        Public UmaKigoCD As String    ''�n�L���R�[�h
        Public SexCD As String     ''���ʃR�[�h
        Public ChokyosiCode As String   ''�����t�R�[�h
        Public ChokyosiName As String   ''�����t��
        Public Futan As String     ''���S�d��
        Public KisyuCode As String    ''�R��R�[�h
        Public KisyuName As String    ''�R�薼
        '�f�[�^�Z�b�g
        Public Sub SetDataB(ByVal bBuff As Byte())
            KettoNum = MidB2S(bBuff, 1, 10)
            Bamei = MidB2S(bBuff, 11, 36)
            UmaKigoCD = MidB2S(bBuff, 47, 2)
            SexCD = MidB2S(bBuff, 49, 1)
            ChokyosiCode = MidB2S(bBuff, 50, 5)
            ChokyosiName = MidB2S(bBuff, 55, 34)
            Futan = MidB2S(bBuff, 89, 3)
            KisyuCode = MidB2S(bBuff, 92, 5)
            KisyuName = MidB2S(bBuff, 97, 34)
        End Sub
    End Structure
    Public Structure JV_RC_RECORD
        Public head As RECORD_ID    ''<���R�[�h�w�b�_�[>
        Public RecInfoKubun As String   ''���R�[�h���ʋ敪
        Public id As RACE_ID     ''<�������ʏ��>
        Public TokuNum As String    ''���ʋ����ԍ�
        Public Hondai As String     ''�������{��
        Public GradeCD As String    ''�O���[�h�R�[�h
        Public SyubetuCD As String    ''������ʃR�[�h
        Public Kyori As String     ''����
        Public TrackCD As String    ''�g���b�N�R�[�h
        Public RecKubun As String    ''���R�[�h�敪
        Public RecTime As String    ''���R�[�h�^�C��
        Public TenkoBaba As TENKO_BABA_INFO  ''�V��E�n����
        Public RecUmaInfo() As RECUMA_INFO  ''<���R�[�h�ێ��n���>
        Public crlf As String     ''���R�[�h��؂�
        '�z��̏�����
        Public Sub Initialize()
            ReDim RecUmaInfo(2)
        End Sub
        '�f�[�^�Z�b�g
        Public Sub SetData(ByRef strBuff As String)
            Initialize()      ''�z��̏�����
            Dim i As Integer
            Dim bBuff As Byte()
            Dim bSize As Long
            bSize = 501
            bBuff = New Byte(bSize) {}
            bBuff = Str2Byte(strBuff)

            head.SetDataB(MidB2B(bBuff, 1, 11))
            RecInfoKubun = MidB2S(bBuff, 12, 1)
            id.SetDataB(MidB2B(bBuff, 13, 16))
            TokuNum = MidB2S(bBuff, 29, 4)
            Hondai = MidB2S(bBuff, 33, 60)
            GradeCD = MidB2S(bBuff, 93, 1)
            SyubetuCD = MidB2S(bBuff, 94, 2)
            Kyori = MidB2S(bBuff, 96, 4)
            TrackCD = MidB2S(bBuff, 100, 2)
            RecKubun = MidB2S(bBuff, 102, 1)
            RecTime = MidB2S(bBuff, 103, 4)
            TenkoBaba.SetDataB(MidB2B(bBuff, 107, 3))
            For i = 0 To 2
                RecUmaInfo(i).SetDataB(MidB2B(bBuff, 110 + (130 * i), 130))
            Next i
            crlf = MidB2S(bBuff, 500, 2)
            bBuff = Nothing
        End Sub
    End Structure

    '****** �P�X�D��H���� ****************************************
    Public Structure JV_HC_HANRO
        Public head As RECORD_ID    ''<���R�[�h�w�b�_�[>
        Public TresenKubun As String   ''�g���Z���敪
        Public ChokyoDate As YMD    ''�����N����
        Public ChokyoTime As String    ''��������
        Public KettoNum As String    ''�����o�^�ԍ�
        Public HaronTime4 As String    ''4�n�����^�C�����v(800M-0M)
        Public LapTime4 As String    ''���b�v�^�C��(800M-600M)
        Public HaronTime3 As String    ''3�n�����^�C�����v(600M-0M)
        Public LapTime3 As String    ''���b�v�^�C��(600M-400M)
        Public HaronTime2 As String    ''2�n�����^�C�����v(400M-0M)
        Public LapTime2 As String    ''���b�v�^�C��(400M-200M)
        Public LapTime1 As String    ''���b�v�^�C��(200M-0M)
        Public crlf As String     ''���R�[�h��؂�
        '�f�[�^�Z�b�g
        Public Sub SetData(ByRef strBuff As String)
            Dim bBuff As Byte()
            Dim bSize As Long
            bSize = 60
            bBuff = New Byte(bSize) {}
            bBuff = Str2Byte(strBuff)

            head.SetDataB(MidB2B(bBuff, 1, 11))
            TresenKubun = MidB2S(bBuff, 12, 1)
            ChokyoDate.SetDataB(MidB2B(bBuff, 13, 8))
            ChokyoTime = MidB2S(bBuff, 21, 4)
            KettoNum = MidB2S(bBuff, 25, 10)
            HaronTime4 = MidB2S(bBuff, 35, 4)
            LapTime4 = MidB2S(bBuff, 39, 3)
            HaronTime3 = MidB2S(bBuff, 42, 4)
            LapTime3 = MidB2S(bBuff, 46, 3)
            HaronTime2 = MidB2S(bBuff, 49, 4)
            LapTime2 = MidB2S(bBuff, 53, 3)
            LapTime1 = MidB2S(bBuff, 56, 3)
            crlf = MidB2S(bBuff, 59, 2)
            bBuff = Nothing
        End Sub
    End Structure

    '****** �Q�O�D�n�̏d ****************************************
    '<�n�̏d���>
    Public Structure BATAIJYU_INFO
        Public Umaban As String     ''�n��
        Public Bamei As String     ''�n��
        Public BaTaijyu As String    ''�n�̏d
        Public ZogenFugo As String    ''��������
        Public ZogenSa As String    ''������
        '�f�[�^�Z�b�g
        Public Sub SetDataB(ByVal bBuff As Byte())
            Umaban = MidB2S(bBuff, 1, 2)
            Bamei = MidB2S(bBuff, 3, 36)
            BaTaijyu = MidB2S(bBuff, 39, 3)
            ZogenFugo = MidB2S(bBuff, 42, 1)
            ZogenSa = MidB2S(bBuff, 43, 3)
        End Sub
    End Structure
    Public Structure JV_WH_BATAIJYU
        Public head As RECORD_ID    ''<���R�[�h�w�b�_�[>
        Public id As RACE_ID     ''<�������ʏ��>
        Public HappyoTime As MDHM    ''���\��������
        Public BataijyuInfo() As BATAIJYU_INFO ''<�n�̏d���>
        Public crlf As String     ''���R�[�h��؂�
        '�z��̏�����
        Public Sub Initialize()
            ReDim BataijyuInfo(17)
        End Sub
        '�f�[�^�Z�b�g
        Public Sub SetData(ByRef strBuff As String)
            Initialize()      ''�z��̏�����
            Dim i As Integer
            Dim bBuff As Byte()
            Dim bSize As Long
            bSize = 847
            bBuff = New Byte(bSize) {}
            bBuff = Str2Byte(strBuff)

            head.SetDataB(MidB2B(bBuff, 1, 11))
            id.SetDataB(MidB2B(bBuff, 12, 16))
            HappyoTime.SetDataB(MidB2B(bBuff, 28, 8))
            For i = 0 To 17
                BataijyuInfo(i).SetDataB(MidB2B(bBuff, 36 + (45 * i), 45))
            Next i
            crlf = MidB2S(bBuff, 846, 2)
            bBuff = Nothing
        End Sub
    End Structure

    '****** �Q�P�D�V��n���� ******************************************
    Public Structure JV_WE_WEATHER
        Public head As RECORD_ID     ''<���R�[�h�w�b�_�[>
        Public id As RACE_ID2      ''<�������ʏ��Q>
        Public HappyoTime As MDHM     ''���\��������
        Public HenkoID As String     ''�ύX����
        Public TenkoBaba As TENKO_BABA_INFO   ''���ݏ�ԏ��
        Public TenkoBabaBefore As TENKO_BABA_INFO   ''�ύX�O��ԏ��
        Public crlf As String     ''���R�[�h��؂�
        '�f�[�^�Z�b�g
        Public Sub SetData(ByRef strBuff As String)
            Dim bBuff As Byte()
            Dim bSize As Long
            bSize = 42
            bBuff = New Byte(bSize) {}
            bBuff = Str2Byte(strBuff)

            head.SetDataB(MidB2B(bBuff, 1, 11))
            id.SetDataB(MidB2B(bBuff, 12, 14))
            HappyoTime.SetDataB(MidB2B(bBuff, 26, 8))
            HenkoID = MidB2S(bBuff, 34, 1)
            TenkoBaba.SetDataB(MidB2B(bBuff, 35, 3))
            TenkoBabaBefore.SetDataB(MidB2B(bBuff, 38, 3))
            crlf = MidB2S(bBuff, 41, 2)
            bBuff = Nothing
        End Sub
    End Structure

    '****** �Q�Q�D�o������E�������O ****************************************
    Public Structure JV_AV_INFO
        Public head As RECORD_ID    ''<���R�[�h�w�b�_�[>
        Public id As RACE_ID     ''<�������ʏ��>
        Public HappyoTime As MDHM    ''���\��������
        Public Umaban As String     ''�n��
        Public Bamei As String     ''�n��
        Public JiyuKubun As String    ''���R�敪
        Public crlf As String     ''���R�[�h��؂�
        '�f�[�^�Z�b�g
        Public Sub SetData(ByRef strBuff As String)
            Dim bBuff As Byte()
            Dim bSize As Long
            bSize = 78
            bBuff = New Byte(bSize) {}
            bBuff = Str2Byte(strBuff)

            head.SetDataB(MidB2B(bBuff, 1, 11))
            id.SetDataB(MidB2B(bBuff, 12, 16))
            HappyoTime.SetDataB(MidB2B(bBuff, 28, 8))
            Umaban = MidB2S(bBuff, 36, 2)
            Bamei = MidB2S(bBuff, 38, 36)
            JiyuKubun = MidB2S(bBuff, 74, 3)
            crlf = MidB2S(bBuff, 77, 2)
            bBuff = Nothing
        End Sub
    End Structure

    '************ �Q�R�D�R��ύX **************************************** 
    '<�ύX���>
    Public Structure JC_INFO
        Public Futan As String     ''���S�d��
        Public KisyuCode As String    ''�R��R�[�h
        Public KisyuName As String    ''�R�薼
        Public MinaraiCD As String    ''�R�茩�K�R�[�h
        '�f�[�^�Z�b�g
        Public Sub SetDataB(ByVal bBuff As Byte())
            Futan = MidB2S(bBuff, 1, 3)
            KisyuCode = MidB2S(bBuff, 4, 5)
            KisyuName = MidB2S(bBuff, 9, 34)
            MinaraiCD = MidB2S(bBuff, 43, 1)
        End Sub
    End Structure
    Public Structure JV_JC_INFO
        Public head As RECORD_ID    ''<���R�[�h�w�b�_�[>
        Public id As RACE_ID     ''<�������ʏ��>
        Public HappyoTime As MDHM    ''���\��������
        Public Umaban As String     ''�n��
        Public Bamei As String     ''�n��
        Public JCInfoAfter As JC_INFO   ''<�ύX����>
        Public JCInfoBefore As JC_INFO   ''<�ύX�O���>
        Public crlf As String     ''���R�[�h��؂�
        '�f�[�^�Z�b�g
        Public Sub SetData(ByRef strBuff As String)
            Dim bBuff As Byte()
            Dim bSize As Long
            bSize = 161
            bBuff = New Byte(bSize) {}
            bBuff = Str2Byte(strBuff)

            head.SetDataB(MidB2B(bBuff, 1, 11))
            id.SetDataB(MidB2B(bBuff, 12, 16))
            HappyoTime.SetDataB(MidB2B(bBuff, 28, 8))
            Umaban = MidB2S(bBuff, 36, 2)
            Bamei = MidB2S(bBuff, 38, 36)
            JCInfoAfter.SetDataB(MidB2B(bBuff, 74, 43))
            JCInfoBefore.SetDataB(MidB2B(bBuff, 117, 43))
            crlf = MidB2S(bBuff, 160, 2)
            bBuff = Nothing
        End Sub
    End Structure

    '************ �Q�R�|�P�D���������ύX **************************************** 
    '<�ύX���>
    Public Structure TC_INFO
        Public Ji As String  ''��
        Public Fun As String  ''��
        '�f�[�^�Z�b�g
        Public Sub SetDataB(ByVal bBuff As Byte())
            Ji = MidB2S(bBuff, 1, 2)
            Fun = MidB2S(bBuff, 3, 2)
        End Sub
    End Structure
    Public Structure JV_TC_INFO
        Public head As RECORD_ID  ''<���R�[�h�w�b�_�[>
        Public id As RACE_ID   ''<�������ʏ��>
        Public HappyoTime As MDHM  ''���\��������
        Public TCInfoAfter As TC_INFO ''<�ύX����>
        Public TCInfoBefore As TC_INFO ''<�ύX�O���>
        Public crlf As String   ''���R�[�h��؂�
        '�f�[�^�Z�b�g
        Public Sub SetData(ByRef strBuff As String)
            Dim bBuff As Byte()
            Dim bSize As Long
            bSize = 45
            bBuff = New Byte(bSize) {}
            bBuff = Str2Byte(strBuff)
            head.SetDataB(MidB2B(bBuff, 1, 11))
            id.SetDataB(MidB2B(bBuff, 12, 16))
            HappyoTime.SetDataB(MidB2B(bBuff, 28, 8))
            TCInfoAfter.SetDataB(MidB2B(bBuff, 36, 4))
            TCInfoBefore.SetDataB(MidB2B(bBuff, 40, 4))
            crlf = MidB2S(bBuff, 44, 2)
            bBuff = Nothing
        End Sub
    End Structure

    '************ �Q�R�|�Q�D�R�[�X�ύX **************************************** 
    '<�ύX���>
    Public Structure CC_INFO
        Public Kyori As String   ''����
        Public TruckCd As String  ''�g���b�N�R�[�h
        '�f�[�^�Z�b�g
        Public Sub SetDataB(ByVal bBuff As Byte())
            Kyori = MidB2S(bBuff, 1, 4)
            TruckCd = MidB2S(bBuff, 5, 2)
        End Sub
    End Structure
    Public Structure JV_CC_INFO
        Public head As RECORD_ID  ''<���R�[�h�w�b�_�[>
        Public id As RACE_ID   ''<�������ʏ��>
        Public HappyoTime As MDHM  ''���\��������
        Public CCInfoAfter As CC_INFO ''<�ύX����>
        Public CCInfoBefore As CC_INFO ''<�ύX�O���>
        Public JiyuCd As String   ''���R�R�[�h
        Public crlf As String   ''���R�[�h��؂�
        '�f�[�^�Z�b�g
        Public Sub SetData(ByRef strBuff As String)
            Dim bBuff As Byte()
            Dim bSize As Long
            bSize = 50
            bBuff = New Byte(bSize) {}
            bBuff = Str2Byte(strBuff)
            head.SetDataB(MidB2B(bBuff, 1, 11))
            id.SetDataB(MidB2B(bBuff, 12, 16))
            HappyoTime.SetDataB(MidB2B(bBuff, 28, 8))
            CCInfoAfter.SetDataB(MidB2B(bBuff, 36, 6))
            CCInfoBefore.SetDataB(MidB2B(bBuff, 42, 6))
            JiyuCd = MidB2S(bBuff, 48, 1)
            crlf = MidB2S(bBuff, 49, 2)
            bBuff = Nothing
        End Sub
    End Structure


    '****** �Q�S�D�f�[�^�}�C�j���O�\�z************************************
    '<�}�C�j���O�\�z>
    Public Structure DM_INFO
        Public Umaban As String     ''�n��
        Public DMTime As String     ''�\�z���j�^�C��
        Public DMGosaP As String    ''�\�z�덷(�M���x)�{
        Public DMGosaM As String    ''�\�z�덷(�M���x)�|
        '�f�[�^�Z�b�g
        Public Sub SetDataB(ByVal bBuff As Byte())
            Umaban = MidB2S(bBuff, 1, 2)
            DMTime = MidB2S(bBuff, 3, 5)
            DMGosaP = MidB2S(bBuff, 8, 4)
            DMGosaM = MidB2S(bBuff, 12, 4)
        End Sub
    End Structure
    Public Structure JV_DM_INFO
        Public head As RECORD_ID    ''<���R�[�h�w�b�_�[>
        Public id As RACE_ID     ''<�������ʏ��>
        Public MakeHM As HM      ''�f�[�^�쐬����
        Public DMInfo() As DM_INFO    ''<�}�C�j���O�\�z>
        Public crlf As String     ''���R�[�h��؂�
        '�z��̏�����
        Public Sub Initialize()
            ReDim DMInfo(17)
        End Sub
        '�f�[�^�Z�b�g
        Public Sub SetData(ByRef strBuff As String)
            Initialize()      ''�z��̏�����							
            Dim i As Integer
            Dim bBuff As Byte()
            Dim bSize As Long
            bSize = 303
            bBuff = New Byte(bSize) {}
            bBuff = Str2Byte(strBuff)

            head.SetDataB(MidB2B(bBuff, 1, 11))
            id.SetDataB(MidB2B(bBuff, 12, 16))
            MakeHM.SetDataB(MidB2B(bBuff, 28, 4))
            For i = 0 To 17
                DMInfo(i).SetDataB(MidB2B(bBuff, 32 + (15 * i), 15))
            Next i
            crlf = MidB2S(bBuff, 302, 2)
            bBuff = Nothing
        End Sub
    End Structure

    '****** �Q�T�D�J�ÃX�P�W���[��************************************
    '<�d�܈ē�>
    Public Structure JYUSYO_INFO
        Public TokuNum As String    ''���ʋ����ԍ�
        Public Hondai As String     ''�������{��
        Public Ryakusyo10 As String    ''����������10��
        Public Ryakusyo6 As String    ''����������6��
        Public Ryakusyo3 As String    ''����������3��
        Public Nkai As String     ''�d�܉�[��N��]
        Public GradeCD As String    ''�O���[�h�R�[�h
        Public SyubetuCD As String    ''������ʃR�[�h
        Public KigoCD As String     ''�����L���R�[�h
        Public JyuryoCD As String    ''�d�ʎ�ʃR�[�h
        Public Kyori As String     ''����
        Public TrackCD As String    ''�g���b�N�R�[�h
        '�f�[�^�Z�b�g
        Public Sub SetDataB(ByVal bBuff As Byte())
            TokuNum = MidB2S(bBuff, 1, 4)
            Hondai = MidB2S(bBuff, 5, 60)
            Ryakusyo10 = MidB2S(bBuff, 65, 20)
            Ryakusyo6 = MidB2S(bBuff, 85, 12)
            Ryakusyo3 = MidB2S(bBuff, 97, 6)
            Nkai = MidB2S(bBuff, 103, 3)
            GradeCD = MidB2S(bBuff, 106, 1)
            SyubetuCD = MidB2S(bBuff, 107, 2)
            KigoCD = MidB2S(bBuff, 109, 3)
            JyuryoCD = MidB2S(bBuff, 112, 1)
            Kyori = MidB2S(bBuff, 113, 4)
            TrackCD = MidB2S(bBuff, 117, 2)
        End Sub
    End Structure
    Public Structure JV_YS_SCHEDULE
        Public head As RECORD_ID    ''<���R�[�h�w�b�_�[>
        Public id As RACE_ID2     ''<�������ʏ��Q>
        Public YoubiCD As String    ''�j���R�[�h
        Public JyusyoInfo() As JYUSYO_INFO  ''<�d�܈ē�>
        Public crlf As String     ''���R�[�h��؂�
        '�z��̏�����
        Public Sub Initialize()
            ReDim JyusyoInfo(2)
        End Sub
        '�f�[�^�Z�b�g
        Public Sub SetData(ByRef strBuff As String)
            Initialize()      ''�z��̏�����	
            Dim bBuff As Byte()
            Dim i As Integer
            Dim bSize As Long
            bSize = 382
            bBuff = New Byte(bSize) {}
            bBuff = Str2Byte(strBuff)

            head.SetDataB(MidB2B(bBuff, 1, 11))
            id.SetDataB(MidB2B(bBuff, 12, 14))
            YoubiCD = MidB2S(bBuff, 26, 1)
            For i = 0 To 2
                JyusyoInfo(i).SetDataB(MidB2B(bBuff, 27 + (118 * i), 118))
            Next i
            crlf = MidB2S(bBuff, 381, 2)
            bBuff = Nothing
        End Sub
    End Structure

    '****** �Q�U�D�����n�s�������i ****************************************
    Public Structure JV_HS_SALE
        Public head As RECORD_ID         ''<���R�[�h�w�b�_�[>
        Public KettoNum As String        ''�����o�^�ԍ�
        Public HansyokuFNum As String    ''���n�ɐB�o�^�ԍ�
        Public HansyokuMNum As String    ''��n�ɐB�o�^�ԍ�
        Public BirthYear As String       ''���N
        Public SaleCode As String        ''��ÎҁE�s��R�[�h
        Public SaleHostName As String    ''��ÎҖ���
        Public SaleName As String        ''�s��̖���
        Public FromDate As YMD           ''�s��̊J�Ê���(�J�n��)
        Public ToDate As YMD             ''�s��̊J�Ê���(�I����)
        Public Barei As String          ''������̋����n�̔N��
        Public Price As String          ''������i
        Public crlf As String            ''���R�[�h��؂�
        '�f�[�^�Z�b�g
        Public Sub SetData(ByRef strBuff As String)
            Dim bBuff As Byte()
            Dim bSize As Long
            bSize = 200
            bBuff = New Byte(bSize) {}
            bBuff = Str2Byte(strBuff)

            head.SetDataB(MidB2B(bBuff, 1, 11))
            KettoNum = MidB2S(bBuff, 12, 10)
            HansyokuFNum = MidB2S(bBuff, 22, 10)
            HansyokuMNum = MidB2S(bBuff, 32, 10)
            BirthYear = MidB2S(bBuff, 42, 4)
            SaleCode = MidB2S(bBuff, 46, 6)
            SaleHostName = MidB2S(bBuff, 52, 40)
            SaleName = MidB2S(bBuff, 92, 80)
            FromDate.SetDataB(MidB2B(bBuff, 172, 8))
            ToDate.SetDataB(MidB2B(bBuff, 180, 8))
            Barei = MidB2S(bBuff, 188, 1)
            Price = MidB2S(bBuff, 189, 10)
            crlf = MidB2S(bBuff, 199, 2)
            bBuff = Nothing
        End Sub
    End Structure

    ''****** �Q�V�D�n���̈Ӗ��R�� ****************************************
    Public Structure JV_HY_BAMEIORIGIN
        Public head As RECORD_ID       ''<���R�[�h�w�b�_�[>
        Public KettoNum As String      ''�����o�^�ԍ�
        Public Bamei As String         ''�n��
        Public Origin As String        ''�n���̈Ӗ��R��
        Public crlf As String          ''���R�[�h��؂�
        '�f�[�^�Z�b�g
        Public Sub SetData(ByRef strBuff As String)
            Dim bBuff As Byte()
            Dim bSize As Long
            bSize = 123
            bBuff = New Byte(bSize) {}
            bBuff = Str2Byte(strBuff)

            head.SetDataB(MidB2B(bBuff, 1, 11))
            KettoNum = MidB2S(bBuff, 12, 10)
            Bamei = MidB2S(bBuff, 22, 36)
            Origin = MidB2S(bBuff, 58, 64)
            crlf = MidB2S(bBuff, 122, 2)
            bBuff = Nothing
        End Sub
    End Structure

    '****** �Q�W�D�o���ʒ��x�� ****************************************

    '<�o���ʒ��x�� �����n���>
    Public Structure JV_CK_UMA
        Public KettoNum As String                         ''�����o�^�ԍ�
        Public Bamei As String                            ''�n��
        Public RuikeiHonsyoHeiti As String                ''���n�{�܋��݌v
        Public RuikeiHonsyoSyogai As String               ''��Q�{�܋��݌v
        Public RuikeiFukaHeichi As String                 ''���n�t���܋��݌v
        Public RuikeiFukaSyogai As String                 ''��Q�t���܋��݌v
        Public RuikeiSyutokuHeichi As String              ''���n�����܋��݌v
        Public RuikeiSyutokuSyogai As String              ''��Q�����܋��݌v
        Public ChakuSogo As CHAKUKAISU3_INFO              ''��������
        Public ChakuChuo As CHAKUKAISU3_INFO              ''�������v����
        Public ChakuKaisuBa() As CHAKUKAISU3_INFO         ''�n��ʒ���
        Public ChakuKaisuJyotai() As CHAKUKAISU3_INFO     ''�n���ԕʒ���
        Public ChakuKaisuSibaKyori() As CHAKUKAISU3_INFO  ''�ŋ����ʒ���
        Public ChakuKaisuDirtKyori() As CHAKUKAISU3_INFO  ''�_�[�g�����ʒ���
        Public ChakuKaisuJyoSiba() As CHAKUKAISU3_INFO    ''���n��ʎŒ���
        Public ChakuKaisuJyoDirt() As CHAKUKAISU3_INFO    ''���n��ʃ_�[�g����
        Public ChakuKaisuJyoSyogai() As CHAKUKAISU3_INFO  ''���n��ʏ�Q����
        Public Kyakusitu() As String                      ''�r���X��
        Public RaceCount As String                        ''�o�^���[�X��
        '�z��̏�����
        Public Sub Initialize()
            ReDim ChakuKaisuBa(6)
            ReDim ChakuKaisuJyotai(11)
            ReDim ChakuKaisuSibaKyori(8)
            ReDim ChakuKaisuDirtKyori(8)
            ReDim ChakuKaisuJyoSiba(9)
            ReDim ChakuKaisuJyoDirt(9)
            ReDim ChakuKaisuJyoSyogai(9)
            ReDim Kyakusitu(3)
        End Sub
        '�f�[�^�Z�b�g
        Public Sub SetDataB(ByVal bBuff As Byte())
            Initialize()      ''�z��̏�����
            KettoNum = MidB2S(bBuff, 1, 10)
            Bamei = MidB2S(bBuff, 11, 36)
            RuikeiHonsyoHeiti = MidB2S(bBuff, 47, 9)
            RuikeiHonsyoSyogai = MidB2S(bBuff, 56, 9)
            RuikeiFukaHeichi = MidB2S(bBuff, 65, 9)
            RuikeiFukaSyogai = MidB2S(bBuff, 74, 9)
            RuikeiSyutokuHeichi = MidB2S(bBuff, 83, 9)
            RuikeiSyutokuSyogai = MidB2S(bBuff, 92, 9)
            ChakuSogo.SetDataB(MidB2B(bBuff, 101, 18))
            ChakuChuo.SetDataB(MidB2B(bBuff, 119, 18))
            Dim i As Integer = 0
            For i = 0 To 6
                ChakuKaisuBa(i).SetDataB(MidB2B(bBuff, 137 + 18 * i, 18))
            Next i
            For i = 0 To 11
                ChakuKaisuJyotai(i).SetDataB(MidB2B(bBuff, 263 + 18 * i, 18))
            Next i
            For i = 0 To 8
                ChakuKaisuSibaKyori(i).SetDataB(MidB2B(bBuff, 479 + 18 * i, 18))
            Next i
            For i = 0 To 8
                ChakuKaisuDirtKyori(i).SetDataB(MidB2B(bBuff, 641 + 18 * i, 18))
            Next i
            For i = 0 To 9
                ChakuKaisuJyoSiba(i).SetDataB(MidB2B(bBuff, 803 + 18 * i, 18))
            Next i
            For i = 0 To 9
                ChakuKaisuJyoDirt(i).SetDataB(MidB2B(bBuff, 983 + 18 * i, 18))
            Next i
            For i = 0 To 9
                ChakuKaisuJyoSyogai(i).SetDataB(MidB2B(bBuff, 1163 + 18 * i, 18))
            Next i
            For i = 0 To 3
                Kyakusitu(i) = MidB2S(bBuff, 1343 + (3 * i), 3)
            Next i
            RaceCount = MidB2S(bBuff, 1355, 3)
        End Sub
    End Structure

    '<�o���ʒ��x�� �{�N�E�݌v���я��>
    Public Structure JV_CK_HON_RUIKEISEI_INFO
        Public SetYear As String                          ''�ݒ�N
        Public HonSyokinHeichi As String                  ''���n�{�܋����v
        Public HonSyokinSyogai As String                  ''��Q�{�܋����v
        Public FukaSyokinHeichi As String                 ''���n�t���܋����v
        Public FukaSyokinSyogai As String                 ''��Q�t���܋����v
        Public ChakuKaisuSiba As CHAKUKAISU5_INFO         ''�Œ���
        Public ChakuKaisuDirt As CHAKUKAISU5_INFO         ''�_�[�g����
        Public ChakuKaisuSyogai As CHAKUKAISU4_INFO       ''��Q����
        Public ChakuKaisuSibaKyori() As CHAKUKAISU4_INFO ''�ŋ����ʒ���
        Public ChakuKaisuDirtKyori() As CHAKUKAISU4_INFO ''�_�[�g�����ʒ���
        Public ChakuKaisuJyoSiba() As CHAKUKAISU4_INFO   ''���n��ʎŒ���
        Public ChakuKaisuJyoDirt() As CHAKUKAISU4_INFO   ''���n��ʃ_�[�g����
        Public ChakuKaisuJyoSyogai() As CHAKUKAISU3_INFO ''���n��ʏ�Q����
        '�z��̏�����
        Public Sub Initialize()
            ReDim ChakuKaisuSibaKyori(8)
            ReDim ChakuKaisuDirtKyori(8)
            ReDim ChakuKaisuJyoSiba(9)
            ReDim ChakuKaisuJyoDirt(9)
            ReDim ChakuKaisuJyoSyogai(9)
        End Sub
        '�f�[�^�Z�b�g
        Public Sub SetDataB(ByVal bBuff As Byte())
            Initialize()      ''�z��̏�����
            SetYear = MidB2S(bBuff, 1, 4)
            HonSyokinHeichi = MidB2S(bBuff, 5, 10)
            HonSyokinSyogai = MidB2S(bBuff, 15, 10)
            FukaSyokinHeichi = MidB2S(bBuff, 25, 10)
            FukaSyokinSyogai = MidB2S(bBuff, 35, 10)
            ChakuKaisuSiba.SetDataB(MidB2B(bBuff, 45, 30))
            ChakuKaisuDirt.SetDataB(MidB2B(bBuff, 75, 30))
            ChakuKaisuSyogai.SetDataB(MidB2B(bBuff, 105, 24))
            Dim i As Integer = 0
            For i = 0 To 8
                ChakuKaisuSibaKyori(i).SetDataB(MidB2B(bBuff, 129 + 24 * i, 24))
            Next i
            For i = 0 To 8
                ChakuKaisuDirtKyori(i).SetDataB(MidB2B(bBuff, 345 + 24 * i, 24))
            Next i
            For i = 0 To 9
                ChakuKaisuJyoSiba(i).SetDataB(MidB2B(bBuff, 561 + 24 * i, 24))
            Next i
            For i = 0 To 9
                ChakuKaisuJyoDirt(i).SetDataB(MidB2B(bBuff, 801 + 24 * i, 24))
            Next i
            For i = 0 To 9
                ChakuKaisuJyoSyogai(i).SetDataB(MidB2B(bBuff, 1041 + 18 * i, 18))
            Next i
        End Sub
    End Structure

    '<�o���ʒ��x�� �R����>
    Public Structure JV_CK_KISYU
        Public KisyuCode As String                 ''�R��R�[�h
        Public KisyuName As String                 ''�R�薼����
        Public HonRuikei() As JV_CK_HON_RUIKEISEI_INFO ''<�{�N�E�݌v���я��>
        '�z��̏�����
        Public Sub Initialize()
            ReDim HonRuikei(1)
        End Sub
        '�f�[�^�Z�b�g
        Public Sub SetDataB(ByVal bBuff As Byte())
            Initialize()      ''�z��̏�����
            KisyuCode = MidB2S(bBuff, 1, 5)
            KisyuName = MidB2S(bBuff, 6, 34)
            Dim i As Integer = 0
            For i = 0 To 1
                HonRuikei(i).SetDataB(MidB2B(bBuff, 40 + 1220 * i, 1220))
            Next i
        End Sub
    End Structure

    '<�o���ʒ��x�� �����t���>
    Public Structure JV_CK_CHOKYOSI
        Public ChokyosiCode As String              ''�����t�R�[�h
        Public ChokyosiName As String              ''�����t������
        Public HonRuikei() As JV_CK_HON_RUIKEISEI_INFO ''<�{�N�E�݌v���я��>
        '�z��̏�����
        Public Sub Initialize()
            ReDim HonRuikei(1)
        End Sub
        '�f�[�^�Z�b�g
        Public Sub SetDataB(ByVal bBuff As Byte())
            Initialize()      ''�z��̏�����
            ChokyosiCode = MidB2S(bBuff, 1, 5)
            ChokyosiName = MidB2S(bBuff, 6, 34)
            Dim i As Integer = 0
            For i = 0 To 1
                HonRuikei(i).SetDataB(MidB2B(bBuff, 40 + 1220 * i, 1220))
            Next i
        End Sub
    End Structure

    '<�o���ʒ��x�� �n����>
    Public Structure JV_CK_BANUSI
        Public BanusiCode As String                ''�n��R�[�h
        Public BanusiName_Co As String             ''�n�喼�i�@�l�i�L�j
        Public BanusiName As String                ''�n�喼�i�@�l�i���j
        Public HonRuikei() As SEI_RUIKEI_INFO     ''<�{�N�E�݌v���я��>
        '�z��̏�����
        Public Sub Initialize()
            ReDim HonRuikei(1)
        End Sub
        '�f�[�^�Z�b�g
        Public Sub SetDataB(ByVal bBuff As Byte())
            Initialize()      ''�z��̏�����
            BanusiCode = MidB2S(bBuff, 1, 6)
            BanusiName_Co = MidB2S(bBuff, 7, 64)
            BanusiName = MidB2S(bBuff, 71, 64)
            Dim i As Integer = 0
            For i = 0 To 1
                HonRuikei(i).SetDataB(MidB2B(bBuff, 135 + 60 * i, 60))
            Next i
        End Sub
    End Structure

    '<�o���ʒ��x�� ���Y�ҏ��>
    Public Structure JV_CK_BREEDER
        Public BreederCode As String               ''���Y�҃R�[�h
        Public BreederName_Co As String            ''���Y�Җ��i�@�l�i�L�j
        Public BreederName As String               ''���Y�Җ��i�@�l�i���j
        Public HonRuikei() As SEI_RUIKEI_INFO     ''<�{�N�E�݌v���я��>
        '�z��̏�����
        Public Sub Initialize()
            ReDim HonRuikei(1)
        End Sub
        '�f�[�^�Z�b�g
        Public Sub SetDataB(ByVal bBuff As Byte())
            Initialize()      ''�z��̏�����
            BreederCode = MidB2S(bBuff, 1, 8)
            BreederName_Co = MidB2S(bBuff, 9, 72)
            BreederName = MidB2S(bBuff, 81, 72)
            Dim i As Integer = 0
            For i = 0 To 1
                HonRuikei(i).SetDataB(MidB2B(bBuff, 153 + 60 * i, 60))
            Next i
        End Sub
    End Structure

    Public Structure JV_CK_CHAKU
        Public head As RECORD_ID                   ''<���R�[�h�w�b�_�[>
        Public id As RACE_ID                       ''<�������ʏ��P>
        Public UmaChaku As JV_CK_UMA               ''<�o���ʒ��x�� �����n���>
        Public KisyuChaku As JV_CK_KISYU           ''<�o���ʒ��x�� �R����>
        Public ChokyoChaku As JV_CK_CHOKYOSI       ''<�o���ʒ��x�� �����t���>
        Public BanusiChaku As JV_CK_BANUSI         ''<�o���ʒ��x�� �n����>
        Public BreederChaku As JV_CK_BREEDER       ''<�o���ʒ��x�� ���Y�ҏ��>
        Public crlf As String                      ''���R�[�h��؂�
        '�f�[�^�Z�b�g
        Public Sub SetData(ByRef strBuff As String)
            Dim bBuff As Byte()
            Dim bSize As Long
            bSize = 6870
            bBuff = New Byte(bSize) {}
            bBuff = Str2Byte(strBuff)

            head.SetDataB(MidB2B(bBuff, 1, 11))
            id.SetDataB(MidB2B(bBuff, 12, 16))
            UmaChaku.SetDataB(MidB2B(bBuff, 28, 1357))
            KisyuChaku.SetDataB(MidB2B(bBuff, 1385, 2479))
            ChokyoChaku.SetDataB(MidB2B(bBuff, 3864, 2479))
            BanusiChaku.SetDataB(MidB2B(bBuff, 6343, 254))
            BreederChaku.SetDataB(MidB2B(bBuff, 6597, 272))
            crlf = MidB2S(bBuff, 6869, 2)
            bBuff = Nothing
        End Sub
    End Structure

    '****** �Q�X�D�n����� ****************************************
    Public Structure JV_BT_KEITO
        Public head As RECORD_ID       ''<���R�[�h�w�b�_�[>
        Public HansyokuNum As String   ''�ɐB�o�^�ԍ�
        Public KeitoId As String       ''�n��ID
        Public KeitoName As String     ''�n����
        Public KeitoEx As String       ''�n������
        Public crlf As String          ''���R�[�h��؂�
        '�f�[�^�Z�b�g
        Public Sub SetData(ByRef strBuff As String)
            Dim bBuff As Byte()
            Dim bSize As Long
            bSize = 6889
            bBuff = New Byte(bSize) {}
            bBuff = Str2Byte(strBuff)

            head.SetDataB(MidB2B(bBuff, 1, 11))
            HansyokuNum = MidB2S(bBuff, 12, 10)
            KeitoId = MidB2S(bBuff, 22, 30)
            KeitoName = MidB2S(bBuff, 52, 36)
            KeitoEx = MidB2S(bBuff, 88, 6800)
            crlf = MidB2S(bBuff, 6888, 2)
            bBuff = Nothing
        End Sub
    End Structure

    '****** �R�O�D�R�[�X��� ****************************************
    Public Structure JV_CS_COURSE
        Public head As RECORD_ID       ''<���R�[�h�w�b�_�[>
        Public JyoCD As String         ''���n��R�[�h
        Public Kyori As String         ''����
        Public TrackCD As String       ''�g���b�N�R�[�h
        Public KaishuDate As YMD       ''�R�[�X���C�N����
        Public CourseEx As String      ''�R�[�X����
        Public crlf As String          ''���R�[�h��؂�
        '�f�[�^�Z�b�g
        Public Sub SetData(ByRef strBuff As String)
            Dim bBuff As Byte()
            Dim bSize As Long
            bSize = 6829
            bBuff = New Byte(bSize) {}
            bBuff = Str2Byte(strBuff)

            head.SetDataB(MidB2B(bBuff, 1, 11))
            JyoCD = MidB2S(bBuff, 12, 2)
            Kyori = MidB2S(bBuff, 14, 4)
            TrackCD = MidB2S(bBuff, 18, 2)
            KaishuDate.SetDataB(MidB2B(bBuff, 20, 8))
            CourseEx = MidB2S(bBuff, 28, 6800)
            crlf = MidB2S(bBuff, 6828, 2)
            bBuff = Nothing
        End Sub
    End Structure

    '****** �R�P�D�ΐ�^�f�[�^�}�C�j���O�\�z************************************
    '<�}�C�j���O�\�z>
    Public Structure TM_INFO
        Public Umaban As String    ''�n��
        Public TMScore As String    ''�\���X�R�A
        '�f�[�^�Z�b�g
        Public Sub SetDataB(ByVal bBuff As Byte())
            Umaban = MidB2S(bBuff, 1, 2)
            TMScore = MidB2S(bBuff, 3, 4)
        End Sub
    End Structure
    Public Structure JV_TM_INFO
        Public head As RECORD_ID      ''<���R�[�h�w�b�_�[>
        Public id As RACE_ID          ''<�������ʏ��>
        Public MakeHM As HM           ''�f�[�^�쐬����
        Public TMInfo() As TM_INFO    ''<�}�C�j���O�\�z>
        Public crlf As String         ''���R�[�h��؂�
        '�z��̏�����
        Public Sub Initialize()
            ReDim TMInfo(17)
        End Sub
        '�f�[�^�Z�b�g
        Public Sub SetData(ByRef strBuff As String)
            Initialize()      ''�z��̏�����							
            Dim i As Integer
            Dim bBuff As Byte()
            Dim bSize As Long
            bSize = 141
            bBuff = New Byte(bSize) {}
            bBuff = Str2Byte(strBuff)

            head.SetDataB(MidB2B(bBuff, 1, 11))
            id.SetDataB(MidB2B(bBuff, 12, 16))
            MakeHM.SetDataB(MidB2B(bBuff, 28, 4))
            For i = 0 To 17
                TMInfo(i).SetDataB(MidB2B(bBuff, 32 + (6 * i), 6))
            Next i
            crlf = MidB2S(bBuff, 140, 2)
            bBuff = Nothing
        End Sub
    End Structure

    '****** �R�Q�D�d����(WIN5)************************************
    '<�d�����Ώۃ��[�X���>
    Public Structure WF_RACE_INFO
        Public JyoCD As String     ''���n��R�[�h
        Public Kaiji As String     ''�J�É�[��N��]
        Public Nichiji As String   ''�J�Ó���[N����]
        Public RaceNum As String   ''���[�X�ԍ�
        '�f�[�^�Z�b�g
        Public Sub SetDataB(ByVal bBuff As Byte())
            JyoCD = MidB2S(bBuff, 1, 2)
            Kaiji = MidB2S(bBuff, 3, 2)
            Nichiji = MidB2S(bBuff, 5, 2)
            RaceNum = MidB2S(bBuff, 7, 2)
        End Sub
    End Structure

    '<�L���[�����>
    Public Structure WF_YUKO_HYO_INFO
        Public Yuko_Hyo As String     ''�L���[��
        '�f�[�^�Z�b�g
        Public Sub SetDataB(ByVal bBuff As Byte())
            Yuko_Hyo = MidB2S(bBuff, 1, 11)
        End Sub
    End Structure

    '<�d�������ߏ��>
    Public Structure WF_PAY_INFO
        Public Kumiban As String     ''�g��
        Public Pay As String         ''�d�������ߋ�
        Public Tekichu_Hyo As String ''�I���[��

        '�f�[�^�Z�b�g
        Public Sub SetDataB(ByVal bBuff As Byte())
            Kumiban = MidB2S(bBuff, 1, 10)
            Pay = MidB2S(bBuff, 11, 9)
            Tekichu_Hyo = MidB2S(bBuff, 20, 10)

        End Sub
    End Structure

    Public Structure JV_WF_INFO
        Public head As RECORD_ID                   ''<���R�[�h�w�b�_�[>
        Public KaisaiDate As YMD                   ''�J�ÔN����
        Public reserved1 As String                 ''�\��
        Public WFRaceInfo() As WF_RACE_INFO        ''<�d�����Ώۃ��[�X���>
        Public reserved2 As String                 ''�\��
        Public Hatsubai_Hyo As String              ''�d���������[��
        Public WFYukoHyoInfo() As WF_YUKO_HYO_INFO ''<�L���[�����>
        Public HenkanFlag As String                ''�Ԋ҃t���O
        Public FuseiritsuFlag As String            ''�s�����t���O
        Public TekichunashiFlag As String          ''�I�����t���O
        Public COShoki As String                   ''�L�����[�I�[�o�[���z����
        Public COZanDaka As String                 ''�L�����[�I�[�o�[���z�c��
        Public WFPayInfo() As WF_PAY_INFO          ''<�d�������ߏ��>
        Public crlf As String                      ''���R�[�h��؂�

        '�z��̏�����
        Public Sub Initialize()
            ReDim WFRaceInfo(4)
            ReDim WFYukoHyoInfo(4)
            ReDim WFPayInfo(242)
        End Sub
        '�f�[�^�Z�b�g
        Public Sub SetData(ByRef strBuff As String)
            Initialize()      ''�z��̏�����
            Dim i As Integer
            Dim bBuff As Byte()
            Dim bSize As Long
            bSize = 7215
            bBuff = New Byte(bSize) {}
            bBuff = Str2Byte(strBuff)

            head.SetDataB(MidB2B(bBuff, 1, 11))
            KaisaiDate.SetDataB(MidB2B(bBuff, 12, 8))
            reserved1 = MidB2S(bBuff, 20, 2)

            For i = 0 To 4
                WFRaceInfo(i).SetDataB(MidB2B(bBuff, 22 + (8 * i), 8))
            Next i

            reserved2 = MidB2S(bBuff, 62, 6)
            Hatsubai_Hyo = MidB2S(bBuff, 68, 11)

            For i = 0 To 4
                WFYukoHyoInfo(i).SetDataB(MidB2B(bBuff, 79 + (11 * i), 11))
            Next i

            HenkanFlag = MidB2S(bBuff, 134, 1)
            FuseiritsuFlag = MidB2S(bBuff, 135, 1)
            TekichunashiFlag = MidB2S(bBuff, 136, 1)
            COShoki = MidB2S(bBuff, 137, 15)
            COZanDaka = MidB2S(bBuff, 152, 15)

            For i = 0 To 242
                WFPayInfo(i).SetDataB(MidB2B(bBuff, 167 + (29 * i), 29))
            Next i

            crlf = MidB2S(bBuff, 7214, 2)
            bBuff = Nothing

        End Sub
    End Structure

    '****** �R�R�D�����n���O���************************************
    Public Structure JV_JG_JOGAIBA
        Public head As RECORD_ID            ''<���R�[�h�w�b�_�[>
        Public id As RACE_ID                ''<�������ʏ��>
        Public KettoNum As String           ''�����o�^�ԍ�
        Public Bamei As String              ''�n��
        Public ShutsubaTohyoJun As String   ''�o�n���[��t����
        Public ShussoKubun As String        ''�o���敪
        Public JogaiJotaiKubun As String    ''���O��ԋ敪
        Public crlf As String               ''���R�[�h��؂�

        '�f�[�^�Z�b�g
        Public Sub SetData(ByRef strBuff As String)
            Dim bBuff As Byte()
            Dim bSize As Long
            bSize = 80
            bBuff = New Byte(bSize) {}
            bBuff = Str2Byte(strBuff)

            head.SetDataB(MidB2B(bBuff, 1, 11))
            id.SetDataB(MidB2B(bBuff, 12, 16))
            KettoNum = MidB2S(bBuff, 28, 10)
            Bamei = MidB2S(bBuff, 38, 36)
            ShutsubaTohyoJun = MidB2S(bBuff, 74, 3)
            ShussoKubun = MidB2S(bBuff, 77, 1)
            JogaiJotaiKubun = MidB2S(bBuff, 78, 1)
            crlf = MidB2S(bBuff, 79, 2)
            bBuff = Nothing
        End Sub
    End Structure

    '****** �R�S�D�E�b�h�`�b�v���� ****************************************
    Public Structure JV_WC_WOOD
        Public head As RECORD_ID    ''<���R�[�h�w�b�_�[>
        Public TresenKubun As String   ''�g���Z���敪
        Public ChokyoDate As YMD    ''�����N����
        Public ChokyoTime As String    ''��������
        Public KettoNum As String    ''�����o�^�ԍ�
        Public Course As String    ''�R�[�X
        Public BabaAround As String    ''�n�����
        Public reserved As String    ''�\��
        Public HaronTime10 As String    ''10�n�����^�C�����v(2000M-0M)
        Public LapTime10 As String    ''���b�v�^�C��(2000M-1800M)
        Public HaronTime9 As String    ''9�n�����^�C�����v(1800M-0M)
        Public LapTime9 As String    ''���b�v�^�C��(1800M-1600M)
        Public HaronTime8 As String    ''8�n�����^�C�����v(1600M-0M)
        Public LapTime8 As String    ''���b�v�^�C��1600M-1400M)
        Public HaronTime7 As String    ''7�n�����^�C�����v(1400M-0M)
        Public LapTime7 As String    ''���b�v�^�C��(1400M-1200M)
        Public HaronTime6 As String    ''6�n�����^�C�����v(1200M-0M)
        Public LapTime6 As String    ''���b�v�^�C��(1200M-1000M)
        Public HaronTime5 As String    ''5�n�����^�C�����v(1000M-0M)
        Public LapTime5 As String    ''���b�v�^�C��(1000M-800M)
        Public HaronTime4 As String    ''4�n�����^�C�����v(800M-0M)
        Public LapTime4 As String    ''���b�v�^�C��(800M-600M)
        Public HaronTime3 As String    ''3�n�����^�C�����v(600M-0M)
        Public LapTime3 As String    ''���b�v�^�C��(600M-400M)
        Public HaronTime2 As String    ''2�n�����^�C�����v(400M-0M)
        Public LapTime2 As String    ''���b�v�^�C��(400M-200M)
        Public LapTime1 As String    ''���b�v�^�C��(200M-0M)
        Public crlf As String     ''���R�[�h��؂�
        '�f�[�^�Z�b�g
        Public Sub SetData(ByRef strBuff As String)
            Dim bBuff As Byte()
            Dim bSize As Long
            bSize = 105
            bBuff = New Byte(bSize) {}
            bBuff = Str2Byte(strBuff)

            head.SetDataB(MidB2B(bBuff, 1, 11))
            TresenKubun = MidB2S(bBuff, 12, 1)
            ChokyoDate.SetDataB(MidB2B(bBuff, 13, 8))
            ChokyoTime = MidB2S(bBuff, 21, 4)
            KettoNum = MidB2S(bBuff, 25, 10)
            Course = MidB2S(bBuff, 35, 1)
            BabaAround = MidB2S(bBuff, 36, 1)
            reserved = MidB2S(bBuff, 37, 1)
            HaronTime10 = MidB2S(bBuff, 38, 4)
            LapTime10 = MidB2S(bBuff, 42, 3)
            HaronTime9 = MidB2S(bBuff, 45, 4)
            LapTime9 = MidB2S(bBuff, 49, 3)
            HaronTime8 = MidB2S(bBuff, 52, 4)
            LapTime8 = MidB2S(bBuff, 56, 3)
            HaronTime7 = MidB2S(bBuff, 59, 4)
            LapTime7 = MidB2S(bBuff, 63, 3)
            HaronTime6 = MidB2S(bBuff, 66, 4)
            LapTime6 = MidB2S(bBuff, 70, 3)
            HaronTime5 = MidB2S(bBuff, 73, 4)
            LapTime5 = MidB2S(bBuff, 77, 3)
            HaronTime4 = MidB2S(bBuff, 80, 4)
            LapTime4 = MidB2S(bBuff, 84, 3)
            HaronTime3 = MidB2S(bBuff, 87, 4)
            LapTime3 = MidB2S(bBuff, 91, 3)
            HaronTime2 = MidB2S(bBuff, 94, 4)
            LapTime2 = MidB2S(bBuff, 98, 3)
            LapTime1 = MidB2S(bBuff, 101, 3)
            crlf = MidB2S(bBuff, 104, 2)
            bBuff = Nothing
        End Sub
    End Structure

End Module
