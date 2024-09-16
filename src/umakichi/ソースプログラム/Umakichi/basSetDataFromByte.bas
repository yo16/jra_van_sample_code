Attribute VB_Name = "basSetDataFromByte"
'
'   �f�[�^�Z�b�g�֐�
'

Option Explicit

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �O���֐�
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   �@�\: �\���̂Ƀf�[�^���Z�b�g����|���ʓo�^�n
'
'   ���l: �Ȃ�
'
    Public Sub SetDataFromByte_TK(ByRef bytBuf() As Byte, ByRef mBuf As JV_TK_TOKUUMA)
    Dim i As Integer                                '' ���[�v�J�E���^
    Dim j As Integer                                '' ���[�v�J�E���^
    Dim k As Integer                                '' ���[�v�J�E���^
    Dim p As Long                                   '' �؂蕪���J�n�ʒu
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMidByte(bytBuf, p, 2)              '' ���R�[�h���
            .DataKubun = IncMidByte(bytBuf, p, 1)               '' �f�[�^�敪
            With .MakeDate
                .Year = IncMidByte(bytBuf, p, 4)                '' �N
                .Month = IncMidByte(bytBuf, p, 2)               '' ��
                .Day = IncMidByte(bytBuf, p, 2)                 '' ��
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMidByte(bytBuf, p, 4)                    '' �J�ÔN
            .MonthDay = IncMidByte(bytBuf, p, 4)                '' �J�Ì���
            .JyoCD = IncMidByte(bytBuf, p, 2)                   '' ���n��R�[�h
            .Kaiji = IncMidByte(bytBuf, p, 2)                   '' �J�É�[��N��]
            .Nichiji = IncMidByte(bytBuf, p, 2)                 '' �J�Ó���[N����]
            .RaceNum = IncMidByte(bytBuf, p, 2)                 '' ���[�X�ԍ�
        End With ' id
        With .RaceInfo
            .YoubiCD = IncMidByte(bytBuf, p, 1)                 '' �j���R�[�h
            .TokuNum = IncMidByte(bytBuf, p, 4)                 '' ���ʋ����ԍ�
            .Hondai = IncMidByte(bytBuf, p, 60)                 '' �������{��
            .Fukudai = IncMidByte(bytBuf, p, 60)                '' ����������
            .Kakko = IncMidByte(bytBuf, p, 60)                  '' �������J�b�R��
            .HondaiEng = IncMidByte(bytBuf, p, 120)             '' �������{�艢��
            .FukudaiEng = IncMidByte(bytBuf, p, 120)            '' ���������艢��
            .KakkoEng = IncMidByte(bytBuf, p, 120)              '' �������J�b�R������
            .Ryakusyo10 = IncMidByte(bytBuf, p, 20)             '' ���������̂P�O��
            .Ryakusyo6 = IncMidByte(bytBuf, p, 12)              '' ���������̂U��
            .Ryakusyo3 = IncMidByte(bytBuf, p, 6)               '' ���������̂R��
            .Kubun = IncMidByte(bytBuf, p, 1)                   '' �������敪
            .Nkai = IncMidByte(bytBuf, p, 3)                    '' �d�܉�[��N��]
        End With ' RaceInfo
        .GradeCD = IncMidByte(bytBuf, p, 1)                     '' �O���[�h�R�[�h
        With .JyokenInfo
            .SyubetuCD = IncMidByte(bytBuf, p, 2)               '' ������ʃR�[�h
            .KigoCD = IncMidByte(bytBuf, p, 3)                  '' �����L���R�[�h
            .JyuryoCD = IncMidByte(bytBuf, p, 1)                '' �d�ʎ�ʃR�[�h
            For j = 0 To 4
                .JyokenCD(j) = IncMidByte(bytBuf, p, 3)         '' ���������R�[�h
            Next j
        End With ' JyokenInfo
        .KYORI = IncMidByte(bytBuf, p, 4)                       '' ����
        .TrackCD = IncMidByte(bytBuf, p, 2)                     '' �g���b�N�R�[�h
        .CourseKubunCD = IncMidByte(bytBuf, p, 2)               '' �R�[�X�敪
        With .HandiDate
            .Year = IncMidByte(bytBuf, p, 4)                    '' �N
            .Month = IncMidByte(bytBuf, p, 2)                   '' ��
            .Day = IncMidByte(bytBuf, p, 2)                     '' ��
        End With ' HandiDate
        .TorokuTosu = IncMidByte(bytBuf, p, 3)                  '' �o�^����
        For i = 0 To 299
            With .TokuUmaInfo(i)
                .num = IncMidByte(bytBuf, p, 3)                 '' �A��
                .KettoNum = IncMidByte(bytBuf, p, 10)           '' �����o�^�ԍ�
                .BAMEI = IncMidByte(bytBuf, p, 36)              '' �n��
                .UmaKigoCD = IncMidByte(bytBuf, p, 2)           '' �n�L���R�[�h
                .SexCD = IncMidByte(bytBuf, p, 1)               '' ���ʃR�[�h
                .TozaiCD = IncMidByte(bytBuf, p, 1)             '' �����t���������R�[�h
                .ChokyosiCode = IncMidByte(bytBuf, p, 5)        '' �����t�R�[�h
                .ChokyosiRyakusyo = IncMidByte(bytBuf, p, 8)    '' �����t������
                .Futan = IncMidByte(bytBuf, p, 3)               '' ���S�d��
                .Koryu = IncMidByte(bytBuf, p, 1)               '' �𗬋敪
            End With ' TokuUmaInfo
        Next i
        .CRLF = IncMidByte(bytBuf, p, 2)                        '' ���R�[�h���
    End With

    End Sub


'
'   �@�\: �\���̂Ƀf�[�^���Z�b�g����|���[�X�ڍ�
'
'   ���l: �Ȃ�
'
    Public Sub SetDataFromByte_RA(ByRef bytBuf() As Byte, ByRef mBuf As JV_RA_RACE)
    Dim i As Integer                                '' ���[�v�J�E���^
    Dim j As Integer                                '' ���[�v�J�E���^
    Dim k As Integer                                '' ���[�v�J�E���^
    Dim p As Long                                   '' �؂蕪���J�n�ʒu
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMidByte(bytBuf, p, 2)              '' ���R�[�h���
            .DataKubun = IncMidByte(bytBuf, p, 1)               '' �f�[�^�敪
            With .MakeDate
                .Year = IncMidByte(bytBuf, p, 4)                '' �N
                .Month = IncMidByte(bytBuf, p, 2)               '' ��
                .Day = IncMidByte(bytBuf, p, 2)                 '' ��
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMidByte(bytBuf, p, 4)                    '' �J�ÔN
            .MonthDay = IncMidByte(bytBuf, p, 4)                '' �J�Ì���
            .JyoCD = IncMidByte(bytBuf, p, 2)                   '' ���n��R�[�h
            .Kaiji = IncMidByte(bytBuf, p, 2)                   '' �J�É�[��N��]
            .Nichiji = IncMidByte(bytBuf, p, 2)                 '' �J�Ó���[N����]
            .RaceNum = IncMidByte(bytBuf, p, 2)                 '' ���[�X�ԍ�
        End With ' id
        With .RaceInfo
            .YoubiCD = IncMidByte(bytBuf, p, 1)                 '' �j���R�[�h
            .TokuNum = IncMidByte(bytBuf, p, 4)                 '' ���ʋ����ԍ�
            .Hondai = IncMidByte(bytBuf, p, 60)                 '' �������{��
            .Fukudai = IncMidByte(bytBuf, p, 60)                '' ����������
            .Kakko = IncMidByte(bytBuf, p, 60)                  '' �������J�b�R��
            .HondaiEng = IncMidByte(bytBuf, p, 120)             '' �������{�艢��
            .FukudaiEng = IncMidByte(bytBuf, p, 120)            '' ���������艢��
            .KakkoEng = IncMidByte(bytBuf, p, 120)              '' �������J�b�R������
            .Ryakusyo10 = IncMidByte(bytBuf, p, 20)             '' ���������̂P�O��
            .Ryakusyo6 = IncMidByte(bytBuf, p, 12)              '' ���������̂U��
            .Ryakusyo3 = IncMidByte(bytBuf, p, 6)               '' ���������̂R��
            .Kubun = IncMidByte(bytBuf, p, 1)                   '' �������敪
            .Nkai = IncMidByte(bytBuf, p, 3)                    '' �d�܉�[��N��]
        End With ' RaceInfo
        .GradeCD = IncMidByte(bytBuf, p, 1)                     '' �O���[�h�R�[�h
        .GradeCDBefore = IncMidByte(bytBuf, p, 1)               '' �ύX�O�O���[�h�R�[�h
        With .JyokenInfo
            .SyubetuCD = IncMidByte(bytBuf, p, 2)               '' ������ʃR�[�h
            .KigoCD = IncMidByte(bytBuf, p, 3)                  '' �����L���R�[�h
            .JyuryoCD = IncMidByte(bytBuf, p, 1)                '' �d�ʎ�ʃR�[�h
            For j = 0 To 4
                .JyokenCD(j) = IncMidByte(bytBuf, p, 3)         '' ���������R�[�h
            Next j
        End With ' JyokenInfo
        .JyokenName = IncMidByte(bytBuf, p, 60)                 '' ������������
        .KYORI = IncMidByte(bytBuf, p, 4)                       '' ����
        .KyoriBefore = IncMidByte(bytBuf, p, 4)                 '' �ύX�O����
        .TrackCD = IncMidByte(bytBuf, p, 2)                     '' �g���b�N�R�[�h
        .TrackCDBefore = IncMidByte(bytBuf, p, 2)               '' �ύX�O�g���b�N�R�[�h
        .CourseKubunCD = IncMidByte(bytBuf, p, 2)               '' �R�[�X�敪
        .CourseKubunCDBefore = IncMidByte(bytBuf, p, 2)         '' �ύX�O�R�[�X�敪
        For i = 0 To 6
            .Honsyokin(i) = IncMidByte(bytBuf, p, 8)            '' �{�܋�
        Next i
        For i = 0 To 4
            .HonsyokinBefore(i) = IncMidByte(bytBuf, p, 8)      '' �ύX�O�{�܋�
        Next i
        For i = 0 To 4
            .Fukasyokin(i) = IncMidByte(bytBuf, p, 8)           '' �t���܋�
        Next i
        For i = 0 To 2
            .FukasyokinBefore(i) = IncMidByte(bytBuf, p, 8)     '' �ύX�O�t���܋�
        Next i
        .HassoTime = IncMidByte(bytBuf, p, 4)                   '' ��������
        .HassoTimeBefore = IncMidByte(bytBuf, p, 4)             '' �ύX�O��������
        .TorokuTosu = IncMidByte(bytBuf, p, 2)                  '' �o�^����
        .SyussoTosu = IncMidByte(bytBuf, p, 2)                  '' �o������
        .NyusenTosu = IncMidByte(bytBuf, p, 2)                  '' ��������
        With .TenkoBaba
            .TenkoCD = IncMidByte(bytBuf, p, 1)                 '' �V��R�[�h
            .SibaBabaCD = IncMidByte(bytBuf, p, 1)              '' �Ŕn���ԃR�[�h
            .DirtBabaCD = IncMidByte(bytBuf, p, 1)              '' �_�[�g�n���ԃR�[�h
        End With ' TenkoBaba
        For i = 0 To 24
            .LapTime(i) = IncMidByte(bytBuf, p, 3)              '' ���b�v�^�C��
        Next i
        .SyogaiMileTime = IncMidByte(bytBuf, p, 4)              '' ��Q�}�C���^�C��
        .HaronTimeS3 = IncMidByte(bytBuf, p, 3)                 '' �O�R�n�����^�C��
        .HaronTimeS4 = IncMidByte(bytBuf, p, 3)                 '' �O�S�n�����^�C��
        .HaronTimeL3 = IncMidByte(bytBuf, p, 3)                 '' ��R�n�����^�C��
        .HaronTimeL4 = IncMidByte(bytBuf, p, 3)                 '' ��S�n�����^�C��
        For i = 0 To 3
            With .CornerInfo(i)
                .Corner = IncMidByte(bytBuf, p, 1)              '' �R�[�i�[
                .Syukaisu = IncMidByte(bytBuf, p, 1)            '' ����
                .Jyuni = IncMidByte(bytBuf, p, 70)              '' �e�ʉߏ���
            End With ' CornerInfo
        Next i
        .RecordUpKubun = IncMidByte(bytBuf, p, 1)               '' ���R�[�h�X�V�敪
        .CRLF = IncMidByte(bytBuf, p, 2)        '' ���R�[�h��؂�
    End With

    End Sub


'
'   �@�\: �\���̂Ƀf�[�^���Z�b�g����|�n�����[�X���
'
'   ���l: �Ȃ�
'
    Public Sub SetDataFromByte_SE(ByRef bytBuf() As Byte, ByRef mBuf As JV_SE_RACE_UMA)
    Dim i As Integer                                '' ���[�v�J�E���^
    Dim p As Long                                   '' �؂蕪���J�n�ʒu
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMidByte(bytBuf, p, 2)      '' ���R�[�h���
            .DataKubun = IncMidByte(bytBuf, p, 1)       '' �f�[�^�敪
            With .MakeDate
                .Year = IncMidByte(bytBuf, p, 4)        '' �N
                .Month = IncMidByte(bytBuf, p, 2)       '' ��
                .Day = IncMidByte(bytBuf, p, 2)         '' ��
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMidByte(bytBuf, p, 4)            '' �J�ÔN
            .MonthDay = IncMidByte(bytBuf, p, 4)        '' �J�Ì���
            .JyoCD = IncMidByte(bytBuf, p, 2)           '' ���n��R�[�h
            .Kaiji = IncMidByte(bytBuf, p, 2)           '' �J�É�[��N��]
            .Nichiji = IncMidByte(bytBuf, p, 2)         '' �J�Ó���[N����]
            .RaceNum = IncMidByte(bytBuf, p, 2)         '' ���[�X�ԍ�
        End With ' id
        .Wakuban = IncMidByte(bytBuf, p, 1)             '' �g��
        .Umaban = IncMidByte(bytBuf, p, 2)              '' �n��
        .KettoNum = IncMidByte(bytBuf, p, 10)           '' �����o�^�ԍ�
        .BAMEI = IncMidByte(bytBuf, p, 36)              '' �n��
        .UmaKigoCD = IncMidByte(bytBuf, p, 2)           '' �n�L���R�[�h
        .SexCD = IncMidByte(bytBuf, p, 1)               '' ���ʃR�[�h
        .HinsyuCD = IncMidByte(bytBuf, p, 1)            '' �i��R�[�h
        .KeiroCD = IncMidByte(bytBuf, p, 2)             '' �ѐF�R�[�h
        .Barei = IncMidByte(bytBuf, p, 2)               '' �n��
        .TozaiCD = IncMidByte(bytBuf, p, 1)             '' ���������R�[�h
        .ChokyosiCode = IncMidByte(bytBuf, p, 5)        '' �����t�R�[�h
        .ChokyosiRyakusyo = IncMidByte(bytBuf, p, 8)    '' �����t������
        .BanusiCode = IncMidByte(bytBuf, p, 6)          '' �n��R�[�h
        .BanusiName = IncMidByte(bytBuf, p, 64)         '' �n�喼
        .Fukusyoku = IncMidByte(bytBuf, p, 60)          '' ���F�W��
        .reserved1 = IncMidByte(bytBuf, p, 60)          '' �\��
        .Futan = IncMidByte(bytBuf, p, 3)               '' ���S�d��
        .FutanBefore = IncMidByte(bytBuf, p, 3)         '' �ύX�O���S�d��
        .BLINKER = IncMidByte(bytBuf, p, 1)             '' �u�����J�[�g�p�敪
        .reserved2 = IncMidByte(bytBuf, p, 1)           '' �\��
        .KisyuCode = IncMidByte(bytBuf, p, 5)           '' �R��R�[�h
        .KisyuCodeBefore = IncMidByte(bytBuf, p, 5)     '' �ύX�O�R��R�[�h
        .KisyuRyakusyo = IncMidByte(bytBuf, p, 8)       '' �R�薼����
        .KisyuRyakusyoBefore = IncMidByte(bytBuf, p, 8) '' �ύX�O�R�薼����
        .MinaraiCD = IncMidByte(bytBuf, p, 1)           '' �R�茩�K�R�[�h
        .MinaraiCDBefore = IncMidByte(bytBuf, p, 1)     '' �ύX�O�R�茩�K�R�[�h
        .BaTaijyu = IncMidByte(bytBuf, p, 3)            '' �n�̏d
        .ZogenFugo = IncMidByte(bytBuf, p, 1)           '' ��������
        .ZogenSa = IncMidByte(bytBuf, p, 3)             '' ������
        .IJyoCD = IncMidByte(bytBuf, p, 1)              '' �ُ�敪�R�[�h
        .NyusenJyuni = IncMidByte(bytBuf, p, 2)         '' ��������
        .KakuteiJyuni = IncMidByte(bytBuf, p, 2)        '' �m�蒅��
        .DochakuKubun = IncMidByte(bytBuf, p, 1)        '' �����敪
        .DochakuTosu = IncMidByte(bytBuf, p, 1)         '' ��������
        .TIME = IncMidByte(bytBuf, p, 4)                '' ���j�^�C��
        .ChakusaCD = IncMidByte(bytBuf, p, 3)           '' �����R�[�h
        .ChakusaCDP = IncMidByte(bytBuf, p, 3)          '' +�����R�[�h
        .ChakusaCDPP = IncMidByte(bytBuf, p, 3)         '' ++�����R�[�h
        .Jyuni1c = IncMidByte(bytBuf, p, 2)             '' 1�R�[�i�[�ł̏���
        .Jyuni2c = IncMidByte(bytBuf, p, 2)             '' 2�R�[�i�[�ł̏���
        .Jyuni3c = IncMidByte(bytBuf, p, 2)             '' 3�R�[�i�[�ł̏���
        .Jyuni4c = IncMidByte(bytBuf, p, 2)             '' 4�R�[�i�[�ł̏���
        .Odds = IncMidByte(bytBuf, p, 4)                '' �P���I�b�Y
        .Ninki = IncMidByte(bytBuf, p, 2)               '' �P���l�C��
        .Honsyokin = IncMidByte(bytBuf, p, 8)           '' �l���{�܋�
        .Fukasyokin = IncMidByte(bytBuf, p, 8)          '' �l���t���܋�
        .reserved3 = IncMidByte(bytBuf, p, 3)           '' �\��
        .reserved4 = IncMidByte(bytBuf, p, 3)           '' �\��
        .HaronTimeL4 = IncMidByte(bytBuf, p, 3)         '' ��S�n�����^�C��
        .HaronTimeL3 = IncMidByte(bytBuf, p, 3)         '' ��R�n�����^�C��
        For i = 0 To 2
            With .ChakuUmaInfo(i)
                .KettoNum = IncMidByte(bytBuf, p, 10)   '' �����o�^�ԍ�
                .BAMEI = IncMidByte(bytBuf, p, 36)      '' �n��
            End With ' ChakuUmaInfo
        Next i
        .TimeDiff = IncMidByte(bytBuf, p, 4)            '' �^�C����
        .RecordUpKubun = IncMidByte(bytBuf, p, 1)       '' ���R�[�h�X�V�敪
        .DMKubun = IncMidByte(bytBuf, p, 1)             '' �}�C�j���O�敪
        .DMTime = IncMidByte(bytBuf, p, 5)              '' �}�C�j���O�\�z���j�^�C��
        .DMGosaP = IncMidByte(bytBuf, p, 4)             '' �\���덷(�M���x)�{
        .DMGosaM = IncMidByte(bytBuf, p, 4)             '' �\���덷(�M���x)�|
        .DMJyuni = IncMidByte(bytBuf, p, 2)             '' �}�C�j���O�\�z����
        .KyakusituKubun = IncMidByte(bytBuf, p, 1)      '' ���񃌁[�X�r������
        .CRLF = IncMidByte(bytBuf, p, 2)                '' ���R�[�h��؂�
    End With

    End Sub


'
'   �@�\: �\���̂Ƀf�[�^���Z�b�g����|����
'
'   ���l: �Ȃ�
'
    Public Sub SetDataFromByte_HR(bytBuf() As Byte, ByRef mBuf As JV_HR_PAY)
    Dim i As Integer                                '' ���[�v�J�E���^
    Dim j As Integer                                '' ���[�v�J�E���^
    Dim k As Integer                                '' ���[�v�J�E���^
    Dim p As Long                                   '' �؂蕪���J�n�ʒu
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMidByte(bytBuf, p, 2)          '' ���R�[�h���
            .DataKubun = IncMidByte(bytBuf, p, 1)           '' �f�[�^�敪
            With .MakeDate
                .Year = IncMidByte(bytBuf, p, 4)            '' �N
                .Month = IncMidByte(bytBuf, p, 2)           '' ��
                .Day = IncMidByte(bytBuf, p, 2)             '' ��
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMidByte(bytBuf, p, 4)                '' �J�ÔN
            .MonthDay = IncMidByte(bytBuf, p, 4)            '' �J�Ì���
            .JyoCD = IncMidByte(bytBuf, p, 2)               '' ���n��R�[�h
            .Kaiji = IncMidByte(bytBuf, p, 2)               '' �J�É�[��N��]
            .Nichiji = IncMidByte(bytBuf, p, 2)             '' �J�Ó���[N����]
            .RaceNum = IncMidByte(bytBuf, p, 2)             '' ���[�X�ԍ�
        End With ' id
        .TorokuTosu = IncMidByte(bytBuf, p, 2)              '' �o�^����
        .SyussoTosu = IncMidByte(bytBuf, p, 2)              '' �o������
        For i = 0 To 8
            .FuseirituFlag(i) = IncMidByte(bytBuf, p, 1)    '' �s�����t���O
        Next i
        For i = 0 To 8
            .TokubaraiFlag(i) = IncMidByte(bytBuf, p, 1)    '' �����t���O
        Next i
        For i = 0 To 8
            .HenkanFlag(i) = IncMidByte(bytBuf, p, 1)       '' �Ԋ҃t���O
        Next i
        For i = 0 To 27
            .HenkanUma(i) = IncMidByte(bytBuf, p, 1)        '' �ԊҔn�ԏ��(�n��01�`28)
        Next i
        For i = 0 To 7
            .HenkanWaku(i) = IncMidByte(bytBuf, p, 1)       '' �ԊҘg�ԏ��(�g��1�`8)
        Next i
        For i = 0 To 7
            .HenkanDoWaku(i) = IncMidByte(bytBuf, p, 1)     '' �Ԋғ��g���(�g��1�`8)
        Next i
        For i = 0 To 2
            With .PayTansyo(i)
                .Umaban = IncMidByte(bytBuf, p, 2)          '' �n��
                .Pay = IncMidByte(bytBuf, p, 9)             '' ���ߋ�
                .Ninki = IncMidByte(bytBuf, p, 2)           '' �l�C��
            End With ' PayTansyo
        Next i
        For i = 0 To 4
            With .PayFukusyo(i)
                .Umaban = IncMidByte(bytBuf, p, 2)          '' �n��
                .Pay = IncMidByte(bytBuf, p, 9)             '' ���ߋ�
                .Ninki = IncMidByte(bytBuf, p, 2)           '' �l�C��
            End With ' PayFukusyo
        Next i
        For i = 0 To 2
            With .PayWakuren(i)
                .Umaban = IncMidByte(bytBuf, p, 2)          '' �n��
                .Pay = IncMidByte(bytBuf, p, 9)             '' ���ߋ�
                .Ninki = IncMidByte(bytBuf, p, 2)           '' �l�C��
            End With ' PayWakuren
        Next i
        For i = 0 To 2
            With .PayUmaren(i)
                .Kumi = IncMidByte(bytBuf, p, 4)            '' �g��
                .Pay = IncMidByte(bytBuf, p, 9)             '' ���ߋ�
                .Ninki = IncMidByte(bytBuf, p, 3)           '' �l�C��
            End With ' PayUmaren
        Next i
        For i = 0 To 6
            With .PayWide(i)
                .Kumi = IncMidByte(bytBuf, p, 4)            '' �g��
                .Pay = IncMidByte(bytBuf, p, 9)             '' ���ߋ�
                .Ninki = IncMidByte(bytBuf, p, 3)           '' �l�C��
            End With ' PayWide
        Next i
        For i = 0 To 2
            With .PayReserved1(i)
                .Kumi = IncMidByte(bytBuf, p, 4)            '' �g��
                .Pay = IncMidByte(bytBuf, p, 9)             '' ���ߋ�
                .Ninki = IncMidByte(bytBuf, p, 3)           '' �l�C��
            End With ' PayReserved1
        Next i
        For i = 0 To 5
            With .PayUmatan(i)
                .Kumi = IncMidByte(bytBuf, p, 4)            '' �g��
                .Pay = IncMidByte(bytBuf, p, 9)             '' ���ߋ�
                .Ninki = IncMidByte(bytBuf, p, 3)           '' �l�C��
            End With ' PayUmatan
        Next i
        For i = 0 To 2
            With .PaySanrenpuku(i)
                .Kumi = IncMidByte(bytBuf, p, 6)            '' �g��
                .Pay = IncMidByte(bytBuf, p, 9)             '' ���ߋ�
                .Ninki = IncMidByte(bytBuf, p, 3)           '' �l�C��
            End With ' PaySanrenpuku
        Next i
        For i = 0 To 5
            With .PaySanrentan(i)
                .Kumi = IncMidByte(bytBuf, p, 6)            '' �g��
                .Pay = IncMidByte(bytBuf, p, 9)             '' ���ߋ�
                .Ninki = IncMidByte(bytBuf, p, 4)           '' �l�C��
            End With ' PaySanrentan
        Next i
        .CRLF = IncMidByte(bytBuf, p, 2)        '' ���R�[�h��؂�
    End With
   
    End Sub


'
'   �@�\: �\���̂Ƀf�[�^���Z�b�g����|�[���i�S�|���j
'
'   ���l: �Ȃ�
'
    Public Sub SetDataFromByte_H1(bytBuf() As Byte, ByRef mBuf As JV_H1_HYOSU_ZENKAKE)
    Dim i As Integer                                '' ���[�v�J�E���^
    Dim j As Integer                                '' ���[�v�J�E���^
    Dim k As Integer                                '' ���[�v�J�E���^
    Dim p As Long                                   '' �؂蕪���J�n�ʒu
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMidByte(bytBuf, p, 2)      '' ���R�[�h���
            .DataKubun = IncMidByte(bytBuf, p, 1)       '' �f�[�^�敪
            With .MakeDate
                .Year = IncMidByte(bytBuf, p, 4)        '' �N
                .Month = IncMidByte(bytBuf, p, 2)       '' ��
                .Day = IncMidByte(bytBuf, p, 2)         '' ��
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMidByte(bytBuf, p, 4)            '' �J�ÔN
            .MonthDay = IncMidByte(bytBuf, p, 4)        '' �J�Ì���
            .JyoCD = IncMidByte(bytBuf, p, 2)           '' ���n��R�[�h
            .Kaiji = IncMidByte(bytBuf, p, 2)           '' �J�É�[��N��]
            .Nichiji = IncMidByte(bytBuf, p, 2)         '' �J�Ó���[N����]
            .RaceNum = IncMidByte(bytBuf, p, 2)         '' ���[�X�ԍ�
        End With ' id
        .TorokuTosu = IncMidByte(bytBuf, p, 2)          '' �o�^����
        .SyussoTosu = IncMidByte(bytBuf, p, 2)          '' �o������
        For i = 0 To 6
            .HatubaiFlag(i) = IncMidByte(bytBuf, p, 1)  '' �����t���O
        Next i
        .FukuChakuBaraiKey = IncMidByte(bytBuf, p, 1)   '' ���������L�[
        For i = 0 To 27
            .HenkanUma(i) = IncMidByte(bytBuf, p, 1)    '' �ԊҔn�ԏ��(�n��01�`28)
        Next i
        For i = 0 To 7
            .HenkanWaku(i) = IncMidByte(bytBuf, p, 1)   '' �ԊҘg�ԏ��(�g��1�`8)
        Next i
        For i = 0 To 7
            .HenkanDoWaku(i) = IncMidByte(bytBuf, p, 1) '' �Ԋғ��g���(�g��1�`8)
        Next i
        For i = 0 To 27
            With .HyoTansyo(i)
                .Umaban = IncMidByte(bytBuf, p, 2)      '' �n��
                .Hyo = IncMidByte(bytBuf, p, 11)        '' �[��
                .Ninki = IncMidByte(bytBuf, p, 2)       '' �l�C
            End With ' HyoTansyo
        Next i
        For i = 0 To 27
            With .HyoFukusyo(i)
                .Umaban = IncMidByte(bytBuf, p, 2)      '' �n��
                .Hyo = IncMidByte(bytBuf, p, 11)        '' �[��
                .Ninki = IncMidByte(bytBuf, p, 2)       '' �l�C
            End With ' HyoFukusyo
        Next i
        For i = 0 To 35
            With .HyoWakuren(i)
                .Umaban = IncMidByte(bytBuf, p, 2)      '' �n��
                .Hyo = IncMidByte(bytBuf, p, 11)        '' �[��
                .Ninki = IncMidByte(bytBuf, p, 2)       '' �l�C
            End With ' HyoWakuren
        Next i
        For i = 0 To 152
            With .HyoUmaren(i)
                .Kumi = IncMidByte(bytBuf, p, 4)        '' �g��
                .Hyo = IncMidByte(bytBuf, p, 11)        '' �[��
                .Ninki = IncMidByte(bytBuf, p, 3)       '' �l�C
            End With ' HyoUmaren
        Next i
        For i = 0 To 152
            With .HyoWide(i)
                .Kumi = IncMidByte(bytBuf, p, 4)        '' �g��
                .Hyo = IncMidByte(bytBuf, p, 11)        '' �[��
                .Ninki = IncMidByte(bytBuf, p, 3)       '' �l�C
            End With ' HyoWide
        Next i
        For i = 0 To 305
            With .HyoUmatan(i)
                .Kumi = IncMidByte(bytBuf, p, 4)        '' �g��
                .Hyo = IncMidByte(bytBuf, p, 11)        '' �[��
                .Ninki = IncMidByte(bytBuf, p, 3)       '' �l�C
            End With ' HyoUmatan
        Next i
        For i = 0 To 815
            With .HyoSanrenpuku(i)
                .Kumi = IncMidByte(bytBuf, p, 6)        '' �g��
                .Hyo = IncMidByte(bytBuf, p, 11)        '' �[��
                .Ninki = IncMidByte(bytBuf, p, 3)       '' �l�C
            End With ' HyoSanrenpuku
        Next i
        For i = 0 To 13
            .HyoTotal(i) = IncMidByte(bytBuf, p, 11)    '' �[�����v
        Next i
        .CRLF = IncMidByte(bytBuf, p, 2)                '' ���R�[�h��؂�
    End With
    
    End Sub


'
'   �@�\: �\���̂Ƀf�[�^���Z�b�g����|�[���U�i�R�A�P�j
'
'   ���l: �Ȃ�
'
    Public Sub SetDataFromByte_H6(bytBuf() As Byte, ByRef mBuf As JV_H6_HYOSU_SANRENTAN)
    Dim i As Integer                                '' ���[�v�J�E���^
    Dim p As Long                                   '' �؂蕪���J�n�ʒu
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMidByte(bytBuf, p, 2)      '' ���R�[�h���
            .DataKubun = IncMidByte(bytBuf, p, 1)       '' �f�[�^�敪
            With .MakeDate
                .Year = IncMidByte(bytBuf, p, 4)        '' �N
                .Month = IncMidByte(bytBuf, p, 2)       '' ��
                .Day = IncMidByte(bytBuf, p, 2)         '' ��
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMidByte(bytBuf, p, 4)            '' �J�ÔN
            .MonthDay = IncMidByte(bytBuf, p, 4)        '' �J�Ì���
            .JyoCD = IncMidByte(bytBuf, p, 2)           '' ���n��R�[�h
            .Kaiji = IncMidByte(bytBuf, p, 2)           '' �J�É�[��N��]
            .Nichiji = IncMidByte(bytBuf, p, 2)         '' �J�Ó���[N����]
            .RaceNum = IncMidByte(bytBuf, p, 2)         '' ���[�X�ԍ�
        End With ' id
        .TorokuTosu = IncMidByte(bytBuf, p, 2)          '' �o�^����
        .SyussoTosu = IncMidByte(bytBuf, p, 2)          '' �o������
        .HatubaiFlag = IncMidByte(bytBuf, p, 1)         '' �����t���O 3�A�P
        For i = 0 To 17
            .HenkanUma(i) = IncMidByte(bytBuf, p, 1)    '' �ԊҔn�ԏ��(�n��01�`18)
        Next i
        For i = 0 To 4895
            With .HyoSanrentan(i)
                .Kumi = IncMidByte(bytBuf, p, 6)        '' �g��
                .Hyo = IncMidByte(bytBuf, p, 11)        '' �[��
                .Ninki = IncMidByte(bytBuf, p, 4)       '' �l�C
            End With ' HyoSanrentan
        Next i
        .TotalHyoSanrentan = IncMidByte(bytBuf, p, 11)    '' 3�A�P�[�����v
        .TotalHyoSanrentanHenkan = IncMidByte(bytBuf, p, 11) '' 3�A�P�Ԋҕ[�����v
        .CRLF = IncMidByte(bytBuf, p, 2)                  '' ���R�[�h��؂�
    End With
    End Sub
    
    
'
'   �@�\: �\���̂Ƀf�[�^���Z�b�g����|�I�b�Y�i�P���g�j
'
'   ���l: �Ȃ�
'
    Public Sub SetDataFromByte_O1(bytBuf() As Byte, ByRef mBuf As JV_O1_ODDS_TANFUKUWAKU)
    Dim i As Integer                                '' ���[�v�J�E���^
    Dim j As Integer                                '' ���[�v�J�E���^
    Dim k As Integer                                '' ���[�v�J�E���^
    Dim p As Long                                   '' �؂蕪���J�n�ʒu
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMidByte(bytBuf, p, 2)      '' ���R�[�h���
            .DataKubun = IncMidByte(bytBuf, p, 1)       '' �f�[�^�敪
            With .MakeDate
                .Year = IncMidByte(bytBuf, p, 4)        '' �N
                .Month = IncMidByte(bytBuf, p, 2)       '' ��
                .Day = IncMidByte(bytBuf, p, 2) '' ��
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMidByte(bytBuf, p, 4)            '' �J�ÔN
            .MonthDay = IncMidByte(bytBuf, p, 4)        '' �J�Ì���
            .JyoCD = IncMidByte(bytBuf, p, 2)           '' ���n��R�[�h
            .Kaiji = IncMidByte(bytBuf, p, 2)           '' �J�É�[��N��]
            .Nichiji = IncMidByte(bytBuf, p, 2)         '' �J�Ó���[N����]
            .RaceNum = IncMidByte(bytBuf, p, 2)         '' ���[�X�ԍ�
        End With ' id
        With .HappyoTime
            .Month = IncMidByte(bytBuf, p, 2)           '' ��
            .Day = IncMidByte(bytBuf, p, 2)             '' ��
            .Hour = IncMidByte(bytBuf, p, 2)            '' ��
            .Minute = IncMidByte(bytBuf, p, 2)          '' ��
        End With ' HappyoTime
        .TorokuTosu = IncMidByte(bytBuf, p, 2)          '' �o�^����
        .SyussoTosu = IncMidByte(bytBuf, p, 2)          '' �o������
        .TansyoFlag = IncMidByte(bytBuf, p, 1)          '' �����t���O
        .FukusyoFlag = IncMidByte(bytBuf, p, 1)         '' �����t���O
        .WakurenFlag = IncMidByte(bytBuf, p, 1)         '' �����t���O�@�g�A
        .FukuChakuBaraiKey = IncMidByte(bytBuf, p, 1)   '' ���������L�[
        For i = 0 To 27
            With .OddsTansyoInfo(i)
                .Umaban = IncMidByte(bytBuf, p, 2)      '' �n��
                .Odds = IncMidByte(bytBuf, p, 4)        '' �I�b�Y
                .Ninki = IncMidByte(bytBuf, p, 2)       '' �l�C��
            End With ' OddsTansyoInfo
        Next i
        For i = 0 To 27
            With .OddsFukusyoInfo(i)
                .Umaban = IncMidByte(bytBuf, p, 2)      '' �n��
                .OddsLow = IncMidByte(bytBuf, p, 4)     '' �Œ�I�b�Y
                .OddsHigh = IncMidByte(bytBuf, p, 4)    '' �ō��I�b�Y
                .Ninki = IncMidByte(bytBuf, p, 2)       '' �l�C��
            End With ' OddsFukusyoInfo
        Next i
        For i = 0 To 35
            With .OddsWakurenInfo(i)
                .Kumi = IncMidByte(bytBuf, p, 2)        '' �g
                .Odds = IncMidByte(bytBuf, p, 5)        '' �I�b�Y
                .Ninki = IncMidByte(bytBuf, p, 2)       '' �l�C��
            End With ' OddsWakurenInfo
        Next i
        .TotalHyosuTansyo = IncMidByte(bytBuf, p, 11)   '' �P���[�����v
        .TotalHyosuFukusyo = IncMidByte(bytBuf, p, 11)  '' �����[�����v
        .TotalHyosuWakuren = IncMidByte(bytBuf, p, 11)  '' �g�A�[�����v
        .CRLF = IncMidByte(bytBuf, p, 2)                '' ���R�[�h��؂�
    End With

    End Sub


'
'   �@�\: �\���̂Ƀf�[�^���Z�b�g����|�I�b�Y�i�n�A�j
'
'   ���l: �Ȃ�
'
    Public Sub SetDataFromByte_O2(bytBuf() As Byte, ByRef mBuf As JV_O2_ODDS_UMAREN)
    Dim i As Integer                                '' ���[�v�J�E���^
    Dim j As Integer                                '' ���[�v�J�E���^
    Dim k As Integer                                '' ���[�v�J�E���^
    Dim p As Long                                   '' �؂蕪���J�n�ʒu
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMidByte(bytBuf, p, 2)      '' ���R�[�h���
            .DataKubun = IncMidByte(bytBuf, p, 1)       '' �f�[�^�敪
            With .MakeDate
                .Year = IncMidByte(bytBuf, p, 4)        '' �N
                .Month = IncMidByte(bytBuf, p, 2)       '' ��
                .Day = IncMidByte(bytBuf, p, 2) '' ��
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMidByte(bytBuf, p, 4)    '' �J�ÔN
            .MonthDay = IncMidByte(bytBuf, p, 4)        '' �J�Ì���
            .JyoCD = IncMidByte(bytBuf, p, 2)   '' ���n��R�[�h
            .Kaiji = IncMidByte(bytBuf, p, 2)   '' �J�É�[��N��]
            .Nichiji = IncMidByte(bytBuf, p, 2) '' �J�Ó���[N����]
            .RaceNum = IncMidByte(bytBuf, p, 2) '' ���[�X�ԍ�
        End With ' id
        With .HappyoTime
            .Month = IncMidByte(bytBuf, p, 2)   '' ��
            .Day = IncMidByte(bytBuf, p, 2)     '' ��
            .Hour = IncMidByte(bytBuf, p, 2)    '' ��
            .Minute = IncMidByte(bytBuf, p, 2)  '' ��
        End With ' HappyoTime
        .TorokuTosu = IncMidByte(bytBuf, p, 2)  '' �o�^����
        .SyussoTosu = IncMidByte(bytBuf, p, 2)  '' �o������
        .UmarenFlag = IncMidByte(bytBuf, p, 1)  '' �����t���O�@�n�A
        For i = 0 To 152
            With .OddsUmarenInfo(i)
                .Kumi = IncMidByte(bytBuf, p, 4)        '' �g��
                .Odds = IncMidByte(bytBuf, p, 6)        '' �I�b�Y
                .Ninki = IncMidByte(bytBuf, p, 3)       '' �l�C��
            End With ' OddsUmarenInfo
        Next i
        .TotalHyosuUmaren = IncMidByte(bytBuf, p, 11)   '' �n�A�[�����v
        .CRLF = IncMidByte(bytBuf, p, 2)        '' ���R�[�h��؂�
    End With

    End Sub


'
'   �@�\: �\���̂Ƀf�[�^���Z�b�g����|�I�b�Y�i���C�h�j
'
'   ���l: �Ȃ�
'
    Public Sub SetDataFromByte_O3(bytBuf() As Byte, ByRef mBuf As JV_O3_ODDS_WIDE)
    Dim i As Integer                                '' ���[�v�J�E���^
    Dim j As Integer                                '' ���[�v�J�E���^
    Dim k As Integer                                '' ���[�v�J�E���^
    Dim p As Long                                   '' �؂蕪���J�n�ʒu
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMidByte(bytBuf, p, 2)      '' ���R�[�h���
            .DataKubun = IncMidByte(bytBuf, p, 1)       '' �f�[�^�敪
            With .MakeDate
                .Year = IncMidByte(bytBuf, p, 4)        '' �N
                .Month = IncMidByte(bytBuf, p, 2)       '' ��
                .Day = IncMidByte(bytBuf, p, 2) '' ��
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMidByte(bytBuf, p, 4)    '' �J�ÔN
            .MonthDay = IncMidByte(bytBuf, p, 4)        '' �J�Ì���
            .JyoCD = IncMidByte(bytBuf, p, 2)   '' ���n��R�[�h
            .Kaiji = IncMidByte(bytBuf, p, 2)   '' �J�É�[��N��]
            .Nichiji = IncMidByte(bytBuf, p, 2) '' �J�Ó���[N����]
            .RaceNum = IncMidByte(bytBuf, p, 2) '' ���[�X�ԍ�
        End With ' id
        With .HappyoTime
            .Month = IncMidByte(bytBuf, p, 2)   '' ��
            .Day = IncMidByte(bytBuf, p, 2)     '' ��
            .Hour = IncMidByte(bytBuf, p, 2)    '' ��
            .Minute = IncMidByte(bytBuf, p, 2)  '' ��
        End With ' HappyoTime
        .TorokuTosu = IncMidByte(bytBuf, p, 2)  '' �o�^����
        .SyussoTosu = IncMidByte(bytBuf, p, 2)  '' �o������
        .WideFlag = IncMidByte(bytBuf, p, 1)    '' �����t���O�@���C�h
        For i = 0 To 152
            With .OddsWideInfo(i)
                .Kumi = IncMidByte(bytBuf, p, 4)        '' �g��
                .OddsLow = IncMidByte(bytBuf, p, 5)     '' �Œ�I�b�Y
                .OddsHigh = IncMidByte(bytBuf, p, 5)    '' �ō��I�b�Y
                .Ninki = IncMidByte(bytBuf, p, 3)       '' �l�C��
            End With ' OddsWideInfo
        Next i
        .TotalHyosuWide = IncMidByte(bytBuf, p, 11)     '' ���C�h�[�����v
        .CRLF = IncMidByte(bytBuf, p, 2)        '' ���R�[�h��؂�
    End With

    End Sub


'
'   �@�\: �\���̂Ƀf�[�^���Z�b�g����|�I�b�Y�i�n�P�j
'
'   ���l: �Ȃ�
'
    Public Sub SetDataFromByte_O4(bytBuf() As Byte, ByRef mBuf As JV_O4_ODDS_UMATAN)
    Dim i As Integer                                '' ���[�v�J�E���^
    Dim j As Integer                                '' ���[�v�J�E���^
    Dim k As Integer                                '' ���[�v�J�E���^
    Dim p As Long                                   '' �؂蕪���J�n�ʒu
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMidByte(bytBuf, p, 2)      '' ���R�[�h���
            .DataKubun = IncMidByte(bytBuf, p, 1)       '' �f�[�^�敪
            With .MakeDate
                .Year = IncMidByte(bytBuf, p, 4)        '' �N
                .Month = IncMidByte(bytBuf, p, 2)       '' ��
                .Day = IncMidByte(bytBuf, p, 2)         '' ��
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMidByte(bytBuf, p, 4)            '' �J�ÔN
            .MonthDay = IncMidByte(bytBuf, p, 4)        '' �J�Ì���
            .JyoCD = IncMidByte(bytBuf, p, 2)           '' ���n��R�[�h
            .Kaiji = IncMidByte(bytBuf, p, 2)           '' �J�É�[��N��]
            .Nichiji = IncMidByte(bytBuf, p, 2)         '' �J�Ó���[N����]
            .RaceNum = IncMidByte(bytBuf, p, 2)         '' ���[�X�ԍ�
        End With ' id
        With .HappyoTime
            .Month = IncMidByte(bytBuf, p, 2)           '' ��
            .Day = IncMidByte(bytBuf, p, 2)             '' ��
            .Hour = IncMidByte(bytBuf, p, 2)            '' ��
            .Minute = IncMidByte(bytBuf, p, 2)          '' ��
        End With ' HappyoTime
        .TorokuTosu = IncMidByte(bytBuf, p, 2)          '' �o�^����
        .SyussoTosu = IncMidByte(bytBuf, p, 2)          '' �o������
        .UmatanFlag = IncMidByte(bytBuf, p, 1)          '' �����t���O�@�n�P
        For i = 0 To 305
            With .OddsUmatanInfo(i)
                .Kumi = IncMidByte(bytBuf, p, 4)        '' �g��
                .Odds = IncMidByte(bytBuf, p, 6)        '' �I�b�Y
                .Ninki = IncMidByte(bytBuf, p, 3)       '' �l�C��
            End With ' OddsUmatanInfo
        Next i
        .TotalHyosuUmatan = IncMidByte(bytBuf, p, 11)   '' �n�P�[�����v
        .CRLF = IncMidByte(bytBuf, p, 2)                '' ���R�[�h��؂�
    End With

    End Sub


'
'   �@�\: �\���̂Ƀf�[�^���Z�b�g����|�I�b�Y�i�R�A���j
'
'   ���l: �Ȃ�
'
    Public Sub SetDataFromByte_O5(bytBuf() As Byte, ByRef mBuf As JV_O5_ODDS_SANREN)
    Dim i As Integer                                '' ���[�v�J�E���^
    Dim j As Integer                                '' ���[�v�J�E���^
    Dim k As Integer                                '' ���[�v�J�E���^
    Dim p As Long                                   '' �؂蕪���J�n�ʒu
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMidByte(bytBuf, p, 2)      '' ���R�[�h���
            .DataKubun = IncMidByte(bytBuf, p, 1)       '' �f�[�^�敪
            With .MakeDate
                .Year = IncMidByte(bytBuf, p, 4)        '' �N
                .Month = IncMidByte(bytBuf, p, 2)       '' ��
                .Day = IncMidByte(bytBuf, p, 2)         '' ��
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMidByte(bytBuf, p, 4)            '' �J�ÔN
            .MonthDay = IncMidByte(bytBuf, p, 4)        '' �J�Ì���
            .JyoCD = IncMidByte(bytBuf, p, 2)           '' ���n��R�[�h
            .Kaiji = IncMidByte(bytBuf, p, 2)               '' �J�É�[��N��]
            .Nichiji = IncMidByte(bytBuf, p, 2)         '' �J�Ó���[N����]
            .RaceNum = IncMidByte(bytBuf, p, 2)         '' ���[�X�ԍ�
        End With ' id
        With .HappyoTime
            .Month = IncMidByte(bytBuf, p, 2)           '' ��
            .Day = IncMidByte(bytBuf, p, 2)             '' ��
            .Hour = IncMidByte(bytBuf, p, 2)            '' ��
            .Minute = IncMidByte(bytBuf, p, 2)          '' ��
        End With ' HappyoTime
        .TorokuTosu = IncMidByte(bytBuf, p, 2)          '' �o�^����
        .SyussoTosu = IncMidByte(bytBuf, p, 2)          '' �o������
        .SanrenpukuFlag = IncMidByte(bytBuf, p, 1)      '' �����t���O�@3�A��
        For i = 0 To 815
            With .OddsSanrenInfo(i)
                .Kumi = IncMidByte(bytBuf, p, 6)        '' �g��
                .Odds = IncMidByte(bytBuf, p, 6)        '' �I�b�Y
                .Ninki = IncMidByte(bytBuf, p, 3)       '' �l�C��
            End With ' OddsSanrenInfo
        Next i
        .TotalHyosuSanrenpuku = IncMidByte(bytBuf, p, 11)       '' 3�A���[�����v
        .CRLF = IncMidByte(bytBuf, p, 2)        '' ���R�[�h��؂�
    End With
   
    End Sub


'
'   �@�\: �\���̂Ƀf�[�^���Z�b�g����|�I�b�Y�i�R�A�P�j
'
'   ���l: �Ȃ�
'
    Public Sub SetDataFromByte_O6(bytBuf() As Byte, ByRef mBuf As JV_O6_ODDS_SANRENTAN)
    Dim i As Integer                                '' ���[�v�J�E���^�[
    Dim p As Long                                   '' �؂蕪���J�n�ʒu
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMidByte(bytBuf, p, 2)        '' ���R�[�h���
            .DataKubun = IncMidByte(bytBuf, p, 1)         '' �f�[�^�敪
            With .MakeDate
                .Year = IncMidByte(bytBuf, p, 4)          '' �N
                .Month = IncMidByte(bytBuf, p, 2)         '' ��
                .Day = IncMidByte(bytBuf, p, 2)           '' ��
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMidByte(bytBuf, p, 4)              '' �J�ÔN
            .MonthDay = IncMidByte(bytBuf, p, 4)          '' �J�Ì���
            .JyoCD = IncMidByte(bytBuf, p, 2)             '' ���n��R�[�h
            .Kaiji = IncMidByte(bytBuf, p, 2)             '' �J�É�[��N��]
            .Nichiji = IncMidByte(bytBuf, p, 2)           '' �J�Ó���[N����]
            .RaceNum = IncMidByte(bytBuf, p, 2)           '' ���[�X�ԍ�
        End With ' id
        With .HappyoTime
            .Month = IncMidByte(bytBuf, p, 2)             '' ��
            .Day = IncMidByte(bytBuf, p, 2)               '' ��
            .Hour = IncMidByte(bytBuf, p, 2)              '' ��
            .Minute = IncMidByte(bytBuf, p, 2)            '' ��
        End With ' HappyoTime
        .TorokuTosu = IncMidByte(bytBuf, p, 2)            '' �o�^����
        .SyussoTosu = IncMidByte(bytBuf, p, 2)            '' �o������
        .SanrentanFlag = IncMidByte(bytBuf, p, 1)         '' �����t���O�@3�A�P
        For i = 0 To 4895
            With .OddsSanrentanInfo(i)
                .Kumi = IncMidByte(bytBuf, p, 6)          '' �g��
                .Odds = IncMidByte(bytBuf, p, 7)          '' �I�b�Y
                .Ninki = IncMidByte(bytBuf, p, 4)         '' �l�C��
            End With
        Next i
        .TotalHyosuSanrentan = IncMidByte(bytBuf, p, 11)  '' 3�A�P�[�����v
        .CRLF = IncMidByte(bytBuf, p, 2)                  '' ���R�[�h��؂�
    End With
    
    End Sub

    
'
'   �@�\: �\���̂Ƀf�[�^���Z�b�g����|�����n�}�X�^
'
'   ���l: �Ȃ�
'
    Public Sub SetDataFromByte_UM(bytBuf() As Byte, ByRef mBuf As JV_UM_UMA)
    Dim i As Integer                                '' ���[�v�J�E���^
    Dim j As Integer                                '' ���[�v�J�E���^
    Dim k As Integer                                '' ���[�v�J�E���^
    Dim p As Long                                   '' �؂蕪���J�n�ʒu
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMidByte(bytBuf, p, 2)      '' ���R�[�h���
            .DataKubun = IncMidByte(bytBuf, p, 1)       '' �f�[�^�敪
            With .MakeDate
                .Year = IncMidByte(bytBuf, p, 4)        '' �N
                .Month = IncMidByte(bytBuf, p, 2)       '' ��
                .Day = IncMidByte(bytBuf, p, 2)         '' ��
            End With ' MakeDate
        End With ' head
        .KettoNum = IncMidByte(bytBuf, p, 10)           '' �����o�^�ԍ�
        .DelKubun = IncMidByte(bytBuf, p, 1)            '' �����n�����敪
        With .RegDate
            .Year = IncMidByte(bytBuf, p, 4)            '' �N
            .Month = IncMidByte(bytBuf, p, 2)           '' ��
            .Day = IncMidByte(bytBuf, p, 2)             '' ��
        End With ' RegDate
        With .DelDate
            .Year = IncMidByte(bytBuf, p, 4)            '' �N
            .Month = IncMidByte(bytBuf, p, 2)           '' ��
            .Day = IncMidByte(bytBuf, p, 2)             '' ��
        End With ' DelDate
        With .BirthDate
            .Year = IncMidByte(bytBuf, p, 4)            '' �N
            .Month = IncMidByte(bytBuf, p, 2)           '' ��
            .Day = IncMidByte(bytBuf, p, 2)             '' ��
        End With ' BirthDate
        .BAMEI = IncMidByte(bytBuf, p, 36)              '' �n��
        .BameiKana = IncMidByte(bytBuf, p, 36)          '' �n�����p�J�i
        .BameiEng = IncMidByte(bytBuf, p, 80)           '' �n������
        .UmaKigoCD = IncMidByte(bytBuf, p, 2)           '' �n�L���R�[�h
        .SexCD = IncMidByte(bytBuf, p, 1)               '' ���ʃR�[�h
        .HinsyuCD = IncMidByte(bytBuf, p, 1)            '' �i��R�[�h
        .KeiroCD = IncMidByte(bytBuf, p, 2)             '' �ѐF�R�[�h
        For i = 0 To 13
            With .Ketto3Info(i)
                .HansyokuNum = IncMidByte(bytBuf, p, 8) '' �ɐB�o�^�ԍ�
                .BAMEI = IncMidByte(bytBuf, p, 36)      '' �n��
            End With ' Ketto3Info
        Next i
        .TozaiCD = IncMidByte(bytBuf, p, 1)             '' ���������R�[�h
        .ChokyosiCode = IncMidByte(bytBuf, p, 5)        '' �����t�R�[�h
        .ChokyosiRyakusyo = IncMidByte(bytBuf, p, 8)    '' �����t������
        .Syotai = IncMidByte(bytBuf, p, 20)             '' ���Ғn�於
        .BreederCode = IncMidByte(bytBuf, p, 6)         '' ���Y�҃R�[�h
        .BreederName = IncMidByte(bytBuf, p, 70)        '' ���Y�Җ�
        .SanchiName = IncMidByte(bytBuf, p, 20)         '' �Y�n��
        .BanusiCode = IncMidByte(bytBuf, p, 6)          '' �n��R�[�h
        .BanusiName = IncMidByte(bytBuf, p, 64)         '' �n�喼
        .RuikeiHonsyoHeiti = IncMidByte(bytBuf, p, 9)   '' ���n�{�܋��݌v
        .RuikeiHonsyoSyogai = IncMidByte(bytBuf, p, 9)  '' ��Q�{�܋��݌v
        .RuikeiFukaHeichi = IncMidByte(bytBuf, p, 9)    '' ���n�t���܋��݌v
        .RuikeiFukaSyogai = IncMidByte(bytBuf, p, 9)    '' ��Q�t���܋��݌v
        .RuikeiSyutokuHeichi = IncMidByte(bytBuf, p, 9) '' ���n�����܋��݌v
        .RuikeiSyutokuSyogai = IncMidByte(bytBuf, p, 9) '' ��Q�����܋��݌v
        With .ChakuSogo
            For j = 0 To 5
                .Chakukaisu(j) = IncMidByte(bytBuf, p, 3)
            Next j
        End With ' ChakuSogo
        With .ChakuChuo
            For j = 0 To 5
                .Chakukaisu(j) = IncMidByte(bytBuf, p, 3)
            Next j
        End With ' ChakuChuo
        For i = 0 To 6
            With .ChakuKaisuBa(i)
                For j = 0 To 5
                    .Chakukaisu(j) = IncMidByte(bytBuf, p, 3)
                Next j
            End With ' ChakuKaisuBa
        Next i
        For i = 0 To 11
            With .ChakuKaisuJyotai(i)
                For j = 0 To 5
                    .Chakukaisu(j) = IncMidByte(bytBuf, p, 3)
                Next j
            End With ' ChakuKaisuJyotai
        Next i
        For i = 0 To 5
            With .ChakuKaisuKyori(i)
                For j = 0 To 5
                    .Chakukaisu(j) = IncMidByte(bytBuf, p, 3)
                Next j
            End With ' ChakuKaisuKyoriu
        Next i
        For i = 0 To 3
            .Kyakusitu(i) = IncMidByte(bytBuf, p, 3)    '' �r���X��
        Next i
        .RaceCount = IncMidByte(bytBuf, p, 3)           '' �o�^���[�X��
        .CRLF = IncMidByte(bytBuf, p, 2)                '' ���R�[�h��؂�
    End With

    End Sub


'
'   �@�\: �\���̂Ƀf�[�^���Z�b�g����|�R��}�X�^
'
'   ���l: �Ȃ�
'
    Public Sub SetDataFromByte_KS(bytBuf() As Byte, ByRef mBuf As JV_KS_KISYU)
    Dim i As Integer                                '' ���[�v�J�E���^
    Dim j As Integer                                '' ���[�v�J�E���^
    Dim k As Integer                                '' ���[�v�J�E���^
    Dim p As Long                                   '' �؂蕪���J�n�ʒu
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMidByte(bytBuf, p, 2)      '' ���R�[�h���
            .DataKubun = IncMidByte(bytBuf, p, 1)       '' �f�[�^�敪
            With .MakeDate
                .Year = IncMidByte(bytBuf, p, 4)        '' �N
                .Month = IncMidByte(bytBuf, p, 2)       '' ��
                .Day = IncMidByte(bytBuf, p, 2) '' ��
            End With ' MakeDate
        End With ' head
        .KisyuCode = IncMidByte(bytBuf, p, 5)   '' �R��R�[�h
        .DelKubun = IncMidByte(bytBuf, p, 1)    '' �R�薕���敪
        With .IssueDate
            .Year = IncMidByte(bytBuf, p, 4)    '' �N
            .Month = IncMidByte(bytBuf, p, 2)   '' ��
            .Day = IncMidByte(bytBuf, p, 2)     '' ��
        End With ' IssueDate
        With .DelDate
            .Year = IncMidByte(bytBuf, p, 4)    '' �N
            .Month = IncMidByte(bytBuf, p, 2)   '' ��
            .Day = IncMidByte(bytBuf, p, 2)     '' ��
        End With ' DelDate
        With .BirthDate
            .Year = IncMidByte(bytBuf, p, 4)    '' �N
            .Month = IncMidByte(bytBuf, p, 2)   '' ��
            .Day = IncMidByte(bytBuf, p, 2)     '' ��
        End With ' BirthDate
        .KisyuName = IncMidByte(bytBuf, p, 34)  '' �R�薼����
        .Reserved = IncMidByte(bytBuf, p, 34)   '' �\��
        .KisyuNameKana = IncMidByte(bytBuf, p, 30)      '' �R�薼���p�J�i
        .KisyuRyakusyo = IncMidByte(bytBuf, p, 8)       '' �R�薼����
        .KisyuNameEng = IncMidByte(bytBuf, p, 80)       '' �R�薼����
        .SexCD = IncMidByte(bytBuf, p, 1)       '' ���ʋ敪
        .SikakuCD = IncMidByte(bytBuf, p, 1)    '' �R�掑�i�R�[�h
        .MinaraiCD = IncMidByte(bytBuf, p, 1)   '' �R�茩�K�R�[�h
        .TozaiCD = IncMidByte(bytBuf, p, 1)     '' �R�蓌�������R�[�h
        .Syotai = IncMidByte(bytBuf, p, 20)     '' ���Ғn�於
        .ChokyosiCode = IncMidByte(bytBuf, p, 5)        '' ���������t�R�[�h
        .ChokyosiRyakusyo = IncMidByte(bytBuf, p, 8)    '' ���������t������
        For i = 0 To 1
            With .HatuKiJyo(i)
                With .Hatukijyoid
                    .Year = IncMidByte(bytBuf, p, 4)    '' �J�ÔN
                    .MonthDay = IncMidByte(bytBuf, p, 4)        '' �J�Ì���
                    .JyoCD = IncMidByte(bytBuf, p, 2)   '' ���n��R�[�h
                    .Kaiji = IncMidByte(bytBuf, p, 2)   '' �J�É�[��N��]
                    .Nichiji = IncMidByte(bytBuf, p, 2) '' �J�Ó���[N����]
                    .RaceNum = IncMidByte(bytBuf, p, 2) '' ���[�X�ԍ�
                End With ' Hatukijyoid
                .SyussoTosu = IncMidByte(bytBuf, p, 2)  '' �o������
                .KettoNum = IncMidByte(bytBuf, p, 10)   '' �����o�^�ԍ�
                .BAMEI = IncMidByte(bytBuf, p, 36)      '' �n��
                .KakuteiJyuni = IncMidByte(bytBuf, p, 2)        '' �m�蒅��
                .IJyoCD = IncMidByte(bytBuf, p, 1)      '' �ُ�敪�R�[�h
            End With ' HatuKiJyo
        Next i
        For i = 0 To 1
            With .HatuSyori(i)
                With .Hatusyoriid
                    .Year = IncMidByte(bytBuf, p, 4)    '' �J�ÔN
                    .MonthDay = IncMidByte(bytBuf, p, 4)        '' �J�Ì���
                    .JyoCD = IncMidByte(bytBuf, p, 2)   '' ���n��R�[�h
                    .Kaiji = IncMidByte(bytBuf, p, 2)   '' �J�É�[��N��]
                    .Nichiji = IncMidByte(bytBuf, p, 2) '' �J�Ó���[N����]
                    .RaceNum = IncMidByte(bytBuf, p, 2) '' ���[�X�ԍ�
                End With ' Hatusyoriid
                .SyussoTosu = IncMidByte(bytBuf, p, 2)  '' �o������
                .KettoNum = IncMidByte(bytBuf, p, 10)   '' �����o�^�ԍ�
                .BAMEI = IncMidByte(bytBuf, p, 36)      '' �n��
            End With ' HatuSyori
        Next i
        For i = 0 To 2
            With .SaikinJyusyo(i)
                With .SaikinJyusyoid
                    .Year = IncMidByte(bytBuf, p, 4)    '' �J�ÔN
                    .MonthDay = IncMidByte(bytBuf, p, 4)        '' �J�Ì���
                    .JyoCD = IncMidByte(bytBuf, p, 2)   '' ���n��R�[�h
                    .Kaiji = IncMidByte(bytBuf, p, 2)   '' �J�É�[��N��]
                    .Nichiji = IncMidByte(bytBuf, p, 2) '' �J�Ó���[N����]
                    .RaceNum = IncMidByte(bytBuf, p, 2) '' ���[�X�ԍ�
                End With ' SaikinJyusyoid
                .Hondai = IncMidByte(bytBuf, p, 60)     '' �������{��
                .Ryakusyo10 = IncMidByte(bytBuf, p, 20) '' ����������10��
                .Ryakusyo6 = IncMidByte(bytBuf, p, 12)  '' ����������6��
                .Ryakusyo3 = IncMidByte(bytBuf, p, 6)   '' ����������3��
                .GradeCD = IncMidByte(bytBuf, p, 1)     '' �O���[�h�R�[�h
                .SyussoTosu = IncMidByte(bytBuf, p, 2)  '' �o������
                .KettoNum = IncMidByte(bytBuf, p, 10)   '' �����o�^�ԍ�
                .BAMEI = IncMidByte(bytBuf, p, 36)      '' �n��
            End With ' SaikinJyusyo
        Next i
        For i = 0 To 2
            With .HonZenRuikei(i)
                .SetYear = IncMidByte(bytBuf, p, 4)     '' �ݒ�N
                .HonSyokinHeichi = IncMidByte(bytBuf, p, 10)    '' ���n�{�܋����v
                .HonSyokinSyogai = IncMidByte(bytBuf, p, 10)    '' ��Q�{�܋����v
                .FukaSyokinHeichi = IncMidByte(bytBuf, p, 10)   '' ���n�t���܋����v
                .FukaSyokinSyogai = IncMidByte(bytBuf, p, 10)   '' ��Q�t���܋����v
                With .ChakuKaisuHeichi
                    For k = 0 To 5
                        .Chakukaisu(k) = IncMidByte(bytBuf, p, 6)
                    Next k
                End With ' ChakuKaisuHeichi
                With .ChakuKaisuSyogai
                    For k = 0 To 5
                        .Chakukaisu(k) = IncMidByte(bytBuf, p, 6)
                    Next k
                End With ' ChakuKaisuSyogai
                For j = 0 To 19
                    With .ChakuKaisuJyo(j)
                        For k = 0 To 5
                            .Chakukaisu(k) = IncMidByte(bytBuf, p, 6)
                        Next k
                    End With ' ChakuKaisuJyo
                Next j
                For j = 0 To 5
                    With .ChakuKaisuKyori(j)
                        For k = 0 To 5
                            .Chakukaisu(k) = IncMidByte(bytBuf, p, 6)
                        Next k
                    End With ' ChakuKaisuKyori
                Next j
            End With ' HonZenRuikei
        Next i
        .CRLF = IncMidByte(bytBuf, p, 2)        '' ���R�[�h��؂�
    End With

    End Sub


'
'   �@�\: �\���̂Ƀf�[�^���Z�b�g����|�����t�}�X�^
'
'   ���l: �Ȃ�
'
    Public Sub SetDataFromByte_CH(bytBuf() As Byte, ByRef mBuf As JV_CH_CHOKYOSI)
    Dim i As Integer                                '' ���[�v�J�E���^
    Dim j As Integer                                '' ���[�v�J�E���^
    Dim k As Integer                                '' ���[�v�J�E���^
    Dim p As Long                                   '' �؂蕪���J�n�ʒu
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMidByte(bytBuf, p, 2)              '' ���R�[�h���
            .DataKubun = IncMidByte(bytBuf, p, 1)               '' �f�[�^�敪
            With .MakeDate
                .Year = IncMidByte(bytBuf, p, 4)                '' �N
                .Month = IncMidByte(bytBuf, p, 2)               '' ��
                .Day = IncMidByte(bytBuf, p, 2)                 '' ��
            End With ' MakeDate
        End With ' head
        .ChokyosiCode = IncMidByte(bytBuf, p, 5)                '' �����t�R�[�h
        .DelKubun = IncMidByte(bytBuf, p, 1)                    '' �����t�����敪
        With .IssueDate
            .Year = IncMidByte(bytBuf, p, 4)                    '' �N
            .Month = IncMidByte(bytBuf, p, 2)                   '' ��
            .Day = IncMidByte(bytBuf, p, 2)                     '' ��
        End With ' IssueDate
        With .DelDate
            .Year = IncMidByte(bytBuf, p, 4)                    '' �N
            .Month = IncMidByte(bytBuf, p, 2)                   '' ��
            .Day = IncMidByte(bytBuf, p, 2)                     '' ��
        End With ' DelDate
        With .BirthDate
            .Year = IncMidByte(bytBuf, p, 4)                    '' �N
            .Month = IncMidByte(bytBuf, p, 2)                   '' ��
            .Day = IncMidByte(bytBuf, p, 2)                     '' ��
        End With ' BirthDate
        .ChokyosiName = IncMidByte(bytBuf, p, 34)               '' �����t������
        .ChokyosiNameKana = IncMidByte(bytBuf, p, 30)           '' �����t�����p�J�i
        .ChokyosiRyakusyo = IncMidByte(bytBuf, p, 8)            '' �����t������
        .ChokyosiNameEng = IncMidByte(bytBuf, p, 80)            '' �����t������
        .SexCD = IncMidByte(bytBuf, p, 1)                       '' ���ʋ敪
        .TozaiCD = IncMidByte(bytBuf, p, 1)                     '' �����t���������R�[�h
        .Syotai = IncMidByte(bytBuf, p, 20)                     '' ���Ғn�於
        For i = 0 To 2
            With .SaikinJyusyo(i)
                With .SaikinJyusyoid
                    .Year = IncMidByte(bytBuf, p, 4)            '' �J�ÔN
                    .MonthDay = IncMidByte(bytBuf, p, 4)        '' �J�Ì���
                    .JyoCD = IncMidByte(bytBuf, p, 2)           '' ���n��R�[�h
                    .Kaiji = IncMidByte(bytBuf, p, 2)           '' �J�É�[��N��]
                    .Nichiji = IncMidByte(bytBuf, p, 2)         '' �J�Ó���[N����]
                    .RaceNum = IncMidByte(bytBuf, p, 2)         '' ���[�X�ԍ�
                End With ' SaikinJyusyoid
                .Hondai = IncMidByte(bytBuf, p, 60)             '' �������{��
                .Ryakusyo10 = IncMidByte(bytBuf, p, 20)         '' ����������10��
                .Ryakusyo6 = IncMidByte(bytBuf, p, 12)          '' ����������6��
                .Ryakusyo3 = IncMidByte(bytBuf, p, 6)           '' ����������3��
                .GradeCD = IncMidByte(bytBuf, p, 1)             '' �O���[�h�R�[�h
                .SyussoTosu = IncMidByte(bytBuf, p, 2)          '' �o������
                .KettoNum = IncMidByte(bytBuf, p, 10)           '' �����o�^�ԍ�
                .BAMEI = IncMidByte(bytBuf, p, 36)              '' �n��
            End With ' SaikinJyusyo
        Next i
        For i = 0 To 2
            With .HonZenRuikei(i)
                .SetYear = IncMidByte(bytBuf, p, 4)             '' �ݒ�N
                .HonSyokinHeichi = IncMidByte(bytBuf, p, 10)    '' ���n�{�܋����v
                .HonSyokinSyogai = IncMidByte(bytBuf, p, 10)    '' ��Q�{�܋����v
                .FukaSyokinHeichi = IncMidByte(bytBuf, p, 10)   '' ���n�t���܋����v
                .FukaSyokinSyogai = IncMidByte(bytBuf, p, 10)   '' ��Q�t���܋����v
                With .ChakuKaisuHeichi
                    For k = 0 To 5
                        .Chakukaisu(k) = IncMidByte(bytBuf, p, 6)
                    Next k
                End With ' ChakuKaisuHeichi
                With .ChakuKaisuSyogai
                    For k = 0 To 5
                        .Chakukaisu(k) = IncMidByte(bytBuf, p, 6)
                    Next k
                End With ' ChakuKaisuSyogai
                For j = 0 To 19
                    With .ChakuKaisuJyo(j)
                        For k = 0 To 5
                            .Chakukaisu(k) = IncMidByte(bytBuf, p, 6)
                        Next k
                    End With ' ChakuKaisuJyo
                Next j
                For j = 0 To 5
                    With .ChakuKaisuKyori(j)
                        For k = 0 To 5
                            .Chakukaisu(k) = IncMidByte(bytBuf, p, 6)
                        Next k
                    End With ' ChakuKaisuKyori
                Next j
            End With ' HonZenRuikei
        Next i
        .CRLF = IncMidByte(bytBuf, p, 2)        '' ���R�[�h��؂�
    End With

    End Sub


'
'   �@�\: �\���̂Ƀf�[�^���Z�b�g����|���Y�҃}�X�^
'
'   ���l: �Ȃ�
'
    Public Sub SetDataFromByte_BR(bytBuf() As Byte, ByRef mBuf As JV_BR_BREEDER)
    Dim i As Integer                                '' ���[�v�J�E���^
    Dim j As Integer                                '' ���[�v�J�E���^
    Dim k As Integer                                '' ���[�v�J�E���^
    Dim p As Long                                   '' �؂蕪���J�n�ʒu
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMidByte(bytBuf, p, 2)              '' ���R�[�h���
            .DataKubun = IncMidByte(bytBuf, p, 1)               '' �f�[�^�敪
            With .MakeDate
                .Year = IncMidByte(bytBuf, p, 4)                '' �N
                .Month = IncMidByte(bytBuf, p, 2)               '' ��
                .Day = IncMidByte(bytBuf, p, 2)                 '' ��
            End With ' MakeDate
        End With ' head
        .BreederCode = IncMidByte(bytBuf, p, 6)                 '' ���Y�҃R�[�h
        .BreederName_Co = IncMidByte(bytBuf, p, 70)             '' ���Y�Җ�(�@�l�i�L�j
        .BreederName = IncMidByte(bytBuf, p, 70)                '' ���Y�Җ�(�@�l�i���j
        .BreederNameKana = IncMidByte(bytBuf, p, 70)            '' ���Y�Җ����p�J�i
        .BreederNameEng = IncMidByte(bytBuf, p, 168)            '' ���Y�Җ�����
        .Address = IncMidByte(bytBuf, p, 20)                    '' ���Y�ҏZ�������Ȗ�
        For i = 0 To 1
            With .HonRuikei(i)
                .SetYear = IncMidByte(bytBuf, p, 4)             '' �ݒ�N
                .HonSyokinTotal = IncMidByte(bytBuf, p, 10)     '' �{�܋����v
                .Fukasyokin = IncMidByte(bytBuf, p, 10)         '' �t���܋����v
                For j = 0 To 5
                    .Chakukaisu(j) = IncMidByte(bytBuf, p, 6)   '' ����
                Next j
            End With ' HonRuikei
        Next i
        .CRLF = IncMidByte(bytBuf, p, 2)                        '' ���R�[�h��؂�
    End With

    End Sub


'
'   �@�\: �\���̂Ƀf�[�^���Z�b�g����|�n��}�X�^
'
'   ���l: �Ȃ�
'
    Public Sub SetDataFromByte_BN(bytBuf() As Byte, ByRef mBuf As JV_BN_BANUSI)
    Dim i As Integer                                '' ���[�v�J�E���^
    Dim j As Integer                                '' ���[�v�J�E���^
    Dim k As Integer                                '' ���[�v�J�E���^
    Dim p As Long                                   '' �؂蕪���J�n�ʒu
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMidByte(bytBuf, p, 2)              '' ���R�[�h���
            .DataKubun = IncMidByte(bytBuf, p, 1)               '' �f�[�^�敪
            With .MakeDate
                .Year = IncMidByte(bytBuf, p, 4)                '' �N
                .Month = IncMidByte(bytBuf, p, 2)               '' ��
                .Day = IncMidByte(bytBuf, p, 2)                 '' ��
            End With ' MakeDate
        End With ' head
        .BanusiCode = IncMidByte(bytBuf, p, 6)                  '' �n��R�[�h
        .BanusiName_Co = IncMidByte(bytBuf, p, 64)              '' �n�喼�i�@�l�i�L�j
        .BanusiName = IncMidByte(bytBuf, p, 64)                 '' �n�喼�i�@�l�i���j
        .BanusiNameKana = IncMidByte(bytBuf, p, 50)             '' �n�喼���p�J�i
        .BanusiNameEng = IncMidByte(bytBuf, p, 100)             '' �n�喼����
        .Fukusyoku = IncMidByte(bytBuf, p, 60)                  '' ���F�W��
        For i = 0 To 1
            With .HonRuikei(i)
                .SetYear = IncMidByte(bytBuf, p, 4)             '' �ݒ�N
                .HonSyokinTotal = IncMidByte(bytBuf, p, 10)     '' �{�܋����v
                .Fukasyokin = IncMidByte(bytBuf, p, 10)         '' �t���܋����v
                For j = 0 To 5
                    .Chakukaisu(j) = IncMidByte(bytBuf, p, 6)   '' ����
                Next j
            End With ' HonRuikei
        Next i
        .CRLF = IncMidByte(bytBuf, p, 2)                        '' ���R�[�h��؂�
    End With

    End Sub

'
'   �@�\: �\���̂Ƀf�[�^���Z�b�g����|�ɐB�n�}�X�^
'
'   ���l: �Ȃ�
'
    Public Sub SetDataFromByte_HN(bytBuf() As Byte, ByRef mBuf As JV_HN_HANSYOKU)
    Dim i As Integer                                '' ���[�v�J�E���^
    Dim j As Integer                                '' ���[�v�J�E���^
    Dim k As Integer                                '' ���[�v�J�E���^
    Dim p As Long                                   '' �؂蕪���J�n�ʒu
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMidByte(bytBuf, p, 2)      '' ���R�[�h���
            .DataKubun = IncMidByte(bytBuf, p, 1)       '' �f�[�^�敪
            With .MakeDate
                .Year = IncMidByte(bytBuf, p, 4)        '' �N
                .Month = IncMidByte(bytBuf, p, 2)       '' ��
                .Day = IncMidByte(bytBuf, p, 2)         '' ��
            End With ' MakeDate
        End With ' head
        .HansyokuNum = IncMidByte(bytBuf, p, 8)         '' �ɐB�o�^�ԍ�
        .Reserved = IncMidByte(bytBuf, p, 8)            '' �\��
        .KettoNum = IncMidByte(bytBuf, p, 10)           '' �����o�^�ԍ�
        .DelKubun = IncMidByte(bytBuf, p, 1)            '' �ɐB�n�����敪
        .BAMEI = IncMidByte(bytBuf, p, 36)              '' �n��
        .BameiKana = IncMidByte(bytBuf, p, 40)          '' �n�����p�J�i
        .BameiEng = IncMidByte(bytBuf, p, 80)           '' �n������
        .BirthYear = IncMidByte(bytBuf, p, 4)           '' ���N
        .SexCD = IncMidByte(bytBuf, p, 1)               '' ���ʃR�[�h
        .HinsyuCD = IncMidByte(bytBuf, p, 1)            '' �i��R�[�h
        .KeiroCD = IncMidByte(bytBuf, p, 2)             '' �ѐF�R�[�h
        .HansyokuMochiKubun = IncMidByte(bytBuf, p, 1)  '' �ɐB�n�����敪
        .ImportYear = IncMidByte(bytBuf, p, 4)          '' �A���N
        .SanchiName = IncMidByte(bytBuf, p, 20)         '' �Y�n��
        .HansyokuFNum = IncMidByte(bytBuf, p, 8)        '' ���n�ɐB�o�^�ԍ�
        .HansyokuMNum = IncMidByte(bytBuf, p, 8)        '' ��n�ɐB�o�^�ԍ�
        .CRLF = IncMidByte(bytBuf, p, 2)                '' ���R�[�h��؂�
    End With

    End Sub


'
'   �@�\: �\���̂Ƀf�[�^���Z�b�g����|�Y��}�X�^
'
'   ���l: �Ȃ�
'
    Public Sub SetDataFromByte_SK(bytBuf() As Byte, ByRef mBuf As JV_SK_SANKU)
    Dim i As Integer                                '' ���[�v�J�E���^
    Dim j As Integer                                '' ���[�v�J�E���^
    Dim k As Integer                                '' ���[�v�J�E���^
    Dim p As Long                                   '' �؂蕪���J�n�ʒu
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMidByte(bytBuf, p, 2)      '' ���R�[�h���
            .DataKubun = IncMidByte(bytBuf, p, 1)       '' �f�[�^�敪
            With .MakeDate
                .Year = IncMidByte(bytBuf, p, 4)        '' �N
                .Month = IncMidByte(bytBuf, p, 2)       '' ��
                .Day = IncMidByte(bytBuf, p, 2)         '' ��
            End With ' MakeDate
        End With ' head
        .KettoNum = IncMidByte(bytBuf, p, 10)           '' �����o�^�ԍ�
        With .BirthDate
            .Year = IncMidByte(bytBuf, p, 4)            '' �N
            .Month = IncMidByte(bytBuf, p, 2)           '' ��
            .Day = IncMidByte(bytBuf, p, 2)             '' ��
        End With ' BirthDate
        .SexCD = IncMidByte(bytBuf, p, 1)               '' ���ʃR�[�h
        .HinsyuCD = IncMidByte(bytBuf, p, 1)            '' �i��R�[�h
        .KeiroCD = IncMidByte(bytBuf, p, 2)             '' �ѐF�R�[�h
        .SankuMochiKubun = IncMidByte(bytBuf, p, 1)     '' �Y����敪
        .ImportYear = IncMidByte(bytBuf, p, 4)          '' �A���N
        .BreederCode = IncMidByte(bytBuf, p, 6)         '' ���Y�҃R�[�h
        .SanchiName = IncMidByte(bytBuf, p, 20)         '' �Y�n��
        For i = 0 To 13
            .HansyokuNum(i) = IncMidByte(bytBuf, p, 8)  '' 3�㌌��
        Next i
    End With

    End Sub

'
'   �@�\: �\���̂Ƀf�[�^���Z�b�g����|���R�[�h�}�X�^
'
'   ���l: �Ȃ�
'
    Public Sub SetDataFromByte_RC(bytBuf() As Byte, ByRef mBuf As JV_RC_RECORD)
    Dim i As Integer                                '' ���[�v�J�E���^
    Dim j As Integer                                '' ���[�v�J�E���^
    Dim k As Integer                                '' ���[�v�J�E���^
    Dim p As Long                                   '' �؂蕪���J�n�ʒu
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMidByte(bytBuf, p, 2)              '' ���R�[�h���
            .DataKubun = IncMidByte(bytBuf, p, 1)               '' �f�[�^�敪
            With .MakeDate
                .Year = IncMidByte(bytBuf, p, 4)                '' �N
                .Month = IncMidByte(bytBuf, p, 2)               '' ��
                .Day = IncMidByte(bytBuf, p, 2)                 '' ��
            End With ' MakeDate
        End With ' head
        .RecInfoKubun = IncMidByte(bytBuf, p, 1)                '' ���R�[�h���ʋ敪
        With .id
            .Year = IncMidByte(bytBuf, p, 4)                    '' �J�ÔN
            .MonthDay = IncMidByte(bytBuf, p, 4)                '' �J�Ì���
            .JyoCD = IncMidByte(bytBuf, p, 2)                   '' ���n��R�[�h
            .Kaiji = IncMidByte(bytBuf, p, 2)                   '' �J�É�[��N��]
            .Nichiji = IncMidByte(bytBuf, p, 2)                 '' �J�Ó���[N����]
            .RaceNum = IncMidByte(bytBuf, p, 2)                 '' ���[�X�ԍ�
        End With ' id
        .TokuNum = IncMidByte(bytBuf, p, 4)                     '' ���ʋ����ԍ�
        .Hondai = IncMidByte(bytBuf, p, 60)                     '' �������{��
        .GradeCD = IncMidByte(bytBuf, p, 1)                     '' �O���[�h�R�[�h
        .SyubetuCD = IncMidByte(bytBuf, p, 2)                   '' ������ʃR�[�h
        .KYORI = IncMidByte(bytBuf, p, 4)                       '' ����
        .TrackCD = IncMidByte(bytBuf, p, 2)                     '' �g���b�N�R�[�h
        .RecKubun = IncMidByte(bytBuf, p, 1)                    '' ���R�[�h�敪
        .RecTime = IncMidByte(bytBuf, p, 4)                     '' ���R�[�h�^�C��
        With .TenkoBaba
            .TenkoCD = IncMidByte(bytBuf, p, 1)                 '' �V��R�[�h
            .SibaBabaCD = IncMidByte(bytBuf, p, 1)              '' �Ŕn���ԃR�[�h
            .DirtBabaCD = IncMidByte(bytBuf, p, 1)              '' �_�[�g�n���ԃR�[�h
        End With ' TenkoBaba
        For i = 0 To 2
            With .RecUmaInfo(i)
                .KettoNum = IncMidByte(bytBuf, p, 10)           '' �����o�^�ԍ�
                .BAMEI = IncMidByte(bytBuf, p, 36)              '' �n��
                .UmaKigoCD = IncMidByte(bytBuf, p, 2)           '' �n�L���R�[�h
                .SexCD = IncMidByte(bytBuf, p, 1)               '' ���ʃR�[�h
                .ChokyosiCode = IncMidByte(bytBuf, p, 5)        '' �����t�R�[�h
                .ChokyosiName = IncMidByte(bytBuf, p, 34)       '' �����t��
                .Futan = IncMidByte(bytBuf, p, 3)               '' ���S�d��
                .KisyuCode = IncMidByte(bytBuf, p, 5)           '' �R��R�[�h
                .KisyuName = IncMidByte(bytBuf, p, 34)          '' �R�薼
            End With ' RecUmaInfo
        Next i
        .CRLF = IncMidByte(bytBuf, p, 2)                        '' ���R�[�h��؂�
    End With

    End Sub


'
'   �@�\: �\���̂Ƀf�[�^���Z�b�g����|��H����
'
'   ���l: �Ȃ�
'
    Public Sub SetDataFromByte_HC(lBuf() As Byte, ByRef mBuf As JV_HC_HANRO)
    Dim i As Integer                                '' ���[�v�J�E���^
    Dim j As Integer                                '' ���[�v�J�E���^
    Dim k As Integer                                '' ���[�v�J�E���^
    Dim p As Long                                   '' �؂蕪���J�n�ʒu
    
    With mBuf
        With .head
            .RecordSpec = Mid$(lBuf, 1, 2)  '' ���R�[�h���
            .DataKubun = Mid$(lBuf, 3, 1)   '' �f�[�^�敪
            With .MakeDate
                .Year = Mid$(lBuf, 4, 4)    '' �N
                .Month = Mid$(lBuf, 8, 2)   '' ��
                .Day = Mid$(lBuf, 10, 2)     '' ��
            End With ' MakeDate
        End With ' head
        .TresenKubun = Mid$(lBuf, 12, 1)     '' �g���Z���敪
        With .ChokyoDate
            .Year = Mid$(lBuf, 13, 4)        '' �N
            .Month = Mid$(lBuf, 17, 2)       '' ��
            .Day = Mid$(lBuf, 19, 2)         '' ��
        End With ' ChokyoDate
        .ChokyoTime = Mid$(lBuf, 21, 4)      '' ��������
        .KettoNum = Mid$(lBuf, 25, 10)       '' �����o�^�ԍ�
        .HaronTime4 = Mid$(lBuf, 35, 4)      '' 4�n�����^�C�����v(800M-0M)
        .LapTime4 = Mid$(lBuf, 39, 3)        '' ���b�v�^�C��(800M-600M)
        .HaronTime3 = Mid$(lBuf, 42, 4)      '' 3�n�����^�C�����v(600M-0M)
        .LapTime3 = Mid$(lBuf, 46, 3)        '' ���b�v�^�C��(600M-400M)
        .HaronTime2 = Mid$(lBuf, 49, 4)      '' 2�n�����^�C�����v(400M-0M)
        .LapTime2 = Mid$(lBuf, 53, 3)        '' ���b�v�^�C��(400M-200M)
        .LapTime1 = Mid$(lBuf, 56, 3)        '' ���b�v�^�C��(200M-0M)
        .CRLF = Mid$(lBuf, 59, 2)            '' ���R�[�h��؂�
    End With

  End Sub


'
'   �@�\: �\���̂Ƀf�[�^���Z�b�g����|�n�̏d
'
'   ���l: �Ȃ�
'
    Public Sub SetDataFromByte_WH(bytBuf() As Byte, ByRef mBuf As JV_WH_BATAIJYU)
    Dim i As Integer                                '' ���[�v�J�E���^
    Dim j As Integer                                '' ���[�v�J�E���^
    Dim k As Integer                                '' ���[�v�J�E���^
    Dim p As Long                                   '' �؂蕪���J�n�ʒu
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMidByte(bytBuf, p, 2)      '' ���R�[�h���
            .DataKubun = IncMidByte(bytBuf, p, 1)       '' �f�[�^�敪
            With .MakeDate
                .Year = IncMidByte(bytBuf, p, 4)        '' �N
                .Month = IncMidByte(bytBuf, p, 2)       '' ��
                .Day = IncMidByte(bytBuf, p, 2)         '' ��
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMidByte(bytBuf, p, 4)            '' �J�ÔN
            .MonthDay = IncMidByte(bytBuf, p, 4)        '' �J�Ì���
            .JyoCD = IncMidByte(bytBuf, p, 2)           '' ���n��R�[�h
            .Kaiji = IncMidByte(bytBuf, p, 2)           '' �J�É�[��N��]
            .Nichiji = IncMidByte(bytBuf, p, 2)         '' �J�Ó���[N����]
            .RaceNum = IncMidByte(bytBuf, p, 2)         '' ���[�X�ԍ�
        End With ' id
        With .HappyoTime
            .Month = IncMidByte(bytBuf, p, 2)           '' ��
            .Day = IncMidByte(bytBuf, p, 2)             '' ��
            .Hour = IncMidByte(bytBuf, p, 2)            '' ��
            .Minute = IncMidByte(bytBuf, p, 2)          '' ��
        End With ' HappyoTime
        For i = 0 To 17
            With .BataijyuInfo(i)
                .Umaban = IncMidByte(bytBuf, p, 2)      '' �n��
                .BAMEI = IncMidByte(bytBuf, p, 36)      '' �n��
                .BaTaijyu = IncMidByte(bytBuf, p, 3)    '' �n�̏d
                .ZogenFugo = IncMidByte(bytBuf, p, 1)   '' ��������
                .ZogenSa = IncMidByte(bytBuf, p, 3)     '' ������
            End With ' BataijyuInfo
        Next i
        .CRLF = IncMidByte(bytBuf, p, 2)                '' ���R�[�h��؂�
    End With

    End Sub


'
'   �@�\: �\���̂Ƀf�[�^���Z�b�g����|�V��n����
'
'   ���l: �Ȃ�
'
    Public Sub SetDataFromByte_WE(bytBuf() As Byte, ByRef mBuf As JV_WE_WEATHER)
    Dim i As Integer                                '' ���[�v�J�E���^
    Dim j As Integer                                '' ���[�v�J�E���^
    Dim k As Integer                                '' ���[�v�J�E���^
    Dim p As Long                                   '' �؂蕪���J�n�ʒu
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMidByte(bytBuf, p, 2)      '' ���R�[�h���
            .DataKubun = IncMidByte(bytBuf, p, 1)       '' �f�[�^�敪
            With .MakeDate
                .Year = IncMidByte(bytBuf, p, 4)        '' �N
                .Month = IncMidByte(bytBuf, p, 2)       '' ��
                .Day = IncMidByte(bytBuf, p, 2)         '' ��
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMidByte(bytBuf, p, 4)            '' �J�ÔN
            .MonthDay = IncMidByte(bytBuf, p, 4)        '' �J�Ì���
            .JyoCD = IncMidByte(bytBuf, p, 2)           '' ���n��R�[�h
            .Kaiji = IncMidByte(bytBuf, p, 2)           '' �J�É�[��N��]
            .Nichiji = IncMidByte(bytBuf, p, 2)         '' �J�Ó���[N����]
        End With ' id
        With .HappyoTime
            .Month = IncMidByte(bytBuf, p, 2)           '' ��
            .Day = IncMidByte(bytBuf, p, 2)             '' ��
            .Hour = IncMidByte(bytBuf, p, 2)            '' ��
            .Minute = IncMidByte(bytBuf, p, 2)          '' ��
        End With ' HappyoTime
        .HenkoID = IncMidByte(bytBuf, p, 1)             '' �ύX����
        With .TenkoBaba
            .TenkoCD = IncMidByte(bytBuf, p, 1)         '' �V��R�[�h
            .SibaBabaCD = IncMidByte(bytBuf, p, 1)      '' �Ŕn���ԃR�[�h
            .DirtBabaCD = IncMidByte(bytBuf, p, 1)      '' �_�[�g�n���ԃR�[�h
        End With ' TenkoBaba
        With .TenkoBabaBefore
            .TenkoCD = IncMidByte(bytBuf, p, 1)         '' �V��R�[�h
            .SibaBabaCD = IncMidByte(bytBuf, p, 1)      '' �Ŕn���ԃR�[�h
            .DirtBabaCD = IncMidByte(bytBuf, p, 1)      '' �_�[�g�n���ԃR�[�h
        End With ' TenkoBabaBefore
        .CRLF = IncMidByte(bytBuf, p, 2)                '' ���R�[�h��؂�
    End With

    End Sub


'
'   �@�\: �\���̂Ƀf�[�^���Z�b�g����|�o������E�������O
'
'   ���l: �Ȃ�
'
    Public Sub SetDataFromByte_AV(bytBuf() As Byte, ByRef mBuf As JV_AV_INFO)
    Dim i As Integer                                '' ���[�v�J�E���^
    Dim j As Integer                                '' ���[�v�J�E���^
    Dim k As Integer                                '' ���[�v�J�E���^
    Dim p As Long                                   '' �؂蕪���J�n�ʒu
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMidByte(bytBuf, p, 2)      '' ���R�[�h���
            .DataKubun = IncMidByte(bytBuf, p, 1)       '' �f�[�^�敪
            With .MakeDate
                .Year = IncMidByte(bytBuf, p, 4)        '' �N
                .Month = IncMidByte(bytBuf, p, 2)       '' ��
                .Day = IncMidByte(bytBuf, p, 2)         '' ��
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMidByte(bytBuf, p, 4)            '' �J�ÔN
            .MonthDay = IncMidByte(bytBuf, p, 4)        '' �J�Ì���
            .JyoCD = IncMidByte(bytBuf, p, 2)           '' ���n��R�[�h
            .Kaiji = IncMidByte(bytBuf, p, 2)           '' �J�É�[��N��]
            .Nichiji = IncMidByte(bytBuf, p, 2)         '' �J�Ó���[N����]
            .RaceNum = IncMidByte(bytBuf, p, 2)         '' ���[�X�ԍ�
        End With ' id
        With .HappyoTime
            .Month = IncMidByte(bytBuf, p, 2)           '' ��
            .Day = IncMidByte(bytBuf, p, 2)             '' ��
            .Hour = IncMidByte(bytBuf, p, 2)            '' ��
            .Minute = IncMidByte(bytBuf, p, 2)          '' ��
        End With ' HappyoTime
        .Umaban = IncMidByte(bytBuf, p, 2)              '' �n��
        .BAMEI = IncMidByte(bytBuf, p, 36)              '' �n��
        .JiyuKubun = IncMidByte(bytBuf, p, 3)           '' ���R�敪
        .CRLF = IncMidByte(bytBuf, p, 2)                '' ���R�[�h��؂�
    End With

    End Sub

'
'   �@�\: �\���̂Ƀf�[�^���Z�b�g����|�R��ύX
'
'   ���l: �Ȃ�
'
    Public Sub SetDataFromByte_JC(bytBuf() As Byte, ByRef mBuf As JV_JC_INFO)
    Dim i As Integer                                '' ���[�v�J�E���^
    Dim j As Integer                                '' ���[�v�J�E���^
    Dim k As Integer                                '' ���[�v�J�E���^
    Dim p As Long                                   '' �؂蕪���J�n�ʒu
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMidByte(bytBuf, p, 2)  '' ���R�[�h���
            .DataKubun = IncMidByte(bytBuf, p, 1)   '' �f�[�^�敪
            With .MakeDate
                .Year = IncMidByte(bytBuf, p, 4)    '' �N
                .Month = IncMidByte(bytBuf, p, 2)   '' ��
                .Day = IncMidByte(bytBuf, p, 2)     '' ��
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMidByte(bytBuf, p, 4)        '' �J�ÔN
            .MonthDay = IncMidByte(bytBuf, p, 4)        '' �J�Ì���
            .JyoCD = IncMidByte(bytBuf, p, 2)       '' ���n��R�[�h
            .Kaiji = IncMidByte(bytBuf, p, 2)       '' �J�É�[��N��]
            .Nichiji = IncMidByte(bytBuf, p, 2)     '' �J�Ó���[N����]
            .RaceNum = IncMidByte(bytBuf, p, 2)     '' ���[�X�ԍ�
        End With ' id
        With .HappyoTime
            .Month = IncMidByte(bytBuf, p, 2)       '' ��
            .Day = IncMidByte(bytBuf, p, 2)         '' ��
            .Hour = IncMidByte(bytBuf, p, 2)        '' ��
            .Minute = IncMidByte(bytBuf, p, 2)      '' ��
        End With ' HappyoTime
        .Umaban = IncMidByte(bytBuf, p, 2)          '' �n��
        .BAMEI = IncMidByte(bytBuf, p, 36)          '' �n��
        With .JCInfoAfter
            .Futan = IncMidByte(bytBuf, p, 3)       '' ���S�d��
            .KisyuCode = IncMidByte(bytBuf, p, 5)   '' �R��R�[�h
            .KisyuName = IncMidByte(bytBuf, p, 34)  '' �R�薼
            .MinaraiCD = IncMidByte(bytBuf, p, 1)   '' �R�茩�K�R�[�h
        End With ' JCInfoAfter
        With .JCInfoBefore
            .Futan = IncMidByte(bytBuf, p, 3)       '' ���S�d��
            .KisyuCode = IncMidByte(bytBuf, p, 5)   '' �R��R�[�h
            .KisyuName = IncMidByte(bytBuf, p, 34)  '' �R�薼
            .MinaraiCD = IncMidByte(bytBuf, p, 1)   '' �R�茩�K�R�[�h
        End With ' JCInfoBefore
        .CRLF = IncMidByte(bytBuf, p, 2)            '' ���R�[�h��؂�
    End With

    End Sub

'
'   �@�\: �\���̂Ƀf�[�^���Z�b�g����|�f�[�^�}�C�j���O�\�z
'
'   ���l: �Ȃ�
'
    Public Sub SetDataFromByte_DM(bytBuf() As Byte, ByRef mBuf As JV_DM_INFO)
    Dim i As Integer                                '' ���[�v�J�E���^
    Dim j As Integer                                '' ���[�v�J�E���^
    Dim k As Integer                                '' ���[�v�J�E���^
    Dim p As Long                                   '' �؂蕪���J�n�ʒu
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMidByte(bytBuf, p, 2)  '' ���R�[�h���
            .DataKubun = IncMidByte(bytBuf, p, 1)   '' �f�[�^�敪
            With .MakeDate
                .Year = IncMidByte(bytBuf, p, 4)    '' �N
                .Month = IncMidByte(bytBuf, p, 2)   '' ��
                .Day = IncMidByte(bytBuf, p, 2)     '' ��
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMidByte(bytBuf, p, 4)        '' �J�ÔN
            .MonthDay = IncMidByte(bytBuf, p, 4)    '' �J�Ì���
            .JyoCD = IncMidByte(bytBuf, p, 2)       '' ���n��R�[�h
            .Kaiji = IncMidByte(bytBuf, p, 2)       '' �J�É�[��N��]
            .Nichiji = IncMidByte(bytBuf, p, 2)     '' �J�Ó���[N����]
            .RaceNum = IncMidByte(bytBuf, p, 2)     '' ���[�X�ԍ�
        End With ' id
        With .MakeHM
            .Hour = IncMidByte(bytBuf, p, 2)        '' ��
            .Minute = IncMidByte(bytBuf, p, 2)      '' ��
        End With ' MakeHM
        For i = 0 To 17
            With .DMInfo(i)
                .Umaban = IncMidByte(bytBuf, p, 2)  '' �n��
                .DMTime = IncMidByte(bytBuf, p, 5)  '' �\�z���j�^�C��
                .DMGosaP = IncMidByte(bytBuf, p, 4) '' �\�z�덷(�M���x)�{
                .DMGosaM = IncMidByte(bytBuf, p, 4) '' �\�z�덷(�M���x)�|
            End With ' DMInfo
        Next i
        .CRLF = IncMidByte(bytBuf, p, 2)            '' ���R�[�h��؂�
    End With

    End Sub


'
'   �@�\: �\���̂Ƀf�[�^���Z�b�g����|�J�ÃX�P�W���[��
'
'   ���l: �Ȃ�
'
    Public Sub SetDataFromByte_YS(bytBuf() As Byte, ByRef mBuf As JV_YS_SCHEDULE)
    Dim i As Integer                                '' ���[�v�J�E���^
    Dim j As Integer                                '' ���[�v�J�E���^
    Dim k As Integer                                '' ���[�v�J�E���^
    Dim p As Long                                   '' �؂蕪���J�n�ʒu
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMidByte(bytBuf, p, 2)      '' ���R�[�h���
            .DataKubun = IncMidByte(bytBuf, p, 1)       '' �f�[�^�敪
            With .MakeDate
                .Year = IncMidByte(bytBuf, p, 4)        '' �N
                .Month = IncMidByte(bytBuf, p, 2)       '' ��
                .Day = IncMidByte(bytBuf, p, 2)         '' ��
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMidByte(bytBuf, p, 4)            '' �J�ÔN
            .MonthDay = IncMidByte(bytBuf, p, 4)        '' �J�Ì���
            .JyoCD = IncMidByte(bytBuf, p, 2)           '' ���n��R�[�h
            .Kaiji = IncMidByte(bytBuf, p, 2)           '' �J�É�[��N��]
            .Nichiji = IncMidByte(bytBuf, p, 2)         '' �J�Ó���[N����]
        End With ' id
        .YoubiCD = IncMidByte(bytBuf, p, 1)             '' �j���R�[�h
        For i = 0 To 2
            With .JyusyoInfo(i)
                .TokuNum = IncMidByte(bytBuf, p, 4)     '' ���ʋ����ԍ�
                .Hondai = IncMidByte(bytBuf, p, 60)     '' �������{��
                .Ryakusyo10 = IncMidByte(bytBuf, p, 20) '' ����������10��
                .Ryakusyo6 = IncMidByte(bytBuf, p, 12)  '' ����������6��
                .Ryakusyo3 = IncMidByte(bytBuf, p, 6)   '' ����������3��
                .Nkai = IncMidByte(bytBuf, p, 3)        '' �d�܉�[��N��]
                .GradeCD = IncMidByte(bytBuf, p, 1)     '' �O���[�h�R�[�h
                .SyubetuCD = IncMidByte(bytBuf, p, 2)   '' ������ʃR�[�h
                .KigoCD = IncMidByte(bytBuf, p, 3)      '' �����L���R�[�h
                .JyuryoCD = IncMidByte(bytBuf, p, 1)    '' �d�ʎ�ʃR�[�h
                .KYORI = IncMidByte(bytBuf, p, 4)       '' ����
                .TrackCD = IncMidByte(bytBuf, p, 2)     '' �g���b�N�R�[�h
            End With ' JyusyoInfo
        Next i
        .CRLF = IncMidByte(bytBuf, p, 2)                '' ���R�[�h��؂�
    End With

    End Sub
    
    
'
'   �@�\: �\���̂Ƀf�[�^���Z�b�g����|���������ύX
'
'   ���l: �Ȃ�
'
    Public Sub SetDataFromByte_TC(bytBuf() As Byte, ByRef mBuf As JV_TC_HASSOU)

        Dim p As Long
    
        p = 1
        With mBuf
            With .head
                .RecordSpec = IncMidByte(bytBuf, p, 2)      '' ���R�[�h���
                .DataKubun = IncMidByte(bytBuf, p, 1)       '' �f�[�^�敪
                With .MakeDate
                    .Year = IncMidByte(bytBuf, p, 4)        '' �N
                    .Month = IncMidByte(bytBuf, p, 2)       '' ��
                    .Day = IncMidByte(bytBuf, p, 2)         '' ��
                End With
            End With
            With .id
                .Year = IncMidByte(bytBuf, p, 4)            '' �J�ÔN
                .MonthDay = IncMidByte(bytBuf, p, 4)        '' �J�Ì���
                .JyoCD = IncMidByte(bytBuf, p, 2)           '' ���n��R�[�h
                .Kaiji = IncMidByte(bytBuf, p, 2)           '' �J�É�[��N��]
                .Nichiji = IncMidByte(bytBuf, p, 2)         '' �J�Ó���[N����]
                .RaceNum = IncMidByte(bytBuf, p, 2)         '' ���[�X�ԍ�
            End With
            With .HappyoTime                                '' ���\��������
                .Month = IncMidByte(bytBuf, p, 2)
                .Day = IncMidByte(bytBuf, p, 2)
                .Hour = IncMidByte(bytBuf, p, 2)
                .Minute = IncMidByte(bytBuf, p, 2)
            End With
            .AtoHassoTime.Hour = IncMidByte(bytBuf, p, 2)            '' �ύX�㎞
            .AtoHassoTime.Minute = IncMidByte(bytBuf, p, 2)          '' �ύX�㕪
            .MaeHassoTime.Hour = IncMidByte(bytBuf, p, 2)            '' �ύX�O��
            .MaeHassoTime.Minute = IncMidByte(bytBuf, p, 2)          '' �ύX�O��
            .CRLF = IncMidByte(bytBuf, p, 2)                '' ���R�[�h��؂�
        End With
        
    End Sub
    
'
'   �@�\: �\���̂Ƀf�[�^���Z�b�g����|�R�[�X�ύX
'
'   ���l: �Ȃ�
'
    Public Sub SetDataFromByte_CC(bytBuf() As Byte, ByRef mBuf As JV_CC_COURSE)
    
        Dim p As Long
        
        p = 1
        With mBuf
            With .head
                .RecordSpec = IncMidByte(bytBuf, p, 2)      '' ���R�[�h���
                .DataKubun = IncMidByte(bytBuf, p, 1)       '' �f�[�^�敪
                With .MakeDate
                    .Year = IncMidByte(bytBuf, p, 4)        '' �N
                    .Month = IncMidByte(bytBuf, p, 2)       '' ��
                    .Day = IncMidByte(bytBuf, p, 2)         '' ��
                End With
            End With
            With .id
                .Year = IncMidByte(bytBuf, p, 4)            '' �J�ÔN
                .MonthDay = IncMidByte(bytBuf, p, 4)        '' �J�Ì���
                .JyoCD = IncMidByte(bytBuf, p, 2)           '' ���n��R�[�h
                .Kaiji = IncMidByte(bytBuf, p, 2)           '' �J�É�[��N��]
                .Nichiji = IncMidByte(bytBuf, p, 2)         '' �J�Ó���[N����]
                .RaceNum = IncMidByte(bytBuf, p, 2)         '' ���[�X�ԍ�
            End With
            With .HappyoTime                                '' ���\��������
                .Month = IncMidByte(bytBuf, p, 2)
                .Day = IncMidByte(bytBuf, p, 2)
                .Hour = IncMidByte(bytBuf, p, 2)
                .Minute = IncMidByte(bytBuf, p, 2)
            End With
            .AtoKyori = IncMidByte(bytBuf, p, 4)            '' �ύX�㋗��
            .AtoTrackCD = IncMidByte(bytBuf, p, 2)          '' �ύX��g���b�N�R�[�h
            .MaeKyori = IncMidByte(bytBuf, p, 4)            '' �ύX�O����
            .MaeTrackCD = IncMidByte(bytBuf, p, 2)          '' �ύX�O�g���b�N�R�[�h
            .JiyuKubun = IncMidByte(bytBuf, p, 1)           '' ���R�R�[�h
            .CRLF = IncMidByte(bytBuf, p, 2)                '' ���R�[�h��؂�
        End With
        
    End Sub
