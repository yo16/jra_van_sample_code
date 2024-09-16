Attribute VB_Name = "basSetDataFromRS"
'
'   �\���̂Ƀ��R�[�h�f�[�^���擾���郂�W���[��
'

Option Explicit

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �O���֐�
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   �@�\: �\���̂Ƀf�[�^���Z�b�g����|RA���R�[�h
'
'   ���l: �Ȃ�
'
Public Sub SetDataFromRS_RA(ByRef rs As ADODB.Recordset, ByRef buf As JV_RA_RACE)
    Dim i As Integer                                '' ���[�v�J�E���^
    Dim j As Integer                                '' ���[�v�J�E���^
    Dim k As Integer                                '' ���[�v�J�E���^
    
    With buf
        With .head
            .RecordSpec = rs("RecordSpec")                                                      '' ���R�[�h���
            .DataKubun = rs("DataKubun")                                                        '' �f�[�^�敪
            With .MakeDate
                .Year = Mid$(rs("MakeDate"), 1, 4)                                               '' �N
                .Month = Mid$(rs("MakeDate"), 5, 2)                                              '' ��
                .Day = Mid$(rs("MakeDate"), 7, 2)                                                '' ��
            End With ' MakeDate
        End With ' head
        With .id
            .Year = rs("Year")
            .MonthDay = rs("MonthDay")
            .JyoCD = rs("JyoCD")
            .Kaiji = rs("Kaiji")
            .Nichiji = rs("Nichiji")
            .RaceNum = rs("RaceNum")
        End With
        With .RaceInfo
            .YoubiCD = rs("YoubiCD")                                                            '' �j���R�[�h
            .TokuNum = rs("TokuNum")                                                            '' ���ʋ����ԍ�
            .Hondai = rs("Hondai")                                                              '' �������{��
            .Fukudai = rs("Fukudai")                                                            '' ����������
            .Kakko = rs("Kakko")                                                                '' �������J�b�R��
            .HondaiEng = rs("HondaiEng")                                                        '' �������{�艢��
            .FukudaiEng = rs("FukudaiEng")                                                      '' ���������艢��
            .KakkoEng = rs("KakkoEng")                                                          '' �������J�b�R������
            .Ryakusyo10 = rs("Ryakusyo10")                                                      '' ���������̂P�O��
            .Ryakusyo6 = rs("Ryakusyo6")                                                        '' ���������̂U��
            .Ryakusyo3 = rs("Ryakusyo3")                                                        '' ���������̂R��
            .Kubun = rs("Kubun")                                                                '' �������敪
            .Nkai = rs("Nkai")                                                                  '' �d�܉�[��N��]
        End With ' RaceInfo
        .GradeCD = rs("GradeCD")                                                                '' �O���[�h�R�[�h
        .GradeCDBefore = rs("GradeCDBefore")                                                    '' �ύX�O�O���[�h�R�[�h
        With .JyokenInfo
            .SyubetuCD = rs("SyubetuCD")                                                        '' ������ʃR�[�h
            .KigoCD = rs("KigoCD")                                                              '' �����L���R�[�h
            .JyuryoCD = rs("JyuryoCD")                                                          '' �d�ʎ�ʃR�[�h
            For j = 0 To 4
                .JyokenCD(j) = rs("JyokenCD" & j + 1)                                           '' ���������R�[�h
            Next j
        End With ' JyokenInfo
        .JyokenName = rs("JyokenName")                                                          '' ������������
        .KYORI = rs("Kyori")                                                                    '' ����
        .KyoriBefore = rs("KyoriBefore")                                                        '' �ύX�O����
        .TrackCD = rs("TrackCD")                                                                '' �g���b�N�R�[�h
        .TrackCDBefore = rs("TrackCDBefore")                                                    '' �ύX�O�g���b�N�R�[�h
        .CourseKubunCD = rs("CourseKubunCD")                                                    '' �R�[�X�敪
        .CourseKubunCDBefore = rs("CourseKubunCDBefore")                                        '' �ύX�O�R�[�X�敪
        For i = 0 To 6
            .Honsyokin(i) = rs("Honsyokin" & i + 1)                                             '' �{�܋�
        Next i
        For i = 0 To 4
            .HonsyokinBefore(i) = rs("HonsyokinBefore" & i + 1)                                 '' �ύX�O�{�܋�
        Next i
        For i = 0 To 4
            .Fukasyokin(i) = rs("Fukasyokin" & i + 1)                                           '' �t���܋�
        Next i
        For i = 0 To 2
            .FukasyokinBefore(i) = rs("FukasyokinBefore" & i + 1)                               '' �ύX�O�t���܋�
        Next i
        .HassoTime = rs("HassoTime")                                                            '' ��������
        .HassoTimeBefore = rs("HassoTimeBefore")                                                '' �ύX�O��������
        .TorokuTosu = rs("TorokuTosu")                                                          '' �o�^����
        .SyussoTosu = rs("SyussoTosu")                                                          '' �o������
        .NyusenTosu = rs("NyusenTosu")                                                          '' ��������
        With .TenkoBaba
            .TenkoCD = rs("TenkoCD")                                                            '' �V��R�[�h
            .SibaBabaCD = rs("SibaBabaCD")                                                      '' �Ŕn���ԃR�[�h
            .DirtBabaCD = rs("DirtBabaCD")                                                      '' �_�[�g�n���ԃR�[�h
        End With ' TenkoBaba
        For i = 0 To 24
            .LapTime(i) = rs("LapTime" & i + 1)                                                 '' ���b�v�^�C��
        Next i
        .SyogaiMileTime = rs("SyogaiMileTime")                                                  '' ��Q�}�C���^�C��
        .HaronTimeS3 = rs("HaronTimeS3")                                                        '' �O�R�n�����^�C��
        .HaronTimeS4 = rs("HaronTimeS4")                                                        '' �O�S�n�����^�C��
        .HaronTimeL3 = rs("HaronTimeL3")                                                        '' ��R�n�����^�C��
        .HaronTimeL4 = rs("HaronTimeL4")                                                        '' ��S�n�����^�C��
        For i = 0 To 3
            With .CornerInfo(i)
                .Corner = rs("Corner" & i + 1)                                                  '' �R�[�i�[
                .Syukaisu = rs("Syukaisu" & i + 1)                                              '' ����
                .Jyuni = rs("Jyuni" & i + 1)                                                    '' �e�ʉߏ���
            End With ' CornerInfo
        Next i
        .RecordUpKubun = rs("RecordUpKubun")                                                    '' ���R�[�h�X�V�敪
        .CRLF = vbCrLf 'CRLF
    End With
End Sub


'
'   �@�\: �\���̂Ƀf�[�^���Z�b�g����|AV���R�[�h
'
'   ���l: �Ȃ�
'
Public Sub SetDataFromRS_AV(ByRef rs As ADODB.Recordset, ByRef buf As JV_AV_INFO)
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer

    With buf
        With .head
            .RecordSpec = rs("RecordSpec")                                                      '' ���R�[�h���
            .DataKubun = rs("DataKubun")                                                       '' �f�[�^�敪
            With .MakeDate
                .Year = Mid$(rs("MakeDate"), 1, 4)                                               '' �N
                .Month = Mid$(rs("MakeDate"), 5, 2)                                              '' ��
                .Day = Mid$(rs("MakeDate"), 7, 2)                                                '' ��
            End With ' MakeDate
        End With ' head
        With .id
            .Year = rs("Year")                                                                 '' �J�ÔN
            .MonthDay = rs("MonthDay")                                                         '' �J�Ì���
            .JyoCD = rs("JyoCD")                                                               '' ���n��R�[�h
            .Kaiji = rs("Kaiji")                                                               '' �J�É�[��N��]
            .Nichiji = rs("Nichiji")                                                           '' �J�Ó���[N����]
            .RaceNum = rs("RaceNum")                                                           '' ���[�X�ԍ�
        End With ' id
        With .HappyoTime
            .Month = Mid$(rs("HappyoTime"), 1, 2)                                               '' ��
            .Day = Mid$(rs("HappyoTime"), 3, 2)                                                 '' ��
            .Hour = Mid$(rs("HappyoTime"), 5, 2)                                                '' ��
            .Minute = Mid$(rs("HappyoTime"), 7, 2)                                              '' ��
        End With ' HappyoTime
        .Umaban = rs("Umaban")                                                                 '' �n��
        .BAMEI = rs("Bamei")                                                                   '' �n��
        .JiyuKubun = rs("JiyuKubun")                                                           '' ���R�敪
        .CRLF = vbCrLf 'CRLF
    End With ' buf
End Sub


'
'   �@�\: �\���̂Ƀf�[�^���Z�b�g����|BN���R�[�h
'
'   ���l: �Ȃ�
'
Public Sub SetDataFromRS_BN(ByRef rs As ADODB.Recordset, ByRef buf As JV_BN_BANUSI)
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer

    With buf
        With .head
            .RecordSpec = rs("RecordSpec")                                                     '' ���R�[�h���
            .DataKubun = rs("DataKubun")                                                       '' �f�[�^�敪
            With .MakeDate
                .Year = Mid$(rs("MakeDate"), 1, 4)                                               '' �N
                .Month = Mid$(rs("MakeDate"), 5, 2)                                              '' ��
                .Day = Mid$(rs("MakeDate"), 7, 2)                                                '' ��
            End With ' MakeDate
        End With ' head
        .BanusiCode = rs("BanusiCode")
        .BanusiName_Co = rs("BanusiName_Co")                    '' �n��R�[�h
        .BanusiName = rs("BanusiName")                                                         '' �n�喼
        .BanusiNameKana = rs("BanusiNameKana")                                                 '' �n�喼���p�J�i
        .BanusiNameEng = rs("BanusiNameEng")                                                   '' �n�喼����
        .Fukusyoku = rs("Fukusyoku")                                                           '' ���F�W��
        With .HonRuikei(0)
            .SetYear = rs("H_SetYear")                                                         '' �ݒ�N
            .HonSyokinTotal = rs("H_HonSyokinTotal")                                           '' �{�܋����v
            .Fukasyokin = rs("H_Fukasyokin")                                                   '' �t���܋����v
            For j = 0 To 5
                .Chakukaisu(j) = rs("H_Chakukaisu" & j + 1)                               '' ����
            Next j
        End With ' HonRuikei(0)
        With .HonRuikei(1)
            .SetYear = rs("R_SetYear")                                                         '' �ݒ�N
            .HonSyokinTotal = rs("R_HonSyokinTotal")                                           '' �{�܋����v
            .Fukasyokin = rs("R_Fukasyokin")                                                   '' �t���܋����v
            For j = 0 To 5
                .Chakukaisu(j) = rs("R_Chakukaisu" & j + 1)                               '' ����
            Next j
        End With ' HonRuikei(1)
        .CRLF = vbCrLf 'CRLF
    End With ' buf
End Sub


'
'   �@�\: �\���̂Ƀf�[�^���Z�b�g����|BR���R�[�h
'
'   ���l: �Ȃ�
'
Public Sub SetDataFromRS_BR(ByRef rs As ADODB.Recordset, ByRef buf As JV_BR_BREEDER)
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer

    With buf
        With .head
            .RecordSpec = rs("RecordSpec")                                                     '' ���R�[�h���
            .DataKubun = rs("DataKubun")                                                       '' �f�[�^�敪
            With .MakeDate
                .Year = Mid$(rs("MakeDate"), 1, 4)                                               '' �N
                .Month = Mid$(rs("MakeDate"), 5, 2)                                              '' ��
                .Day = Mid$(rs("MakeDate"), 7, 2)                                                '' ��
            End With ' MakeDate
        End With ' head
        .BreederCode = rs("BreederCode")                                                       '' ���Y�҃R�[�h
        .BreederName_Co = rs("BreederName_Co")
        .BreederName = rs("BreederName")                                                       '' ���Y�Җ�
        .BreederNameKana = rs("BreederNameKana")                                               '' ���Y�Җ����p�J�i
        .BreederNameEng = rs("BreederNameEng")                                                 '' ���Y�Җ�����
        .Address = rs("Address")                                                               '' ���Y�ҏZ�������Ȗ�
        With .HonRuikei(0)
            .SetYear = rs("H_SetYear")                                                         '' �ݒ�N
            .HonSyokinTotal = rs("H_HonSyokinTotal")                                           '' �{�܋����v
            .Fukasyokin = rs("H_Fukasyokin")                                                   '' �t���܋����v
            For j = 0 To 5
                .Chakukaisu(j) = rs("H_Chakukaisu" & j + 1)                               '' ����
            Next j
        End With ' HonRuikei(0)
        With .HonRuikei(1)
            .SetYear = rs("R_SetYear")                                                         '' �ݒ�N
            .HonSyokinTotal = rs("R_HonSyokinTotal")                                           '' �{�܋����v
            .Fukasyokin = rs("R_Fukasyokin")                                                   '' �t���܋����v
            For j = 0 To 5
                .Chakukaisu(j) = rs("R_Chakukaisu" & j + 1)                               '' ����
            Next j
        End With ' HonRuikei(1)
        .CRLF = vbCrLf 'CRLF
    End With ' buf
End Sub


'
'   �@�\: �\���̂Ƀf�[�^���Z�b�g����|CH���R�[�h
'
'   ���l: �Ȃ�
'
Public Sub SetDataFromRS_CH(ByRef rs As ADODB.Recordset, ByRef buf As JV_CH_CHOKYOSI)
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer

    With buf
        With .head
            .RecordSpec = rs("RecordSpec")                                                     '' ���R�[�h���
            .DataKubun = rs("DataKubun")                                                       '' �f�[�^�敪
            With .MakeDate
                .Year = Mid$(rs("MakeDate"), 1, 4)                                               '' �N
                .Month = Mid$(rs("MakeDate"), 5, 2)                                              '' ��
                .Day = Mid$(rs("MakeDate"), 7, 2)                                                '' ��
            End With ' MakeDate
        End With ' head
        .ChokyosiCode = rs("ChokyosiCode")                                                     '' �����t�R�[�h
        .DelKubun = rs("DelKubun")                                                             '' �����t�����敪
        With .IssueDate
                .Year = Mid$(rs("IssueDate"), 1, 4)                                               '' �N
                .Month = Mid$(rs("IssueDate"), 5, 2)                                              '' ��
                .Day = Mid$(rs("IssueDate"), 7, 2)                                                '' ��
        End With ' IssueDate
        With .DelDate
                .Year = Mid$(rs("DelDate"), 1, 4)                                               '' �N
                .Month = Mid$(rs("DelDate"), 5, 2)                                              '' ��
                .Day = Mid$(rs("DelDate"), 7, 2)                                                '' ��
        End With ' DelDate
        With .BirthDate
                .Year = Mid$(rs("BirthDate"), 1, 4)                                               '' �N
                .Month = Mid$(rs("BirthDate"), 5, 2)                                              '' ��
                .Day = Mid$(rs("BirthDate"), 7, 2)                                                '' ��
        End With ' BirthDate
        .ChokyosiName = rs("ChokyosiName")                                                     '' �����t������
        .ChokyosiNameKana = rs("ChokyosiNameKana")                                             '' �����t�����p�J�i
        .ChokyosiRyakusyo = rs("ChokyosiRyakusyo")                                             '' �����t������
        .ChokyosiNameEng = rs("ChokyosiNameEng")                                               '' �����t������
        .SexCD = rs("SexCD")                                                                   '' ���ʋ敪
        .TozaiCD = rs("TozaiCD")                                                               '' �����t���������R�[�h
        .Syotai = rs("Syotai")                                                                 '' ���Ғn�於
        For i = 0 To 2
            With .SaikinJyusyo(i)
                With .SaikinJyusyoid
                    .Year = Mid$(rs("SaikinJyusyo" & i + 1 & "SaikinJyusyoid"), 1, 4)
                    .MonthDay = Mid$(rs("SaikinJyusyo" & i + 1 & "SaikinJyusyoid"), 5, 4)
                    .JyoCD = Mid$(rs("SaikinJyusyo" & i + 1 & "SaikinJyusyoid"), 9, 2)
                    .Kaiji = Mid$(rs("SaikinJyusyo" & i + 1 & "SaikinJyusyoid"), 11, 2)
                    .Nichiji = Mid$(rs("SaikinJyusyo" & i + 1 & "SaikinJyusyoid"), 13, 2)
                    .RaceNum = Mid$(rs("SaikinJyusyo" & i + 1 & "SaikinJyusyoid"), 15, 2)
                End With ' SaikinJyusyoid
                .Hondai = rs("SaikinJyusyo" & i + 1 & "Hondai")                                '' �������{��
                .Ryakusyo10 = rs("SaikinJyusyo" & i + 1 & "Ryakusyo10")                        '' ����������10��
                .Ryakusyo6 = rs("SaikinJyusyo" & i + 1 & "Ryakusyo6")                          '' ����������6��
                .Ryakusyo3 = rs("SaikinJyusyo" & i + 1 & "Ryakusyo3")                          '' ����������3��
                .GradeCD = rs("SaikinJyusyo" & i + 1 & "GradeCD")                              '' �O���[�h�R�[�h
                .SyussoTosu = rs("SaikinJyusyo" & i + 1 & "SyussoTosu")                        '' �o������
                .KettoNum = rs("SaikinJyusyo" & i + 1 & "KettoNum")                            '' �����o�^�ԍ�
                .BAMEI = rs("SaikinJyusyo" & i + 1 & "Bamei")                                  '' �n��
            End With ' SaikinJyusyo
        Next i

    End With ' buf
End Sub
        

'
'   �@�\: �\���̂Ƀf�[�^���Z�b�g����|CH_SEISEKI���R�[�h
'
'   ���l: �Ȃ�
'
Public Sub SetDataFromRS_CH_SEISEKI(ByRef rs As ADODB.Recordset, ByRef buf As JV_CH_CHOKYOSI)
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer

    With buf
        .ChokyosiCode = rs("ChokyosiCode")                                                 '' �����t�R�[�h
        For i = 0 To 2
            rs.Filter = "Num='" & i & "'"  ' ADO Recordset Function
            With .HonZenRuikei(i)
                .SetYear = rs("SetYear")                                                       '' �ݒ�N
                .HonSyokinHeichi = rs("HonSyokinHeichi")                                       '' ���n�{�܋����v
                .HonSyokinSyogai = rs("HonSyokinSyogai")                                       '' ��Q�{�܋����v
                .FukaSyokinHeichi = rs("FukaSyokinHeichi")                                     '' ���n�t���܋����v
                .FukaSyokinSyogai = rs("FukaSyokinSyogai")                                     '' ��Q�t���܋����v
                With .ChakuKaisuHeichi
                    For k = 0 To 5
                        .Chakukaisu(k) = rs("HeichiChakukaisu" & k + 1)
                    Next k
                End With ' ChakuKaisuHeichi
                With .ChakuKaisuSyogai
                    For k = 0 To 5
                        .Chakukaisu(k) = rs("SyogaiChakukaisu" & k + 1)
                    Next k
                End With ' ChakuKaisuSyogai
                For j = 0 To 5
                    With .ChakuKaisuKyori(j)
                        For k = 0 To 5
                            .Chakukaisu(k) = rs("Kyori" & j + 1 & "Chakukaisu" & k + 1)
                        Next k
                    End With ' ChakuKaisuKyori
                Next j
            End With ' HonZenRuikei
            With .HonZenRuikei(i)
                For j = 0 To 19
                    With .ChakuKaisuJyo(j)
                        For k = 0 To 5
                            .Chakukaisu(k) = rs("Jyo" & j + 1 & "Chakukaisu" & k + 1)
                        Next k
                    End With ' ChakuKaisuJyo
                Next j
            End With ' HonZenRuikei
        Next i
        .CRLF = vbCrLf 'CRLF
    End With ' buf
End Sub


'
'   �@�\: �\���̂Ƀf�[�^���Z�b�g����|DM���R�[�h
'
'   ���l: �Ȃ�
'
Public Sub SetDataFromRS_DM(ByRef rs As ADODB.Recordset, ByRef buf As JV_DM_INFO)
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer

    With buf
        With .head
            .RecordSpec = rs("RecordSpec")                                                     '' ���R�[�h���
            .DataKubun = rs("DataKubun")                                                       '' �f�[�^�敪
            With .MakeDate
                .Year = Mid$(rs("MakeDate"), 1, 4)                                               '' �N
                .Month = Mid$(rs("MakeDate"), 5, 2)                                              '' ��
                .Day = Mid$(rs("MakeDate"), 7, 2)                                                '' ��
            End With ' MakeDate
        End With ' head
        With .id
            .Year = rs("Year")                                                                 '' �J�ÔN
            .MonthDay = rs("MonthDay")                                                         '' �J�Ì���
            .JyoCD = rs("JyoCD")                                                               '' ���n��R�[�h
            .Kaiji = rs("Kaiji")                                                               '' �J�É�[��N��]
            .Nichiji = rs("Nichiji")                                                           '' �J�Ó���[N����]
            .RaceNum = rs("RaceNum")                                                           '' ���[�X�ԍ�
        End With ' id
        With .MakeHM
            .Hour = Mid$(rs("MakeHM"), 1, 2)                                                               '' ��
            .Minute = Mid$(rs("MakeHM"), 3, 2)                                                           '' ��
        End With ' MakeHM
        For i = 0 To 17
            With .DMInfo(i)
                .Umaban = rs("Umaban" & i + 1)                                                       '' �n��
                .DMTime = rs("DMTime" & i + 1)                                                        '' �\�z���j�^�C��
                .DMGosaP = rs("DMGosaP" & i + 1)                                                      '' �\�z�덷(�M���x)�{
                .DMGosaM = rs("DMGosaM" & i + 1)                                                      '' �\�z�덷(�M���x)�|
            End With ' DMInfo
        Next i
        .CRLF = vbCrLf 'CRLF
    End With ' buf
End Sub


'
'   �@�\: �\���̂Ƀf�[�^���Z�b�g����|H1���R�[�h
'
'   ���l: �Ȃ�
'
Public Sub SetDataFromRS_H1_(ByRef rs As ADODB.Recordset, ByRef buf As JV_H1_HYOSU_ZENKAKE)
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer

    With buf
        With .head
            .RecordSpec = rs("RecordSpec")                                                     '' ���R�[�h���
            .DataKubun = rs("DataKubun")                                                       '' �f�[�^�敪
            With .MakeDate
                .Year = Mid$(rs("MakeDate"), 1, 4)                                               '' �N
                .Month = Mid$(rs("MakeDate"), 5, 2)                                              '' ��
                .Day = Mid$(rs("MakeDate"), 7, 2)                                                '' ��
            End With ' MakeDate
        End With ' head
        With .id
            .Year = rs("Year")                                                                 '' �J�ÔN
            .MonthDay = rs("MonthDay")                                                         '' �J�Ì���
            .JyoCD = rs("JyoCD")                                                               '' ���n��R�[�h
            .Kaiji = rs("Kaiji")                                                               '' �J�É�[��N��]
            .Nichiji = rs("Nichiji")                                                           '' �J�Ó���[N����]
            .RaceNum = rs("RaceNum")                                                           '' ���[�X�ԍ�
        End With ' id
        .TorokuTosu = rs("TorokuTosu")                                                         '' �o�^����
        .SyussoTosu = rs("SyussoTosu")                                                         '' �o������
        For i = 0 To 6
            .HatubaiFlag(i) = rs("HatubaiFlag" & i + 1)                                   '' �����t���O
        Next i
        .FukuChakuBaraiKey = rs("FukuChakuBaraiKey")                                           '' ���������L�[
        For i = 0 To 27
            .HenkanUma(i) = rs("HenkanUma" & i + 1)                                       '' �ԊҔn�ԏ��(�n��01�`28)
        Next i
        For i = 0 To 7
            .HenkanWaku(i) = rs("HenkanWaku" & i + 1)                                     '' �ԊҘg�ԏ��(�g��1�`8)
        Next i
        For i = 0 To 7
            .HenkanDoWaku(i) = rs("HenkanDoWaku" & i + 1)                                 '' �Ԋғ��g���(�g��1�`8)
        Next i
        For i = 0 To 13
            .HyoTotal(i) = rs("HyoTotal" & i + 1)                                         '' �[�����v
        Next i
        
        For i = 0 To 27
            With .HyoTansyo(i)
                .Umaban = rs("Umaban")                                                         '' �n��
                .Hyo = rs("Hyo")                                                               '' �[��
                .Ninki = rs("Ninki")                                                           '' �l�C
            End With ' HyoTansyo
            With .HyoFukusyo(i)
                .Hyo = rs("Hyo")                                                               '' �[��
                .Ninki = rs("Ninki")                                                           '' �l�C
            End With ' HyoFukusyo
        Next i
    
        For i = 0 To 35
            With .HyoWakuren(i)
                .Umaban = rs("Kumi")                                                         '' �n��
                .Hyo = rs("Hyo")                                                               '' �[��
                .Ninki = rs("Ninki")                                                           '' �l�C
            End With ' HyoWakuren
        Next i
        
        For i = 0 To 152
            With .HyoUmaren(i)
                .Kumi = rs("Kumi")                                                         '' �n��
                .Hyo = rs("Hyo")                                                               '' �[��
                .Ninki = rs("Ninki")                                                           '' �l�C
            End With ' HyoUmaren
            With .HyoWide(i)
                .Hyo = rs("Hyo")                                                               '' �[��
                .Ninki = rs("Ninki")                                                           '' �l�C
            End With ' HyoWide
        Next i
    
        For i = 0 To 305
            With .HyoUmatan(i)
                .Kumi = rs("Kumi")                                                        '' �n��
                .Hyo = rs("Hyo")                                                               '' �[��
                .Ninki = rs("Ninki")                                                           '' �l�C
            End With ' HyoUmatan
        Next i
    
        For i = 0 To 815
            With .HyoSanrenpuku(i)
                .Kumi = rs("Kumi")                                                         '' �n��
                .Hyo = rs("Hyo")                                                               '' �[��
                .Ninki = rs("Ninki")                                                           '' �l�C
            End With ' HyoSanrenpuku
        Next i

    End With ' buf
End Sub


'
'   �@�\: �\���̂Ƀf�[�^���Z�b�g����|H1���R�[�h(HYOSU_TANPUKU)
'
'   ���l: �Ȃ�
'
Public Sub SetDataFromRS_H1_TANPUKU(ByRef rs As ADODB.Recordset, ByRef buf As JV_H1_HYOSU_ZENKAKE)
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer

    With buf
        With .head
            .RecordSpec = rs("RecordSpec")                                                     '' ���R�[�h���
            .DataKubun = rs("DataKubun")                                                       '' �f�[�^�敪
            With .MakeDate
                .Year = Mid$(rs("MakeDate"), 1, 4)                                               '' �N
                .Month = Mid$(rs("MakeDate"), 5, 2)                                              '' ��
                .Day = Mid$(rs("MakeDate"), 7, 2)                                                '' ��
            End With ' MakeDate
        End With ' head
        With .id
            .Year = rs("Year")                                                                 '' �J�ÔN
            .MonthDay = rs("MonthDay")                                                         '' �J�Ì���
            .JyoCD = rs("JyoCD")                                                               '' ���n��R�[�h
            .Kaiji = rs("Kaiji")                                                               '' �J�É�[��N��]
            .Nichiji = rs("Nichiji")                                                           '' �J�Ó���[N����]
            .RaceNum = rs("RaceNum")                                                           '' ���[�X�ԍ�
        End With ' id
        .TorokuTosu = rs("TorokuTosu")                                                         '' �o�^����
        .SyussoTosu = rs("SyussoTosu")                                                         '' �o������
        For i = 0 To 6
            .HatubaiFlag(i) = rs("HatubaiFlag" & i + 1)                                   '' �����t���O
        Next i
        .FukuChakuBaraiKey = rs("FukuChakuBaraiKey")                                           '' ���������L�[
        For i = 0 To 27
            .HenkanUma(i) = rs("HenkanUma" & i + 1)                                       '' �ԊҔn�ԏ��(�n��01�`28)
        Next i
        For i = 0 To 7
            .HenkanWaku(i) = rs("HenkanWaku" & i + 1)                                     '' �ԊҘg�ԏ��(�g��1�`8)
        Next i
        For i = 0 To 7
            .HenkanDoWaku(i) = rs("HenkanDoWaku" & i + 1)                                 '' �Ԋғ��g���(�g��1�`8)
        Next i
        For i = 0 To 13
            .HyoTotal(i) = rs("HyoTotal" & i + 1)                                         '' �[�����v
        Next i
        
        For i = 0 To 27
            With .HyoTansyo(i)
                .Umaban = rs("Umaban")                                                         '' �n��
                .Hyo = rs("Hyo")                                                               '' �[��
                .Ninki = rs("Ninki")                                                           '' �l�C
            End With ' HyoTansyo
            With .HyoFukusyo(i)
                .Hyo = rs("Hyo")                                                               '' �[��
                .Ninki = rs("Ninki")                                                           '' �l�C
            End With ' HyoFukusyo
        Next i
    
        For i = 0 To 35
            With .HyoWakuren(i)
                .Umaban = rs("Kumi")                                                         '' �n��
                .Hyo = rs("Hyo")                                                               '' �[��
                .Ninki = rs("Ninki")                                                           '' �l�C
            End With ' HyoWakuren
        Next i
        
        For i = 0 To 152
            With .HyoUmaren(i)
                .Kumi = rs("Kumi")                                                         '' �n��
                .Hyo = rs("Hyo")                                                               '' �[��
                .Ninki = rs("Ninki")                                                           '' �l�C
            End With ' HyoUmaren
            With .HyoWide(i)
                .Hyo = rs("Hyo")                                                               '' �[��
                .Ninki = rs("Ninki")                                                           '' �l�C
            End With ' HyoWide
        Next i
    
        For i = 0 To 305
            With .HyoUmatan(i)
                .Kumi = rs("Kumi")                                                        '' �n��
                .Hyo = rs("Hyo")                                                               '' �[��
                .Ninki = rs("Ninki")                                                           '' �l�C
            End With ' HyoUmatan
        Next i
    
        For i = 0 To 815
            With .HyoSanrenpuku(i)
                .Kumi = rs("Kumi")                                                         '' �n��
                .Hyo = rs("Hyo")                                                               '' �[��
                .Ninki = rs("Ninki")                                                           '' �l�C
            End With ' HyoSanrenpuku
        Next i

    End With ' buf
End Sub


'
'   �@�\: �\���̂Ƀf�[�^���Z�b�g����|H1���R�[�h(WAKU)
'
'   ���l: �Ȃ�
'
Public Sub SetDataFromRS_H1_WAKU(ByRef rs As ADODB.Recordset, ByRef buf As JV_H1_HYOSU_ZENKAKE)
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer

    With buf
        With .head
            .RecordSpec = rs("RecordSpec")                                                     '' ���R�[�h���
            .DataKubun = rs("DataKubun")                                                       '' �f�[�^�敪
            With .MakeDate
                .Year = Mid$(rs("MakeDate"), 1, 4)                                               '' �N
                .Month = Mid$(rs("MakeDate"), 5, 2)                                              '' ��
                .Day = Mid$(rs("MakeDate"), 7, 2)                                                '' ��
            End With ' MakeDate
        End With ' head
        With .id
            .Year = rs("Year")                                                                 '' �J�ÔN
            .MonthDay = rs("MonthDay")                                                         '' �J�Ì���
            .JyoCD = rs("JyoCD")                                                               '' ���n��R�[�h
            .Kaiji = rs("Kaiji")                                                               '' �J�É�[��N��]
            .Nichiji = rs("Nichiji")                                                           '' �J�Ó���[N����]
            .RaceNum = rs("RaceNum")                                                           '' ���[�X�ԍ�
        End With ' id
        .TorokuTosu = rs("TorokuTosu")                                                         '' �o�^����
        .SyussoTosu = rs("SyussoTosu")                                                         '' �o������
        For i = 0 To 6
            .HatubaiFlag(i) = rs("HatubaiFlag" & i + 1)                                   '' �����t���O
        Next i
        .FukuChakuBaraiKey = rs("FukuChakuBaraiKey")                                           '' ���������L�[
        For i = 0 To 27
            .HenkanUma(i) = rs("HenkanUma" & i + 1)                                       '' �ԊҔn�ԏ��(�n��01�`28)
        Next i
        For i = 0 To 7
            .HenkanWaku(i) = rs("HenkanWaku" & i + 1)                                     '' �ԊҘg�ԏ��(�g��1�`8)
        Next i
        For i = 0 To 7
            .HenkanDoWaku(i) = rs("HenkanDoWaku" & i + 1)                                 '' �Ԋғ��g���(�g��1�`8)
        Next i
        For i = 0 To 13
            .HyoTotal(i) = rs("HyoTotal" & i + 1)                                         '' �[�����v
        Next i
        
        For i = 0 To 27
            With .HyoTansyo(i)
                .Umaban = rs("Umaban")                                                         '' �n��
                .Hyo = rs("Hyo")                                                               '' �[��
                .Ninki = rs("Ninki")                                                           '' �l�C
            End With ' HyoTansyo
            With .HyoFukusyo(i)
                .Hyo = rs("Hyo")                                                               '' �[��
                .Ninki = rs("Ninki")                                                           '' �l�C
            End With ' HyoFukusyo
        Next i
    
        For i = 0 To 35
            With .HyoWakuren(i)
                .Umaban = rs("Kumi")                                                         '' �n��
                .Hyo = rs("Hyo")                                                               '' �[��
                .Ninki = rs("Ninki")                                                           '' �l�C
            End With ' HyoWakuren
        Next i
        
        For i = 0 To 152
            With .HyoUmaren(i)
                .Kumi = rs("Kumi")                                                         '' �n��
                .Hyo = rs("Hyo")                                                               '' �[��
                .Ninki = rs("Ninki")                                                           '' �l�C
            End With ' HyoUmaren
            With .HyoWide(i)
                .Hyo = rs("Hyo")                                                               '' �[��
                .Ninki = rs("Ninki")                                                           '' �l�C
            End With ' HyoWide
        Next i
    
        For i = 0 To 305
            With .HyoUmatan(i)
                .Kumi = rs("Kumi")                                                        '' �n��
                .Hyo = rs("Hyo")                                                               '' �[��
                .Ninki = rs("Ninki")                                                           '' �l�C
            End With ' HyoUmatan
        Next i
    
        For i = 0 To 815
            With .HyoSanrenpuku(i)
                .Kumi = rs("Kumi")                                                         '' �n��
                .Hyo = rs("Hyo")                                                               '' �[��
                .Ninki = rs("Ninki")                                                           '' �l�C
            End With ' HyoSanrenpuku
        Next i

    End With ' buf
End Sub


'
'   �@�\: �\���̂Ƀf�[�^���Z�b�g����|H1���R�[�h(UMAREN)
'
'   ���l: �Ȃ�
'
Public Sub SetDataFromRS_H1_UMAREN(ByRef rs As ADODB.Recordset, ByRef buf As JV_H1_HYOSU_ZENKAKE)
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer

    With buf
        With .head
            .RecordSpec = rs("RecordSpec")                                                     '' ���R�[�h���
            .DataKubun = rs("DataKubun")                                                       '' �f�[�^�敪
            With .MakeDate
                .Year = Mid$(rs("MakeDate"), 1, 4)                                               '' �N
                .Month = Mid$(rs("MakeDate"), 5, 2)                                              '' ��
                .Day = Mid$(rs("MakeDate"), 7, 2)                                                '' ��
            End With ' MakeDate
        End With ' head
        With .id
            .Year = rs("Year")                                                                 '' �J�ÔN
            .MonthDay = rs("MonthDay")                                                         '' �J�Ì���
            .JyoCD = rs("JyoCD")                                                               '' ���n��R�[�h
            .Kaiji = rs("Kaiji")                                                               '' �J�É�[��N��]
            .Nichiji = rs("Nichiji")                                                           '' �J�Ó���[N����]
            .RaceNum = rs("RaceNum")                                                           '' ���[�X�ԍ�
        End With ' id
        .TorokuTosu = rs("TorokuTosu")                                                         '' �o�^����
        .SyussoTosu = rs("SyussoTosu")                                                         '' �o������
        For i = 0 To 6
            .HatubaiFlag(i) = rs("HatubaiFlag" & i + 1)                                   '' �����t���O
        Next i
        .FukuChakuBaraiKey = rs("FukuChakuBaraiKey")                                           '' ���������L�[
        For i = 0 To 27
            .HenkanUma(i) = rs("HenkanUma" & i + 1)                                       '' �ԊҔn�ԏ��(�n��01�`28)
        Next i
        For i = 0 To 7
            .HenkanWaku(i) = rs("HenkanWaku" & i + 1)                                     '' �ԊҘg�ԏ��(�g��1�`8)
        Next i
        For i = 0 To 7
            .HenkanDoWaku(i) = rs("HenkanDoWaku" & i + 1)                                 '' �Ԋғ��g���(�g��1�`8)
        Next i
        For i = 0 To 13
            .HyoTotal(i) = rs("HyoTotal" & i + 1)                                         '' �[�����v
        Next i
        
        For i = 0 To 27
            With .HyoTansyo(i)
                .Umaban = rs("Umaban")                                                         '' �n��
                .Hyo = rs("Hyo")                                                               '' �[��
                .Ninki = rs("Ninki")                                                           '' �l�C
            End With ' HyoTansyo
            With .HyoFukusyo(i)
                .Hyo = rs("Hyo")                                                               '' �[��
                .Ninki = rs("Ninki")                                                           '' �l�C
            End With ' HyoFukusyo
        Next i
    
        For i = 0 To 35
            With .HyoWakuren(i)
                .Umaban = rs("Kumi")                                                         '' �n��
                .Hyo = rs("Hyo")                                                               '' �[��
                .Ninki = rs("Ninki")                                                           '' �l�C
            End With ' HyoWakuren
        Next i
        
        For i = 0 To 152
            With .HyoUmaren(i)
                .Kumi = rs("Kumi")                                                         '' �n��
                .Hyo = rs("Hyo")                                                               '' �[��
                .Ninki = rs("Ninki")                                                           '' �l�C
            End With ' HyoUmaren
            With .HyoWide(i)
                .Hyo = rs("Hyo")                                                               '' �[��
                .Ninki = rs("Ninki")                                                           '' �l�C
            End With ' HyoWide
        Next i
    
        For i = 0 To 305
            With .HyoUmatan(i)
                .Kumi = rs("Kumi")                                                        '' �n��
                .Hyo = rs("Hyo")                                                               '' �[��
                .Ninki = rs("Ninki")                                                           '' �l�C
            End With ' HyoUmatan
        Next i
    
        For i = 0 To 815
            With .HyoSanrenpuku(i)
                .Kumi = rs("Kumi")                                                         '' �n��
                .Hyo = rs("Hyo")                                                               '' �[��
                .Ninki = rs("Ninki")                                                           '' �l�C
            End With ' HyoSanrenpuku
        Next i

    End With ' buf
End Sub


'
'   �@�\: �\���̂Ƀf�[�^���Z�b�g����|H1���R�[�h(WIDE)
'
'   ���l: �Ȃ�
'
Public Sub SetDataFromRS_H1_WIDE(ByRef rs As ADODB.Recordset, ByRef buf As JV_H1_HYOSU_ZENKAKE)
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer

    With buf
        With .head
            .RecordSpec = rs("RecordSpec")                                                     '' ���R�[�h���
            .DataKubun = rs("DataKubun")                                                       '' �f�[�^�敪
            With .MakeDate
                .Year = Mid$(rs("MakeDate"), 1, 4)                                               '' �N
                .Month = Mid$(rs("MakeDate"), 5, 2)                                              '' ��
                .Day = Mid$(rs("MakeDate"), 7, 2)                                                '' ��
            End With ' MakeDate
        End With ' head
        With .id
            .Year = rs("Year")                                                                 '' �J�ÔN
            .MonthDay = rs("MonthDay")                                                         '' �J�Ì���
            .JyoCD = rs("JyoCD")                                                               '' ���n��R�[�h
            .Kaiji = rs("Kaiji")                                                               '' �J�É�[��N��]
            .Nichiji = rs("Nichiji")                                                           '' �J�Ó���[N����]
            .RaceNum = rs("RaceNum")                                                           '' ���[�X�ԍ�
        End With ' id
        .TorokuTosu = rs("TorokuTosu")                                                         '' �o�^����
        .SyussoTosu = rs("SyussoTosu")                                                         '' �o������
        For i = 0 To 6
            .HatubaiFlag(i) = rs("HatubaiFlag" & i + 1)                                   '' �����t���O
        Next i
        .FukuChakuBaraiKey = rs("FukuChakuBaraiKey")                                           '' ���������L�[
        For i = 0 To 27
            .HenkanUma(i) = rs("HenkanUma" & i + 1)                                       '' �ԊҔn�ԏ��(�n��01�`28)
        Next i
        For i = 0 To 7
            .HenkanWaku(i) = rs("HenkanWaku" & i + 1)                                     '' �ԊҘg�ԏ��(�g��1�`8)
        Next i
        For i = 0 To 7
            .HenkanDoWaku(i) = rs("HenkanDoWaku" & i + 1)                                 '' �Ԋғ��g���(�g��1�`8)
        Next i
        For i = 0 To 13
            .HyoTotal(i) = rs("HyoTotal" & i + 1)                                         '' �[�����v
        Next i
        
        For i = 0 To 27
            With .HyoTansyo(i)
                .Umaban = rs("Umaban")                                                         '' �n��
                .Hyo = rs("Hyo")                                                               '' �[��
                .Ninki = rs("Ninki")                                                           '' �l�C
            End With ' HyoTansyo
            With .HyoFukusyo(i)
                .Hyo = rs("Hyo")                                                               '' �[��
                .Ninki = rs("Ninki")                                                           '' �l�C
            End With ' HyoFukusyo
        Next i
    
        For i = 0 To 35
            With .HyoWakuren(i)
                .Umaban = rs("Kumi")                                                         '' �n��
                .Hyo = rs("Hyo")                                                               '' �[��
                .Ninki = rs("Ninki")                                                           '' �l�C
            End With ' HyoWakuren
        Next i
        
        For i = 0 To 152
            With .HyoUmaren(i)
                .Kumi = rs("Kumi")                                                         '' �n��
                .Hyo = rs("Hyo")                                                               '' �[��
                .Ninki = rs("Ninki")                                                           '' �l�C
            End With ' HyoUmaren
            With .HyoWide(i)
                .Hyo = rs("Hyo")                                                               '' �[��
                .Ninki = rs("Ninki")                                                           '' �l�C
            End With ' HyoWide
        Next i
    
        For i = 0 To 305
            With .HyoUmatan(i)
                .Kumi = rs("Kumi")                                                        '' �n��
                .Hyo = rs("Hyo")                                                               '' �[��
                .Ninki = rs("Ninki")                                                           '' �l�C
            End With ' HyoUmatan
        Next i
    
        For i = 0 To 815
            With .HyoSanrenpuku(i)
                .Kumi = rs("Kumi")                                                         '' �n��
                .Hyo = rs("Hyo")                                                               '' �[��
                .Ninki = rs("Ninki")                                                           '' �l�C
            End With ' HyoSanrenpuku
        Next i

    End With ' buf
End Sub


'
'   �@�\: �\���̂Ƀf�[�^���Z�b�g����|H1���R�[�h(UMATAN)
'
'   ���l: �Ȃ�
'
Public Sub SetDataFromRS_H1_UMATAN(ByRef rs As ADODB.Recordset, ByRef buf As JV_H1_HYOSU_ZENKAKE)
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer

    With buf
        With .head
            .RecordSpec = rs("RecordSpec")                                                     '' ���R�[�h���
            .DataKubun = rs("DataKubun")                                                       '' �f�[�^�敪
            With .MakeDate
                .Year = Mid$(rs("MakeDate"), 1, 4)                                               '' �N
                .Month = Mid$(rs("MakeDate"), 5, 2)                                              '' ��
                .Day = Mid$(rs("MakeDate"), 7, 2)                                                '' ��
            End With ' MakeDate
        End With ' head
        With .id
            .Year = rs("Year")                                                                 '' �J�ÔN
            .MonthDay = rs("MonthDay")                                                         '' �J�Ì���
            .JyoCD = rs("JyoCD")                                                               '' ���n��R�[�h
            .Kaiji = rs("Kaiji")                                                               '' �J�É�[��N��]
            .Nichiji = rs("Nichiji")                                                           '' �J�Ó���[N����]
            .RaceNum = rs("RaceNum")                                                           '' ���[�X�ԍ�
        End With ' id
        .TorokuTosu = rs("TorokuTosu")                                                         '' �o�^����
        .SyussoTosu = rs("SyussoTosu")                                                         '' �o������
        For i = 0 To 6
            .HatubaiFlag(i) = rs("HatubaiFlag" & i + 1)                                   '' �����t���O
        Next i
        .FukuChakuBaraiKey = rs("FukuChakuBaraiKey")                                           '' ���������L�[
        For i = 0 To 27
            .HenkanUma(i) = rs("HenkanUma" & i + 1)                                       '' �ԊҔn�ԏ��(�n��01�`28)
        Next i
        For i = 0 To 7
            .HenkanWaku(i) = rs("HenkanWaku" & i + 1)                                     '' �ԊҘg�ԏ��(�g��1�`8)
        Next i
        For i = 0 To 7
            .HenkanDoWaku(i) = rs("HenkanDoWaku" & i + 1)                                 '' �Ԋғ��g���(�g��1�`8)
        Next i
        For i = 0 To 13
            .HyoTotal(i) = rs("HyoTotal" & i + 1)                                         '' �[�����v
        Next i
        
        For i = 0 To 27
            With .HyoTansyo(i)
                .Umaban = rs("Umaban")                                                         '' �n��
                .Hyo = rs("Hyo")                                                               '' �[��
                .Ninki = rs("Ninki")                                                           '' �l�C
            End With ' HyoTansyo
            With .HyoFukusyo(i)
                .Hyo = rs("Hyo")                                                               '' �[��
                .Ninki = rs("Ninki")                                                           '' �l�C
            End With ' HyoFukusyo
        Next i
    
        For i = 0 To 35
            With .HyoWakuren(i)
                .Umaban = rs("Kumi")                                                         '' �n��
                .Hyo = rs("Hyo")                                                               '' �[��
                .Ninki = rs("Ninki")                                                           '' �l�C
            End With ' HyoWakuren
        Next i
        
        For i = 0 To 152
            With .HyoUmaren(i)
                .Kumi = rs("Kumi")                                                         '' �n��
                .Hyo = rs("Hyo")                                                               '' �[��
                .Ninki = rs("Ninki")                                                           '' �l�C
            End With ' HyoUmaren
            With .HyoWide(i)
                .Hyo = rs("Hyo")                                                               '' �[��
                .Ninki = rs("Ninki")                                                           '' �l�C
            End With ' HyoWide
        Next i
    
        For i = 0 To 305
            With .HyoUmatan(i)
                .Kumi = rs("Kumi")                                                        '' �n��
                .Hyo = rs("Hyo")                                                               '' �[��
                .Ninki = rs("Ninki")                                                           '' �l�C
            End With ' HyoUmatan
        Next i
    
        For i = 0 To 815
            With .HyoSanrenpuku(i)
                .Kumi = rs("Kumi")                                                         '' �n��
                .Hyo = rs("Hyo")                                                               '' �[��
                .Ninki = rs("Ninki")                                                           '' �l�C
            End With ' HyoSanrenpuku
        Next i

    End With ' buf
End Sub


'
'   �@�\: �\���̂Ƀf�[�^���Z�b�g����|H1���R�[�h(SANREN)
'
'   ���l: �Ȃ�
'
Public Sub SetDataFromRS_H1_SANREN(ByRef rs As ADODB.Recordset, ByRef buf As JV_H1_HYOSU_ZENKAKE)
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer

    With buf
        With .head
            .RecordSpec = rs("RecordSpec")                                                     '' ���R�[�h���
            .DataKubun = rs("DataKubun")                                                       '' �f�[�^�敪
            With .MakeDate
                .Year = Mid$(rs("MakeDate"), 1, 4)                                               '' �N
                .Month = Mid$(rs("MakeDate"), 5, 2)                                              '' ��
                .Day = Mid$(rs("MakeDate"), 7, 2)                                                '' ��
            End With ' MakeDate
        End With ' head
        With .id
            .Year = rs("Year")                                                                 '' �J�ÔN
            .MonthDay = rs("MonthDay")                                                         '' �J�Ì���
            .JyoCD = rs("JyoCD")                                                               '' ���n��R�[�h
            .Kaiji = rs("Kaiji")                                                               '' �J�É�[��N��]
            .Nichiji = rs("Nichiji")                                                           '' �J�Ó���[N����]
            .RaceNum = rs("RaceNum")                                                           '' ���[�X�ԍ�
        End With ' id
        .TorokuTosu = rs("TorokuTosu")                                                         '' �o�^����
        .SyussoTosu = rs("SyussoTosu")                                                         '' �o������
        For i = 0 To 6
            .HatubaiFlag(i) = rs("HatubaiFlag" & i + 1)                                   '' �����t���O
        Next i
        .FukuChakuBaraiKey = rs("FukuChakuBaraiKey")                                           '' ���������L�[
        For i = 0 To 27
            .HenkanUma(i) = rs("HenkanUma" & i + 1)                                       '' �ԊҔn�ԏ��(�n��01�`28)
        Next i
        For i = 0 To 7
            .HenkanWaku(i) = rs("HenkanWaku" & i + 1)                                     '' �ԊҘg�ԏ��(�g��1�`8)
        Next i
        For i = 0 To 7
            .HenkanDoWaku(i) = rs("HenkanDoWaku" & i + 1)                                 '' �Ԋғ��g���(�g��1�`8)
        Next i
        For i = 0 To 13
            .HyoTotal(i) = rs("HyoTotal" & i + 1)                                         '' �[�����v
        Next i
        
        For i = 0 To 27
            With .HyoTansyo(i)
                .Umaban = rs("Umaban")                                                         '' �n��
                .Hyo = rs("Hyo")                                                               '' �[��
                .Ninki = rs("Ninki")                                                           '' �l�C
            End With ' HyoTansyo
            With .HyoFukusyo(i)
                .Hyo = rs("Hyo")                                                               '' �[��
                .Ninki = rs("Ninki")                                                           '' �l�C
            End With ' HyoFukusyo
        Next i
    
        For i = 0 To 35
            With .HyoWakuren(i)
                .Umaban = rs("Kumi")                                                         '' �n��
                .Hyo = rs("Hyo")                                                               '' �[��
                .Ninki = rs("Ninki")                                                           '' �l�C
            End With ' HyoWakuren
        Next i
        
        For i = 0 To 152
            With .HyoUmaren(i)
                .Kumi = rs("Kumi")                                                         '' �n��
                .Hyo = rs("Hyo")                                                               '' �[��
                .Ninki = rs("Ninki")                                                           '' �l�C
            End With ' HyoUmaren
            With .HyoWide(i)
                .Hyo = rs("Hyo")                                                               '' �[��
                .Ninki = rs("Ninki")                                                           '' �l�C
            End With ' HyoWide
        Next i
    
        For i = 0 To 305
            With .HyoUmatan(i)
                .Kumi = rs("Kumi")                                                        '' �n��
                .Hyo = rs("Hyo")                                                               '' �[��
                .Ninki = rs("Ninki")                                                           '' �l�C
            End With ' HyoUmatan
        Next i
    
        For i = 0 To 815
            With .HyoSanrenpuku(i)
                .Kumi = rs("Kumi")                                                         '' �n��
                .Hyo = rs("Hyo")                                                               '' �[��
                .Ninki = rs("Ninki")                                                           '' �l�C
            End With ' HyoSanrenpuku
        Next i

    End With ' buf
End Sub


'
'   �@�\: �\���̂Ƀf�[�^���Z�b�g����|HC���R�[�h
'
'   ���l: �Ȃ�
'
Public Sub SetDataFromRS_HC(ByRef rs As ADODB.Recordset, ByRef buf As JV_HC_HANRO)
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer

    With buf
        With .head
            .RecordSpec = rs("RecordSpec")                              '' ���R�[�h���
            .DataKubun = rs("DataKubun")                                '' �f�[�^�敪
            With .MakeDate
                .Year = Mid$(rs("MakeDate"), 1, 4)                       '' �N
                .Month = Mid$(rs("MakeDate"), 5, 2)                      '' ��
                .Day = Mid$(rs("MakeDate"), 7, 2)                        '' ��
            End With ' MakeDate
        End With ' head
        .TresenKubun = rs("TresenKubun")                                '' �g���Z���敪
        With .ChokyoDate
            .Year = Mid$(rs("ChokyoDate"), 1, 4)                         '' �N
            .Month = Mid$(rs("ChokyoDate"), 5, 2)                        '' ��
            .Day = Mid$(rs("ChokyoDate"), 7, 2)                          '' ��
        End With ' ChokyoDate
        .ChokyoTime = rs("ChokyoTime")                                  '' ��������
        .KettoNum = rs("KettoNum")                                      '' �����o�^�ԍ�
        .HaronTime4 = rs("HaronTime4")                                  '' 4�n�����^�C�����v(800M-0M)
        .LapTime4 = rs("LapTime4")                                      '' ���b�v�^�C��(800M-600M)
        .HaronTime3 = rs("HaronTime3")                                  '' 3�n�����^�C�����v(600M-0M)
        .LapTime3 = rs("LapTime3")                                      '' ���b�v�^�C��(600M-400M)
        .HaronTime2 = rs("HaronTime2")                                  '' 2�n�����^�C�����v(400M-0M)
        .LapTime2 = rs("LapTime2")                                      '' ���b�v�^�C��(400M-200M)
        .LapTime1 = rs("LapTime1")                                      '' ���b�v�^�C��(200M-0M)
        .CRLF = vbCrLf 'CRLF
    End With ' buf
End Sub


'
'   �@�\: �\���̂Ƀf�[�^���Z�b�g����|HN���R�[�h
'
'   ���l: �Ȃ�
'
Public Sub SetDataFromRS_HN(ByRef rs As ADODB.Recordset, ByRef buf As JV_HN_HANSYOKU)
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer

    With buf
        With .head
            .RecordSpec = rs("RecordSpec")                              '' ���R�[�h���
            .DataKubun = rs("DataKubun")                                '' �f�[�^�敪
            With .MakeDate
                .Year = Mid$(rs("MakeDate"), 1, 4)                       '' �N
                .Month = Mid$(rs("MakeDate"), 5, 2)                      '' ��
                .Day = Mid$(rs("MakeDate"), 7, 2)                        '' ��
            End With ' MakeDate
        End With ' head
        .HansyokuNum = rs("HansyokuNum")                                '' �ɐB�o�^�ԍ�
        .Reserved = rs("reserved")                                      '' �\��
        .KettoNum = rs("KettoNum")                                      '' �����o�^�ԍ�
        .DelKubun = rs("DelKubun")                                      '' �ɐB�n�����敪
        .BAMEI = rs("Bamei")                                            '' �n��
        .BameiKana = rs("BameiKana")                                    '' �n�����p�J�i
        .BameiEng = rs("BameiEng")                                      '' �n������
        .BirthYear = rs("BirthYear")                                    '' ���N
        .SexCD = rs("SexCD")                                            '' ���ʃR�[�h
        .HinsyuCD = rs("HinsyuCD")                                      '' �i��R�[�h
        .KeiroCD = rs("KeiroCD")                                        '' �ѐF�R�[�h
        .HansyokuMochiKubun = rs("HansyokuMochiKubun")                  '' �ɐB�n�����敪
        .ImportYear = rs("ImportYear")                                  '' �A���N
        .SanchiName = rs("SanchiName")                                  '' �Y�n��
        .HansyokuFNum = rs("HansyokuFNum")                              '' ���n�ɐB�o�^�ԍ�
        .HansyokuMNum = rs("HansyokuMNum")                              '' ��n�ɐB�o�^�ԍ�
        .CRLF = vbCrLf 'CRLF
    End With ' buf
End Sub


'
'   �@�\: �\���̂Ƀf�[�^���Z�b�g����|HR���R�[�h
'
'   ���l: �Ȃ�
'
Public Sub SetDataFromRS_HR(ByRef rs As ADODB.Recordset, ByRef buf As JV_HR_PAY)
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer

    With buf
        With .head
            .RecordSpec = rs("RecordSpec")                              '' ���R�[�h���
            .DataKubun = rs("DataKubun")                                '' �f�[�^�敪
            With .MakeDate
                .Year = Mid$(rs("MakeDate"), 1, 4)                       '' �N
                .Month = Mid$(rs("MakeDate"), 5, 2)                      '' ��
                .Day = Mid$(rs("MakeDate"), 7, 2)                        '' ��
            End With ' MakeDate
        End With ' head
        With .id
            .Year = rs("Year")                                          '' �J�ÔN
            .MonthDay = rs("MonthDay")                                  '' �J�Ì���
            .JyoCD = rs("JyoCD")                                        '' ���n��R�[�h
            .Kaiji = rs("Kaiji")
            .Nichiji = rs("Nichiji")                                    '' �J�Ó���[N����]
            .RaceNum = rs("RaceNum")                                    '' ���[�X�ԍ�
        End With ' id
        .TorokuTosu = rs("TorokuTosu")                                  '' �o�^����
        .SyussoTosu = rs("SyussoTosu")                                  '' �o������
        For i = 0 To 8
            .FuseirituFlag(i) = rs("FuseirituFlag" & i + 1)             '' �s�����t���O
        Next i
        For i = 0 To 8
            .TokubaraiFlag(i) = rs("TokubaraiFlag" & i + 1)             '' �����t���O
        Next i
        For i = 0 To 8
            .HenkanFlag(i) = rs("HenkanFlag" & i + 1)                   '' �Ԋ҃t���O
        Next i
        For i = 0 To 27
            .HenkanUma(i) = rs("HenkanUma" & i + 1)                     '' �ԊҔn�ԏ��(�n��01�`28)
        Next i
        For i = 0 To 7
            .HenkanWaku(i) = rs("HenkanWaku" & i + 1)                   '' �ԊҘg�ԏ��(�g��1�`8)
        Next i
        For i = 0 To 7
            .HenkanDoWaku(i) = rs("HenkanDoWaku" & i + 1)               '' �Ԋғ��g���(�g��1�`8)
        Next i
        For i = 0 To 2
            With .PayTansyo(i)
                .Umaban = rs("PayTansyoUmaban" & i + 1)                 '' �n��
                .Pay = rs("PayTansyoPay" & i + 1)                       '' ���ߋ�
                .Ninki = rs("PayTansyoNinki" & i + 1)                   '' �l�C��
            End With ' PayTansyo
        Next i
        For i = 0 To 4
            With .PayFukusyo(i)
                .Umaban = rs("PayFukusyoUmaban" & i + 1)                '' �n��
                .Pay = rs("PayFukusyoPay" & i + 1)                      '' ���ߋ�
                .Ninki = rs("PayFukusyoNinki" & i + 1)                  '' �l�C��
            End With ' PayFukusyo
        Next i
        For i = 0 To 2
            With .PayWakuren(i)
                .Umaban = rs("PayWakurenKumiban" & i + 1)               '' �n��
                .Pay = rs("PayWakurenPay" & i + 1)                      '' ���ߋ�
                .Ninki = rs("PayWakurenNinki" & i + 1)                  '' �l�C��
            End With ' PayWakuren
        Next i
        For i = 0 To 2
            With .PayUmaren(i)
                .Kumi = rs("PayUmarenKumiban" & i + 1)                  '' �g��
                .Pay = rs("PayUmarenPay" & i + 1)                       '' ���ߋ�
                .Ninki = rs("PayUmarenNinki" & i + 1)                   '' �l�C��
            End With ' PayUmaren
        Next i
        For i = 0 To 6
            With .PayWide(i)
                .Kumi = rs("PayWideKumiban" & i + 1)                    '' �g��
                .Pay = rs("PayWidePay" & i + 1)                         '' ���ߋ�
                .Ninki = rs("PayWideNinki" & i + 1)                     '' �l�C��
            End With ' PayWide
        Next i
        For i = 0 To 2
            With .PayReserved1(i)
                .Kumi = rs("PayReserved1Kumiban" & i + 1)               '' �g��
                .Pay = rs("PayReserved1Pay" & i + 1)                    '' ���ߋ�
                .Ninki = rs("PayReserved1Ninki" & i + 1)                '' �l�C��
            End With ' PayReserved1
        Next i
        For i = 0 To 5
            With .PayUmatan(i)
                .Kumi = rs("PayUmatanKumiban" & i + 1)                  '' �g��
                .Pay = rs("PayUmatanPay" & i + 1)                       '' ���ߋ�
                .Ninki = rs("PayUmatanNinki" & i + 1)                   '' �l�C��
            End With ' PayUmatan
        Next i
        For i = 0 To 2
            With .PaySanrenpuku(i)
                .Kumi = rs("PaySanrenpukuKumiban" & i + 1)              '' �g��
                .Pay = rs("PaySanrenpukuPay" & i + 1)                   '' ���ߋ�
                .Ninki = rs("PaySanrenpukuNinki" & i + 1)               '' �l�C��
            End With ' PaySanrenpuku
        Next i
        For i = 0 To 5
            With .PaySanrentan(i)
                .Kumi = rs("PaySanrentanKumiban" & i + 1)               '' �g��
                .Pay = rs("PaySanrentanPay" & i + 1)                    '' ���ߋ�
                .Ninki = rs("PaySanrentanNinki" & i + 1)                '' �l�C��
            End With ' PaySanrentan
            
        Next i
        .CRLF = vbCrLf 'CRLF
    End With ' buf
End Sub


'
'   �@�\: �\���̂Ƀf�[�^���Z�b�g����|JC���R�[�h
'
'   ���l: �Ȃ�
'
Public Sub SetDataFromRS_JC(ByRef rs As ADODB.Recordset, ByRef buf As JV_JC_INFO)
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer

    With buf
        With .head
            .RecordSpec = rs("RecordSpec")                              '' ���R�[�h���
            .DataKubun = rs("DataKubun")                                '' �f�[�^�敪
            With .MakeDate
                .Year = Mid$(rs("MakeDate"), 1, 4)                       '' �N
                .Month = Mid$(rs("MakeDate"), 5, 2)                      '' ��
                .Day = Mid$(rs("MakeDate"), 7, 2)                        '' ��
            End With ' MakeDate
        End With ' head
        With .id
            .Year = rs("Year")                                          '' �J�ÔN
            .MonthDay = rs("MonthDay")                                  '' �J�Ì���
            .JyoCD = rs("JyoCD")                                        '' ���n��R�[�h
            .Kaiji = rs("Kaiji")                                        '' �J�É�[��N��]
            .Nichiji = rs("Nichiji")                                    '' �J�Ó���[N����]
            .RaceNum = rs("RaceNum")                                    '' ���[�X�ԍ�
        End With ' id
        With .HappyoTime
            .Month = Mid$(rs("HappyoTime"), 1, 2)                        '' ��
            .Day = Mid$(rs("HappyoTime"), 3, 2)                          '' ��
            .Hour = Mid$(rs("HappyoTime"), 5, 2)                         '' ��
            .Minute = Mid$(rs("HappyoTime"), 7, 2)                       '' ��
        End With ' HappyoTime
        .Umaban = rs("Umaban")                                          '' �n��
        .BAMEI = rs("Bamei")                                            '' �n��
        With .JCInfoAfter
            .Futan = rs("AtoFutan")                                        '' ���S�d��
            .KisyuCode = rs("AtoKisyuCode")                                '' �R��R�[�h
            .KisyuName = rs("AtoKisyuName")                                '' �R�薼
            .MinaraiCD = rs("AtoMinaraiCD")                                '' �R�茩�K�R�[�h
        End With ' JCInfoAfter
        With .JCInfoBefore
            .Futan = rs("MaeFutan")                                        '' ���S�d��
            .KisyuCode = rs("MaeKisyuCode")                                '' �R��R�[�h
            .KisyuName = rs("MaeKisyuName")                                '' �R�薼
            .MinaraiCD = rs("MaeMinaraiCD")                                '' �R�茩�K�R�[�h
        End With ' JCInfoBefore
        .CRLF = vbCrLf 'CRLF
    End With ' buf
End Sub


'
'   �@�\: �\���̂Ƀf�[�^���Z�b�g����|KS���R�[�h
'
'   ���l: �Ȃ�
'
Public Sub SetDataFromRS_KS(ByRef rs As ADODB.Recordset, ByRef buf As JV_KS_KISYU)
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer

    With buf
        With .head
            .RecordSpec = rs("RecordSpec")                              '' ���R�[�h���
            .DataKubun = rs("DataKubun")                                '' �f�[�^�敪
            With .MakeDate
                .Year = Mid$(rs("MakeDate"), 1, 4)                       '' �N
                .Month = Mid$(rs("MakeDate"), 5, 2)                      '' ��
                .Day = Mid$(rs("MakeDate"), 7, 2)                        '' ��
            End With ' MakeDate
        End With ' head
        .KisyuCode = rs("KisyuCode")                                    '' �R��R�[�h
        .DelKubun = rs("DelKubun")                                      '' �R�薕���敪
        With .IssueDate
            .Year = Mid$(rs("IssueDate"), 1, 4)                          '' �N
            .Month = Mid$(rs("IssueDate"), 5, 2)                         '' ��
            .Day = Mid$(rs("IssueDate"), 7, 2)                           '' ��
        End With ' IssueDate
        With .DelDate
            .Year = Mid$(rs("DelDate"), 1, 4)                            '' �N
            .Month = Mid$(rs("DelDate"), 5, 2)                           '' ��
            .Day = Mid$(rs("DelDate"), 7, 2)                             '' ��
        End With ' DelDate
        With .BirthDate
            .Year = Mid$(rs("BirthDate"), 1, 4)                          '' �N
            .Month = Mid$(rs("BirthDate"), 5, 2)                         '' ��
            .Day = Mid$(rs("BirthDate"), 7, 2)                           '' ��
        End With ' BirthDate
        .KisyuName = rs("KisyuName")                                    '' �R�薼����
        .Reserved = rs("reserved")                                      '' �\��
        .KisyuNameKana = rs("KisyuNameKana")                            '' �R�薼���p�J�i
        .KisyuRyakusyo = rs("KisyuRyakusyo")                            '' �R�薼����
        .KisyuNameEng = rs("KisyuNameEng")                              '' �R�薼����
        .SexCD = rs("SexCD")                                            '' ���ʋ敪
        .SikakuCD = rs("SikakuCD")                                      '' �R�掑�i�R�[�h
        .MinaraiCD = rs("MinaraiCD")                                    '' �R�茩�K�R�[�h
        .TozaiCD = rs("TozaiCD")                                        '' �R�蓌�������R�[�h
        .Syotai = rs("Syotai")                                          '' ���Ғn�於
        .ChokyosiCode = rs("ChokyosiCode")                              '' ���������t�R�[�h
        .ChokyosiRyakusyo = rs("ChokyosiRyakusyo")                      '' ���������t������
        For i = 0 To 1
            With .HatuKiJyo(i)
                With .Hatukijyoid
                    .Year = Mid$(rs("HatuKiJyo" & i + 1 & "Hatukijyoid"), 1, 4)
                    .MonthDay = Mid$(rs("HatuKiJyo" & i + 1 & "Hatukijyoid"), 5, 4)
                    .JyoCD = Mid$(rs("HatuKiJyo" & i + 1 & "Hatukijyoid"), 9, 2)
                    .Kaiji = Mid$(rs("HatuKiJyo" & i + 1 & "Hatukijyoid"), 11, 2)
                    .Nichiji = Mid$(rs("HatuKiJyo" & i + 1 & "Hatukijyoid"), 13, 2)
                    .RaceNum = Mid$(rs("HatuKiJyo" & i + 1 & "Hatukijyoid"), 15, 2)
                End With ' Hatukijyoid
                .SyussoTosu = rs("HatuKiJyo" & i + 1 & "SyussoTosu")            '' �o������
                .KettoNum = rs("HatuKiJyo" & i + 1 & "KettoNum")                '' �����o�^�ԍ�
                .BAMEI = rs("HatuKiJyo" & i + 1 & "Bamei")                      '' �n��
                .KakuteiJyuni = rs("HatuKiJyo" & i + 1 & "KakuteiJyuni")        '' �m�蒅��
                .IJyoCD = rs("HatuKiJyo" & i + 1 & "IJyoCD")                    '' �ُ�敪�R�[�h
            End With ' HatuKiJyo
        Next i
        For i = 0 To 1
            With .HatuSyori(i)
                With .Hatusyoriid
                    .Year = Mid$(rs("HatuSyori" & i + 1 & "Hatusyoriid"), 1, 4)
                    .MonthDay = Mid$(rs("HatuSyori" & i + 1 & "Hatusyoriid"), 5, 4)
                    .JyoCD = Mid$(rs("HatuSyori" & i + 1 & "Hatusyoriid"), 9, 2)
                    .Kaiji = Mid$(rs("HatuSyori" & i + 1 & "Hatusyoriid"), 11, 2)
                    .Nichiji = Mid$(rs("HatuSyori" & i + 1 & "Hatusyoriid"), 13, 2)
                    .RaceNum = Mid$(rs("HatuSyori" & i + 1 & "Hatusyoriid"), 15, 2)
                End With ' Hatusyoriid
                .SyussoTosu = rs("HatuSyori" & i + 1 & "SyussoTosu")            '' �o������
                .KettoNum = rs("HatuSyori" & i + 1 & "KettoNum")                '' �����o�^�ԍ�
                .BAMEI = rs("HatuSyori" & i + 1 & "Bamei")                      '' �n��
            End With ' HatuSyori
        Next i
        For i = 0 To 2
            With .SaikinJyusyo(i)
                With .SaikinJyusyoid
                    .Year = Mid$(rs("SaikinJyusyo" & i + 1 & "SaikinJyusyoid"), 1, 4)
                    .MonthDay = Mid$(rs("SaikinJyusyo" & i + 1 & "SaikinJyusyoid"), 5, 4)
                    .JyoCD = Mid$(rs("SaikinJyusyo" & i + 1 & "SaikinJyusyoid"), 9, 2)
                    .Kaiji = Mid$(rs("SaikinJyusyo" & i + 1 & "SaikinJyusyoid"), 11, 2)
                    .Nichiji = Mid$(rs("SaikinJyusyo" & i + 1 & "SaikinJyusyoid"), 13, 2)
                    .RaceNum = Mid$(rs("SaikinJyusyo" & i + 1 & "SaikinJyusyoid"), 15, 2)
                End With ' SaikinJyusyoid
                .Hondai = rs("SaikinJyusyo" & i + 1 & "Hondai")                 '' �������{��
                .Ryakusyo10 = rs("SaikinJyusyo" & i + 1 & "Ryakusyo10")         '' ����������10��
                .Ryakusyo6 = rs("SaikinJyusyo" & i + 1 & "Ryakusyo6")           '' ����������6��
                .Ryakusyo3 = rs("SaikinJyusyo" & i + 1 & "Ryakusyo3")           '' ����������3��
                .GradeCD = rs("SaikinJyusyo" & i + 1 & "GradeCD")               '' �O���[�h�R�[�h
                .SyussoTosu = rs("SaikinJyusyo" & i + 1 & "SyussoTosu")         '' �o������
                .KettoNum = rs("SaikinJyusyo" & i + 1 & "KettoNum")             '' �����o�^�ԍ�
                .BAMEI = rs("SaikinJyusyo" & i + 1 & "Bamei")                   '' �n��
            End With ' SaikinJyusyo
        Next i
    End With ' buf
End Sub
        

'
'   �@�\: �\���̂Ƀf�[�^���Z�b�g����|KS_SEISEKI���R�[�h
'
'   ���l: �Ȃ�
'
Public Sub SetDataFromRS_KS_SEISEKI(ByRef rs As ADODB.Recordset, ByRef buf As JV_KS_KISYU)
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer

    With buf
        .KisyuCode = rs("KisyuCode")                                        '' �����t�R�[�h
        For i = 0 To 2
            rs.Filter = "Num='" & i & "'"  ' ADO Recordset Function
            With .HonZenRuikei(i)
                .SetYear = rs("SetYear")                                    '' �ݒ�N
                .HonSyokinHeichi = rs("HonSyokinHeichi")                    '' ���n�{�܋����v
                .HonSyokinSyogai = rs("HonSyokinSyogai")                    '' ��Q�{�܋����v
                .FukaSyokinHeichi = rs("FukaSyokinHeichi")                  '' ���n�t���܋����v
                .FukaSyokinSyogai = rs("FukaSyokinSyogai")                  '' ��Q�t���܋����v
                With .ChakuKaisuHeichi
                    For k = 0 To 5
                        .Chakukaisu(k) = rs("HeichiChakukaisu" & k + 1)
                    Next k
                End With ' ChakuKaisuHeichi
                With .ChakuKaisuSyogai
                    For k = 0 To 5
                        .Chakukaisu(k) = rs("SyogaiChakukaisu" & k + 1)
                    Next k
                End With ' ChakuKaisuSyogai
                For j = 0 To 5
                    With .ChakuKaisuKyori(j)
                        For k = 0 To 5
                            .Chakukaisu(k) = rs("Kyori" & j + 1 & "Chakukaisu" & k + 1)
                        Next k
                    End With ' ChakuKaisuKyori
                Next j
            End With ' HonZenRuikei
            With .HonZenRuikei(i)
                For j = 0 To 19
                    With .ChakuKaisuJyo(j)
                        For k = 0 To 5
                            .Chakukaisu(k) = rs("Jyo" & j + 1 & "Chakukaisu" & k + 1)
                        Next k
                    End With ' ChakuKaisuJyo
                Next j
            End With ' HonZenRuikei
        Next i
        .CRLF = vbCrLf 'CRLF
    End With ' buf
End Sub


'
'   �@�\: �\���̂Ƀf�[�^���Z�b�g����|O1���R�[�h
'
'   ���l: �Ȃ�
'
Public Sub SetDataFromRS_O1(ByRef rs As ADODB.Recordset, ByRef buf As JV_O1_ODDS_TANFUKUWAKU)
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim dic_O1_Tan As Scripting.Dictionary
    Dim dic_O1_Fuku As Scripting.Dictionary
    Dim dic_O1_Waku As Scripting.Dictionary
    Dim strKey As String
    
    Set dic_O1_Tan = New Dictionary
    Set dic_O1_Fuku = New Dictionary
    Set dic_O1_Waku = New Dictionary
    
    With buf
        With .head
            .RecordSpec = rs("RecordSpec")                          '' ���R�[�h���
            .DataKubun = rs("DataKubun")                            '' �f�[�^�敪
            With .MakeDate
                .Year = Mid$(rs("MakeDate"), 1, 4)                   '' �N
                .Month = Mid$(rs("MakeDate"), 5, 2)                  '' ��
                .Day = Mid$(rs("MakeDate"), 7, 2)                    '' ��
            End With ' MakeDate
        End With ' head
        With .id
            .Year = rs("Year")                                      '' �J�ÔN
            .MonthDay = rs("MonthDay")                              '' �J�Ì���
            .JyoCD = rs("JyoCD")                                    '' ���n��R�[�h
            .Kaiji = rs("Kaiji")                                    '' �J�É�[��N��]
            .Nichiji = rs("Nichiji")                                '' �J�Ó���[N����]
            .RaceNum = rs("RaceNum")                                '' ���[�X�ԍ�
        End With ' id
        With .HappyoTime
            .Month = Mid$(rs("HappyoTime"), 1, 2)                    '' ��
            .Day = Mid$(rs("HappyoTime"), 3, 2)                      '' ��
            .Hour = Mid$(rs("HappyoTime"), 5, 2)                     '' ��
            .Minute = Mid$(rs("HappyoTime"), 7, 2)                   '' ��
        End With ' HappyoTime
        .TorokuTosu = rs("TorokuTosu")                              '' �o�^����
        .SyussoTosu = rs("SyussoTosu")                              '' �o������
        .TansyoFlag = rs("TansyoFlag")                              '' �����t���O
        .FukusyoFlag = rs("FukusyoFlag")                            '' �����t���O
        .WakurenFlag = rs("WakurenFlag")                            '' �����t���O�@�g�A
        .FukuChakuBaraiKey = rs("FukuChakuBaraiKey")                '' ���������L�[
        .TotalHyosuTansyo = rs("TotalHyosuTansyo")                  '' �P���[�����v
        .TotalHyosuFukusyo = rs("TotalHyosuFukusyo")                '' �����[�����v
        .TotalHyosuWakuren = rs("TotalHyosuWakuren")                '' �g�A�[�����v
        
        Call MakeDic(dic_O1_Tan, rs, "TanData", 28, 2, 8)
        Call MakeDic(dic_O1_Fuku, rs, "FukuData", 28, 2, 12)
        For i = 0 To 27
            strKey = Format$(i + 1, "00")
            With .OddsTansyoInfo(i)
                If dic_O1_Tan.Exists(strKey) Then
                    .Umaban = strKey                                '' �n��
                    .Odds = Mid$(dic_O1_Tan.item(strKey), 1, 4)      '' �I�b�Y
                    .Ninki = Mid$(dic_O1_Tan.item(strKey), 5, 2)     '' �l�C��
                Else
                    .Umaban = Space(2)
                    .Odds = Space(4)
                    .Ninki = Space(2)
                End If
            End With ' OddsTansyoInfo
            With .OddsFukusyoInfo(i)
                If dic_O1_Fuku.Exists(strKey) Then
                    .Umaban = strKey
                    .OddsLow = Mid$(dic_O1_Fuku.item(strKey), 1, 4)  '' �Œ�I�b�Y
                    .OddsHigh = Mid$(dic_O1_Fuku.item(strKey), 5, 4) '' �ō��I�b�Y
                    .Ninki = Mid$(dic_O1_Fuku.item(strKey), 9, 2)    '' �l�C��
                Else
                    .Umaban = Space(2)
                    .OddsLow = Space(4)
                    .OddsHigh = Space(4)
                    .Ninki = Space(2)
                End If
            End With ' OddsFukusyoInfo
        Next i
        
        Call MakeDic(dic_O1_Waku, rs, "WakuData", 36, 2, 9)
        k = 0
        For i = 1 To 8
            For j = i To 8
                strKey = Format$(i * 10 + j, "00")
                With .OddsWakurenInfo(k)
                    If dic_O1_Waku.Exists(strKey) Then
                        .Kumi = strKey                              '' �g
                        .Odds = Mid$(dic_O1_Waku.item(strKey), 1, 5) '' �I�b�Y
                        .Ninki = Mid$(dic_O1_Waku.item(strKey), 6, 2) '' �l�C��
                    Else
                        .Kumi = Space(2)
                        .Odds = Space(5)
                        .Ninki = Space(2)
                    End If
                End With ' OddsWakurenInfo
            k = k + 1
            Next j
        Next i
        .CRLF = vbCrLf 'CRLF
    End With ' buf
    
    Set dic_O1_Tan = Nothing
    Set dic_O1_Fuku = Nothing
    Set dic_O1_Waku = Nothing
    
End Sub


'
'   �@�\: �\���̂Ƀf�[�^���Z�b�g����|O2���R�[�h
'
'   ���l: �Ȃ�
'
Public Sub SetDataFromRS_O2(ByRef rs As ADODB.Recordset, ByRef buf As JV_O2_ODDS_UMAREN)
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim dic_O2 As Scripting.Dictionary
    Dim strKey As String

    Set dic_O2 = New Dictionary
    
    With buf
        With .head
            .RecordSpec = rs("RecordSpec")                          '' ���R�[�h���
            .DataKubun = rs("DataKubun")                            '' �f�[�^�敪
            With .MakeDate
                .Year = Mid$(rs("MakeDate"), 1, 4)                   '' �N
                .Month = Mid$(rs("MakeDate"), 5, 2)                  '' ��
                .Day = Mid$(rs("MakeDate"), 7, 2)                    '' ��
            End With ' MakeDate
        End With ' head
        With .id
            .Year = rs("Year")                                      '' �J�ÔN
            .MonthDay = rs("MonthDay")                              '' �J�Ì���
            .JyoCD = rs("JyoCD")                                    '' ���n��R�[�h
            .Kaiji = rs("Kaiji")                                    '' �J�É�[��N��]
            .Nichiji = rs("Nichiji")                                '' �J�Ó���[N����]
            .RaceNum = rs("RaceNum")                                '' ���[�X�ԍ�
        End With ' id
        With .HappyoTime
            .Month = Mid$(rs("HappyoTime"), 1, 2)                    '' ��
            .Day = Mid$(rs("HappyoTime"), 3, 2)                      '' ��
            .Hour = Mid$(rs("HappyoTime"), 5, 2)                     '' ��
            .Minute = Mid$(rs("HappyoTime"), 7, 2)                   '' ��
        End With ' HappyoTime
        .TorokuTosu = rs("TorokuTosu")                              '' �o�^����
        .SyussoTosu = rs("SyussoTosu")                              '' �o������
        .UmarenFlag = rs("UmarenFlag")                              '' �����t���O�@�n�A
        .TotalHyosuUmaren = rs("TotalHyosuUmaren")                  '' �n�A�[�����v
        
        Call MakeDic(dic_O2, rs, "Data", 153, 4, 13)
        k = 0
        For i = 1 To 17
            For j = i + 1 To 18
                strKey = Format$(i * 100 + j, "0000")
                With .OddsUmarenInfo(k)
                    If dic_O2.Exists(strKey) Then
                        .Kumi = strKey                              '' �g
                        .Odds = Mid$(dic_O2.item(strKey), 1, 6)      '' �I�b�Y
                        .Ninki = Mid$(dic_O2.item(strKey), 7, 3)     '' �l�C��
                    Else
                        .Kumi = Space(4)
                        .Odds = Space(6)
                        .Ninki = Space(3)
                    End If
                End With ' OddsUmarenInfo
                k = k + 1
            Next j
        Next i
        .CRLF = vbCrLf 'CRLF
    End With ' buf
    
    Set dic_O2 = Nothing
End Sub


'
'   �@�\: �\���̂Ƀf�[�^���Z�b�g����|O3���R�[�h
'
'   ���l: �Ȃ�
'
Public Sub SetDataFromRS_O3(ByRef rs As ADODB.Recordset, ByRef buf As JV_O3_ODDS_WIDE)
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim dic_O3 As Scripting.Dictionary
    Dim strKey As String
    
    Set dic_O3 = New Dictionary
    
    With buf
        With .head
            .RecordSpec = rs("RecordSpec")                          '' ���R�[�h���
            .DataKubun = rs("DataKubun")                            '' �f�[�^�敪
            With .MakeDate
                .Year = Mid$(rs("MakeDate"), 1, 4)                   '' �N
                .Month = Mid$(rs("MakeDate"), 5, 2)                  '' ��
                .Day = Mid$(rs("MakeDate"), 7, 2)                    '' ��
            End With ' MakeDate
        End With ' head
        With .id
            .Year = rs("Year")                                      '' �J�ÔN
            .MonthDay = rs("MonthDay")                              '' �J�Ì���
            .JyoCD = rs("JyoCD")                                    '' ���n��R�[�h
            .Kaiji = rs("Kaiji")                                    '' �J�É�[��N��]
            .Nichiji = rs("Nichiji")                                '' �J�Ó���[N����]
            .RaceNum = rs("RaceNum")                                '' ���[�X�ԍ�
        End With ' id
        With .HappyoTime
            .Month = Mid$(rs("HappyoTime"), 1, 2)                    '' ��
            .Day = Mid$(rs("HappyoTime"), 3, 2)                      '' ��
            .Hour = Mid$(rs("HappyoTime"), 5, 2)                     '' ��
            .Minute = Mid$(rs("HappyoTime"), 7, 2)                   '' ��
        End With ' HappyoTime
        .TorokuTosu = rs("TorokuTosu")                              '' �o�^����
        .SyussoTosu = rs("SyussoTosu")                              '' �o������
        .WideFlag = rs("WideFlag")                                  '' �����t���O�@���C�h
        .TotalHyosuWide = rs("TotalHyosuWide")                      '' ���C�h�[�����v
                
        Call MakeDic(dic_O3, rs, "Data", 153, 4, 17)
        k = 0
        For i = 1 To 17
            For j = i + 1 To 18
                strKey = Format$(i * 100 + j, "0000")
                With .OddsWideInfo(k)
                    If dic_O3.Exists(strKey) Then
                        .Kumi = strKey                              '' �g��
                        .OddsLow = Mid$(dic_O3.item(strKey), 1, 5)   '' �Œ�I�b�Y
                        .OddsHigh = Mid$(dic_O3.item(strKey), 6, 5)  '' �ō��I�b�Y
                        .Ninki = Mid$(dic_O3.item(strKey), 11, 3)    '' �l�C��
                    Else
                        .Kumi = Space(4)
                        .OddsLow = Space(5)
                        .OddsHigh = Space(5)
                        .Ninki = Space(3)
                    End If
                End With ' OddsWideInfo
                k = k + 1
            Next j
        Next i
        .CRLF = vbCrLf 'CRLF
    End With ' buf
    
    Set dic_O3 = Nothing
End Sub


'
'   �@�\: �\���̂Ƀf�[�^���Z�b�g����|O4���R�[�h
'
'   ���l: �Ȃ�
'
Public Sub SetDataFromRS_O4(ByRef rs As ADODB.Recordset, ByRef buf As JV_O4_ODDS_UMATAN)
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim dic_O4 As Scripting.Dictionary
    Dim strKey As String
    
    Set dic_O4 = New Dictionary

    With buf
        With .head
            .RecordSpec = rs("RecordSpec")                              '' ���R�[�h���
            .DataKubun = rs("DataKubun")                                '' �f�[�^�敪
            With .MakeDate
                .Year = Mid$(rs("MakeDate"), 1, 4)                       '' �N
                .Month = Mid$(rs("MakeDate"), 5, 2)                      '' ��
                .Day = Mid$(rs("MakeDate"), 7, 2)                        '' ��
            End With ' MakeDate
        End With ' head
        With .id
            .Year = rs("Year")                                          '' �J�ÔN
            .MonthDay = rs("MonthDay")                                  '' �J�Ì���
            .JyoCD = rs("JyoCD")                                        '' ���n��R�[�h
            .Kaiji = rs("Kaiji")                                        '' �J�É�[��N��]
            .Nichiji = rs("Nichiji")                                    '' �J�Ó���[N����]
            .RaceNum = rs("RaceNum")                                    '' ���[�X�ԍ�
        End With ' id
        With .HappyoTime
            .Month = Mid$(rs("HappyoTime"), 1, 2)                        '' ��
            .Day = Mid$(rs("HappyoTime"), 3, 2)                          '' ��
            .Hour = Mid$(rs("HappyoTime"), 5, 2)                         '' ��
            .Minute = Mid$(rs("HappyoTime"), 7, 2)                       '' ��
        End With ' HappyoTime
        .TorokuTosu = rs("TorokuTosu")                                  '' �o�^����
        .SyussoTosu = rs("SyussoTosu")                                  '' �o������
        .UmatanFlag = rs("UmatanFlag")                                  '' �����t���O�@�n�P
        .TotalHyosuUmatan = rs("TotalHyosuUmatan")                      '' �n�P�[�����v
        
        Call MakeDic(dic_O4, rs, "Data", 306, 4, 13)
        k = 0
        For i = 1 To 18
            For j = 1 To 18
                If (j <> i) Then
                    strKey = Format$(i * 100 + j, "0000")
                    With .OddsUmatanInfo(k)
                        If dic_O4.Exists(strKey) Then
                            .Kumi = strKey                              '' �g��
                            .Odds = Mid$(dic_O4.item(strKey), 1, 6)      '' �I�b�Y
                            .Ninki = Mid$(dic_O4.item(strKey), 7, 3)     '' �l�C��
                        Else
                            .Kumi = Space(4)
                            .Odds = Space(6)
                            .Ninki = Space(3)
                        End If
                    End With ' OddsUmatanInfo
                    k = k + 1
                End If
                
            Next j
        Next i
        .CRLF = vbCrLf 'CRLF
    End With ' buf
    
    Set dic_O4 = Nothing
End Sub


'
'   �@�\: �\���̂Ƀf�[�^���Z�b�g����|O5���R�[�h
'
'   ���l: �Ȃ�
'
Public Sub SetDataFromRS_O5(ByRef rs As ADODB.Recordset, ByRef buf As JV_O5_ODDS_SANREN)
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim p As Long
    Dim dic_O5 As Scripting.Dictionary
    Dim strKey As String
    
    Set dic_O5 = New Dictionary
    
    With buf
        With .head
            .RecordSpec = rs("RecordSpec")                              '' ���R�[�h���
            .DataKubun = rs("DataKubun")                                '' �f�[�^�敪
            With .MakeDate
                .Year = Mid$(rs("MakeDate"), 1, 4)                       '' �N
                .Month = Mid$(rs("MakeDate"), 5, 2)                      '' ��
                .Day = Mid$(rs("MakeDate"), 7, 2)                        '' ��
            End With ' MakeDate
        End With ' head
        With .id
            .Year = rs("Year")                                          '' �J�ÔN
            .MonthDay = rs("MonthDay")                                  '' �J�Ì���
            .JyoCD = rs("JyoCD")                                        '' ���n��R�[�h
            .Kaiji = rs("Kaiji")                                        '' �J�É�[��N��]
            .Nichiji = rs("Nichiji")                                    '' �J�Ó���[N����]
            .RaceNum = rs("RaceNum")                                    '' ���[�X�ԍ�
        End With ' id
        With .HappyoTime
            .Month = Mid$(rs("HappyoTime"), 1, 2)                        '' ��
            .Day = Mid$(rs("HappyoTime"), 3, 2)                          '' ��
            .Hour = Mid$(rs("HappyoTime"), 5, 2)                         '' ��
            .Minute = Mid$(rs("HappyoTime"), 7, 2)                       '' ��
        End With ' HappyoTime
        .TorokuTosu = rs("TorokuTosu")                                  '' �o�^����
        .SyussoTosu = rs("SyussoTosu")                                  '' �o������
        .SanrenpukuFlag = rs("SanrenFlag")                              '' �����t���O�@3�A��
        .TotalHyosuSanrenpuku = rs("TotalHyosuSanren")                  '' 3�A���[�����v
        
        Call MakeDic(dic_O5, rs, "Data", 816, 6, 15)
        p = 0
        For i = 1 To 16
            For j = i + 1 To 17
                For k = j + 1 To 18
                    strKey = Format$(i * 10000 + j * 100 + k, "000000")
                    With .OddsSanrenInfo(p)
                        If dic_O5.Exists(strKey) Then
                            .Kumi = strKey                              '' �g��
                            .Odds = Mid$(dic_O5.item(strKey), 1, 6)      '' �I�b�Y
                            .Ninki = Mid$(dic_O5.item(strKey), 7, 3)     '' �l�C��
                        Else
                            .Kumi = Space(6)
                            .Odds = Space(6)
                            .Ninki = Space(3)
                        End If
                    End With ' OddsSanrenInfo
                    p = p + 1
                Next k
            Next j
        Next i
        .CRLF = vbCrLf 'CRLF
    End With ' buf
    
    Set dic_O5 = Nothing
End Sub


'
'   �@�\: �\���̂Ƀf�[�^���Z�b�g����|RC���R�[�h
'
'   ���l: �Ȃ�
'
Public Sub SetDataFromRS_RC(ByRef rs As ADODB.Recordset, ByRef buf As JV_RC_RECORD)
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer

    With buf
        With .head
            .RecordSpec = rs("RecordSpec")                                                     '' ���R�[�h���
            .DataKubun = rs("DataKubun")                                                       '' �f�[�^�敪
            With .MakeDate
                .Year = Mid$(rs("MakeDate"), 1, 4)                                               '' �N
                .Month = Mid$(rs("MakeDate"), 5, 2)                                              '' ��
                .Day = Mid$(rs("MakeDate"), 7, 2)                                                '' ��
            End With ' MakeDate
        End With ' head
        .RecInfoKubun = rs("RecInfoKubun")                                                     '' ���R�[�h���ʋ敪
        With .id
            .Year = rs("Year")                                                                 '' �J�ÔN
            .MonthDay = rs("MonthDay")                                                         '' �J�Ì���
            .JyoCD = rs("JyoCD")                                                               '' ���n��R�[�h
            .Kaiji = rs("Kaiji")                                                               '' �J�É�[��N��]
            .Nichiji = rs("Nichiji")                                                           '' �J�Ó���[N����]
            .RaceNum = rs("RaceNum")                                                           '' ���[�X�ԍ�
        End With ' id
        .TokuNum = Mid$(rs("TokuNum_SyubetuCD"), 1, 4)                                                             '' ���ʋ����ԍ�
        .SyubetuCD = Mid$(rs("TokuNum_SyubetuCD"), 5, 2)                                        '' ������ʃR�[�h
        .Hondai = rs("Hondai")                                                                 '' �������{��
        .GradeCD = rs("GradeCD")                                                               '' �O���[�h�R�[�h
        .KYORI = rs("Kyori")                                                                   '' ����
        .TrackCD = rs("TrackCD")                                                               '' ������ʃR�[�h
        .RecKubun = rs("RecKubun")                                                             '' ���R�[�h�敪
        .RecTime = rs("RecTime")                                                               '' ���R�[�h�^�C��
        With .TenkoBaba
            .TenkoCD = rs("TenkoCD")                                                           '' �V��R�[�h
            .SibaBabaCD = rs("SibaBabaCD")                                                     '' �Ŕn���ԃR�[�h
            .DirtBabaCD = rs("DirtBabaCD")                                                     '' �_�[�g�n���ԃR�[�h
        End With ' TenkoBaba
        With .RecUmaInfo(0)
            .KettoNum = rs("RecUmaKettoNum1")                                                  '' �����o�^�ԍ�
            .BAMEI = rs("RecUmaBamei1")                                                        '' �n��
            .UmaKigoCD = rs("RecUmaUmaKigoCD1")                                                '' �n�L���R�[�h
            .SexCD = rs("RecUmaSexCD1")                                                        '' ���ʃR�[�h
            .ChokyosiCode = rs("RecUmaChokyosiCode1")                                          '' �����t�R�[�h
            .ChokyosiName = rs("RecUmaChokyosiName1")                                          '' �����t��
            .Futan = rs("RecUmaFutan1")                                                        '' ���S�d��
            .KisyuCode = rs("RecUmaKisyuCode1")                                                '' �R��R�[�h
            .KisyuName = rs("RecUmaKisyuName1")                                                '' �R�薼
        End With ' RecUmaInfo
        With .RecUmaInfo(1)
            .KettoNum = rs("RecUmaKettoNum2")                                                  '' �����o�^�ԍ�
            .BAMEI = rs("RecUmaBamei2")                                                        '' �n��
            .UmaKigoCD = rs("RecUmaUmaKigoCD2")                                                '' �n�L���R�[�h
            .SexCD = rs("RecUmaSexCD2")                                                        '' ���ʃR�[�h
            .ChokyosiCode = rs("RecUmaChokyosiCode2")                                          '' �����t�R�[�h
            .ChokyosiName = rs("RecUmaChokyosiName2")                                          '' �����t��
            .Futan = rs("RecUmaFutan2")                                                        '' ���S�d��
            .KisyuCode = rs("RecUmaKisyuCode2")                                                '' �R��R�[�h
            .KisyuName = rs("RecUmaKisyuName2")                                                '' �R�薼
        End With ' RecUmaInfo
        With .RecUmaInfo(2)
            .KettoNum = rs("RecUmaKettoNum3")                                                  '' �����o�^�ԍ�
            .BAMEI = rs("RecUmaBamei3")                                                        '' �n��
            .UmaKigoCD = rs("RecUmaUmaKigoCD3")                                                '' �n�L���R�[�h
            .SexCD = rs("RecUmaSexCD3")                                                        '' ���ʃR�[�h
            .ChokyosiCode = rs("RecUmaChokyosiCode3")                                          '' �����t�R�[�h
            .ChokyosiName = rs("RecUmaChokyosiName3")                                          '' �����t��
            .Futan = rs("RecUmaFutan3")                                                        '' ���S�d��
            .KisyuCode = rs("RecUmaKisyuCode3")                                                '' �R��R�[�h
            .KisyuName = rs("RecUmaKisyuName3")                                                '' �R�薼
        End With ' RecUmaInfo
        .CRLF = vbCrLf 'CRLF
    End With ' buf
End Sub


'
'   �@�\: �\���̂Ƀf�[�^���Z�b�g����|SE���R�[�h
'
'   ���l: �Ȃ�
'
Public Sub SetDataFromRS_SE(ByRef rs As ADODB.Recordset, ByRef buf As JV_SE_RACE_UMA)
On Error GoTo ErrorHandler
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer

    With buf
        With .head
            .RecordSpec = rs("RecordSpec")                                                     '' ���R�[�h���
            .DataKubun = rs("DataKubun")                                                       '' �f�[�^�敪
            With .MakeDate
                .Year = Mid$(rs("MakeDate"), 1, 4)                                               '' �N
                .Month = Mid$(rs("MakeDate"), 5, 2)                                              '' ��
                .Day = Mid$(rs("MakeDate"), 7, 2)                                                '' ��
            End With ' MakeDate
        End With ' head
        With .id
            .Year = rs("Year")                                                                 '' �J�ÔN
            .MonthDay = rs("MonthDay")                                                         '' �J�Ì���
            .JyoCD = rs("JyoCD")                                                               '' ���n��R�[�h
            .Kaiji = rs("Kaiji")                                                               '' �J�É�[��N��]
            .Nichiji = rs("Nichiji")                                                           '' �J�Ó���[N����]
            .RaceNum = rs("RaceNum")                                                           '' ���[�X�ԍ�
        End With ' id
        .Wakuban = rs("Wakuban")                                                               '' �g��
        .Umaban = rs("Umaban")                                                                 '' �n��
        .KettoNum = rs("KettoNum")                                                             '' �����o�^�ԍ�
        .BAMEI = rs("Bamei")                                                                   '' �n��
        .UmaKigoCD = rs("UmaKigoCD")                                                           '' �n�L���R�[�h
        .SexCD = rs("SexCD")                                                                   '' ���ʃR�[�h
        .HinsyuCD = rs("HinsyuCD")                                                             '' �i��R�[�h
        .KeiroCD = rs("KeiroCD")                                                               '' �ѐF�R�[�h
        .Barei = rs("Barei")                                                                   '' �n��
        .TozaiCD = rs("TozaiCD")                                                               '' ���������R�[�h
        .ChokyosiCode = rs("ChokyosiCode")                                                     '' �����t�R�[�h
        .ChokyosiRyakusyo = rs("ChokyosiRyakusyo")                                             '' �����t������
        .BanusiCode = rs("BanusiCode")                                                         '' �n��R�[�h
        .BanusiName = rs("BanusiName")                                                         '' �n�喼
        .Fukusyoku = rs("Fukusyoku")                                                           '' ���F�W��
        .reserved1 = rs("reserved1")                                                           '' �\��
        .Futan = rs("Futan")                                                                   '' ���S�d��
        .FutanBefore = rs("FutanBefore")                                                       '' �ύX�O���S�d��
        .BLINKER = rs("Blinker")                                                               '' �u�����J�[�g�p�敪
        .reserved2 = rs("reserved2")                                                           '' �\��
        .KisyuCode = rs("KisyuCode")                                                           '' �R��R�[�h
        .KisyuCodeBefore = rs("KisyuCodeBefore")                                               '' �ύX�O�R��R�[�h
        .KisyuRyakusyo = rs("KisyuRyakusyo")                                                   '' �R�薼����
        .KisyuRyakusyoBefore = rs("KisyuRyakusyoBefore")                                       '' �ύX�O�R�薼����
        .MinaraiCD = rs("MinaraiCD")                                                           '' �R�茩�K�R�[�h
        .MinaraiCDBefore = rs("MinaraiCDBefore")                                               '' �ύX�O�R�茩�K�R�[�h
        .BaTaijyu = rs("BaTaijyu")                                                             '' �n�̏d
        .ZogenFugo = rs("ZogenFugo")                                                           '' ��������
        .ZogenSa = rs("ZogenSa")                                                               '' ������
        .IJyoCD = rs("IJyoCD")                                                                 '' �ُ�敪�R�[�h
        .NyusenJyuni = rs("NyusenJyuni")                                                       '' ��������
        .KakuteiJyuni = rs("KakuteiJyuni")                                                     '' �m�蒅��
        .DochakuKubun = rs("DochakuKubun")                                                     '' �����敪
        .DochakuTosu = rs("DochakuTosu")                                                       '' ��������
        .TIME = rs("Time")                                                                     '' ���j�^�C��
        .ChakusaCD = rs("ChakusaCD")                                                           '' �����R�[�h
        .ChakusaCDP = rs("ChakusaCDP")                                                         '' +�����R�[�h
        .ChakusaCDPP = rs("ChakusaCDPP")                                                       '' ++�����R�[�h
        .Jyuni1c = rs("Jyuni1c")                                                               '' 1�R�[�i�[�ł̏���
        .Jyuni2c = rs("Jyuni2c")                                                               '' 2�R�[�i�[�ł̏���
        .Jyuni3c = rs("Jyuni3c")                                                               '' 3�R�[�i�[�ł̏���
        .Jyuni4c = rs("Jyuni4c")                                                               '' 4�R�[�i�[�ł̏���
        .Odds = rs("Odds")                                                                     '' �P���I�b�Y
        .Ninki = rs("Ninki")                                                                   '' �P���l�C��
        .Honsyokin = rs("Honsyokin")                                                           '' �l���{�܋�
        .Fukasyokin = rs("Fukasyokin")                                                         '' �l���t���܋�
        .reserved3 = rs("reserved3")                                                           '' �\��
        .reserved4 = rs("reserved4")                                                           '' �\��
        .HaronTimeL4 = rs("HaronTimeL4")                                                       '' ��S�n�����^�C��
        .HaronTimeL3 = rs("HaronTimeL3")                                                       '' ��R�n�����^�C��
        For i = 0 To 2
            With .ChakuUmaInfo(i)
                .KettoNum = rs("KettoNum" & i + 1)              '' �����o�^�ԍ�                                                    '' �����o�^�ԍ��i����n1�j
                .BAMEI = rs("Bamei" & i + 1)                    '' �n��                                            '' �n��
            End With ' ChakuUmaInfo
        Next i
        .TimeDiff = rs("TimeDiff")                                                             '' �^�C����
        .RecordUpKubun = rs("RecordUpKubun")                                                   '' ���R�[�h�X�V�敪
        .DMKubun = rs("DMKubun")                                                               '' �}�C�j���O�敪
        .DMTime = rs("DMTime")                                                                 '' �}�C�j���O�\�z���j�^�C��
        .DMGosaP = rs("DMGosaP")                                                               '' �\���덷(�M���x)�{
        .DMGosaM = rs("DMGosaM")                                                               '' �\���덷(�M���x)�|
        .DMJyuni = rs("DMJyuni")                                                               '' �}�C�j���O�\�z����
        .KyakusituKubun = rs("KyakusituKubun")                                                 '' ���񃌁[�X�r������
        .CRLF = vbCrLf 'CRLF
    End With ' buf
    
    Exit Sub
ErrorHandler:
    If Err.Number = 3265 Then
        Resume Next
    End If
    gApp.ErrLog
End Sub


'
'   �@�\: �\���̂Ƀf�[�^���Z�b�g����|SE���R�[�h(A)
'
'   ���l: �Ȃ�
'
Public Sub SetDataFromRS_SE_A(ByRef rs As ADODB.Recordset, ByRef buf As JV_SE_RACE_UMA)
On Error GoTo ErrorHandler
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer

    With buf
        With .head
            .RecordSpec = rs("RecordSpec")                                                     '' ���R�[�h���
            .DataKubun = rs("DataKubun")                                                       '' �f�[�^�敪
            With .MakeDate
                .Year = Mid$(rs("MakeDate"), 1, 4)                                               '' �N
                .Month = Mid$(rs("MakeDate"), 5, 2)                                              '' ��
                .Day = Mid$(rs("MakeDate"), 7, 2)                                                '' ��
            End With ' MakeDate
        End With ' head
        With .id
            .Year = rs("Year")                                                                 '' �J�ÔN
            .MonthDay = rs("MonthDay")                                                         '' �J�Ì���
            .JyoCD = rs("JyoCD")                                                               '' ���n��R�[�h
            .Kaiji = rs("Kaiji")                                                               '' �J�É�[��N��]
            .Nichiji = rs("Nichiji")                                                           '' �J�Ó���[N����]
            .RaceNum = rs("RaceNum")                                                           '' ���[�X�ԍ�
        End With ' id
        .Wakuban = rs("Wakuban")                                                               '' �g��
        .Umaban = rs("Umaban")                                                                 '' �n��
        .KettoNum = rs("KettoNum")                                                             '' �����o�^�ԍ�
        .BAMEI = rs("Bamei")                                                                   '' �n��
        .UmaKigoCD = rs("UmaKigoCD")                                                           '' �n�L���R�[�h
        .SexCD = rs("SexCD")                                                                   '' ���ʃR�[�h
        .HinsyuCD = rs("HinsyuCD")                                                             '' �i��R�[�h
        .KeiroCD = rs("KeiroCD")                                                               '' �ѐF�R�[�h
        .Barei = rs("Barei")                                                                   '' �n��
        .TozaiCD = rs("TozaiCD")                                                               '' ���������R�[�h
        .ChokyosiCode = rs("ChokyosiCode")                                                     '' �����t�R�[�h
        .ChokyosiRyakusyo = rs("ChokyosiRyakusyo")                                             '' �����t������
        .BanusiCode = rs("BanusiCode")                                                         '' �n��R�[�h
        .BanusiName = rs("BanusiName")                                                         '' �n�喼
        .Fukusyoku = rs("Fukusyoku")                                                           '' ���F�W��
        .reserved1 = rs("reserved1")                                                           '' �\��

    End With ' buf
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �\���̂Ƀf�[�^���Z�b�g����|SE���R�[�h(B)
'
'   ���l: �Ȃ�
'
Public Sub SetDataFromRS_SE_B(ByRef rs As ADODB.Recordset, ByRef buf As JV_SE_RACE_UMA)
On Error GoTo ErrorHandler
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer

    With buf
        .Futan = rs("Futan")                                                                   '' ���S�d��
        .FutanBefore = rs("FutanBefore")                                                       '' �ύX�O���S�d��
        .BLINKER = rs("Blinker")                                                               '' �u�����J�[�g�p�敪
        .reserved2 = rs("reserved2")                                                           '' �\��
        .KisyuCode = rs("KisyuCode")                                                           '' �R��R�[�h
        .KisyuCodeBefore = rs("KisyuCodeBefore")                                               '' �ύX�O�R��R�[�h
        .KisyuRyakusyo = rs("KisyuRyakusyo")                                                   '' �R�薼����
        .KisyuRyakusyoBefore = rs("KisyuRyakusyoBefore")                                       '' �ύX�O�R�薼����
        .MinaraiCD = rs("MinaraiCD")                                                           '' �R�茩�K�R�[�h
        .MinaraiCDBefore = rs("MinaraiCDBefore")                                               '' �ύX�O�R�茩�K�R�[�h
        .BaTaijyu = rs("BaTaijyu")                                                             '' �n�̏d
        .ZogenFugo = rs("ZogenFugo")                                                           '' ��������
        .ZogenSa = rs("ZogenSa")                                                               '' ������
        .IJyoCD = rs("IJyoCD")                                                                 '' �ُ�敪�R�[�h
        .NyusenJyuni = rs("NyusenJyuni")                                                       '' ��������
        .KakuteiJyuni = rs("KakuteiJyuni")                                                     '' �m�蒅��
        .DochakuKubun = rs("DochakuKubun")                                                     '' �����敪
        .DochakuTosu = rs("DochakuTosu")                                                       '' ��������
        .TIME = rs("Time")                                                                     '' ���j�^�C��
        .ChakusaCD = rs("ChakusaCD")                                                           '' �����R�[�h
        .ChakusaCDP = rs("ChakusaCDP")                                                         '' +�����R�[�h
        .ChakusaCDPP = rs("ChakusaCDPP")                                                       '' ++�����R�[�h
        .Jyuni1c = rs("Jyuni1c")                                                               '' 1�R�[�i�[�ł̏���
        .Jyuni2c = rs("Jyuni2c")                                                               '' 2�R�[�i�[�ł̏���
        .Jyuni3c = rs("Jyuni3c")                                                               '' 3�R�[�i�[�ł̏���
        .Jyuni4c = rs("Jyuni4c")                                                               '' 4�R�[�i�[�ł̏���
        .Odds = rs("Odds")                                                                     '' �P���I�b�Y
        .Ninki = rs("Ninki")                                                                   '' �P���l�C��
        .Honsyokin = rs("Honsyokin")                                                           '' �l���{�܋�
        .Fukasyokin = rs("Fukasyokin")                                                         '' �l���t���܋�
        .reserved3 = rs("reserved3")                                                           '' �\��
        .reserved4 = rs("reserved4")                                                           '' �\��
        .HaronTimeL4 = rs("HaronTimeL4")                                                       '' ��S�n�����^�C��
        .HaronTimeL3 = rs("HaronTimeL3")                                                       '' ��R�n�����^�C��
        For i = 0 To 2
            With .ChakuUmaInfo(i)
                .KettoNum = rs("KettoNum" & i + 1)              '' �����o�^�ԍ�                                                    '' �����o�^�ԍ��i����n1�j
                .BAMEI = rs("Bamei" & i + 1)                    '' �n��                                            '' �n��
            End With ' ChakuUmaInfo
        Next i
        .TimeDiff = rs("TimeDiff")                                                             '' �^�C����
        .RecordUpKubun = rs("RecordUpKubun")                                                   '' ���R�[�h�X�V�敪
        .DMKubun = rs("DMKubun")                                                               '' �}�C�j���O�敪
        .DMTime = rs("DMTime")                                                                 '' �}�C�j���O�\�z���j�^�C��
        .DMGosaP = rs("DMGosaP")                                                               '' �\���덷(�M���x)�{
        .DMGosaM = rs("DMGosaM")                                                               '' �\���덷(�M���x)�|
        .DMJyuni = rs("DMJyuni")                                                               '' �}�C�j���O�\�z����
        .KyakusituKubun = rs("KyakusituKubun")                                                 '' ���񃌁[�X�r������
        .CRLF = vbCrLf 'CRLF
    End With ' buf
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �\���̂Ƀf�[�^���Z�b�g����|SK���R�[�h
'
'   ���l: �Ȃ�
'
Public Sub SetDataFromRS_SK(ByRef rs As ADODB.Recordset, ByRef buf As JV_SK_SANKU)
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer

    With buf
        With .head
            .RecordSpec = rs("RecordSpec")                                                     '' ���R�[�h���
            .DataKubun = rs("DataKubun")                                                       '' �f�[�^�敪
            With .MakeDate
                .Year = Mid$(rs("MakeDate"), 1, 4)                                               '' �N
                .Month = Mid$(rs("MakeDate"), 5, 2)                                              '' ��
                .Day = Mid$(rs("MakeDate"), 7, 2)                                                '' ��
            End With ' MakeDate
        End With ' head
        .KettoNum = rs("KettoNum")                                                             '' �����o�^�ԍ�
        With .BirthDate
            .Year = Mid$(rs("BirthDate"), 1, 4)                                               '' �N
            .Month = Mid$(rs("BirthDate"), 5, 2)                                              '' ��
            .Day = Mid$(rs("BirthDate"), 7, 2)                                                '' ��
        End With ' BirthDate
        .SexCD = rs("SexCD")                                                                   '' ���ʃR�[�h
        .HinsyuCD = rs("HinsyuCD")                                                             '' �i��R�[�h
        .KeiroCD = rs("KeiroCD")                                                               '' �ѐF�R�[�h
        .SankuMochiKubun = rs("SankuMochiKubun")                                               '' �Y����敪
        .ImportYear = rs("ImportYear")                                                         '' �A���N
        .BreederCode = rs("BreederCode")                                                       '' ���Y�҃R�[�h
        .SanchiName = rs("SanchiName")                                                         '' �Y�n��
        .HansyokuNum(0) = rs("FNum")                                                           '' ���ɐB�o�^�ԍ�
        .HansyokuNum(1) = rs("MNum")                                                           '' ��ɐB�o�^�ԍ�
        .HansyokuNum(2) = rs("FFNum")                                                          '' �����ɐB�o�^�ԍ�
        .HansyokuNum(3) = rs("FMNum")                                                          '' ����ɐB�o�^�ԍ�
        .HansyokuNum(4) = rs("MFNum")                                                          '' �ꕃ�ɐB�o�^�ԍ�
        .HansyokuNum(5) = rs("MMNum")                                                          '' ���ɐB�o�^�ԍ�
        .HansyokuNum(6) = rs("FFFNum")                                                         '' �������ɐB�o�^�ԍ�
        .HansyokuNum(7) = rs("FFMNum")                                                         '' ������ɐB�o�^�ԍ�
        .HansyokuNum(8) = rs("FMFNum")                                                         '' ���ꕃ�ɐB�o�^�ԍ�
        .HansyokuNum(9) = rs("FMMNum")                                                         '' �����ɐB�o�^�ԍ�
        .HansyokuNum(10) = rs("MFFNum")                                                        '' �ꕃ���ɐB�o�^�ԍ�
        .HansyokuNum(11) = rs("MFMNum")                                                        '' �ꕃ��ɐB�o�^�ԍ�
        .HansyokuNum(12) = rs("MMFNum")                                                        '' ��ꕃ�ɐB�o�^�ԍ�
        .HansyokuNum(13) = rs("MMMNum")                                                        '' ����ɐB�o�^�ԍ�
        .CRLF = vbCrLf 'CRLF
    End With ' buf
End Sub


'
'   �@�\: �\���̂Ƀf�[�^���Z�b�g����|TK���R�[�h
'
'   ���l: �Ȃ�
'
Public Sub SetDataFromRS_TK(ByRef rs As ADODB.Recordset, ByRef buf As JV_TK_TOKUUMA)
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer

    With buf
        With .head
            .RecordSpec = rs("RecordSpec")                                                      '' ���R�[�h���
            .DataKubun = rs("DataKubun")                                                        '' �f�[�^�敪
            With .MakeDate
                .Year = Mid$(rs("MakeDate"), 1, 4)                                              '' �N
                .Month = Mid$(rs("MakeDate"), 5, 2)                                             '' ��
                .Day = Mid$(rs("MakeDate"), 7, 2)                                               '' ��
            End With ' MakeDate
        End With ' head
        With .id
            .Year = rs("Year")                                                                 '' �J�ÔN
            .MonthDay = rs("MonthDay")                                                         '' �J�Ì���
            .JyoCD = rs("JyoCD")                                                               '' ���n��R�[�h
            .Kaiji = rs("Kaiji")                                                               '' �J�É�[��N��]
            .Nichiji = rs("Nichiji")                                                           '' �J�Ó���[N����]
            .RaceNum = rs("RaceNum")                                                           '' ���[�X�ԍ�
        End With ' id
        With .RaceInfo
            .YoubiCD = rs("YoubiCD")                                                            '' �j���R�[�h
            .TokuNum = rs("TokuNum")                                                            '' ���ʋ����ԍ�
            .Hondai = rs("Hondai")                                                              '' �������{��
            .Fukudai = rs("Fukudai")                                                            '' ����������
            .Kakko = rs("Kakko")                                                                '' �������J�b�R��
            .HondaiEng = rs("HondaiEng")                                                        '' �������{�艢��
            .FukudaiEng = rs("FukudaiEng")                                                      '' ���������艢��
            .KakkoEng = rs("KakkoEng")                                                          '' �������J�b�R������
            .Ryakusyo10 = rs("Ryakusyo10")                                                      '' ���������̂P�O��
            .Ryakusyo6 = rs("Ryakusyo6")                                                        '' ���������̂U��
            .Ryakusyo3 = rs("Ryakusyo3")                                                        '' ���������̂R��
            .Kubun = rs("Kubun")                                                                '' �������敪
            .Nkai = rs("Nkai")                                                                  '' �d�܉�[��N��]
        End With ' RaceInfo
        .GradeCD = rs("GradeCD")                                                                '' �O���[�h�R�[�h
        With .JyokenInfo
            .SyubetuCD = rs("SyubetuCD")                                                        '' ������ʃR�[�h
            .KigoCD = rs("KigoCD")                                                              '' �����L���R�[�h
            .JyuryoCD = rs("JyuryoCD")                                                          '' �d�ʎ�ʃR�[�h
            For j = 0 To 4
                .JyokenCD(j) = rs("JyokenCD" & j + 1)                                          '' ���������R�[�h
            Next j
        End With ' JyokenInfo
        .KYORI = rs("Kyori")                                                                    '' ����
        .TrackCD = rs("TrackCD")                                                                '' �g���b�N�R�[�h
        .CourseKubunCD = rs("CourseKubunCD")                                                    '' �R�[�X�敪
        With .HandiDate
            .Year = Mid$(rs("HandiDate"), 1, 4)                                              '' �N
            .Month = Mid$(rs("HandiDate"), 5, 2)                                             '' ��
            .Day = Mid$(rs("HandiDate"), 7, 2)                                               '' ��
        End With ' HandiDate
        .TorokuTosu = rs("TorokuTosu")                                                          '' �o�^����
        .CRLF = vbCrLf 'CRLF
    End With ' buf

End Sub


'
'   �@�\: �\���̂Ƀf�[�^���Z�b�g����|TK_UMAINFO���R�[�h
'
'   ���l: �Ȃ�
'
Public Sub SetDataFromRS_TK_UMAINFO(ByRef rs As ADODB.Recordset, ByRef buf As JV_TK_TOKUUMA)
    Dim i As Integer
    
    ' �S���ēo�^����
    rs.MoveFirst
    i = 0
    Do While Not rs.EOF
        With buf
            With .TokuUmaInfo(i)
                .num = rs("Num")                                                '' �A��
                .KettoNum = rs("KettoNum")                                      '' �����o�^�ԍ�
                .BAMEI = rs("Bamei")                                            '' �n��
                .UmaKigoCD = rs("UmaKigoCD")                                    '' �n�L���R�[�h
                .SexCD = rs("SexCD")                                            '' ���ʃR�[�h
                .TozaiCD = rs("TozaiCD")                                        '' �����t���������R�[�h
                .ChokyosiCode = rs("ChokyosiCode")                              '' �����t�R�[�h
                .ChokyosiRyakusyo = rs("ChokyosiRyakusyo")                      '' �����t������
                .Futan = rs("Futan")                                            '' ���S�d��
                .Koryu = rs("Koryu")                                            '' �𗬋敪
            End With ' TokuUmaInfo
            rs.MoveNext
        End With
        i = i + 1
    Loop
    
End Sub


'
'   �@�\: �\���̂Ƀf�[�^���Z�b�g����|UM���R�[�h
'
'   ���l: �Ȃ�
'
Public Sub SetDataFromRS_UM(ByRef rs As ADODB.Recordset, ByRef buf As JV_UM_UMA)
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer

    With buf
        With .head
            .RecordSpec = rs("RecordSpec")                                                     '' ���R�[�h���
            .DataKubun = rs("DataKubun")                                                       '' �f�[�^�敪
            With .MakeDate
                .Year = Mid$(rs("MakeDate"), 1, 4)                                              '' �N
                .Month = Mid$(rs("MakeDate"), 5, 2)                                             '' ��
                .Day = Mid$(rs("MakeDate"), 7, 2)                                               '' ��
            End With ' MakeDate
        End With ' head
        .KettoNum = rs("KettoNum")                                                             '' �����o�^�ԍ�
        .DelKubun = rs("DelKubun")                                                             '' �����n�����敪
        With .RegDate
            .Year = Mid$(rs("RegDate"), 1, 4)                                                    '' �N����
            .Month = Mid$(rs("RegDate"), 5, 2)                                                           '' �N����
            .Day = Mid$(rs("RegDate"), 7, 2)                                                             '' �N����
        End With ' RegDate
        With .DelDate
            .Year = Mid$(rs("DelDate"), 1, 4)                                                    '' �N����
            .Month = Mid$(rs("DelDate"), 5, 2)                                                           '' �N����
            .Day = Mid$(rs("DelDate"), 7, 2)                                                             '' �N����
        End With ' DelDate
        With .BirthDate
            .Year = Mid$(rs("BirthDate"), 1, 4)                                                    '' �N����
            .Month = Mid$(rs("BirthDate"), 5, 2)                                                           '' �N����
            .Day = Mid$(rs("BirthDate"), 7, 2)                                                             '' �N����
        End With ' BirthDate
        .BAMEI = rs("Bamei")                                                                   '' �n��
        .BameiKana = rs("BameiKana")                                                           '' �n�����p�J�i
        .BameiEng = rs("BameiEng")                                                             '' �n������
        .UmaKigoCD = rs("UmaKigoCD")                                                           '' �n�L���R�[�h
        .SexCD = rs("SexCD")                                                                   '' ���ʃR�[�h
        .HinsyuCD = rs("HinsyuCD")                                                             '' �i��R�[�h
        .KeiroCD = rs("KeiroCD")                                                               '' �ѐF�R�[�h
        For i = 0 To 13
            With .Ketto3Info(i)
                .HansyokuNum = rs("Ketto3InfoHansyokuNum" & i + 1)                                             '' �ɐB�o�^�ԍ�
                .BAMEI = rs("Ketto3InfoBamei" & i + 1)                                                           '' �n��
            End With ' Ketto3Info
        Next i
        .TozaiCD = rs("TozaiCD")                                                               '' ���������R�[�h
        .ChokyosiCode = rs("ChokyosiCode")                                                     '' �����t�R�[�h
        .ChokyosiRyakusyo = rs("ChokyosiRyakusyo")                                             '' �����t������
        .Syotai = rs("Syotai")                                                                 '' ���Ғn�於
        .BreederCode = rs("BreederCode")                                                       '' ���Y�҃R�[�h
        .BreederName = rs("BreederName")                                                       '' ���Y�Җ�
        .SanchiName = rs("SanchiName")                                                         '' �Y�n��
        .BanusiCode = rs("BanusiCode")                                                         '' �n��R�[�h
        .BanusiName = rs("BanusiName")                                                         '' �n�喼
        .RuikeiHonsyoHeiti = rs("RuikeiHonsyoHeiti")                                           '' ���n�{�܋��݌v
        .RuikeiHonsyoSyogai = rs("RuikeiHonsyoSyogai")                                         '' ��Q�{�܋��݌v
        .RuikeiFukaHeichi = rs("RuikeiFukaHeichi")                                             '' ���n�t���܋��݌v
        .RuikeiFukaSyogai = rs("RuikeiFukaSyogai")                                             '' ��Q�t���܋��݌v
        .RuikeiSyutokuHeichi = rs("RuikeiSyutokuHeichi")                                       '' ���n�����܋��݌v
        .RuikeiSyutokuSyogai = rs("RuikeiSyutokuSyogai")                                       '' ��Q�����܋��݌v
        With .ChakuSogo
            For i = 0 To 5
                .Chakukaisu(i) = rs("SogoChakukaisu" & i + 1)
            Next i
        End With ' ChakuSogo
        With .ChakuChuo
            For i = 0 To 5
                .Chakukaisu(i) = rs("ChuoChakukaisu" & i + 1)
            Next i
        End With ' ChakuChuo
        For i = 0 To 6
            With .ChakuKaisuBa(i)
                For j = 0 To 5
                    .Chakukaisu(j) = rs("Ba" & i + 1 & "Chakukaisu" & j + 1)
                Next j
            End With ' ChakuKaisuBa
        Next i
        For i = 0 To 11
            With .ChakuKaisuJyotai(i)
                For j = 0 To 5
                    .Chakukaisu(j) = rs("Jyotai" & i + 1 & "Chakukaisu" & j + 1)
                Next j
            End With ' ChakuKaisuJyotai
        Next i
        For i = 0 To 5
            With .ChakuKaisuKyori(i)
                For j = 0 To 5
                    .Chakukaisu(j) = rs("Kyori" & i + 1 & "Chakukaisu" & j + 1)
                Next j
            End With ' ChakuKaisuKyori
        Next i
        For i = 0 To 3
            .Kyakusitu(i) = rs("Kyakusitu" & i + 1)                                                    '' �r���X��
        Next i
        .RaceCount = rs("RaceCount")                                                           '' �o�^���[�X��
        .CRLF = vbCrLf 'CRLF
    End With ' buf
End Sub


'
'   �@�\: �\���̂Ƀf�[�^���Z�b�g����|WE���R�[�h
'
'   ���l: �Ȃ�
'
Public Sub SetDataFromRS_WE(ByRef rs As ADODB.Recordset, ByRef buf As JV_WE_WEATHER)
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer

    With buf
        With .head
            .RecordSpec = rs("RecordSpec")                                                     '' ���R�[�h���
            .DataKubun = rs("DataKubun")                                                       '' �f�[�^�敪
            With .MakeDate
                .Year = Mid$(rs("MakeDate"), 1, 4)                                               '' �N
                .Month = Mid$(rs("MakeDate"), 5, 2)                                              '' ��
                .Day = Mid$(rs("MakeDate"), 7, 2)                                                '' ��
            End With ' MakeDate
        End With ' head
        With .id
            .Year = rs("Year")                                                                 '' �J�ÔN
            .MonthDay = rs("MonthDay")                                                         '' �J�Ì���
            .JyoCD = rs("JyoCD")                                                               '' ���n��R�[�h
            .Kaiji = rs("Kaiji")                                                               '' �J�É�[��N��]
            .Nichiji = rs("Nichiji")                                                           '' �J�Ó���[N����]
        End With ' id
        With .HappyoTime
            .Month = Mid$(rs("HappyoTime"), 1, 2)                                               '' ��
            .Day = Mid$(rs("HappyoTime"), 3, 2)                                                 '' ��
            .Hour = Mid$(rs("HappyoTime"), 5, 2)                                                '' ��
            .Minute = Mid$(rs("HappyoTime"), 7, 2)                                              '' ��
        End With ' HappyoTime
        .HenkoID = rs("HenkoID")                                                               '' �ύX����
        With .TenkoBaba
            .TenkoCD = rs("AtoTenkoCD")                                                           '' �V��R�[�h
            .SibaBabaCD = rs("AtoSibaBabaCD")                                                     '' �Ŕn���ԃR�[�h
            .DirtBabaCD = rs("AtoDirtBabaCD")                                                     '' �_�[�g�n���ԃR�[�h
        End With ' TenkoBaba
        With .TenkoBabaBefore
            .TenkoCD = rs("MaeTenkoCD")                                                           '' �V��R�[�h
            .SibaBabaCD = rs("MaeSibaBabaCD")                                                     '' �Ŕn���ԃR�[�h
            .DirtBabaCD = rs("MaeDirtBabaCD")                                                     '' �_�[�g�n���ԃR�[�h
        End With ' TenkoBabaBefore
        .CRLF = vbCrLf 'CRLF
    End With ' buf
End Sub


'
'   �@�\: �\���̂Ƀf�[�^���Z�b�g����|WH���R�[�h
'
'   ���l: �Ȃ�
'
Public Sub SetDataFromRS_WH(ByRef rs As ADODB.Recordset, ByRef buf As JV_WH_BATAIJYU)
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer

    With buf
        With .head
            .RecordSpec = rs("RecordSpec")                                                     '' ���R�[�h���
            .DataKubun = rs("DataKubun")                                                       '' �f�[�^�敪
            With .MakeDate
                .Year = Mid$(rs("MakeDate"), 1, 4)                                               '' �N
                .Month = Mid$(rs("MakeDate"), 5, 2)                                              '' ��
                .Day = Mid$(rs("MakeDate"), 7, 2)                                                '' ��
            End With ' MakeDate
        End With ' head
        With .id
            .Year = rs("Year")                                                                 '' �J�ÔN
            .MonthDay = rs("MonthDay")                                                         '' �J�Ì���
            .JyoCD = rs("JyoCD")                                                               '' ���n��R�[�h
            .Kaiji = rs("Kaiji")                                                               '' �J�É�[��N��]
            .Nichiji = rs("Nichiji")                                                           '' �J�Ó���[N����]
            .RaceNum = rs("RaceNum")                                                           '' ���[�X�ԍ�
        End With ' id
        With .HappyoTime
            .Month = Mid$(rs("HappyoTime"), 1, 2)                                               '' ��
            .Day = Mid$(rs("HappyoTime"), 3, 2)                                                 '' ��
            .Hour = Mid$(rs("HappyoTime"), 5, 2)                                                '' ��
            .Minute = Mid$(rs("HappyoTime"), 7, 2)                                              '' ��
        End With ' HappyoTime
        For i = 0 To 17
            With .BataijyuInfo(i)
                .Umaban = rs("Umaban" & i + 1)                                            '' �n��
                .BAMEI = rs("Bamei" & i + 1)                                              '' �n��
                .BaTaijyu = rs("BaTaijyu" & i + 1)                                        '' �n�̏d
                .ZogenFugo = rs("ZogenFugo" & i + 1)                                      '' ��������
                .ZogenSa = rs("ZogenSa" & i + 1)                                          '' ������
            End With ' BataijyuInfo
        Next i
        .CRLF = vbCrLf 'CRLF
    End With ' buf
End Sub


'
'   �@�\: �\���̂Ƀf�[�^���Z�b�g����|YS���R�[�h
'
'   ���l: �Ȃ�
'
Public Sub SetDataFromRS_YS(ByRef rs As ADODB.Recordset, ByRef buf As JV_YS_SCHEDULE)
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer

    With buf
        With .head
            .RecordSpec = rs("RecordSpec")                                                     '' ���R�[�h���
            .DataKubun = rs("DataKubun")                                                       '' �f�[�^�敪
            With .MakeDate
                .Year = Mid$(rs("MakeDate"), 1, 4)                                               '' �N
                .Month = Mid$(rs("MakeDate"), 5, 2)                                              '' ��
                .Day = Mid$(rs("MakeDate"), 7, 2)                                                '' ��
            End With ' MakeDate
        End With ' head
        With .id
            .Year = rs("Year")                                                                 '' �J�ÔN
            .MonthDay = rs("MonthDay")                                                         '' �J�Ì���
            .JyoCD = rs("JyoCD")                                                               '' ���n��R�[�h
            .Kaiji = rs("Kaiji")                                                               '' �J�É�[��N��]
            .Nichiji = rs("Nichiji")                                                           '' �J�Ó���[N����]
        End With ' id
        .YoubiCD = rs("YoubiCD")                                                               '' �j���R�[�h
        With .JyusyoInfo(0)
            .TokuNum = rs("Jyusyo1TokuNum")                                                    '' ���ʋ����ԍ�
            .Hondai = rs("Jyusyo1Hondai")                                                      '' �������{��
            .Ryakusyo10 = rs("Jyusyo1Ryakusyo10")                                              '' ����������10��
            .Ryakusyo6 = rs("Jyusyo1Ryakusyo6")                                                '' ����������6��
            .Ryakusyo3 = rs("Jyusyo1Ryakusyo3")                                                '' ����������3��
            .Nkai = rs("Jyusyo1Nkai")                                                          '' �d�܉�[��N��]
            .GradeCD = rs("Jyusyo1GradeCD")                                                    '' �O���[�h�R�[�h
            .SyubetuCD = rs("Jyusyo1SyubetuCD")                                                '' ������ʃR�[�h
            .KigoCD = rs("Jyusyo1KigoCD")                                                      '' �����L���R�[�h
            .JyuryoCD = rs("Jyusyo1JyuryoCD")                                                  '' �d�ʎ�ʃR�[�h
            .KYORI = rs("Jyusyo1Kyori")                                                        '' ����
            .TrackCD = rs("Jyusyo1TrackCD")                                                    '' �g���b�N�R�[�h
        End With ' JyusyoInfo(0)
        With .JyusyoInfo(1)
            .TokuNum = rs("Jyusyo2TokuNum")                                                    '' ���ʋ����ԍ�
            .Hondai = rs("Jyusyo2Hondai")                                                      '' �������{��
            .Ryakusyo10 = rs("Jyusyo2Ryakusyo10")                                              '' ����������10��
            .Ryakusyo6 = rs("Jyusyo2Ryakusyo6")                                                '' ����������6��
            .Ryakusyo3 = rs("Jyusyo2Ryakusyo3")                                                '' ����������3��
            .Nkai = rs("Jyusyo2Nkai")                                                          '' �d�܉�[��N��]
            .GradeCD = rs("Jyusyo2GradeCD")                                                    '' �O���[�h�R�[�h
            .SyubetuCD = rs("Jyusyo2SyubetuCD")                                                '' ������ʃR�[�h
            .KigoCD = rs("Jyusyo2KigoCD")                                                      '' �����L���R�[�h
            .JyuryoCD = rs("Jyusyo2JyuryoCD")                                                  '' �d�ʎ�ʃR�[�h
            .KYORI = rs("Jyusyo2Kyori")                                                        '' ����
            .TrackCD = rs("Jyusyo2TrackCD")                                                    '' �g���b�N�R�[�h
        End With ' JyusyoInfo(1)
        With .JyusyoInfo(2)
            .TokuNum = rs("Jyusyo3TokuNum")                                                    '' ���ʋ����ԍ�
            .Hondai = rs("Jyusyo3Hondai")                                                      '' �������{��
            .Ryakusyo10 = rs("Jyusyo3Ryakusyo10")                                              '' ����������10��
            .Ryakusyo6 = rs("Jyusyo3Ryakusyo6")                                                '' ����������6��
            .Ryakusyo3 = rs("Jyusyo3Ryakusyo3")                                                '' ����������3��
            .Nkai = rs("Jyusyo3Nkai")                                                          '' �d�܉�[��N��]
            .GradeCD = rs("Jyusyo3GradeCD")                                                    '' �O���[�h�R�[�h
            .SyubetuCD = rs("Jyusyo3SyubetuCD")                                                '' ������ʃR�[�h
            .KigoCD = rs("Jyusyo3KigoCD")                                                      '' �����L���R�[�h
            .JyuryoCD = rs("Jyusyo3JyuryoCD")                                                  '' �d�ʎ�ʃR�[�h
            .KYORI = rs("Jyusyo3Kyori")                                                        '' ����
            .TrackCD = rs("Jyusyo3TrackCD")                                                    '' �g���b�N�R�[�h
        End With ' JyusyoInfo(2)
        .CRLF = vbCrLf 'CRLF
    End With ' buf
End Sub


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �����֐�
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   �@�\: ���R�[�h�Z�b�g���玫�����쐬����
'
'   ���l: �Ȃ�
'
Private Sub MakeDic(ByRef dic As Dictionary, _
    rs As ADODB.Recordset, _
    field As String, _
    numBlocks As Long, _
    keyLen As Long, _
    blkLen As Long)

    Dim i As Long
    Dim p As Long
    Dim buf As String

    If IsNull(rs(field)) Then
        Exit Sub
    End If
    
    buf = rs(field)
    
    p = 1
    For i = 0 To numBlocks - 1
        If p > Len(buf) Then Exit For
        Call dic.Add(Mid$(buf, p, keyLen), Mid$(buf, p + keyLen, blkLen - keyLen))
        p = p + blkLen
    Next i

End Sub
