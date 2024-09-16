Attribute VB_Name = "JVLink_Stluct"
Option Explicit
Option Base 0


    '========================================================================
    '  JRA-VAN Data Lab. JV-Data�\���� Ver1.0.5��
    '
    '
    '   �쐬: JRA-VAN �\�t�g�E�F�A�H�[  2003�N6��3��
    '
    '========================================================================
    '   (C) Copyright Turf Media System Co.,Ltd. 2003 All rights reserved
    '========================================================================
    '   Ver1.0.5�������ɏC���������Ă���܂��B

    '''''''''''''''''''' ���ʍ\���� ''''''''''''''''''''

    '<�N����>
    Private Type YMD
        Year   As String                    ''�N
        Month  As String                    ''��
        Day    As String                    ''��
    End Type


    '<�����b>
    Private Type HMS
        Hour   As String                    ''��
        Minute As String                    ''��
        Second As String                    ''�b
    End Type


    '<����>
    Private Type HM
        Hour As String                      ''��
        Minute As String                    ''��
    End Type


    '<��������>
    Private Type MDHM
        Month As String                     ''��
        Day As String                       ''��
        Hour As String                      ''��
        Minute As String                    ''��
    End Type


    '<���R�[�h�w�b�_>
    Private Type RECORD_ID
        RecordSpec As String                ''���R�[�h���
        DataKubun As String                 ''�f�[�^�敪
        MakeDate As YMD                     ''�f�[�^�쐬�N����
    End Type


    '<�������ʏ��P>
    Private Type RACE_ID
        Year As String                      ''�J�ÔN
        MonthDay As String                  ''�J�Ì���
        JyoCD As String                     ''���n��R�[�h
        Kaiji As String                     ''�J�É�[��N��]
        Nichiji As String                   ''�J�Ó���[N����]
        RaceNum As String                   ''���[�X�ԍ�
    End Type


    '<�������ʏ��Q>
    Private Type RACE_ID2
        Year As String                      ''�J�ÔN
        MonthDay As String                  ''�J�Ì���
        JyoCD As String                     ''���n��R�[�h
        Kaiji As String                     ''�J�É�[��N��]
        Nichiji As String                   ''�J�Ó���[N����]
    End Type


    '<���񐔁i�T�C�Y3byte�j>
    Private Type CHAKUKAISU3_INFO
        Chakukaisu(5) As String
    End Type


    '<���񐔁i�T�C�Y6byte�j>
    Private Type CHAKUKAISU6_INFO
        Chakukaisu(5) As String
    End Type


    '<�{�N�E�݌v���я��>
    Private Type SEI_RUIKEI_INFO
        SetYear As String                   ''�ݒ�N
        HonSyokinTotal As String            ''�{�܋����v
        Fukasyokin As String                ''�t���܋����v
        Chakukaisu(5) As String             ''����
    End Type


    '<�ŋߏd�܏������>
    Private Type SAIKIN_JYUSYO_INFO
        SaikinJyusyoid As RACE_ID           ''<�N��������R>
        Hondai As String                    ''�������{��
        Ryakusyo10 As String                ''����������10��
        Ryakusyo6 As String                 ''����������6��
        Ryakusyo3 As String                 ''����������3��
        GradeCD As String                   ''�O���[�h�R�[�h
        SyussoTosu As String                ''�o������
        KettoNum As String                  ''�����o�^�ԍ�
        BAMEI As String                     ''�n��
    End Type


    '<�{�N�E�O�N�E�݌v���я��>
    Private Type HON_ZEN_RUIKEISEI_INFO
        SetYear As String                          ''�ݒ�N
        HonSyokinHeichi As String                  ''���n�{�܋����v
        HonSyokinSyogai As String                  ''��Q�{�܋����v
        FukaSyokinHeichi As String                 ''���n�t���܋����v
        FukaSyokinSyogai As String                 ''��Q�t���܋����v
        ChakuKaisuHeichi As CHAKUKAISU6_INFO       ''���n����
        ChakuKaisuSyogai As CHAKUKAISU6_INFO       ''��Q����
        ChakuKaisuJyo(19) As CHAKUKAISU6_INFO      ''���n��ʒ���
        ChakuKaisuKyori(5) As CHAKUKAISU6_INFO     ''�����ʒ���
    End Type


    '<���[�X���>
    Private Type RACE_INFO
        YoubiCD As String                   ''�j���R�[�h
        TokuNum As String                   ''���ʋ����ԍ�
        Hondai As String                    ''�������{��
        Fukudai As String                   ''����������
        Kakko As String                     ''�������J�b�R��
        HondaiEng As String                 ''�������{�艢��
        FukudaiEng As String                ''���������艢��
        KakkoEng As String                  ''�������J�b�R������
        Ryakusyo10 As String                ''���������̂P�O��
        Ryakusyo6 As String                 ''���������̂U��
        Ryakusyo3 As String                 ''���������̂R��
        Kubun As String                     ''�������敪
        Nkai As String                      ''�d�܉�[��N��]
    End Type


    '<�V��E�n����>
    Private Type TENKO_BABA_INFO
        TenkoCD As String                   ''�V��R�[�h
        SibaBabaCD As String                ''�Ŕn���ԃR�[�h
        DirtBabaCD As String                ''�_�[�g�n���ԃR�[�h
    End Type


    '<���������R�[�h>
    Private Type RACE_JYOKEN
        SyubetuCD As String                 ''������ʃR�[�h
        KigoCD As String                    ''�����L���R�[�h
        JyuryoCD As String                  ''�d�ʎ�ʃR�[�h
        JyokenCD(4) As String               ''���������R�[�h
    End Type


    '''''''''''''''''''' �f�[�^�\���� ''''''''''''''''''''

   '****** �P�D���ʓo�^�n ****************************************
    
    '<�o�^�n�����>
    Private Type TOKUUMA_INFO
        num As String                       ''�A��
        KettoNum As String                  ''�����o�^�ԍ�
        BAMEI As String                     ''�n��
        UmaKigoCD As String                 ''�n�L���R�[�h
        SexCD As String                     ''���ʃR�[�h
        TozaiCD As String                   ''�����t���������R�[�h
        ChokyosiCode As String              ''�����t�R�[�h
        ChokyosiRyakusyo As String          ''�����t������
        Futan As String                     ''���S�d��
        Koryu As String                     ''�𗬋敪
    End Type

    Public Type JV_TK_TOKUUMA
        head As RECORD_ID                   ''<���R�[�h�w�b�_�[>
        id As RACE_ID                       ''<�������ʏ��P>
        RaceInfo As RACE_INFO               ''<���[�X���>
        GradeCD As String                   ''�O���[�h�R�[�h
        JyokenInfo As RACE_JYOKEN           ''<���������R�[�h>
        KYORI As String                     ''����
        TrackCD As String                   ''�g���b�N�R�[�h
        CourseKubunCD As String             ''�R�[�X�敪
        HandiDate As YMD                    ''�n���f���\��
        TorokuTosu As String                ''�o�^����
        TokuUmaInfo(299) As TOKUUMA_INFO    ''<�o�^�n�����>
        CRLF As String                      ''���R�[�h���
        
    End Type

    '****** �Q�D���[�X�ڍ� ****************************************

    '<�R�[�i�[�ʉߏ���>
    Private Type CORNER_INFO
        Corner As String                    ''�R�[�i�[
        Syukaisu As String                  ''����
        Jyuni As String                    ''�e�ʉߏ���
       
    End Type

    Public Type JV_RA_RACE
        head As RECORD_ID                   ''<���R�[�h�w�b�_�[>
        id As RACE_ID                       ''<�������ʏ��P>
        RaceInfo As RACE_INFO               ''<���[�X���>
        GradeCD As String                   ''�O���[�h�R�[�h
        GradeCDBefore As String             ''�ύX�O�O���[�h�R�[�h
        JyokenInfo As RACE_JYOKEN           ''<���������R�[�h>
        JyokenName As String                ''������������
        KYORI As String                     ''����
        KyoriBefore As String               ''�ύX�O����
        TrackCD As String                    ''�g���b�N�R�[�h
        TrackCDBefore As String             ''�ύX�O�g���b�N�R�[�h
        CourseKubunCD As String             ''�R�[�X�敪
        CourseKubunCDBefore As String       ''�ύX�O�R�[�X�敪
        Honsyokin(6) As String              ''�{�܋�
        HonsyokinBefore(4) As String        ''�ύX�O�{�܋�
        Fukasyokin(4) As String             ''�t���܋�
        FukasyokinBefore(2) As String       ''�ύX�O�t���܋�
        HassoTime As String                 ''��������
        HassoTimeBefore As String           ''�ύX�O��������
        TorokuTosu As String                ''�o�^����
        SyussoTosu As String                ''�o������
        NyusenTosu As String                ''��������
        TenkoBaba As TENKO_BABA_INFO        ''�V��E�n���ԃR�[�h
        LapTime(24) As String               ''���b�v�^�C��
        SyogaiMileTime As String            ''��Q�}�C���^�C��
        HaronTimeS3 As String               ''�O�R�n�����^�C��
        HaronTimeS4 As String               ''�O�S�n�����^�C��
        HaronTimeL3 As String               ''��R�n�����^�C��
        HaronTimeL4 As String               ''��S�n�����^�C��
        CornerInfo(3) As CORNER_INFO        ''<�R�[�i�[�ʉߏ���>
        RecordUpKubun As String             ''���R�[�h�X�V�敪
        CRLF As String                      ''���R�[�h��؂�
    End Type


    '****** �R�D�n�����[�X��� ****************************************

    '<1���n(����n)���>
    Private Type CHAKUUMA_INFO
        KettoNum As String                  ''�����o�^�ԍ�
        BAMEI As String                     ''�n��
    End Type

    Public Type JV_SE_RACE_UMA
        head As RECORD_ID                   ''<���R�[�h�w�b�_�[>
        id As RACE_ID                       ''<�������ʏ��P>
        Wakuban As String                   ''�g��
        Umaban As String                    ''�n��
        KettoNum As String                  ''�����o�^�ԍ�
        BAMEI As String                     ''�n��
        UmaKigoCD As String                 ''�n�L���R�[�h
        SexCD As String                     ''���ʃR�[�h
        HinsyuCD As String                  ''�i��R�[�h
        KeiroCD As String                   ''�ѐF�R�[�h
        Barei As String                     ''�n��
        TozaiCD As String                   ''���������R�[�h
        ChokyosiCode As String              ''�����t�R�[�h
        ChokyosiRyakusyo As String          ''�����t������
        BanusiCode As String                ''�n��R�[�h
        BanusiName As String                ''�n�喼
        Fukusyoku As String                 ''���F�W��
        reserved1 As String                 ''�\��
        Futan As String                     ''���S�d��
        FutanBefore As String               ''�ύX�O���S�d��
        BLINKER As String                   ''�u�����J�[�g�p�敪
        reserved2 As String                 ''�\��
        KisyuCode As String                 ''�R��R�[�h
        KisyuCodeBefore As String           ''�ύX�O�R��R�[�h
        KisyuRyakusyo As String             ''�R�薼����
        KisyuRyakusyoBefore As String       ''�ύX�O�R�薼����
        MinaraiCD As String                 ''�R�茩�K�R�[�h
        MinaraiCDBefore As String           ''�ύX�O�R�茩�K�R�[�h
        BaTaijyu As String                  ''�n�̏d
        ZogenFugo As String                 ''��������
        ZogenSa As String                   ''������
        IJyoCD As String                    ''�ُ�敪�R�[�h
        NyusenJyuni As String               ''��������
        KakuteiJyuni As String              ''�m�蒅��
        DochakuKubun As String              ''�����敪
        DochakuTosu As String               ''��������
        TIME As String                      ''���j�^�C��
        ChakusaCD As String                 ''�����R�[�h
        ChakusaCDP As String                ''+�����R�[�h
        ChakusaCDPP As String               ''++�����R�[�h
        Jyuni1c As String                   ''1�R�[�i�[�ł̏���
        Jyuni2c As String                   ''2�R�[�i�[�ł̏���
        Jyuni3c As String                   ''3�R�[�i�[�ł̏���
        Jyuni4c As String                   ''4�R�[�i�[�ł̏���
        Odds As String                      ''�P���I�b�Y
        Ninki As String                     ''�P���l�C��
        Honsyokin As String                 ''�l���{�܋�
        Fukasyokin As String                ''�l���t���܋�
        reserved3 As String                 ''�\��
        reserved4 As String                 ''�\��
        HaronTimeL4 As String               ''��S�n�����^�C��
        HaronTimeL3 As String               ''��R�n�����^�C��
        ChakuUmaInfo(2) As CHAKUUMA_INFO    ''<1���n(����n)���>
        TimeDiff As String                  ''�^�C����
        RecordUpKubun As String             ''���R�[�h�X�V�敪
        DMKubun As String                   ''�}�C�j���O�敪
        DMTime As String                    ''�}�C�j���O�\�z���j�^�C��
        DMGosaP As String                   ''�\���덷(�M���x)�{
        DMGosaM As String                   ''�\���덷(�M���x)�|
        DMJyuni As String                   ''�}�C�j���O�\�z����
        KyakusituKubun As String            ''���񃌁[�X�r������
        CRLF As String                      ''���R�[�h��؂�
    End Type


    '****** �S�D���� ****************************************

    '<���ߏ��P �P�E���E�g>
    Private Type PAY_INFO1
        Umaban As String                    ''�n��
        Pay As String                       ''���ߋ�
        Ninki As String                     ''�l�C��
    End Type

    '<���ߏ��Q �n�A�E���C�h�E�\���E�n�P>
    Private Type PAY_INFO2
        Kumi As String                      ''�g��
        Pay As String                       ''���ߋ�
        Ninki As String                     ''�l�C��
    End Type

    '<���ߏ��R �R�A��>
    Private Type PAY_INFO3
        Kumi As String                      ''�g��
        Pay As String                       ''���ߋ�
        Ninki As String                     ''�l�C��
    End Type

    '<���ߏ��S �\��>
    Private Type PAY_INFO4
        Kumi As String                      ''�g��
        Pay As String                       ''���ߋ�
        Ninki As String                     ''�l�C��
    End Type

    Public Type JV_HR_PAY
        head As RECORD_ID                   ''<���R�[�h�w�b�_�[>
        id As RACE_ID                       ''<�������ʏ��P>
        TorokuTosu As String                ''�o�^����
        SyussoTosu As String                ''�o������
        FuseirituFlag(8) As String          ''�s�����t���O
        TokubaraiFlag(8) As String          ''�����t���O
        HenkanFlag(8) As String             ''�Ԋ҃t���O
        HenkanUma(27) As String             ''�ԊҔn�ԏ��(�n��01�`28)
        HenkanWaku(7) As String             ''�ԊҘg�ԏ��(�g��1�`8)
        HenkanDoWaku(7) As String           ''�Ԋғ��g���(�g��1�`8)
        PayTansyo(2) As PAY_INFO1           ''<�P������>
        PayFukusyo(4) As PAY_INFO1          ''<��������>
        PayWakuren(2) As PAY_INFO1          ''<�g�A����>
        PayUmaren(2) As PAY_INFO2           ''<�n�A����>
        PayWide(6) As PAY_INFO2             ''<���C�h����>
        PayReserved1(2) As PAY_INFO2        ''<�\��>
        PayUmatan(5) As PAY_INFO2           ''<�n�P����>
        PaySanrenpuku(2) As PAY_INFO3       ''<3�A������>
'        PayReserved2(5) As PAY_INFO4        ''<�\��>
        PaySanrentan(5) As PAY_INFO4        ''<3�A�P����>
        CRLF As String                      ''���R�[�h��؂�
    End Type


    '****** �T�D�[���i�S�|���j****************************************

    '<�[�����P �P�E���E�g>
    Private Type HYO_INFO1
        Umaban As String                    ''�n��
        Hyo As String                       ''�[��
        Ninki As String                     ''�l�C
    End Type

    '<�[�����Q �n�A�E���C�h�E�n�P>
    Private Type HYO_INFO2
        Kumi As String                      ''�g��
        Hyo As String                       ''�[��
        Ninki As String                     ''�l�C
    End Type

    '<�[�����R �R�A���[��>
    Private Type HYO_INFO3
        Kumi As String                      ''�g��
        Hyo As String                       ''�[��
        Ninki As String                     ''�l�C
    End Type

    '<�[�����S �\��>
    Private Type HYO_INFO4
        Kumi As String                      ''�g��
        Hyo As String                       ''�[��
        Ninki As String                     ''�l�C
    End Type

    Public Type JV_H1_HYOSU_ZENKAKE
        head As RECORD_ID                   ''<���R�[�h�w�b�_�[>
        id As RACE_ID                       ''<�������ʏ��P>
        TorokuTosu As String                ''�o�^����
        SyussoTosu As String                ''�o������
        HatubaiFlag(6) As String            ''�����t���O
        FukuChakuBaraiKey As String         ''���������L�[
        HenkanUma(27) As String             ''�ԊҔn�ԏ��(�n��01�`28)
        HenkanWaku(7) As String             ''�ԊҘg�ԏ��(�g��1�`8)
        HenkanDoWaku(7) As String           ''�Ԋғ��g���(�g��1�`8)
        HyoTansyo(27) As HYO_INFO1          ''<�P���[��>
        HyoFukusyo(27) As HYO_INFO1         ''<�����[��>
        HyoWakuren(35) As HYO_INFO1         ''<�g�A�[��>
        HyoUmaren(152) As HYO_INFO2         ''<�n�A�[��>
        HyoWide(152) As HYO_INFO2           ''<���C�h�[��>
        HyoUmatan(305) As HYO_INFO2         ''<�n�P�[��>
        HyoSanrenpuku(815) As HYO_INFO3     ''<3�A���[��>
        HyoTotal(13) As String              ''�[�����v
        CRLF As String                      ''���R�[�h��؂�
    End Type
    
    
    '****** �T.�`. �[���U�i�R�A�P�j****************************************
    
    '<3�A�P�[��>
    Private Type HYO_INFO
        Kumi As String                      ''�g��
        Hyo As String                       ''�[��
        Ninki As String                     ''�l�C
    End Type
    
    Public Type JV_H6_HYOSU_SANRENTAN
        head As RECORD_ID                   ''<���R�[�h�w�b�_�[>
        id As RACE_ID                       ''<�������ʏ��P>
        TorokuTosu As String                ''�o�^����
        SyussoTosu As String                ''�o������
        HatubaiFlag As String               ''�����t���O 3�A�P
        HenkanUma(17) As String             ''�ԊҔn�ԏ��(�n��01�`18)
        HyoSanrentan(4895) As HYO_INFO      ''<3�A�P�[��>
        TotalHyoSanrentan As String         ''3�A�P�[�����v
        TotalHyoSanrentanHenkan As String   ''3�A�P�Ԋҕ[�����v
        CRLF As String                      ''���R�[�h��؂�
    End Type
    
    
    '****** �U�D�I�b�Y�i�P���g�j****************************************

    '<�P���I�b�Y>
    Private Type ODDS_TANSYO_INFO
        Umaban As String                    ''�n��
        Odds As String                      ''�I�b�Y
        Ninki As String                     ''�l�C��
    End Type

    '<�����I�b�Y>
    Private Type ODDS_FUKUSYO_INFO
        Umaban As String                    ''�n��
        OddsLow As String                   ''�Œ�I�b�Y
        OddsHigh As String                  ''�ō��I�b�Y
        Ninki As String                     ''�l�C��
    End Type

    '<�g�A�I�b�Y>
    Private Type ODDS_WAKUREN_INFO
        Kumi As String                      ''�g
        Odds As String                      ''�I�b�Y
        Ninki As String                     ''�l�C��
    End Type

    Public Type JV_O1_ODDS_TANFUKUWAKU
        head As RECORD_ID                   ''<���R�[�h�w�b�_�[>
        id As RACE_ID                       ''<�������ʏ��P>
        HappyoTime As MDHM                  ''���\��������
        TorokuTosu As String                ''�o�^����
        SyussoTosu As String                ''�o������
        TansyoFlag As String                ''�����t���O �P��
        FukusyoFlag As String               ''�����t���O ����
        WakurenFlag As String               ''�����t���O�@�g�A
        FukuChakuBaraiKey As String         ''���������L�[
        OddsTansyoInfo(27) As ODDS_TANSYO_INFO    ''<�P���I�b�Y>
        OddsFukusyoInfo(27) As ODDS_FUKUSYO_INFO  ''<�����[���I�b�Y>
        OddsWakurenInfo(35) As ODDS_WAKUREN_INFO  ''<�g�A�[���I�b�Y>
        TotalHyosuTansyo As String                ''�P���[�����v
        TotalHyosuFukusyo As String         ''�����[�����v
        TotalHyosuWakuren As String         ''�g�A�[�����v
        CRLF As String                      ''���R�[�h��؂�
    End Type


    '****** �V�D�I�b�Y�i�n�A�j****************************************

    '<�n�A�I�b�Y>
    Private Type ODDS_UMAREN_INFO
        Kumi As String                      ''�g��
        Odds As String                      ''�I�b�Y
        Ninki As String                     ''�l�C��
    End Type

    Public Type JV_O2_ODDS_UMAREN
        head As RECORD_ID                   ''<���R�[�h�w�b�_�[>
        id As RACE_ID                       ''<�������ʏ��P>
        HappyoTime As MDHM                  ''���\��������
        TorokuTosu As String                ''�o�^����
        SyussoTosu As String                ''�o������
        UmarenFlag As String                ''�����t���O�@�n�A
        OddsUmarenInfo(152) As ODDS_UMAREN_INFO   ''<�n�A�I�b�Y>
        TotalHyosuUmaren As String          ''�n�A�[�����v
        CRLF As String                      ''���R�[�h��؂�
    End Type


    '****** �W�D�I�b�Y�i���C�h�j****************************************

    '<���C�h�I�b�Y>
    Private Type ODDS_WIDE_INFO
        Kumi As String                      ''�g��
        OddsLow As String                   ''�Œ�I�b�Y
        OddsHigh As String                  ''�ō��I�b�Y
        Ninki As String                     ''�l�C��
    End Type

    Public Type JV_O3_ODDS_WIDE
        head As RECORD_ID                   ''<���R�[�h�w�b�_�[>
        id As RACE_ID                       ''<�������ʏ��P>
        HappyoTime As MDHM                  ''���\��������
        TorokuTosu As String                ''�o�^����
        SyussoTosu As String                ''�o������
        WideFlag As String                  ''�����t���O�@���C�h
        OddsWideInfo(152) As ODDS_WIDE_INFO ''<���C�h�I�b�Y>
        TotalHyosuWide As String            ''���C�h�[�����v
        CRLF As String                      ''���R�[�h��؂�
    End Type


    '****** �X�D�I�b�Y�i�n�P�j ****************************************

    '<�n�P�I�b�Y>
    Private Type ODDS_UMATAN_INFO
        Kumi As String                      ''�g��
        Odds As String                      ''�I�b�Y
        Ninki As String                     ''�l�C��
    End Type

    Public Type JV_O4_ODDS_UMATAN
        head As RECORD_ID                   ''<���R�[�h�w�b�_�[>
        id As RACE_ID                       ''<�������ʏ��P>
        HappyoTime As MDHM                  ''���\��������
        TorokuTosu As String                ''�o�^����
        SyussoTosu As String                ''�o������
        UmatanFlag As String                ''�����t���O�@�n�P
        OddsUmatanInfo(305) As ODDS_UMATAN_INFO ''<�n�P�I�b�Y>
        TotalHyosuUmatan As String          ''�n�P�[�����v
        CRLF As String                      ''���R�[�h��؂�
    End Type


    '****** �P�O�D�I�b�Y�i�R�A���j***************************************

    '<3�A���I�b�Y>
    Private Type ODDS_SANREN_INFO
        Kumi As String                      ''�g��
        Odds As String                      ''�I�b�Y
        Ninki As String                     ''�l�C��
    End Type

    Public Type JV_O5_ODDS_SANREN
        head As RECORD_ID                   ''<���R�[�h�w�b�_�[>
        id As RACE_ID                       ''<�������ʏ��P>
        HappyoTime As MDHM                  ''���\��������
        TorokuTosu As String                ''�o�^����
        SyussoTosu As String                ''�o������
        SanrenpukuFlag As String            ''�����t���O�@3�A��
        OddsSanrenInfo(815) As ODDS_SANREN_INFO ''<3�A���I�b�Y>
        TotalHyosuSanrenpuku As String          ''3�A���[�����v
        CRLF As String                          ''���R�[�h��؂�
    End Type
    
    
    '****** �P�O.�`.�@�I�b�Y�i�R�A�P�j***************************************
    
    '<3�A�P�I�b�Y>
    Private Type ODDS_SANRENTAN_INFO
        Kumi As String                      ''�g��
        Odds As String                      ''�I�b�Y
        Ninki As String                     ''�l�C��
    End Type
    
    Public Type JV_O6_ODDS_SANRENTAN
        head As RECORD_ID                   ''<���R�[�h�w�b�_�[>
        id As RACE_ID                       ''<�������ʏ��P>
        HappyoTime As MDHM                  ''���\��������
        TorokuTosu As String                ''�o�^����
        SyussoTosu As String                ''�o������
        SanrentanFlag As String             ''�����t���O�@3�A�P
        OddsSanrentanInfo(4895) As ODDS_SANRENTAN_INFO ''<3�A�P�I�b�Y>
        TotalHyosuSanrentan As String       ''3�A�P�[�����v
        CRLF As String                      ''���R�[�h��؂�
    End Type
    

    '****** �P�P�D�����n�}�X�^ ****************************************

    '<�R�㌌�����>
    Private Type KETTO3_INFO
        HansyokuNum As String               ''�ɐB�o�^�ԍ�
        BAMEI As String                     ''�n��
    End Type

    Public Type JV_UM_UMA
        head As RECORD_ID                   ''<���R�[�h�w�b�_�[>
        KettoNum As String                  ''�����o�^�ԍ�
        DelKubun As String                  ''�����n�����敪
        RegDate As YMD                      ''�����n�o�^�N����
        DelDate As YMD                      ''�����n�����N����
        BirthDate As YMD                    ''���N����
        BAMEI As String                     ''�n��
        BameiKana As String                 ''�n�����p�J�i
        BameiEng As String                  ''�n������
        UmaKigoCD As String                 ''�n�L���R�[�h
        SexCD As String                     ''���ʃR�[�h
        HinsyuCD As String                  ''�i��R�[�h
        KeiroCD As String                   ''�ѐF�R�[�h
        Ketto3Info(13) As KETTO3_INFO       ''<3�㌌�����>
        TozaiCD As String                   ''���������R�[�h
        ChokyosiCode As String              ''�����t�R�[�h
        ChokyosiRyakusyo As String          ''�����t������
        Syotai As String                    ''���Ғn�於
        BreederCode As String               ''���Y�҃R�[�h
        BreederName As String              ''���Y�Җ�
        SanchiName As String                ''�Y�n��
        BanusiCode As String                ''�n��R�[�h
        BanusiName As String                ''�n�喼
        RuikeiHonsyoHeiti As String         ''���n�{�܋��݌v
        RuikeiHonsyoSyogai As String        ''��Q�{�܋��݌v
        RuikeiFukaHeichi As String          ''���n�t���܋��݌v
        RuikeiFukaSyogai As String          ''��Q�t���܋��݌v
        RuikeiSyutokuHeichi As String       ''���n�����܋��݌v
        RuikeiSyutokuSyogai As String       ''��Q�����܋��݌v
        ChakuSogo As CHAKUKAISU3_INFO       ''��������
        ChakuChuo As CHAKUKAISU3_INFO       ''�������v����
        ChakuKaisuBa(6) As CHAKUKAISU3_INFO ''�n��ʒ���
        ChakuKaisuJyotai(11) As CHAKUKAISU3_INFO      ''�n���ԕʒ���
        ChakuKaisuKyori(5) As CHAKUKAISU3_INFO        ''�����ʒ���
        Kyakusitu(3) As String              ''�r���X��
        RaceCount As String                 ''�o�^���[�X��
        CRLF As String                      ''���R�[�h��؂�
    End Type


    '****** �P�Q�D�R��}�X�^ ****************************************

    '<���R����>
    Private Type HATUKIJYO_INFO
        Hatukijyoid As RACE_ID              ''�N��������R
        SyussoTosu As String                ''�o������
        KettoNum As String                  ''�����o�^�ԍ�
        BAMEI As String                     ''�n��
        KakuteiJyuni As String              ''�m�蒅��
        IJyoCD As String                    ''�ُ�敪�R�[�h
    End Type

    '<���������>
    Private Type HATUSYORI_INFO
        Hatusyoriid As RACE_ID              ''�N��������R
        SyussoTosu As String                ''�o������
        KettoNum As String                  ''�����o�^�ԍ�
        BAMEI As String                     ''�n��
    End Type

    Public Type JV_KS_KISYU
        head As RECORD_ID                   ''<���R�[�h�w�b�_�[>
        KisyuCode As String                 ''�R��R�[�h
        DelKubun As String                  ''�R�薕���敪
        IssueDate As YMD                    ''�R��Ƌ���t�N����
        DelDate As YMD                      ''�R��Ƌ������N����
        BirthDate As YMD                    ''���N����
        KisyuName As String                 ''�R�薼����
        Reserved As String                  ''�\��
        KisyuNameKana As String             ''�R�薼���p�J�i
        KisyuRyakusyo As String             ''�R�薼����
        KisyuNameEng As String              ''�R�薼����
        SexCD As String                     ''���ʋ敪
        SikakuCD As String                  ''�R�掑�i�R�[�h
        MinaraiCD As String                 ''�R�茩�K�R�[�h
        TozaiCD As String                   ''�R�蓌�������R�[�h
        Syotai As String                    ''���Ғn�於
        ChokyosiCode As String              ''���������t�R�[�h
        ChokyosiRyakusyo As String          ''���������t������
        HatuKiJyo(1) As HATUKIJYO_INFO      ''<���R����>
        HatuSyori(1) As HATUSYORI_INFO      ''<���������>
        SaikinJyusyo(2) As SAIKIN_JYUSYO_INFO     ''<�ŋߏd�܏������>
        HonZenRuikei(2) As HON_ZEN_RUIKEISEI_INFO ''<�{�N�E�O�N�E�݌v���я��>
        CRLF As String                           ''���R�[�h��؂�
    End Type


    '****** �P�R�D�����t�}�X�^ ****************************************

    Public Type JV_CH_CHOKYOSI
        head As RECORD_ID                   ''<���R�[�h�w�b�_�[>
        ChokyosiCode As String              ''�����t�R�[�h
        DelKubun As String                  ''�����t�����敪
        IssueDate As YMD                    ''�����t�Ƌ���t�N����
        DelDate As YMD                      ''�����t�Ƌ������N����
        BirthDate As YMD                    ''���N����
        ChokyosiName As String              ''�����t������
        ChokyosiNameKana As String          ''�����t�����p�J�i
        ChokyosiRyakusyo As String          ''�����t������
        ChokyosiNameEng As String           ''�����t������
        SexCD As String                     ''���ʋ敪
        TozaiCD As String                   ''�����t���������R�[�h
        Syotai As String                    ''���Ғn�於
        SaikinJyusyo(2) As SAIKIN_JYUSYO_INFO     ''<�ŋߏd�܏������>
        HonZenRuikei(2) As HON_ZEN_RUIKEISEI_INFO ''<�{�N�E�O�N�E�݌v���я��>
        CRLF As String                      ''���R�[�h��؂�
    End Type


    '******�P�S�D���Y�҃}�X�^ ****************************************

    Public Type JV_BR_BREEDER
        head As RECORD_ID                   ''<���R�[�h�w�b�_�[>
        BreederCode As String               ''���Y�҃R�[�h
        BreederName_Co As String            ''���Y�Җ�(�@�l�i�L)
        BreederName As String               ''���Y�Җ�(�@�l�i��)
        BreederNameKana As String           ''���Y�Җ����p�J�i
        BreederNameEng As String            ''���Y�Җ�����
        Address As String                   ''���Y�ҏZ�������Ȗ�
        HonRuikei(1) As SEI_RUIKEI_INFO     ''<�{�N�E�݌v���я��>
        CRLF As String                      ''���R�[�h��؂�
    End Type


    '****** �P�T�D�n��}�X�^ ****************************************

    Public Type JV_BN_BANUSI
        head As RECORD_ID                   ''<���R�[�h�w�b�_�[>
        BanusiCode As String                ''�n��R�[�h
        BanusiName_Co As String             ''�n�喼(�@�l�i�L)
        BanusiName As String                ''�n�喼(�@�l�i��)
        BanusiNameKana As String            ''�n�喼���p�J�i
        BanusiNameEng As String             ''�n�喼����
        Fukusyoku As String                 ''���F�W��
        HonRuikei(1) As SEI_RUIKEI_INFO     ''<�{�N�E�݌v���я��>
        CRLF As String                      ''���R�[�h��؂�
    End Type


    '****** �P�U�D�ɐB�n�}�X�^ ****************************************

    Public Type JV_HN_HANSYOKU
        head As RECORD_ID                   ''<���R�[�h�w�b�_�[>
        HansyokuNum As String               ''�ɐB�o�^�ԍ�
        Reserved As String                  ''�\��
        KettoNum As String                  ''�����o�^�ԍ�
        DelKubun As String                  ''�ɐB�n�����敪
        BAMEI As String                     ''�n��
        BameiKana As String                 ''�n�����p�J�i
        BameiEng As String                  ''�n������
        BirthYear As String                 ''���N
        SexCD As String                     ''���ʃR�[�h
        HinsyuCD As String                  ''�i��R�[�h
        KeiroCD As String                   ''�ѐF�R�[�h
        HansyokuMochiKubun As String        ''�ɐB�n�����敪
        ImportYear As String                ''�A���N
        SanchiName As String                ''�Y�n��
        HansyokuFNum As String              ''���n�ɐB�o�^�ԍ�
        HansyokuMNum As String              ''��n�ɐB�o�^�ԍ�
        CRLF As String                      ''���R�[�h��؂�
    End Type


    '****** �P�V�D�Y��}�X�^ ****************************************

    Public Type JV_SK_SANKU
        head As RECORD_ID                   ''<���R�[�h�w�b�_�[>
        KettoNum As String                  ''�����o�^�ԍ�
        BirthDate As YMD                    ''���N����
        SexCD As String                     ''���ʃR�[�h
        HinsyuCD As String                  ''�i��R�[�h
        KeiroCD As String                   ''�ѐF�R�[�h
        SankuMochiKubun As String           ''�Y����敪
        ImportYear As String                ''�A���N
        BreederCode As String               ''���Y�҃R�[�h
        SanchiName As String                ''�Y�n��
        HansyokuNum(13) As String           ''3�㌌�� �ɐB�o�^�ԍ�
        CRLF As String                      ''���R�[�h��؂�
    End Type


    '****** �P�W�D���R�[�h�}�X�^ ****************************************

    '<���R�[�h�ێ��n���>
    Private Type RECUMA_INFO
        KettoNum As String                  ''�����o�^�ԍ�
        BAMEI As String                     ''�n��
        UmaKigoCD As String                 ''�n�L���R�[�h
        SexCD As String                     ''���ʃR�[�h
        ChokyosiCode As String              ''�����t�R�[�h
        ChokyosiName As String              ''�����t��
        Futan As String                     ''���S�d��
        KisyuCode As String                 ''�R��R�[�h
        KisyuName As String                 ''�R�薼
    End Type

    Public Type JV_RC_RECORD
        head As RECORD_ID                   ''<���R�[�h�w�b�_�[>
        RecInfoKubun As String              ''���R�[�h���ʋ敪
        id As RACE_ID                       ''<�������ʏ��P>
        TokuNum As String                   ''���ʋ����ԍ�
        Hondai As String                    ''�������{��
        GradeCD As String                   ''�O���[�h�R�[�h
        SyubetuCD As String                 ''������ʃR�[�h
        KYORI As String                     ''����
        TrackCD As String                   ''�g���b�N�R�[�h
        RecKubun As String                  ''���R�[�h�敪
        RecTime As String                   ''���R�[�h�^�C��
        TenkoBaba As TENKO_BABA_INFO        ''�V��E�n����
        RecUmaInfo(2) As RECUMA_INFO        ''<���R�[�h�ێ��n���>
        CRLF As String                      ''���R�[�h��؂�
    End Type


    '****** �P�X�D��H���� ****************************************

    Public Type JV_HC_HANRO
        head As RECORD_ID                   ''<���R�[�h�w�b�_�[>
        TresenKubun As String               ''�g���Z���敪
        ChokyoDate As YMD                   ''�����N����
        ChokyoTime As String                ''��������
        KettoNum As String                  ''�����o�^�ԍ�
        HaronTime4 As String                ''4�n�����^�C�����v(800M-0M)
        LapTime4 As String                  ''���b�v�^�C��(800M-600M)
        HaronTime3 As String                ''3�n�����^�C�����v(600M-0M)
        LapTime3 As String                  ''���b�v�^�C��(600M-400M)
        HaronTime2 As String                ''2�n�����^�C�����v(400M-0M)
        LapTime2 As String                  ''���b�v�^�C��(400M-200M)
        LapTime1 As String                  ''���b�v�^�C��(200M-0M)
        CRLF As String                      ''���R�[�h��؂�
    End Type


    '****** �Q�O�D�n�̏d ****************************************

    '<�n�̏d���>
    Private Type BATAIJYU_INFO
        Umaban As String                    ''�n��
        BAMEI As String                     ''�n��
        BaTaijyu As String                  ''�n�̏d
        ZogenFugo As String                 ''��������
        ZogenSa As String                   ''������
    End Type

    Public Type JV_WH_BATAIJYU
        head As RECORD_ID                   ''<���R�[�h�w�b�_�[>
        id As RACE_ID                       ''<�������ʏ��P>
        HappyoTime As MDHM                  ''���\��������
        BataijyuInfo(17) As BATAIJYU_INFO   ''<�n�̏d���>
        CRLF As String                      ''���R�[�h��؂�
    End Type


    '****** �Q�P�D�V��n���� ******************************************

    Public Type JV_WE_WEATHER
        head As RECORD_ID                   ''<���R�[�h�w�b�_�[>
        id As RACE_ID2                      ''<�������ʏ��Q>
        HappyoTime As MDHM                  ''���\��������
        HenkoID As String                   ''�ύX����
        TenkoBaba As TENKO_BABA_INFO        ''���ݏ�ԏ��
        TenkoBabaBefore As TENKO_BABA_INFO  ''�ύX�O��ԏ��
        CRLF As String                      ''���R�[�h��؂�
       
    End Type

    '****** �Q�Q�D�o������E�������O ****************************************

    Public Type JV_AV_INFO
        head As RECORD_ID                   ''<���R�[�h�w�b�_�[>
        id As RACE_ID                       ''<�������ʏ��P>
        HappyoTime As MDHM                  ''���\��������
        Umaban As String                    ''�n��
        BAMEI As String                     ''�n��
        JiyuKubun As String                 ''���R�敪
        CRLF As String                      ''���R�[�h��؂�
      
    End Type

    '************ �Q�R�D�R��ύX ****************************************

    '<�ύX���>
    Private Type JC_INFO
        Futan As String                     ''���S�d��
        KisyuCode As String                 ''�R��R�[�h
        KisyuName As String                 ''�R�薼
        MinaraiCD As String                 ''�R�茩�K�R�[�h
       
    End Type

    Public Type JV_JC_INFO
        head As RECORD_ID                   ''<���R�[�h�w�b�_�[>
        id As RACE_ID                       ''<�������ʏ��P>
        HappyoTime As MDHM                  ''���\��������
        Umaban As String                    ''�n��
        BAMEI As String                     ''�n��
        JCInfoAfter As JC_INFO              ''<�ύX����>
        JCInfoBefore As JC_INFO             ''<�ύX�O���>
        CRLF As String                      ''���R�[�h��؂�
    End Type


    '****** �Q�S�D�f�[�^�}�C�j���O�\�z***********************************

    '<�}�C�j���O�\�z>
    Private Type DM_INFO
        Umaban As String                    ''�n��
        DMTime As String                    ''�\�z���j�^�C��
        DMGosaP As String                   ''�\�z�덷(�M���x)�{
        DMGosaM As String                   ''�\�z�덷(�M���x)�|
    End Type

    Public Type JV_DM_INFO
        head As RECORD_ID                   ''<���R�[�h�w�b�_�[>
        id As RACE_ID                       ''<�������ʏ��P>
        MakeHM As HM                        ''�f�[�^�쐬����
        DMInfo(17) As DM_INFO               ''<�}�C�j���O�\�z>
        CRLF As String                      ''���R�[�h��؂�
    End Type


    '****** �Q�T�D�J�ÃX�P�W���[��************************************

    '<�d�܈ē�>
    Private Type JYUSYO_INFO
        TokuNum As String                   ''���ʋ����ԍ�
        Hondai As String                    ''�������{��
        Ryakusyo10 As String                ''����������10��
        Ryakusyo6 As String                 ''����������6��
        Ryakusyo3 As String                 ''����������3��
        Nkai As String                      ''�d�܉�[��N��]
        GradeCD As String                   ''�O���[�h�R�[�h
        SyubetuCD As String                 ''������ʃR�[�h
        KigoCD As String                    ''�����L���R�[�h
        JyuryoCD As String                  ''�d�ʎ�ʃR�[�h
        KYORI As String                     ''����
        TrackCD As String                   ''�g���b�N�R�[�h
    End Type

    Public Type JV_YS_SCHEDULE
        head As RECORD_ID                   ''<���R�[�h�w�b�_�[>
        id As RACE_ID2                      ''<�������ʏ��Q>
        YoubiCD As String                   ''�j���R�[�h
        JyusyoInfo(2) As JYUSYO_INFO        ''<�d�܈ē�>
        CRLF As String                      ''���R�[�h��؂�
    End Type
    
    '****** �Q�U�D���������ύX************************************

    Public Type JV_TC_HASSOU
        head As RECORD_ID                   ''<���R�[�h�w�b�_�[>
        id As RACE_ID                       ''<�������ʏ��P>
        HappyoTime As MDHM                  ''���\��������
        AtoHassoTime As HM                  ''�ύX�㎞��
        MaeHassoTime As HM                  ''�ύX�O����
        CRLF As String                      ''���R�[�h��؂�
    End Type


    '****** �Q�V�D�R�[�X�ύX************************************

    Public Type JV_CC_COURSE
        head As RECORD_ID                   ''<���R�[�h�w�b�_�[>
        id As RACE_ID                      ''<�������ʏ��P>
        HappyoTime As MDHM                  ''���\��������
        AtoKyori As String                  ''�ύX�㋗��
        AtoTrackCD As String                ''�ύX��g���b�N�R�[�h
        MaeKyori As String                  ''�ύX�O����
        MaeTrackCD As String                ''�ύX�O�g���b�N�R�[�h
        JiyuKubun As String                 ''���R�R�[�h
        CRLF As String                      ''���R�[�h��؂�
    End Type
    
    
    
    '''''''''''''''''''' �f�[�^�Z�b�g�֐� '''''''''''''''''''''''''''
    
   '****** �P�D���ʓo�^�n ****************************************
    
    Public Sub SetData_TK(ByRef lBUf As String, ByRef mBuf As JV_TK_TOKUUMA)
    Dim bytBuf() As Byte                            '' Byte��ŏ������邽�߂̃o�b�t�@
    Dim i As Integer                                '' ���[�v�J�E���^
    Dim j As Integer                                '' ���[�v�J�E���^
    Dim k As Integer                                '' ���[�v�J�E���^
    Dim p As Long                                   '' �؂蕪���J�n�ʒu
    
    bytBuf = StrConv(lBUf, vbFromUnicode)
    
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

    '�o�b�t�@�̈���
    Erase bytBuf
       
    End Sub

    '****** �Q�D���[�X�ڍ� ****************************************
    Public Sub SetData_RA(ByRef lBUf As String, ByRef mBuf As JV_RA_RACE)
    Dim bytBuf() As Byte                            '' Byte��ŏ������邽�߂̃o�b�t�@
    Dim i As Integer                                '' ���[�v�J�E���^
    Dim j As Integer                                '' ���[�v�J�E���^
    Dim k As Integer                                '' ���[�v�J�E���^
    Dim p As Long                                   '' �؂蕪���J�n�ʒu
    
    bytBuf = StrConv(lBUf, vbFromUnicode)
    
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

    '�o�b�t�@�̈���
    Erase bytBuf
     
    End Sub


    '****** �R�D�n�����[�X��� ****************************************

    Public Sub SetData_SE(ByRef lBUf As String, ByRef mBuf As JV_SE_RACE_UMA)
    Dim bytBuf() As Byte                            '' Byte��ŏ������邽�߂̃o�b�t�@
    Dim i As Integer                                '' ���[�v�J�E���^
    Dim p As Long                                   '' �؂蕪���J�n�ʒu
    
    bytBuf = StrConv(lBUf, vbFromUnicode)
    
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


    '****** �S�D���� ****************************************

    Public Sub SetData_HR(lBUf As String, ByRef mBuf As JV_HR_PAY)
    Dim i As Integer                                '' ���[�v�J�E���^
    Dim j As Integer                                '' ���[�v�J�E���^
    Dim k As Integer                                '' ���[�v�J�E���^
    Dim p As Long                                   '' �؂蕪���J�n�ʒu
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMid(lBUf, p, 2)          '' ���R�[�h���
            .DataKubun = IncMid(lBUf, p, 1)           '' �f�[�^�敪
            With .MakeDate
                .Year = IncMid(lBUf, p, 4)            '' �N
                .Month = IncMid(lBUf, p, 2)           '' ��
                .Day = IncMid(lBUf, p, 2)             '' ��
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMid(lBUf, p, 4)                '' �J�ÔN
            .MonthDay = IncMid(lBUf, p, 4)            '' �J�Ì���
            .JyoCD = IncMid(lBUf, p, 2)               '' ���n��R�[�h
            .Kaiji = IncMid(lBUf, p, 2)               '' �J�É�[��N��]
            .Nichiji = IncMid(lBUf, p, 2)             '' �J�Ó���[N����]
            .RaceNum = IncMid(lBUf, p, 2)             '' ���[�X�ԍ�
        End With ' id
        .TorokuTosu = IncMid(lBUf, p, 2)              '' �o�^����
        .SyussoTosu = IncMid(lBUf, p, 2)              '' �o������
        For i = 0 To 8
            .FuseirituFlag(i) = IncMid(lBUf, p, 1)    '' �s�����t���O
        Next i
        For i = 0 To 8
            .TokubaraiFlag(i) = IncMid(lBUf, p, 1)    '' �����t���O
        Next i
        For i = 0 To 8
            .HenkanFlag(i) = IncMid(lBUf, p, 1)       '' �Ԋ҃t���O
        Next i
        For i = 0 To 27
            .HenkanUma(i) = IncMid(lBUf, p, 1)        '' �ԊҔn�ԏ��(�n��01�`28)
        Next i
        For i = 0 To 7
            .HenkanWaku(i) = IncMid(lBUf, p, 1)       '' �ԊҘg�ԏ��(�g��1�`8)
        Next i
        For i = 0 To 7
            .HenkanDoWaku(i) = IncMid(lBUf, p, 1)     '' �Ԋғ��g���(�g��1�`8)
        Next i
        For i = 0 To 2
            With .PayTansyo(i)
                .Umaban = IncMid(lBUf, p, 2)          '' �n��
                .Pay = IncMid(lBUf, p, 9)             '' ���ߋ�
                .Ninki = IncMid(lBUf, p, 2)           '' �l�C��
            End With ' PayTansyo
        Next i
        For i = 0 To 4
            With .PayFukusyo(i)
                .Umaban = IncMid(lBUf, p, 2)          '' �n��
                .Pay = IncMid(lBUf, p, 9)             '' ���ߋ�
                .Ninki = IncMid(lBUf, p, 2)           '' �l�C��
            End With ' PayFukusyo
        Next i
        For i = 0 To 2
            With .PayWakuren(i)
                .Umaban = IncMid(lBUf, p, 2)          '' �n��
                .Pay = IncMid(lBUf, p, 9)             '' ���ߋ�
                .Ninki = IncMid(lBUf, p, 2)           '' �l�C��
            End With ' PayWakuren
        Next i
        For i = 0 To 2
            With .PayUmaren(i)
                .Kumi = IncMid(lBUf, p, 4)            '' �g��
                .Pay = IncMid(lBUf, p, 9)             '' ���ߋ�
                .Ninki = IncMid(lBUf, p, 3)           '' �l�C��
            End With ' PayUmaren
        Next i
        For i = 0 To 6
            With .PayWide(i)
                .Kumi = IncMid(lBUf, p, 4)            '' �g��
                .Pay = IncMid(lBUf, p, 9)             '' ���ߋ�
                .Ninki = IncMid(lBUf, p, 3)           '' �l�C��
            End With ' PayWide
        Next i
        For i = 0 To 2
            With .PayReserved1(i)
                .Kumi = IncMid(lBUf, p, 4)            '' �g��
                .Pay = IncMid(lBUf, p, 9)             '' ���ߋ�
                .Ninki = IncMid(lBUf, p, 3)           '' �l�C��
            End With ' PayReserved1
        Next i
        For i = 0 To 5
            With .PayUmatan(i)
                .Kumi = IncMid(lBUf, p, 4)            '' �g��
                .Pay = IncMid(lBUf, p, 9)             '' ���ߋ�
                .Ninki = IncMid(lBUf, p, 3)           '' �l�C��
            End With ' PayUmatan
        Next i
        For i = 0 To 2
            With .PaySanrenpuku(i)
                .Kumi = IncMid(lBUf, p, 6)            '' �g��
                .Pay = IncMid(lBUf, p, 9)             '' ���ߋ�
                .Ninki = IncMid(lBUf, p, 3)           '' �l�C��
            End With ' PaySanrenpuku
        Next i
        For i = 0 To 5
            With .PaySanrentan(i)
                .Kumi = IncMid(lBUf, p, 6)            '' �g��
                .Pay = IncMid(lBUf, p, 9)             '' ���ߋ�
                .Ninki = IncMid(lBUf, p, 4)           '' �l�C��
            End With ' PaySanrentan
        Next i
        .CRLF = IncMid(lBUf, p, 2)        '' ���R�[�h��؂�
    End With
   
    End Sub


    '****** �T�D�[���i�S�|���j****************************************

    Public Sub SetData_H1(lBUf As String, ByRef mBuf As JV_H1_HYOSU_ZENKAKE)
    Dim i As Integer                                '' ���[�v�J�E���^
    Dim j As Integer                                '' ���[�v�J�E���^
    Dim k As Integer                                '' ���[�v�J�E���^
    Dim p As Long                                   '' �؂蕪���J�n�ʒu
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMid(lBUf, p, 2)      '' ���R�[�h���
            .DataKubun = IncMid(lBUf, p, 1)       '' �f�[�^�敪
            With .MakeDate
                .Year = IncMid(lBUf, p, 4)        '' �N
                .Month = IncMid(lBUf, p, 2)       '' ��
                .Day = IncMid(lBUf, p, 2)         '' ��
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMid(lBUf, p, 4)            '' �J�ÔN
            .MonthDay = IncMid(lBUf, p, 4)        '' �J�Ì���
            .JyoCD = IncMid(lBUf, p, 2)           '' ���n��R�[�h
            .Kaiji = IncMid(lBUf, p, 2)           '' �J�É�[��N��]
            .Nichiji = IncMid(lBUf, p, 2)         '' �J�Ó���[N����]
            .RaceNum = IncMid(lBUf, p, 2)         '' ���[�X�ԍ�
        End With ' id
        .TorokuTosu = IncMid(lBUf, p, 2)          '' �o�^����
        .SyussoTosu = IncMid(lBUf, p, 2)          '' �o������
        For i = 0 To 6
            .HatubaiFlag(i) = IncMid(lBUf, p, 1)  '' �����t���O
        Next i
        .FukuChakuBaraiKey = IncMid(lBUf, p, 1)   '' ���������L�[
        For i = 0 To 27
            .HenkanUma(i) = IncMid(lBUf, p, 1)    '' �ԊҔn�ԏ��(�n��01�`28)
        Next i
        For i = 0 To 7
            .HenkanWaku(i) = IncMid(lBUf, p, 1)   '' �ԊҘg�ԏ��(�g��1�`8)
        Next i
        For i = 0 To 7
            .HenkanDoWaku(i) = IncMid(lBUf, p, 1) '' �Ԋғ��g���(�g��1�`8)
        Next i
        For i = 0 To 27
            With .HyoTansyo(i)
                .Umaban = IncMid(lBUf, p, 2)      '' �n��
                .Hyo = IncMid(lBUf, p, 11)        '' �[��
                .Ninki = IncMid(lBUf, p, 2)       '' �l�C
            End With ' HyoTansyo
        Next i
        For i = 0 To 27
            With .HyoFukusyo(i)
                .Umaban = IncMid(lBUf, p, 2)      '' �n��
                .Hyo = IncMid(lBUf, p, 11)        '' �[��
                .Ninki = IncMid(lBUf, p, 2)       '' �l�C
            End With ' HyoFukusyo
        Next i
        For i = 0 To 35
            With .HyoWakuren(i)
                .Umaban = IncMid(lBUf, p, 2)      '' �n��
                .Hyo = IncMid(lBUf, p, 11)        '' �[��
                .Ninki = IncMid(lBUf, p, 2)       '' �l�C
            End With ' HyoWakuren
        Next i
        For i = 0 To 152
            With .HyoUmaren(i)
                .Kumi = IncMid(lBUf, p, 4)        '' �g��
                .Hyo = IncMid(lBUf, p, 11)        '' �[��
                .Ninki = IncMid(lBUf, p, 3)       '' �l�C
            End With ' HyoUmaren
        Next i
        For i = 0 To 152
            With .HyoWide(i)
                .Kumi = IncMid(lBUf, p, 4)        '' �g��
                .Hyo = IncMid(lBUf, p, 11)        '' �[��
                .Ninki = IncMid(lBUf, p, 3)       '' �l�C
            End With ' HyoWide
        Next i
        For i = 0 To 305
            With .HyoUmatan(i)
                .Kumi = IncMid(lBUf, p, 4)        '' �g��
                .Hyo = IncMid(lBUf, p, 11)        '' �[��
                .Ninki = IncMid(lBUf, p, 3)       '' �l�C
            End With ' HyoUmatan
        Next i
        For i = 0 To 815
            With .HyoSanrenpuku(i)
                .Kumi = IncMid(lBUf, p, 6)        '' �g��
                .Hyo = IncMid(lBUf, p, 11)        '' �[��
                .Ninki = IncMid(lBUf, p, 3)       '' �l�C
            End With ' HyoSanrenpuku
        Next i
        For i = 0 To 13
            .HyoTotal(i) = IncMid(lBUf, p, 11)    '' �[�����v
        Next i
        .CRLF = IncMid(lBUf, p, 2)                '' ���R�[�h��؂�
    End With
    
    End Sub


    '****** �T.�`. �[���U�i�R�A�P�j****************************************
    Public Sub SetData_H6(lBUf As String, ByRef mBuf As JV_H6_HYOSU_SANRENTAN)
    Dim i As Integer                                '' ���[�v�J�E���^
    Dim p As Long                                   '' �؂蕪���J�n�ʒu
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMid(lBUf, p, 2)      '' ���R�[�h���
            .DataKubun = IncMid(lBUf, p, 1)       '' �f�[�^�敪
            With .MakeDate
                .Year = IncMid(lBUf, p, 4)        '' �N
                .Month = IncMid(lBUf, p, 2)       '' ��
                .Day = IncMid(lBUf, p, 2)         '' ��
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMid(lBUf, p, 4)            '' �J�ÔN
            .MonthDay = IncMid(lBUf, p, 4)        '' �J�Ì���
            .JyoCD = IncMid(lBUf, p, 2)           '' ���n��R�[�h
            .Kaiji = IncMid(lBUf, p, 2)           '' �J�É�[��N��]
            .Nichiji = IncMid(lBUf, p, 2)         '' �J�Ó���[N����]
            .RaceNum = IncMid(lBUf, p, 2)         '' ���[�X�ԍ�
        End With ' id
        .TorokuTosu = IncMid(lBUf, p, 2)          '' �o�^����
        .SyussoTosu = IncMid(lBUf, p, 2)          '' �o������
        .HatubaiFlag = IncMid(lBUf, p, 1)         '' �����t���O 3�A�P
        For i = 0 To 17
            .HenkanUma(i) = IncMid(lBUf, p, 1)    '' �ԊҔn�ԏ��(�n��01�`18)
        Next i
        For i = 0 To 4895
            With .HyoSanrentan(i)
                .Kumi = IncMid(lBUf, p, 6)        '' �g��
                .Hyo = IncMid(lBUf, p, 11)        '' �[��
                .Ninki = IncMid(lBUf, p, 4)       '' �l�C
            End With ' HyoSanrentan
        Next i
        .TotalHyoSanrentan = IncMid(lBUf, p, 11)    '' 3�A�P�[�����v
        .TotalHyoSanrentanHenkan = IncMid(lBUf, p, 11) '' 3�A�P�Ԋҕ[�����v
        .CRLF = IncMid(lBUf, p, 2)                  '' ���R�[�h��؂�
    End With
    End Sub
    
    
    '****** �U�D�I�b�Y�i�P���g�j****************************************

    Public Sub SetData_O1(lBUf As String, ByRef mBuf As JV_O1_ODDS_TANFUKUWAKU)
    Dim i As Integer                                '' ���[�v�J�E���^
    Dim j As Integer                                '' ���[�v�J�E���^
    Dim k As Integer                                '' ���[�v�J�E���^
    Dim p As Long                                   '' �؂蕪���J�n�ʒu
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMid(lBUf, p, 2)      '' ���R�[�h���
            .DataKubun = IncMid(lBUf, p, 1)       '' �f�[�^�敪
            With .MakeDate
                .Year = IncMid(lBUf, p, 4)        '' �N
                .Month = IncMid(lBUf, p, 2)       '' ��
                .Day = IncMid(lBUf, p, 2) '' ��
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMid(lBUf, p, 4)            '' �J�ÔN
            .MonthDay = IncMid(lBUf, p, 4)        '' �J�Ì���
            .JyoCD = IncMid(lBUf, p, 2)           '' ���n��R�[�h
            .Kaiji = IncMid(lBUf, p, 2)           '' �J�É�[��N��]
            .Nichiji = IncMid(lBUf, p, 2)         '' �J�Ó���[N����]
            .RaceNum = IncMid(lBUf, p, 2)         '' ���[�X�ԍ�
        End With ' id
        With .HappyoTime
            .Month = IncMid(lBUf, p, 2)           '' ��
            .Day = IncMid(lBUf, p, 2)             '' ��
            .Hour = IncMid(lBUf, p, 2)            '' ��
            .Minute = IncMid(lBUf, p, 2)          '' ��
        End With ' HappyoTime
        .TorokuTosu = IncMid(lBUf, p, 2)          '' �o�^����
        .SyussoTosu = IncMid(lBUf, p, 2)          '' �o������
        .TansyoFlag = IncMid(lBUf, p, 1)          '' �����t���O
        .FukusyoFlag = IncMid(lBUf, p, 1)         '' �����t���O
        .WakurenFlag = IncMid(lBUf, p, 1)         '' �����t���O�@�g�A
        .FukuChakuBaraiKey = IncMid(lBUf, p, 1)   '' ���������L�[
        For i = 0 To 27
            With .OddsTansyoInfo(i)
                .Umaban = IncMid(lBUf, p, 2)      '' �n��
                .Odds = IncMid(lBUf, p, 4)        '' �I�b�Y
                .Ninki = IncMid(lBUf, p, 2)       '' �l�C��
            End With ' OddsTansyoInfo
        Next i
        For i = 0 To 27
            With .OddsFukusyoInfo(i)
                .Umaban = IncMid(lBUf, p, 2)      '' �n��
                .OddsLow = IncMid(lBUf, p, 4)     '' �Œ�I�b�Y
                .OddsHigh = IncMid(lBUf, p, 4)    '' �ō��I�b�Y
                .Ninki = IncMid(lBUf, p, 2)       '' �l�C��
            End With ' OddsFukusyoInfo
        Next i
        For i = 0 To 35
            With .OddsWakurenInfo(i)
                .Kumi = IncMid(lBUf, p, 2)        '' �g
                .Odds = IncMid(lBUf, p, 5)        '' �I�b�Y
                .Ninki = IncMid(lBUf, p, 2)       '' �l�C��
            End With ' OddsWakurenInfo
        Next i
        .TotalHyosuTansyo = IncMid(lBUf, p, 11)   '' �P���[�����v
        .TotalHyosuFukusyo = IncMid(lBUf, p, 11)  '' �����[�����v
        .TotalHyosuWakuren = IncMid(lBUf, p, 11)  '' �g�A�[�����v
        .CRLF = IncMid(lBUf, p, 2)                '' ���R�[�h��؂�
    End With

    End Sub


    '****** �V�D�I�b�Y�i�n�A�j****************************************

    Public Sub SetData_O2(lBUf As String, ByRef mBuf As JV_O2_ODDS_UMAREN)
    Dim i As Integer                                '' ���[�v�J�E���^
    Dim j As Integer                                '' ���[�v�J�E���^
    Dim k As Integer                                '' ���[�v�J�E���^
    Dim p As Long                                   '' �؂蕪���J�n�ʒu
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMid(lBUf, p, 2)      '' ���R�[�h���
            .DataKubun = IncMid(lBUf, p, 1)       '' �f�[�^�敪
            With .MakeDate
                .Year = IncMid(lBUf, p, 4)        '' �N
                .Month = IncMid(lBUf, p, 2)       '' ��
                .Day = IncMid(lBUf, p, 2) '' ��
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMid(lBUf, p, 4)    '' �J�ÔN
            .MonthDay = IncMid(lBUf, p, 4)        '' �J�Ì���
            .JyoCD = IncMid(lBUf, p, 2)   '' ���n��R�[�h
            .Kaiji = IncMid(lBUf, p, 2)   '' �J�É�[��N��]
            .Nichiji = IncMid(lBUf, p, 2) '' �J�Ó���[N����]
            .RaceNum = IncMid(lBUf, p, 2) '' ���[�X�ԍ�
        End With ' id
        With .HappyoTime
            .Month = IncMid(lBUf, p, 2)   '' ��
            .Day = IncMid(lBUf, p, 2)     '' ��
            .Hour = IncMid(lBUf, p, 2)    '' ��
            .Minute = IncMid(lBUf, p, 2)  '' ��
        End With ' HappyoTime
        .TorokuTosu = IncMid(lBUf, p, 2)  '' �o�^����
        .SyussoTosu = IncMid(lBUf, p, 2)  '' �o������
        .UmarenFlag = IncMid(lBUf, p, 1)  '' �����t���O�@�n�A
        For i = 0 To 152
            With .OddsUmarenInfo(i)
                .Kumi = IncMid(lBUf, p, 4)        '' �g��
                .Odds = IncMid(lBUf, p, 6)        '' �I�b�Y
                .Ninki = IncMid(lBUf, p, 3)       '' �l�C��
            End With ' OddsUmarenInfo
        Next i
        .TotalHyosuUmaren = IncMid(lBUf, p, 11)   '' �n�A�[�����v
        .CRLF = IncMid(lBUf, p, 2)        '' ���R�[�h��؂�
    End With

    End Sub


    '****** �W�D�I�b�Y�i���C�h�j****************************************

    Public Sub SetData_O3(lBUf As String, ByRef mBuf As JV_O3_ODDS_WIDE)
    Dim i As Integer                                '' ���[�v�J�E���^
    Dim j As Integer                                '' ���[�v�J�E���^
    Dim k As Integer                                '' ���[�v�J�E���^
    Dim p As Long                                   '' �؂蕪���J�n�ʒu
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMid(lBUf, p, 2)      '' ���R�[�h���
            .DataKubun = IncMid(lBUf, p, 1)       '' �f�[�^�敪
            With .MakeDate
                .Year = IncMid(lBUf, p, 4)        '' �N
                .Month = IncMid(lBUf, p, 2)       '' ��
                .Day = IncMid(lBUf, p, 2) '' ��
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMid(lBUf, p, 4)    '' �J�ÔN
            .MonthDay = IncMid(lBUf, p, 4)        '' �J�Ì���
            .JyoCD = IncMid(lBUf, p, 2)   '' ���n��R�[�h
            .Kaiji = IncMid(lBUf, p, 2)   '' �J�É�[��N��]
            .Nichiji = IncMid(lBUf, p, 2) '' �J�Ó���[N����]
            .RaceNum = IncMid(lBUf, p, 2) '' ���[�X�ԍ�
        End With ' id
        With .HappyoTime
            .Month = IncMid(lBUf, p, 2)   '' ��
            .Day = IncMid(lBUf, p, 2)     '' ��
            .Hour = IncMid(lBUf, p, 2)    '' ��
            .Minute = IncMid(lBUf, p, 2)  '' ��
        End With ' HappyoTime
        .TorokuTosu = IncMid(lBUf, p, 2)  '' �o�^����
        .SyussoTosu = IncMid(lBUf, p, 2)  '' �o������
        .WideFlag = IncMid(lBUf, p, 1)    '' �����t���O�@���C�h
        For i = 0 To 152
            With .OddsWideInfo(i)
                .Kumi = IncMid(lBUf, p, 4)        '' �g��
                .OddsLow = IncMid(lBUf, p, 5)     '' �Œ�I�b�Y
                .OddsHigh = IncMid(lBUf, p, 5)    '' �ō��I�b�Y
                .Ninki = IncMid(lBUf, p, 3)       '' �l�C��
            End With ' OddsWideInfo
        Next i
        .TotalHyosuWide = IncMid(lBUf, p, 11)     '' ���C�h�[�����v
        .CRLF = IncMid(lBUf, p, 2)        '' ���R�[�h��؂�
    End With

    End Sub


    '****** �X�D�I�b�Y�i�n�P�j ****************************************

    Public Sub SetData_O4(lBUf As String, ByRef mBuf As JV_O4_ODDS_UMATAN)
    Dim i As Integer                                '' ���[�v�J�E���^
    Dim j As Integer                                '' ���[�v�J�E���^
    Dim k As Integer                                '' ���[�v�J�E���^
    Dim p As Long                                   '' �؂蕪���J�n�ʒu
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMid(lBUf, p, 2)      '' ���R�[�h���
            .DataKubun = IncMid(lBUf, p, 1)       '' �f�[�^�敪
            With .MakeDate
                .Year = IncMid(lBUf, p, 4)        '' �N
                .Month = IncMid(lBUf, p, 2)       '' ��
                .Day = IncMid(lBUf, p, 2)         '' ��
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMid(lBUf, p, 4)            '' �J�ÔN
            .MonthDay = IncMid(lBUf, p, 4)        '' �J�Ì���
            .JyoCD = IncMid(lBUf, p, 2)           '' ���n��R�[�h
            .Kaiji = IncMid(lBUf, p, 2)           '' �J�É�[��N��]
            .Nichiji = IncMid(lBUf, p, 2)         '' �J�Ó���[N����]
            .RaceNum = IncMid(lBUf, p, 2)         '' ���[�X�ԍ�
        End With ' id
        With .HappyoTime
            .Month = IncMid(lBUf, p, 2)           '' ��
            .Day = IncMid(lBUf, p, 2)             '' ��
            .Hour = IncMid(lBUf, p, 2)            '' ��
            .Minute = IncMid(lBUf, p, 2)          '' ��
        End With ' HappyoTime
        .TorokuTosu = IncMid(lBUf, p, 2)          '' �o�^����
        .SyussoTosu = IncMid(lBUf, p, 2)          '' �o������
        .UmatanFlag = IncMid(lBUf, p, 1)          '' �����t���O�@�n�P
        For i = 0 To 305
            With .OddsUmatanInfo(i)
                .Kumi = IncMid(lBUf, p, 4)        '' �g��
                .Odds = IncMid(lBUf, p, 6)        '' �I�b�Y
                .Ninki = IncMid(lBUf, p, 3)       '' �l�C��
            End With ' OddsUmatanInfo
        Next i
        .TotalHyosuUmatan = IncMid(lBUf, p, 11)   '' �n�P�[�����v
        .CRLF = IncMid(lBUf, p, 2)                '' ���R�[�h��؂�
    End With

    End Sub


    '****** �P�O�D�I�b�Y�i�R�A���j***************************************

    Public Sub SetData_O5(lBUf As String, ByRef mBuf As JV_O5_ODDS_SANREN)
    Dim i As Integer                                '' ���[�v�J�E���^
    Dim j As Integer                                '' ���[�v�J�E���^
    Dim k As Integer                                '' ���[�v�J�E���^
    Dim p As Long                                   '' �؂蕪���J�n�ʒu
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMid(lBUf, p, 2)      '' ���R�[�h���
            .DataKubun = IncMid(lBUf, p, 1)       '' �f�[�^�敪
            With .MakeDate
                .Year = IncMid(lBUf, p, 4)        '' �N
                .Month = IncMid(lBUf, p, 2)       '' ��
                .Day = IncMid(lBUf, p, 2)         '' ��
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMid(lBUf, p, 4)            '' �J�ÔN
            .MonthDay = IncMid(lBUf, p, 4)        '' �J�Ì���
            .JyoCD = IncMid(lBUf, p, 2)           '' ���n��R�[�h
            .Kaiji = IncMid(lBUf, p, 2)               '' �J�É�[��N��]
            .Nichiji = IncMid(lBUf, p, 2)         '' �J�Ó���[N����]
            .RaceNum = IncMid(lBUf, p, 2)         '' ���[�X�ԍ�
        End With ' id
        With .HappyoTime
            .Month = IncMid(lBUf, p, 2)           '' ��
            .Day = IncMid(lBUf, p, 2)             '' ��
            .Hour = IncMid(lBUf, p, 2)            '' ��
            .Minute = IncMid(lBUf, p, 2)          '' ��
        End With ' HappyoTime
        .TorokuTosu = IncMid(lBUf, p, 2)          '' �o�^����
        .SyussoTosu = IncMid(lBUf, p, 2)          '' �o������
        .SanrenpukuFlag = IncMid(lBUf, p, 1)      '' �����t���O�@3�A��
        For i = 0 To 815
            With .OddsSanrenInfo(i)
                .Kumi = IncMid(lBUf, p, 6)        '' �g��
                .Odds = IncMid(lBUf, p, 6)        '' �I�b�Y
                .Ninki = IncMid(lBUf, p, 3)       '' �l�C��
            End With ' OddsSanrenInfo
        Next i
        .TotalHyosuSanrenpuku = IncMid(lBUf, p, 11)       '' 3�A���[�����v
        .CRLF = IncMid(lBUf, p, 2)        '' ���R�[�h��؂�
    End With
   
    End Sub


    '****** �P�O.�`.�@�I�b�Y�i�R�A�P�j***************************************

    Public Sub SetData_O6(lBUf As String, ByRef mBuf As JV_O6_ODDS_SANRENTAN)
    Dim i As Integer                                '' ���[�v�J�E���^�[
    Dim p As Long                                   '' �؂蕪���J�n�ʒu
    
    p = 1
    With mBuf
        With .head
            .RecordSpec = IncMid(lBUf, p, 2)        '' ���R�[�h���
            .DataKubun = IncMid(lBUf, p, 1)         '' �f�[�^�敪
            With .MakeDate
                .Year = IncMid(lBUf, p, 4)          '' �N
                .Month = IncMid(lBUf, p, 2)         '' ��
                .Day = IncMid(lBUf, p, 2)           '' ��
            End With ' MakeDate
        End With ' head
        With .id
            .Year = IncMid(lBUf, p, 4)              '' �J�ÔN
            .MonthDay = IncMid(lBUf, p, 4)          '' �J�Ì���
            .JyoCD = IncMid(lBUf, p, 2)             '' ���n��R�[�h
            .Kaiji = IncMid(lBUf, p, 2)             '' �J�É�[��N��]
            .Nichiji = IncMid(lBUf, p, 2)           '' �J�Ó���[N����]
            .RaceNum = IncMid(lBUf, p, 2)           '' ���[�X�ԍ�
        End With ' id
        With .HappyoTime
            .Month = IncMid(lBUf, p, 2)             '' ��
            .Day = IncMid(lBUf, p, 2)               '' ��
            .Hour = IncMid(lBUf, p, 2)              '' ��
            .Minute = IncMid(lBUf, p, 2)            '' ��
        End With ' HappyoTime
        .TorokuTosu = IncMid(lBUf, p, 2)            '' �o�^����
        .SyussoTosu = IncMid(lBUf, p, 2)            '' �o������
        .SanrentanFlag = IncMid(lBUf, p, 1)         '' �����t���O�@3�A�P
        For i = 0 To 4895
            With .OddsSanrentanInfo(i)
                .Kumi = IncMid(lBUf, p, 6)          '' �g��
                .Odds = IncMid(lBUf, p, 7)          '' �I�b�Y
                .Ninki = IncMid(lBUf, p, 4)         '' �l�C��
            End With
        Next i
        .TotalHyosuSanrentan = IncMid(lBUf, p, 11)  '' 3�A�P�[�����v
        .CRLF = IncMid(lBUf, p, 2)                  '' ���R�[�h��؂�
    End With
    
    End Sub

    
    '****** �P�P�D�����n�}�X�^ ****************************************

    Public Sub SetData_UM(ByVal lBUf As String, ByRef mBuf As JV_UM_UMA)
    Dim bytBuf() As Byte                            '' Byte��ŏ������邽�߂̃o�b�t�@
    
    Dim i As Integer                                '' ���[�v�J�E���^
    Dim j As Integer                                '' ���[�v�J�E���^
    Dim k As Integer                                '' ���[�v�J�E���^
    Dim p As Long                                   '' �؂蕪���J�n�ʒu
    
    bytBuf = StrConv(lBUf, vbFromUnicode)
    
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

    '�o�b�t�@�̈���
    Erase bytBuf
   
    End Sub


    '****** �P�Q�D�R��}�X�^ ****************************************

    Public Sub SetData_KS(lBUf As String, ByRef mBuf As JV_KS_KISYU)
    Dim bytBuf() As Byte                            '' Byte��ŏ������邽�߂̃o�b�t�@
    
    Dim i As Integer                                '' ���[�v�J�E���^
    Dim j As Integer                                '' ���[�v�J�E���^
    Dim k As Integer                                '' ���[�v�J�E���^
    Dim p As Long                                   '' �؂蕪���J�n�ʒu
    
    bytBuf = StrConv(lBUf, vbFromUnicode)
    
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

    '�o�b�t�@�̈���
    Erase bytBuf
    
    End Sub


    '****** �P�R�D�����t�}�X�^ ****************************************

    Public Sub SetData_CH(lBUf As String, ByRef mBuf As JV_CH_CHOKYOSI)
    Dim bytBuf() As Byte                            '' Byte��ŏ������邽�߂̃o�b�t�@
    
    Dim i As Integer                                '' ���[�v�J�E���^
    Dim j As Integer                                '' ���[�v�J�E���^
    Dim k As Integer                                '' ���[�v�J�E���^
    Dim p As Long                                   '' �؂蕪���J�n�ʒu
    
    bytBuf = StrConv(lBUf, vbFromUnicode)
    
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

    '�o�b�t�@�̈���
    Erase bytBuf
   
    End Sub


    '******�P�S�D���Y�҃}�X�^ ****************************************

    Public Sub SetData_BR(lBUf As String, ByRef mBuf As JV_BR_BREEDER)
    Dim bytBuf() As Byte                            '' Byte��ŏ������邽�߂̃o�b�t�@
    
    Dim i As Integer                                '' ���[�v�J�E���^
    Dim j As Integer                                '' ���[�v�J�E���^
    Dim k As Integer                                '' ���[�v�J�E���^
    Dim p As Long                                   '' �؂蕪���J�n�ʒu
    
    bytBuf = StrConv(lBUf, vbFromUnicode)
    
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

    '�o�b�t�@�̈���
    Erase bytBuf
    
    End Sub


    '****** �P�T�D�n��}�X�^ ****************************************

    Public Sub SetData_BN(lBUf As String, ByRef mBuf As JV_BN_BANUSI)
    Dim bytBuf() As Byte                            '' Byte��ŏ������邽�߂̃o�b�t�@
    
    Dim i As Integer                                '' ���[�v�J�E���^
    Dim j As Integer                                '' ���[�v�J�E���^
    Dim k As Integer                                '' ���[�v�J�E���^
    Dim p As Long                                   '' �؂蕪���J�n�ʒu
    
    bytBuf = StrConv(lBUf, vbFromUnicode)
    
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

    '�o�b�t�@�̈���
    Erase bytBuf
   
    End Sub

    '****** �P�U�D�ɐB�n�}�X�^ ****************************************

    Public Sub SetData_HN(lBUf As String, ByRef mBuf As JV_HN_HANSYOKU)
    Dim bytBuf() As Byte                            '' Byte��ŏ������邽�߂̃o�b�t�@
    
    Dim i As Integer                                '' ���[�v�J�E���^
    Dim j As Integer                                '' ���[�v�J�E���^
    Dim k As Integer                                '' ���[�v�J�E���^
    Dim p As Long                                   '' �؂蕪���J�n�ʒu
    
    bytBuf = StrConv(lBUf, vbFromUnicode)
    
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

    '�o�b�t�@�̈���
    Erase bytBuf
   
    End Sub


    '****** �P�V�D�Y��}�X�^ ****************************************

    Public Sub SetData_SK(lBUf As String, ByRef mBuf As JV_SK_SANKU)
    Dim bytBuf() As Byte                            '' Byte��ŏ������邽�߂̃o�b�t�@
    
    Dim i As Integer                                '' ���[�v�J�E���^
    Dim j As Integer                                '' ���[�v�J�E���^
    Dim k As Integer                                '' ���[�v�J�E���^
    Dim p As Long                                   '' �؂蕪���J�n�ʒu
    
    bytBuf = StrConv(lBUf, vbFromUnicode)
    
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

    '�o�b�t�@�̈���
    Erase bytBuf
    
    End Sub

    '****** �P�W�D���R�[�h�}�X�^ ****************************************

    Public Sub SetData_RC(lBUf As String, ByRef mBuf As JV_RC_RECORD)
    Dim bytBuf() As Byte                            '' Byte��ŏ������邽�߂̃o�b�t�@
    
    Dim i As Integer                                '' ���[�v�J�E���^
    Dim j As Integer                                '' ���[�v�J�E���^
    Dim k As Integer                                '' ���[�v�J�E���^
    Dim p As Long                                   '' �؂蕪���J�n�ʒu
    
    bytBuf = StrConv(lBUf, vbFromUnicode)
    
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

    '�o�b�t�@�̈���
    Erase bytBuf

    End Sub


    '****** �P�X�D��H���� ****************************************

    Public Sub SetData_HC(lBUf As String, ByRef mBuf As JV_HC_HANRO)
    Dim i As Integer                                '' ���[�v�J�E���^
    Dim j As Integer                                '' ���[�v�J�E���^
    Dim k As Integer                                '' ���[�v�J�E���^
    Dim p As Long                                   '' �؂蕪���J�n�ʒu
    
    With mBuf
        With .head
            .RecordSpec = Mid$(lBUf, 1, 2)  '' ���R�[�h���
            .DataKubun = Mid$(lBUf, 3, 1)   '' �f�[�^�敪
            With .MakeDate
                .Year = Mid$(lBUf, 4, 4)    '' �N
                .Month = Mid$(lBUf, 8, 2)   '' ��
                .Day = Mid$(lBUf, 10, 2)     '' ��
            End With ' MakeDate
        End With ' head
        .TresenKubun = Mid$(lBUf, 12, 1)     '' �g���Z���敪
        With .ChokyoDate
            .Year = Mid$(lBUf, 13, 4)        '' �N
            .Month = Mid$(lBUf, 17, 2)       '' ��
            .Day = Mid$(lBUf, 19, 2)         '' ��
        End With ' ChokyoDate
        .ChokyoTime = Mid$(lBUf, 21, 4)      '' ��������
        .KettoNum = Mid$(lBUf, 25, 10)       '' �����o�^�ԍ�
        .HaronTime4 = Mid$(lBUf, 35, 4)      '' 4�n�����^�C�����v(800M-0M)
        .LapTime4 = Mid$(lBUf, 39, 3)        '' ���b�v�^�C��(800M-600M)
        .HaronTime3 = Mid$(lBUf, 42, 4)      '' 3�n�����^�C�����v(600M-0M)
        .LapTime3 = Mid$(lBUf, 46, 3)        '' ���b�v�^�C��(600M-400M)
        .HaronTime2 = Mid$(lBUf, 49, 4)      '' 2�n�����^�C�����v(400M-0M)
        .LapTime2 = Mid$(lBUf, 53, 3)        '' ���b�v�^�C��(400M-200M)
        .LapTime1 = Mid$(lBUf, 56, 3)        '' ���b�v�^�C��(200M-0M)
        .CRLF = Mid$(lBUf, 59, 2)            '' ���R�[�h��؂�
    End With

  End Sub


    '****** �Q�O�D�n�̏d ****************************************

    Public Sub SetData_WH(lBUf As String, ByRef mBuf As JV_WH_BATAIJYU)
    Dim bytBuf() As Byte                            '' Byte��ŏ������邽�߂̃o�b�t�@
    
    Dim i As Integer                                '' ���[�v�J�E���^
    Dim j As Integer                                '' ���[�v�J�E���^
    Dim k As Integer                                '' ���[�v�J�E���^
    Dim p As Long                                   '' �؂蕪���J�n�ʒu
    
    bytBuf = StrConv(lBUf, vbFromUnicode)
    
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

    '�o�b�t�@�̈���
    Erase bytBuf
   
    End Sub


    '****** �Q�P�D�V��n���� ******************************************

    Public Sub SetData_WE(lBUf As String, ByRef mBuf As JV_WE_WEATHER)
    Dim bytBuf() As Byte                            '' Byte��ŏ������邽�߂̃o�b�t�@
    
    Dim i As Integer                                '' ���[�v�J�E���^
    Dim j As Integer                                '' ���[�v�J�E���^
    Dim k As Integer                                '' ���[�v�J�E���^
    Dim p As Long                                   '' �؂蕪���J�n�ʒu
    
    bytBuf = StrConv(lBUf, vbFromUnicode)
    
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

    '�o�b�t�@�̈���
    Erase bytBuf
    
    End Sub


    '****** �Q�Q�D�o������E�������O ****************************************

    Public Sub SetData_AV(lBUf As String, ByRef mBuf As JV_AV_INFO)
    Dim bytBuf() As Byte                            '' Byte��ŏ������邽�߂̃o�b�t�@
    
    Dim i As Integer                                '' ���[�v�J�E���^
    Dim j As Integer                                '' ���[�v�J�E���^
    Dim k As Integer                                '' ���[�v�J�E���^
    Dim p As Long                                   '' �؂蕪���J�n�ʒu
    
    bytBuf = StrConv(lBUf, vbFromUnicode)
    
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

    '�o�b�t�@�̈���
    Erase bytBuf
    
    End Sub

    '************ �Q�R�D�R��ύX ****************************************
  
    Public Sub SetData_JC(lBUf As String, ByRef mBuf As JV_JC_INFO)
    Dim bytBuf() As Byte                            '' Byte��ŏ������邽�߂̃o�b�t�@
    
    Dim i As Integer                                '' ���[�v�J�E���^
    Dim j As Integer                                '' ���[�v�J�E���^
    Dim k As Integer                                '' ���[�v�J�E���^
    Dim p As Long                                   '' �؂蕪���J�n�ʒu
    
    bytBuf = StrConv(lBUf, vbFromUnicode)
    
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

    '�o�b�t�@�̈���
    Erase bytBuf
   
    End Sub

    '****** �Q�S�D�f�[�^�}�C�j���O�\�z***********************************
    
    Public Sub SetData_DM(lBUf As String, ByRef mBuf As JV_DM_INFO)
    Dim bytBuf() As Byte                            '' Byte��ŏ������邽�߂̃o�b�t�@
    
    Dim i As Integer                                '' ���[�v�J�E���^
    Dim j As Integer                                '' ���[�v�J�E���^
    Dim k As Integer                                '' ���[�v�J�E���^
    Dim p As Long                                   '' �؂蕪���J�n�ʒu
    
    bytBuf = StrConv(lBUf, vbFromUnicode)
    
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

    '�o�b�t�@�̈���
    Erase bytBuf
    
    End Sub


    '****** �Q�T�D�J�ÃX�P�W���[��************************************
    
    Public Sub SetData_YS(lBUf As String, ByRef mBuf As JV_YS_SCHEDULE)
    Dim bytBuf() As Byte                            '' Byte��ŏ������邽�߂̃o�b�t�@
    
    Dim i As Integer                                '' ���[�v�J�E���^
    Dim j As Integer                                '' ���[�v�J�E���^
    Dim k As Integer                                '' ���[�v�J�E���^
    Dim p As Long                                   '' �؂蕪���J�n�ʒu
    
    bytBuf = StrConv(lBUf, vbFromUnicode)
    
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

    '�o�b�t�@�̈���
    Erase bytBuf
    
    End Sub
    
    
    '****** �Q�U�D���������ύX************************************
    Public Sub SetData_TC(lBUf As String, ByRef mBuf As JV_TC_HASSOU)

        Dim bytBuf() As Byte                            '' Byte��ŏ������邽�߂̃o�b�t�@
        Dim p As Long
    
        bytBuf = StrConv(lBUf, vbFromUnicode)
        
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
        
        '�o�b�t�@�̈�J��
        Erase bytBuf
        
    End Sub
    
    '****** �Q�V�D�R�[�X�ύX************************************
    Public Sub SetData_CC(lBUf As String, ByRef mBuf As JV_CC_COURSE)
    
        Dim bytBuf() As Byte
        Dim p As Long
        
        bytBuf = StrConv(lBUf, vbFromUnicode)
        
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
        
        '�o�b�t�@�̈�J��
        Erase bytBuf
        
    End Sub
    
    '------------------------------------------------------------------------
    '�@�@�o�C�g�z����o�C�g���Ő؏o��
    '------------------------------------------------------------------------
    Public Function IncMidByte(ByRef vBuf() As Byte, p As Long, ByVal length As Long) As String
        IncMidByte = StrConv(MidB$(vBuf, p, length), vbUnicode)
        p = p + length
    End Function
        
    '------------------------------------------------------------------------
    '�@�@������̐؂�o��
    '------------------------------------------------------------------------
    Public Function IncMid(ByRef buf As String, p As Long, ByVal length As Long) As String
        IncMid = Mid$(buf, p, length)
        p = p + length
    End Function

