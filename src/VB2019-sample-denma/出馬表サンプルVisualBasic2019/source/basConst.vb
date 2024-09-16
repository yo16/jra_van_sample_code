Option Explicit On 

Module basConst
    '========================================================================
    '  JRA-VAN Data Lab.�v���O���~���O�p�[�c�uPublic�ϐ���`�t�@�C���v
    '
    '
    '   �쐬: JRA-VAN �\�t�g�E�F�A�H�[  2004�N12��28��
    '
    '========================================================================
    '   (C) Copyright Turf Media System Co.,Ltd. 2003 All rights reserved
    '========================================================================

    ' -----�f�[�^�敪-----
    ' ���[�X�ڍ�
    Public Const ID_RACE As String = "RA"

    ' �n�����[�X���
    Public Const ID_RACE_UMA As String = "SE"

    ' �����n�}�X�^
    Public Const ID_UMA As String = "UM"

    ' �R��}�X�^
    Public Const ID_KISYU As String = "KS"

    ' ���ʓo�^�n
    Public Const ID_TOKU As String = "TK"

    ' ����
    Public Const ID_HARAI As String = "HR"

    ' �[��1
    Public Const ID_HYOSU As String = "H1"

    ' �[��6(3�A�P)
    Public Const ID_HYOSU_SANRENTAN As String = "H6"

    ' �I�b�Y1(�P���g)
    Public Const ID_ODDS_TANPUKU As String = "O1"

    ' �I�b�Y2(�n�A)
    Public Const ID_ODDS_UMAREN As String = "O2"

    ' �I�b�Y3(���C�h)
    Public Const ID_ODDS_WIDE As String = "O3"

    ' �I�b�Y4(�n�P)
    Public Const ID_ODDS_UMATAN As String = "O4"

    ' �I�b�Y5(3�A��)
    Public Const ID_ODDS_SANREN As String = "O5"

    ' �I�b�Y6(3�A�P)
    Public Const ID_ODDS_SANRENTAN As String = "O6"

    ' �����t�}�X�^
    Public Const ID_CHOKYO As String = "CH"

    ' ���Y�҃}�X�^
    Public Const ID_SEISAN As String = "BR"

    ' �n��}�X�^
    Public Const ID_BANUSI As String = "BN"

    ' �ɐB�n�}�X�^
    Public Const ID_HANSYOKU As String = "HN"

    ' �Y��}�X�^
    Public Const ID_SANKU As String = "SK"

    ' ���R�[�h�}�X�^
    Public Const ID_RECORD As String = "RC"

    ' ��H����
    Public Const ID_HANRO As String = "HC"

    ' �n�̏d
    Public Const ID_BATAIJYU As String = "WH"

    ' �V��n����
    Public Const ID_TENKO_BABA As String = "WE"

    ' �o������E�������O
    Public Const ID_TORIKESI_JYOGAI As String = "AV"

    ' �R��ύX
    Public Const ID_KISYU_CHANGE As String = "JC"

    ' �f�[�^�}�C�j���O�\�z
    Public Const ID_MINING As String = "DM"

    ' �J�ÃX�P�W���[��
    Public Const ID_SCHEDULE As String = "YS"

    ' ���������ύX
    Public Const ID_HASSOU_JIKOKU_CHANGE As String = "TC"

    ' �R�[�X�ύX
    Public Const ID_COURSE_CHANGE As String = "CC"


    ' -----JV-Link�X�e�[�^�X-----

    ' JVRead - ����ǂݍ���
    Public Const ST_READ_SUCCESS As Integer = 0

    ' JVRead - �G���[ 
    Public Const ST_READ_ERR As Integer = -2

    ' JVRead - �t�@�C�����X�g�ǂݍ��ݏI�� 
    Public Const ST_READ_EOL As Integer = 0

    ' JVRead - �t�@�C���̋�؂� 
    Public Const ST_READ_EOF As Integer = -1

    ' JVRead - �_�E�����[�h�� 
    Public Const ST_READ_DOWNLOAD_NOW As Integer = -3


    ' -----�R�[�h�ϊ�-----

    ' ���n��R�[�h
    Public Const CV_JO_CD As String = "2001"

    ' �j���R�[�h
    Public Const CV_WD_CD As String = "2002"

    ' �O���[�h�R�[�h
    Public Const CV_GR_CD As String = "2003"

    ' ������ʃR�[�h
    Public Const CV_RS_CD As String = "2005"

    ' �����L���R�[�h
    Public Const CV_RK_CD As String = "2006"

    ' ���������R�[�h
    Public Const CV_RJ_CD As String = "2007"

    ' �d�ʎ�ʃR�[�h
    Public Const CV_WH_CD As String = "2008"

    ' �g���b�N�R�[�h
    Public Const CV_TR_CD As String = "2009"

    ' �n���ԃR�[�h
    Public Const CV_BC_CD As String = "2010"

    ' �V��R�[�h
    Public Const CV_WE_CD As String = "2011"

    ' �ُ�敪�R�[�h
    Public Const CV_IR_CD As String = "2101"

    ' �����R�[�h
    Public Const CV_TS_CD As String = "2102"

    ' �i��R�[�h
    Public Const CV_HS_CD As String = "2201"

    ' ���ʃR�[�h
    Public Const CV_SX_CD As String = "2202"

    ' �ѐF�R�[�h
    Public Const CV_FC_CD As String = "2203"

    ' �n�L���R�[�h
    Public Const CV_UK_CD As String = "2204"

    ' ���������R�[�h
    Public Const CV_TZ_CD As String = "2301"

    ' �R�掑�i�R�[�h
    Public Const CV_KQ_CD As String = "2302"

    ' �R�茩�K�R�[�h
    Public Const CV_KM_CD As String = "2303"


    ' -----�f�[�^�敪-----

    ' �o���n���\(�ؗj)
    Public Const KB_THU As String = "1"

    ' �o�n�\(���E�y�j)
    Public Const KB_FRI As String = "2"

    ' ���ё���(3���܂Ŋm��)
    Public Const KB_S3 As String = "3"

    ' ���ё���(5���܂Ŋm��)
    Public Const KB_S5 As String = "4"

    ' ���ё���(�S�n�����m��)
    Public Const KB_SALL As String = "5"

    ' ���ё���(�S�n����+��Ű�ʉߏ�)
    Public Const KB_SCOR As String = "6"

    ' ����(���j)
    Public Const KB_MON As String = "7"

    ' �n�����n
    Public Const KB_LKL As String = "A"

    ' �C�O���ۃ��[�X
    Public Const KB_FOR As String = "B"

    ' ���[�X���~
    Public Const KB_CAN As String = "9"

    ' �Y���f�[�^�폜
    Public Const KB_DEL As String = "0"


End Module