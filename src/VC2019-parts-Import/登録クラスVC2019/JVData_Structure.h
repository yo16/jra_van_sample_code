#ifndef __JV_DATA_STRUCT
#define __JV_DATA_STRUCT


//========================================================================
//  JRA-VAN Data Lab. JV-Data�\����
//
//
//   �쐬: JRA-VAN �\�t�g�E�F�A�H�[
//
//========================================================================
//   (C) Copyright JRA SYSTEM SERVICE CO.,LTD. 2007 All rights reserved
//========================================================================


//////////////////// ���ʍ\���� ////////////////////


//<�N����>
typedef struct
{
		char		Year[4];                   //�N
		char		Month[2];                  //��
		char		Day[2];                    //��
}				_YMD;


//<�����b>
typedef struct
{
		char		Hour[2];                   //��
		char		Minute[2];                 //��
		char		Second[2];                 //�b
}				_HMS;


 //<����>
typedef struct
{
		char		Hour[2];                   //��
		char		Minute[2];                 //��
}				_HM;


//<��������>
typedef struct
{
		char		Month[2];                  //��
		char		Day[2];                    //��
		char		Hour[2];                   //��
		char		Minute[2];                 //��
}				_MDHM;


//<���R�[�h�w�b�_>
typedef struct
{
		char		RecordSpec[2];             //���R�[�h���
		char		DataKubun[1];              //�f�[�^�敪
		_YMD		MakeDate;                  //�f�[�^�쐬�N����
}				_RECORD_ID;


//<�������ʏ��P>
typedef struct
{
		char		Year[4];                   //�J�ÔN
		char		MonthDay[4];               //�J�Ì���
		char		JyoCD[2];                  //���n��R�[�h
		char		Kaiji[2];                  //�J�É�[��N��]
		char		Nichiji[2];                //�J�Ó���[N����]
		char		RaceNum[2];                //���[�X�ԍ�
}				_RACE_ID;


//<�������ʏ��Q>
typedef struct
{
		char		Year[4];                   //�J�ÔN
		char		MonthDay[4];               //�J�Ì���
		char		JyoCD[2];                  //���n��R�[�h
		char		Kaiji[2];                  //�J�É�[��N��]
		char		Nichiji[2];                //�J�Ó���[N����]
}				_RACE_ID2;


//<���񐔁i�T�C�Y3byte�j>
typedef struct
{
		char		Chakukaisu[6][3];
}				_CHAKUKAISU3_INFO;

//<���񐔁i�T�C�Y4byte�j>
typedef struct
{
		char		Chakukaisu[6][4];
}				_CHAKUKAISU4_INFO;

//<���񐔁i�T�C�Y5byte�j>
typedef struct
{
		char		Chakukaisu[6][5];
}				_CHAKUKAISU5_INFO;

//<���񐔁i�T�C�Y6byte�j>
typedef struct
{
		char		Chakukaisu[6][6];
}				_CHAKUKAISU6_INFO;


//<�{�N�E�݌v���я��>
typedef struct
{
		char		SetYear[4];                //�ݒ�N
		char		HonSyokinTotal[10];        //�{�܋����v
		char		FukaSyokin[10];            //�t���܋����v
		char		ChakuKaisu[6][6];          //����
}				_SEI_RUIKEI_INFO;


//<�ŋߏd�܏������>
typedef struct
{
		_RACE_ID	SaikinJyusyoid;             //<�N��������R>
		char		Hondai[60];                 //�������{��
		char		Ryakusyo10[20];             //����������10��
		char		Ryakusyo6[12];              //����������6��
		char		Ryakusyo3[6];               //����������3��
		char		GradeCD[1];                 //�O���[�h�R�[�h
		char		SyussoTosu[2];              //�o������
		char		KettoNum[10];               //�����o�^�ԍ�
		char		Bamei[36];                  //�n��
}				_SAIKIN_JYUSYO_INFO;



//<�{�N�E�O�N�E�݌v���я��>
typedef struct
{
		char		SetYear[4];                 //�ݒ�N
		char		HonSyokinHeichi[10];        //���n�{�܋����v
		char		HonSyokinSyogai[10];        //��Q�{�܋����v
		char		FukaSyokinHeichi[10];       //���n�t���܋����v
		char		FukaSyokinSyogai[10];       //��Q�t���܋����v
		_CHAKUKAISU6_INFO		ChakuKaisuHeichi;     //���n����
		_CHAKUKAISU6_INFO		ChakuKaisuSyogai;     //��Q����
		_CHAKUKAISU6_INFO		ChakuKaisuJyo[20];    //���n��ʒ���
		_CHAKUKAISU6_INFO		ChakuKaisuKyori[6];   //�����ʒ���
}				_HON_ZEN_RUIKEISEI_INFO;


//<���[�X���>
typedef struct
{
		char		YoubiCD[1];                //�j���R�[�h
		char		TokuNum[4];                //���ʋ����ԍ�
		char		Hondai[60];                //�������{��
		char		Fukudai[60];               //����������
		char		Kakko[60];                 //�������J�b�R��
		char		HondaiEng[120];            //�������{�艢��
		char		FukudaiEng[120];           //���������艢��
		char		KakkoEng[120];             //�������J�b�R������
		char		Ryakusyo10[20];            //���������̂P�O��
		char		Ryakusyo6[12];             //���������̂U��
		char		Ryakusyo3[6];              //���������̂R��
		char		Kubun[1];                  //�������敪
		char		Nkai[3];                   //�d�܉�[��N��]
}				_RACE_INFO;

//<�V��E�n����>
typedef struct
{
		char		TenkoCD[1];                //�V��R�[�h
		char		SibaBabaCD[1];             //�Ŕn���ԃR�[�h
		char		DirtBabaCD[1];             //�_�[�g�n���ԃR�[�h
}				_TENKO_BABA_INFO;


//<���������R�[�h>
typedef struct
{
		char		SyubetuCD[2];              //������ʃR�[�h
		char		KigoCD[3];                 //�����L���R�[�h
		char		JyuryoCD[1];               //�d�ʎ�ʃR�[�h
		char		JyokenCD[5][3];            //���������R�[�h
}				_RACE_JYOKEN;


//<�R��ύX���>
typedef struct
{
		char		Futan[3];                 //���S�d��
		char		KisyuCode[5];             //�R��R�[�h
		char		KisyuName[34];            //�R�薼
		char		MinaraiCD[1];             //�R�茩�K�R�[�h
}				_JC_INFO;


//<���������ύX���>
typedef struct
{
		char		Ji[2];                 	 //��
		char		Fun[2];             	 //��
}				_TC_INFO;


//<�R�[�X�ύX���>
typedef struct
{
		char		Kyori[4];                //����
		char		TruckCD[2];              //�g���b�N�R�[�h
}				_CC_INFO;


//////////////////// �f�[�^�\���� ////////////////////


//**** �P�D���ʓo�^�n ****************************************
typedef struct{
		_RECORD_ID	head;                     //<���R�[�h�w�b�_�[>
		_RACE_ID	id;                       //<�������ʏ��P>
		_RACE_INFO	RaceInfo;                 //<���[�X���>
		char		GradeCD[1];               //�O���[�h�R�[�h
		_RACE_JYOKEN	JyokenInfo;               //<���������R�[�h>
		char		Kyori[4];                 //����
		char		TrackCD[2];               //�g���b�N�R�[�h
		char		CourseKubunCD[2];         //�R�[�X�敪
		_YMD		HandiDate;                //�n���f���\��
		char		TorokuTosu[3];            //�o�^����

		struct _TOKUUMA_INFO                      //<�o�^�n�����>
                {
				char		Num[3];                    //�A��
				char		KettoNum[10];              //�����o�^�ԍ�
				char		Bamei[36];                 //�n��
				char		UmaKigoCD[2];              //�n�L���R�[�h
				char		SexCD[1];                  //���ʃR�[�h
				char		TozaiCD[1];                //�����t���������R�[�h
				char		ChokyosiCode[5];           //�����t�R�[�h
				char		ChokyosiRyakusyo[8];       //�����t������
				char		Futan[3];                  //���S�d��
				char		Koryu[1];                  //�𗬋敪
		}				TokuUmaInfo[300];

		char		crlf[2];                  //���R�[�h���
}				JV_TK_TOKUUMA;


//****** �Q�D���[�X�ڍ� ****************************************
typedef struct
{
		_RECORD_ID	head;                     //<���R�[�h�w�b�_�[>
		_RACE_ID	id;                       //<�������ʏ��P>
		_RACE_INFO	RaceInfo;                 //<���[�X���>
		char		GradeCD[1];               //�O���[�h�R�[�h
		char		GradeCDBefore[1];         //�ύX�O�O���[�h�R�[�h
		_RACE_JYOKEN	JyokenInfo;               //<���������R�[�h>
		char		JyokenName[60];           //������������
		char		Kyori[4];                 //����
		char		KyoriBefore[4];           //�ύX�O����
		char		TrackCD[2];               //�g���b�N�R�[�h
		char		TrackCDBefore[2];         //�ύX�O�g���b�N�R�[�h
		char		CourseKubunCD[2];         //�R�[�X�敪
		char		CourseKubunCDBefore[2];   //�ύX�O�R�[�X�敪
		char		Honsyokin[7][8];          //�{�܋�
		char		HonsyokinBefore[5][8];    //�ύX�O�{�܋�
		char		Fukasyokin[5][8];         //�t���܋�
		char		FukasyokinBefore[3][8];   //�ύX�O�t���܋�
		char		HassoTime[4];             //��������
		char		HassoTimeBefore[4];       //�ύX�O��������
		char		TorokuTosu[2];            //�o�^����
		char		SyussoTosu[2];            //�o������
		char		NyusenTosu[2];            //��������
		_TENKO_BABA_INFO        TenkoBaba;        //�V��E�n���ԃR�[�h
		char		LapTime[25][3];           //���b�v�^�C��
		char		SyogaiMileTime[4];        //��Q�}�C���^�C��
		char		HaronTimeS3[3];           //�O�R�n�����^�C��
		char		HaronTimeS4[3];           //�O�S�n�����^�C��
		char		HaronTimeL3[3];           //��R�n�����^�C��
		char		HaronTimeL4[3];           //��S�n�����^�C��

		struct _CORNER_INFO                       //<�R�[�i�[�ʉߏ���>
                {
				char		Corner[1];                //�R�[�i�[
				char		Syukaisu[1];              //����
				char		Jyuni[70];                 //�e�ʉߏ���
		}				CornerInfo[4];

		char		RecordUpKubun[1];         //���R�[�h�X�V�敪
		char		crlf[2];                  //���R�[�h��؂�
}				JV_RA_RACE;


//****** �R�D�n�����[�X��� ****************************************
typedef struct
{
		_RECORD_ID	head;                     //<���R�[�h�w�b�_�[>
		_RACE_ID	id;                       //<�������ʏ��P>
		char		Wakuban[1];               //�g��
		char		Umaban[2];                //�n��
		char		KettoNum[10];             //�����o�^�ԍ�
		char		Bamei[36];                //�n��
		char		UmaKigoCD[2];             //�n�L���R�[�h
		char		SexCD[1];                 //���ʃR�[�h
		char		HinsyuCD[1];              //�i��R�[�h
		char		KeiroCD[2];               //�ѐF�R�[�h
		char		Barei[2];                 //�n��
		char		TozaiCD[1];               //���������R�[�h
		char		ChokyosiCode[5];          //�����t�R�[�h
		char		ChokyosiRyakusyo[8];      //�����t������
		char		BanusiCode[6];            //�n��R�[�h
		char		BanusiName[64];           //�n�喼
		char		Fukusyoku[60];            //���F�W��
		char		reserved1[60];            //�\��
		char		Futan[3];                 //���S�d��
		char		FutanBefore[3];           //�ύX�O���S�d��
		char		Blinker[1];               //�u�����J�[�g�p�敪
		char		reserved2[1];             //�\��
		char		KisyuCode[5];             //�R��R�[�h
		char		KisyuCodeBefore[5];       //�ύX�O�R��R�[�h
		char		KisyuRyakusyo[8];         //�R�薼����
		char		KisyuRyakusyoBefore[8];   //�ύX�O�R�薼����
		char		MinaraiCD[1];             //�R�茩�K�R�[�h
		char		MinaraiCDBefore[1];       //�ύX�O�R�茩�K�R�[�h
		char		BaTaijyu[3];              //�n�̏d
		char		ZogenFugo[1];             //��������
		char		ZogenSa[3];               //������
		char		IJyoCD[1];                //�ُ�敪�R�[�h
		char		NyusenJyuni[2];           //��������
		char		KakuteiJyuni[2];          //�m�蒅��
		char		DochakuKubun[1];          //�����敪
		char		DochakuTosu[1];           //��������
		char		Time[4];                  //���j�^�C��
		char		ChakusaCD[3];             //�����R�[�h
		char		ChakusaCDP[3];            //+�����R�[�h
		char		ChakusaCDPP[3];           //++�����R�[�h
		char		Jyuni1c[2];               //1�R�[�i�[�ł̏���
		char		Jyuni2c[2];               //2�R�[�i�[�ł̏���
		char		Jyuni3c[2];               //3�R�[�i�[�ł̏���
		char		Jyuni4c[2];               //4�R�[�i�[�ł̏���
		char		Odds[4];                  //�P���I�b�Y
		char		Ninki[2];                 //�P���l�C��
		char		Honsyokin[8];             //�l���{�܋�
		char		Fukasyokin[8];            //�l���t���܋�
		char		reserved3[3];             //�\��
		char		reserved4[3];             //�\��
		char		HaronTimeL4[3];           //��S�n�����^�C��
		char		HaronTimeL3[3];           //��R�n�����^�C��

		struct _CHAKUUMA_INFO                     //<1���n[����n]���>
                {
				char		KettoNum[10];     //�����o�^�ԍ�
				char		Bamei[36];        //�n��
		}			ChakuUmaInfo[3];

		char		TimeDiff[4];              //�^�C����
		char		RecordUpKubun[1];         //���R�[�h�X�V�敪
		char		DMKubun[1];               //�}�C�j���O�敪
		char		DMTime[5];                //�}�C�j���O�\�z���j�^�C��
		char		DMGosaP[4];               //�\���덷[�M���x]�{
		char		DMGosaM[4];               //�\���덷[�M���x]�|
		char		DMJyuni[2];               //�}�C�j���O�\�z����
		char		KyakusituKubun[1];        //���񃌁[�X�r������
		char		crlf[2];                  //���R�[�h��؂�
}				JV_SE_RACE_UMA;


//****** �S�D���� ****************************************
//<���ߏ��P �P�E���E�g>
typedef struct
{
		char		Umaban[2];                //�n��
		char		Pay[9];                   //���ߋ�
		char		Ninki[2];                 //�l�C�� 
}			_PAY_INFO1;


//<���ߏ��Q �n�A�E���C�h�E�\���E�n�P>
typedef struct
{
		char		Kumi[4];                  //�g��
		char		Pay[9];                   //���ߋ�
		char		Ninki[3];                 //�l�C�� 
}			_PAY_INFO2;


//<���ߏ��R �R�A��>
typedef struct{ 
		char		Kumi[6];                  //�g��
		char		Pay[9];                   //���ߋ�
		char		Ninki[3];                 //�l�C�� 
}			_PAY_INFO3;


//<���ߏ��S �R�A�P>
typedef struct
{
		char		Kumi[6];                  //�g��
		char		Pay[9];                   //���ߋ�
		char		Ninki[4];                 //�l�C��
}			_PAY_INFO4;


typedef struct
{
		_RECORD_ID	head;                     //<���R�[�h�w�b�_�[>
		_RACE_ID	id;                       //<�������ʏ��P>
		char		TorokuTosu[2];            //�o�^����
		char		SyussoTosu[2];            //�o������
		char		FuseirituFlag[9][1];      //�s�����t���O
		char		TokubaraiFlag[9][1];      //�����t���O
		char		HenkanFlag[9][1];         //�Ԋ҃t���O
		char		HenkanUma[28][1];         //�ԊҔn�ԏ��[�n��01�`28]
		char		HenkanWaku[8][1];         //�ԊҘg�ԏ��[�g��1�`8]
		char		HenkanDoWaku[8][1];       //�Ԋғ��g���[�g��1�`8]
		_PAY_INFO1		PayTansyo[3];         //<�P������>
		_PAY_INFO1		PayFukusyo[5];        //<��������>
		_PAY_INFO1		PayWakuren[3];        //<�g�A����>
		_PAY_INFO2		PayUmaren[3];         //<�n�A����>
		_PAY_INFO2		PayWide[7];           //<���C�h����>
		_PAY_INFO2		PayReserved1[3];      //<�\��>
		_PAY_INFO2		PayUmatan[6];         //<�n�P����>
		_PAY_INFO3		PaySanrenpuku[3];     //<3�A������>
		_PAY_INFO4		PaySanrentan[6];      //<3�A�P����>
		char		crlf[2];                  //���R�[�h��؂�
}				JV_HR_PAY;


//****** �T�D�[���i�S�|���j****************************************
//<�[�����P �P�E���E�g>
typedef struct
{
		char		Umaban[2];                //�n��
		char		Hyo[11];                  //�[��
		char		Ninki[2];                 //�l�C
}				_HYO_INFO1;


//<�[�����Q �n�A�E���C�h�E�n�P>
typedef struct
{
		char		Kumi[4];                  //�g��     
		char		Hyo[11];                  //�[��
		char		Ninki[3];                 //�l�C
}				_HYO_INFO2;


//<�[�����R �R�A���[��>
typedef struct
{
		char		Kumi[6];                  //�g��     
		char		Hyo[11];                  //�[��
		char		Ninki[3];                 //�l�C
}				_HYO_INFO3;


//<�[�����S �R�A�P�[��>
typedef struct
{
		char		Kumi[6];                  //�g��     
		char		Hyo[11];                  //�[��
		char		Ninki[4];                 //�l�C
}				_HYO_INFO4;


typedef struct
{
		_RECORD_ID	head;                     //<���R�[�h�w�b�_�[>
		_RACE_ID	id;                       //<�������ʏ��P>
		char		TorokuTosu[2];            //�o�^����
		char		SyussoTosu[2];            //�o������
		char		HatubaiFlag[7][1];        //�����t���O�@
		char		FukuChakuBaraiKey[1];     //���������L�[
		char		HenkanUma[28][1];         //�ԊҔn�ԏ��[�n��01�`28]
		char		HenkanWaku[8][1];         //�ԊҘg�ԏ��[�g��1�`8]
		char		HenkanDoWaku[8][1];       //�Ԋғ��g���[�g��1�`8]
		_HYO_INFO1	HyoTansyo[28];            //<�P���[��>
		_HYO_INFO1	HyoFukusyo[28];           //<�����[��>
		_HYO_INFO1	HyoWakuren[36];           //<�g�A�[��>
		_HYO_INFO2	HyoUmaren[153];           //<�n�A�[��>
		_HYO_INFO2	HyoWide[153];             //<���C�h�[��>
		_HYO_INFO2	HyoUmatan[306];           //<�n�P�[��>
		_HYO_INFO3	HyoSanrenpuku[816];       //<3�A���[��>
		char		HyoTotal[14][11];         //�[�����v
		char		crlf[2];                  //���R�[�h��؂�
}				JV_H1_HYOSU_ZENKAKE;

typedef struct
{
		_RECORD_ID	head;                     //<���R�[�h�w�b�_�[>
		_RACE_ID	id;                       //<�������ʏ��P>
		char		TorokuTosu[2];            //�o�^����
		char		SyussoTosu[2];            //�o������
		char		HatubaiFlag[1];        	  //�����t���O
		char		HenkanUma[18][1];         //�ԊҔn�ԏ��[�n��01�`18]
		_HYO_INFO4	HyoSanrentan[4896];       //<3�A�P�[��>
		char		HyoTotal[2][11];         //�[�����v
		char		crlf[2];                  //���R�[�h��؂�
}				JV_H6_HYOSU_SANRENTAN;

//****** �U�D�I�b�Y�i�P���g�j****************************************
typedef struct
{
		_RECORD_ID	head;                     //<���R�[�h�w�b�_�[>
		_RACE_ID	id;                       //<�������ʏ��P>
		_MDHM		HappyoTime;               //���\��������
		char		TorokuTosu[2];            //�o�^����
		char		SyussoTosu[2];            //�o������
		char		TansyoFlag[1];            //�����t���O�@�P��
		char		FukusyoFlag[1];           //�����t���O�@����
		char		WakurenFlag[1];           //�����t���O�@�g�A
		char		FukuChakuBaraiKey[1];     //���������L�[

		struct _ODDS_TANSYO_INFO                  //<�P���I�b�Y>
                {
				char		Umaban[2];                //�n��
				char		Odds[4];                  //�I�b�Y
				char		Ninki[2];                 //�l�C��
		}			OddsTansyoInfo[28];

		struct _ODDS_FUKUSYO_INFO                 //<�����I�b�Y>
                {
				char		Umaban[2];                //�n��
				char		OddsLow[4];               //�Œ�I�b�Y
				char		OddsHigh[4];              //�ō��I�b�Y
				char		Ninki[2];                 //�l�C��
		}			OddsFukusyoInfo[28];

		struct _ODDS_WAKUREN_INFO                 //<�g�A�I�b�Y>
                {
				char		Kumi[2];                  //�g
				char		Odds[5];                  //�I�b�Y
				char		Ninki[2];                 //�l�C��
		}			OddsWakurenInfo[36];

		char		TotalHyosuTansyo[11];     //�P���[�����v
		char		TotalHyosuFukusyo[11];    //�����[�����v
		char		TotalHyosuWakuren[11];    //�g�A�[�����v
		char		crlf[2];                  //���R�[�h��؂�
}				JV_O1_ODDS_TANFUKUWAKU;


//****** �V�D�I�b�Y�i�n�A�j****************************************
typedef struct
{
		_RECORD_ID	head;                     //<���R�[�h�w�b�_�[>
		_RACE_ID	id;                       //<�������ʏ��P>
		_MDHM		HappyoTime;               //���\��������
		char		TorokuTosu[2];            //�o�^����
		char		SyussoTosu[2];            //�o������
		char		UmarenFlag[1];            //�����t���O�@�n�A

		struct _ODDS_UMAREN_INFO                  //<�n�A�I�b�Y>
                {
			char		Kumi[4];                  //�g��
			char		Odds[6];                  //�I�b�Y
			char		Ninki[3];                 //�l�C��
		}			OddsUmarenInfo[153];

		char		TotalHyosuUmaren[11];     //�n�A�[�����v
		char		crlf[2];                  //���R�[�h��؂�
}				JV_O2_ODDS_UMAREN;

//****** �W�D�I�b�Y�i���C�h�j****************************************
typedef struct
{
		_RECORD_ID	head;                     //<���R�[�h�w�b�_�[>
		_RACE_ID	id;                       //<�������ʏ��P>
		_MDHM		HappyoTime;               //���\��������
		char		TorokuTosu[2];            //�o�^����
		char		SyussoTosu[2];            //�o������
		char		WideFlag[1];              //�����t���O ���C�h

		struct _ODDS_WIDE_INFO                    //<���C�h�I�b�Y>
                {
				char		Kumi[4];                  //�g��
				char		OddsLow[5];               //�Œ�I�b�Y
				char		OddsHigh[5];              //�ō��I�b�Y
				char		Ninki[3];                 //�l�C��
		}			OddsWideInfo[153];

		char		TotalHyosuWide[11];       //���C�h�[�����v
		char		crlf[2];                  //���R�[�h��؂�
}				JV_O3_ODDS_WIDE;


//****** �X�D�I�b�Y�i�n�P�j ****************************************
typedef struct
{
		_RECORD_ID	head;                     //<���R�[�h�w�b�_�[>
		_RACE_ID	id;                       //<�������ʏ��P>
		_MDHM		HappyoTime;               //���\��������
		char		TorokuTosu[2];            //�o�^����
		char		SyussoTosu[2];            //�o������
		char		UmatanFlag[1];            //�����t���O�@�n�P

		struct _ODDS_UMATAN_INFO                  //<�n�P�I�b�Y>
                {
				char		Kumi[4];                  //�g��
				char		Odds[6];                  //�I�b�Y
				char		Ninki[3];                 //�l�C��
		}			OddsUmatanInfo[306];

		char		TotalHyosuUmatan[11];     //�n�P�[�����v
		char		crlf[2];                  //���R�[�h��؂�
}				JV_O4_ODDS_UMATAN;


//****** �P�O�D�I�b�Y�i�R�A���j****************************************
typedef struct
{
		_RECORD_ID	head;                     //<���R�[�h�w�b�_�[>
		_RACE_ID	id;                       //<�������ʏ��P>
		_MDHM		HappyoTime;               //���\��������
		char		TorokuTosu[2];            //�o�^����
		char		SyussoTosu[2];            //�o������
		char		SanrenpukuFlag[1];        //�����t���O�@3�A��

		struct _ODDS_SANREN_INFO                  //<3�A���I�b�Y>
                {
				char		Kumi[6];              //�g��
				char		Odds[6];              //�I�b�Y
				char		Ninki[3];             //�l�C��
		}			OddsSanrenInfo[816];

		char		TotalHyosuSanrenpuku[11]; //3�A���[�����v
		char		crlf[2];                  //���R�[�h��؂�
}				JV_O5_ODDS_SANREN;


//****** �P�O�|�P�D�I�b�Y�i�R�A�P�j****************************************
typedef struct
{
		_RECORD_ID	head;                     //<���R�[�h�w�b�_�[>
		_RACE_ID	id;                       //<�������ʏ��P>
		_MDHM		HappyoTime;               //���\��������
		char		TorokuTosu[2];            //�o�^����
		char		SyussoTosu[2];            //�o������
		char		SanrentanFlag[1];         //�����t���O�@3�A�P

		struct _ODDS_SANRENTAN_INFO           //<3�A�P�I�b�Y>
                {
				char		Kumi[6];              //�g��
				char		Odds[7];              //�I�b�Y
				char		Ninki[4];             //�l�C��
		}			OddsSanrentanInfo[4896];

		char		TotalHyosuSanrentan[11];  //3�A�P�[�����v
		char		crlf[2];                  //���R�[�h��؂�
}				JV_O6_ODDS_SANRENTAN;


//****** �P�P�D�����n�}�X�^ ****************************************
typedef struct
{
		_RECORD_ID	head;            //<���R�[�h�w�b�_�[>
		char		KettoNum[10];             //�����o�^�ԍ�
		char		DelKubun[1];              //�����n�����敪
		_YMD		RegDate;                  //�����n�o�^�N����
		_YMD		DelDate;                  //�����n�����N����
		_YMD		BirthDate;                //���N����
		char		Bamei[36];                //�n��
		char		BameiKana[36];            //�n�����p�J�i
		char		BameiEng[60];             //�n������
		char		ZaikyuFlag[1];            //JRA�{�ݍ݂��イ�t���O
		char		Reserved[19];             //�\��
		char		UmaKigoCD[2];             //�n�L���R�[�h
		char		SexCD[1];                 //���ʃR�[�h
		char		HinsyuCD[1];              //�i��R�[�h
		char		KeiroCD[2];               //�ѐF�R�[�h

		struct _KETTO3_INFO                       //<�R�㌌�����>
                {
				char		HansyokuNum[10];           //�ɐB�o�^�ԍ�
				char		Bamei[36];                //�n��
		}			Ketto3Info[14];

		char		TozaiCD[1];               //���������R�[�h
		char		ChokyosiCode[5];          //�����t�R�[�h
		char		ChokyosiRyakusyo[8];      //�����t������
		char		Syotai[20];               //���Ғn�於
		char		BreederCode[8];           //���Y�҃R�[�h
		char		BreederName[72];          //���Y�Җ�
		char		SanchiName[20];           //�Y�n��
		char		BanusiCode[6];            //�n��R�[�h
		char		BanusiName[64];           //�n�喼
		char		RuikeiHonsyoHeiti[9];     //���n�{�܋��݌v
		char		RuikeiHonsyoSyogai[9];    //��Q�{�܋��݌v
		char		RuikeiFukaHeichi[9];      //���n�t���܋��݌v
		char		RuikeiFukaSyogai[9];      //��Q�t���܋��݌v
		char		RuikeiSyutokuHeichi[9];   //���n�����܋��݌v
		char		RuikeiSyutokuSyogai[9];   //��Q�����܋��݌v
		_CHAKUKAISU3_INFO		ChakuSogo;               //��������
		_CHAKUKAISU3_INFO		ChakuChuo;               //�������v����
		_CHAKUKAISU3_INFO		ChakuKaisuBa[7];         //�n��ʒ���
		_CHAKUKAISU3_INFO		ChakuKaisuJyotai[12];    //�n���ԕʒ���
		_CHAKUKAISU3_INFO		ChakuKaisuKyori[6];      //�����ʒ���
		char		Kyakusitu[4][3];          //�r���X��
		char		RaceCount[3];             //�o�^���[�X��
		char		crlf[2];                  //���R�[�h��؂�
}				JV_UM_UMA;


//****** �P�Q�D�R��}�X�^ ****************************************
typedef struct
{
		_RECORD_ID	head;                     //<���R�[�h�w�b�_�[>
		char		KisyuCode[5];             //�R��R�[�h
		char		DelKubun[1];              //�R�薕���敪
		_YMD		IssueDate;                //�R��Ƌ���t�N����
		_YMD		DelDate;                  //�R��Ƌ������N����
		_YMD		BirthDate;                //���N����
		char		KisyuName[34];            //�R�薼����
		char		reserved[34];             //�\��
		char		KisyuNameKana[30];        //�R�薼���p�J�i
		char		KisyuRyakusyo[8];         //�R�薼����
		char		KisyuNameEng[80];         //�R�薼����
		char		SexCD[1];                 //���ʋ敪
		char		SikakuCD[1];              //�R�掑�i�R�[�h
		char		MinaraiCD[1];             //�R�茩�K�R�[�h
		char		TozaiCD[1];               //�R�蓌�������R�[�h
		char		Syotai[20];               //���Ғn�於
		char		ChokyosiCode[5];          //���������t�R�[�h
		char		ChokyosiRyakusyo[8];      //���������t������

		struct _HATUKIJYO_INFO                    //<���R����>
                {
				_RACE_ID		Hatukijyoid;      //�N��������R
				char		SyussoTosu[2];            //�o������
				char		KettoNum[10];             //�����o�^�ԍ�
				char		Bamei[36];                //�n��
				char		KakuteiJyuni[2];          //�m�蒅��
				char		IJyoCD[1];                //�ُ�敪�R�[�h
		}			HatuKiJyo[2];
		

		struct _HATUSYORI_INFO                    //<���������>
                {
				_RACE_ID	Hatusyoriid;              //�N��������R
				char		SyussoTosu[2];            //�o������
				char		KettoNum[10];             //�����o�^�ԍ�
				char		Bamei[36];                //�n��
		}				HatuSyori[2];

		_SAIKIN_JYUSYO_INFO     SaikinJyusyo[3];      //<�ŋߏd�܏������>
		_HON_ZEN_RUIKEISEI_INFO	HonZenRuikei[3];      //<�{�N�E�O�N�E�݌v���я��>
		char		crlf[2];           //���R�[�h��؂�
}				JV_KS_KISYU;


//****** �P�R�D�����t�}�X�^ ****************************************
typedef struct
{
		_RECORD_ID	head;                     //<���R�[�h�w�b�_�[>
		char		ChokyosiCode[5];          //�����t�R�[�h
		char		DelKubun[1];              //�����t�����敪
		_YMD		IssueDate;                //�����t�Ƌ���t�N����
		_YMD		DelDate;                  //�����t�Ƌ������N����
		_YMD		BirthDate;                //���N����
		char		ChokyosiName[34];         //�����t������
		char		ChokyosiNameKana[30];     //�����t�����p�J�i
		char		ChokyosiRyakusyo[8];      //�����t������
		char		ChokyosiNameEng[80];      //�����t������
		char		SexCD[1];                 //���ʋ敪
		char		TozaiCD[1];               //�����t���������R�[�h
		char		Syotai[20];               //���Ғn�於
		_SAIKIN_JYUSYO_INFO     SaikinJyusyo[3];  //<�ŋߏd�܏������>
		_HON_ZEN_RUIKEISEI_INFO HonZenRuikei[3];  //<�{�N�E�O�N�E�݌v���я��>
		char		crlf[2];                  //���R�[�h��؂�
}				JV_CH_CHOKYOSI;


//******�P�S�D���Y�҃}�X�^ ****************************************
typedef struct
{
		_RECORD_ID	head;                     //<���R�[�h�w�b�_�[>
		char		BreederCode[8];           //���Y�҃R�[�h
		char		BreederName_Co[72];       //���Y�Җ�(�@�l�i�L)
		char		BreederName[72];          //���Y�Җ�(�@�l�i��)
		char		BreederNameKana[72];      //���Y�Җ����p�J�i
		char		BreederNameEng[168];      //���Y�Җ�����
		char		Address[20];              //���Y�ҏZ�������Ȗ�
		_SEI_RUIKEI_INFO        HonRuikei[2];     //<�{�N�E�݌v���я��>
		char		crlf[2];                  //���R�[�h��؂�
}				JV_BR_BREEDER;


//****** �P�T�D�n��}�X�^ ****************************************
typedef struct
{
		_RECORD_ID	head;                     //<���R�[�h�w�b�_�[>
		char		BanusiCode[6];            //�n��R�[�h
		char		BanusiName_Co[64];           //�n�喼(�@�l�i�L)
		char		BanusiName[64];           //�n�喼(�@�l�i��)
		char		BanusiNameKana[50];       //�n�喼���p�J�i
		char		BanusiNameEng[100];       //�n�喼����
		char		Fukusyoku[60];            //���F�W��
		_SEI_RUIKEI_INFO        HonRuikei[2];     //<�{�N�E�݌v���я��>
		char		crlf[2];                  //���R�[�h��؂�
}				JV_BN_BANUSI;


//****** �P�U�D�ɐB�n�}�X�^ ****************************************
typedef struct
{
		_RECORD_ID	head;                     //<���R�[�h�w�b�_�[>
		char		HansyokuNum[10];           //�ɐB�o�^�ԍ�
		char		reserved[8];              //�\��
		char		KettoNum[10];             //�����o�^�ԍ�
		char		DelKubun[1];              //�ɐB�n�����敪(���݂͗\���Ƃ��Ďg�p)
		char		Bamei[36];                //�n��
		char		BameiKana[40];            //�n�����p�J�i
		char		BameiEng[80];             //�n������
		char		BirthYear[4];             //���N
		char		SexCD[1];                 //���ʃR�[�h
		char		HinsyuCD[1];              //�i��R�[�h
		char		KeiroCD[2];               //�ѐF�R�[�h
		char		HansyokuMochiKubun[1];    //�ɐB�n�����敪
		char		ImportYear[4];            //�A���N
		char		SanchiName[20];           //�Y�n��
		char		HansyokuFNum[10];          //���n�ɐB�o�^�ԍ�
		char		HansyokuMNum[10];          //��n�ɐB�o�^�ԍ�
		char		crlf[2];                  //���R�[�h��؂�
}				JV_HN_HANSYOKU;


//****** �P�V�D�Y��}�X�^ ****************************************
typedef struct
{
		_RECORD_ID	head;                     //<���R�[�h�w�b�_�[>
		char		KettoNum[10];             //�����o�^�ԍ�
		_YMD		BirthDate;                //���N����
		char		SexCD[1];                 //���ʃR�[�h
		char		HinsyuCD[1];              //�i��R�[�h
		char		KeiroCD[2];               //�ѐF�R�[�h
		char		SankuMochiKubun[1];       //�Y����敪
		char		ImportYear[4];            //�A���N
		char		BreederCode[8];           //���Y�҃R�[�h
		char		SanchiName[20];           //�Y�n��
		char		HansyokuNum[14][10];       //3�㌌�� �ɐB�o�^�ԍ�
		char		crlf[2];                  //���R�[�h��؂�
}				JV_SK_SANKU;


//****** �P�W�D���R�[�h�}�X�^ ****************************************
typedef struct
{
		_RECORD_ID	head;                     //<���R�[�h�w�b�_�[>
		char		RecInfoKubun[1];          //���R�[�h���ʋ敪
		_RACE_ID	id;                       //<�������ʏ��P>
		char		TokuNum[4];               //���ʋ����ԍ�
		char		Hondai[60];               //�������{��
		char		GradeCD[1];               //�O���[�h�R�[�h
		char		SyubetuCD[2];             //������ʃR�[�h
		char		Kyori[4];                 //����
		char		TrackCD[2];               //�g���b�N�R�[�h
		char		RecKubun[1];              //���R�[�h�敪
		char		RecTime[4];               //���R�[�h�^�C��
		_TENKO_BABA_INFO		TenkoBaba;        //�V��E�n����

		struct _RECUMA_INFO                       //<���R�[�h�ێ��n���>
                {
				char		KettoNum[10];             //�����o�^�ԍ�
				char		Bamei[36];                //�n��
				char		UmaKigoCD[2];             //�n�L���R�[�h
				char		SexCD[1];                 //���ʃR�[�h
				char		ChokyosiCode[5];          //�����t�R�[�h
				char		ChokyosiName[34];         //�����t��
				char		Futan[3];                 //���S�d��
				char		KisyuCode[5];             //�R��R�[�h
				char		KisyuName[34];            //�R�薼
		}			RecUmaInfo[3];
		
		char		crlf[2];                   //���R�[�h��؂�
}				JV_RC_RECORD;


//****** �P�X�D��H���� ****************************************
typedef struct
{
		_RECORD_ID	head;                     //<���R�[�h�w�b�_�[>
		char		TresenKubun[1];           //�g���Z���敪
		_YMD		ChokyoDate;               //�����N����
		char		ChokyoTime[4];            //��������
		char		KettoNum[10];             //�����o�^�ԍ�
		char		HaronTime4[4];            //4�n�����^�C�����v[800M-0M]
		char		LapTime4[3];              //���b�v�^�C��[800M-600M]
		char		HaronTime3[4];            //3�n�����^�C�����v[600M-0M]
		char		LapTime3[3];              //���b�v�^�C��[600M-400M]
		char		HaronTime2[4];            //2�n�����^�C�����v[400M-0M]
		char		LapTime2[3];              //���b�v�^�C��[400M-200M]
		char		LapTime1[3];              //���b�v�^�C��[200M-0M]
		char		crlf[2];                  //���R�[�h��؂�
}				JV_HC_HANRO;


//****** �Q�O�D�n�̏d ****************************************
typedef struct
{
		_RECORD_ID	head;                     //<���R�[�h�w�b�_�[>
		_RACE_ID	id;                       //<�������ʏ��P>
		_MDHM		HappyoTime;               //���\��������

		struct _BATAIJYU_INFO                     //<�n�̏d���>
                {
				char		Umaban[2];                //�n��
				char		Bamei[36];                //�n��
				char		BaTaijyu[3];              //�n�̏d
				char		ZogenFugo[1];             //��������
				char		ZogenSa[3];               //������
		}				BataijyuInfo[18];

		char		crlf[2];                   //���R�[�h��؂�
}				JV_WH_BATAIJYU;


//****** �Q�P�D�V��n���� ******************************************
typedef struct
{
		_RECORD_ID	head;                     //<���R�[�h�w�b�_�[>
		_RACE_ID2	id;                       //<�������ʏ��Q>
		_MDHM		HappyoTime;               //���\��������
		char		HenkoID[1];               //�ύX����
		_TENKO_BABA_INFO		TenkoBaba;        //���ݏ�ԏ��
		_TENKO_BABA_INFO		TenkoBabaBefore;  //�ύX�O��ԏ��
		char		crlf[2];                  //���R�[�h��؂�
}				JV_WE_WEATHER;


//****** �Q�Q�D�o������E�������O ****************************************
typedef struct
{
		_RECORD_ID	head;                     //<���R�[�h�w�b�_�[>
		_RACE_ID	id;                       //<�������ʏ��P>
		_MDHM		HappyoTime;               //���\��������
		char		Umaban[2];                //�n��
		char		Bamei[36];                //�n��
		char		JiyuKubun[3];             //���R�敪
		char		crlf[2];                  //���R�[�h��؂�
}				JV_AV_INFO;


//************ �Q�R�D�R��ύX **************************************** 
typedef struct
{
		_RECORD_ID	head;                     //<���R�[�h�w�b�_�[>
		_RACE_ID	id;                       //<�������ʏ��P>
		_MDHM		HappyoTime;               //���\��������
		char		Umaban[2];                //�n��
		char		Bamei[36];                //�n��
		_JC_INFO 	JCInfoAfter;              //<�ύX����>
		_JC_INFO	JCInfoBefore;             //<�ύX�O���>
		char		crlf[2];                  //���R�[�h��؂�
}				JV_JC_INFO;


//************ �Q�R�|�P�D���������ύX **************************************** 
typedef struct
{
		_RECORD_ID	head;                     //<���R�[�h�w�b�_�[>
		_RACE_ID	id;                       //<�������ʏ��P>
		_MDHM		HappyoTime;               //���\��������
		_TC_INFO 	TCInfoAfter;              //<�ύX����>
		_TC_INFO	TCInfoBefore;             //<�ύX�O���>
		char		crlf[2];                  //���R�[�h��؂�
}				JV_TC_INFO;


//************ �Q�R�|�Q�D�R�[�X�ύX **************************************** 
typedef struct
{
		_RECORD_ID	head;                     //<���R�[�h�w�b�_�[>
		_RACE_ID	id;                       //<�������ʏ��P>
		_MDHM		HappyoTime;               //���\��������
		_CC_INFO 	CCInfoAfter;              //<�ύX����>
		_CC_INFO	CCInfoBefore;             //<�ύX�O���>
		char		JiyuCD[1];                //���R�R�[�h
		char		crlf[2];                  //���R�[�h��؂�
}				JV_CC_INFO;


//****** �Q�S�D�f�[�^�}�C�j���O�\�z************************************
typedef struct
{
		_RECORD_ID	head;                     //<���R�[�h�w�b�_�[>
		_RACE_ID	id;                       //<�������ʏ��P>
		_HM		MakeHM;                   //�f�[�^�쐬����

		struct _DM_INFO                           //<�}�C�j���O�\�z>
                {
				char		Umaban[2];                //�n��
				char		DMTime[5];                //�\�z���j�^�C��
				char		DMGosaP[4];               //�\�z�덷[�M���x]�{
				char		DMGosaM[4];               //�\�z�덷[�M���x]�|
		}			DMInfo[18];

		char		crlf[2];                   //���R�[�h��؂�
}				JV_DM_INFO;


//****** �Q�T�D�J�ÃX�P�W���[��************************************
typedef struct
{
               _RECORD_ID       head;                     //<���R�[�h�w�b�_�[>
               _RACE_ID2        id;                       //<�������ʏ��Q>
               char             YoubiCD[1];               //�j���R�[�h

               struct _JYUSYO_INFO                        //<�d�܈ē�>
               { 
                                char            TokuNum[4];             //���ʋ����ԍ�
                                char            Hondai[60];              //�������{��
                                char            Ryakusyo10[20];          //����������10��
                                char            Ryakusyo6[12];           //����������6��
                                char            Ryakusyo3[6];            //����������3��
                                char            Nkai[3];                 //�d�܉�[��N��]
                                char            GradeCD[1];              //�O���[�h�R�[�h
                                char            SyubetuCD[2];            //������ʃR�[�h
                                char            KigoCD[3];               //�����L���R�[�h
                                char            JyuryoCD[1];             //�d�ʎ�ʃR�[�h
                                char            Kyori[4];                //����
                                char            TrackCD[2];              //�g���b�N�R�[�h
               }                        JyusyoInfo[3];
 
               char             crlf[2];                  //���R�[�h��؂�
}				JV_YS_SCHEDULE;


//****** �Q�U�D�����n�s�������i ****************************************
typedef struct
{
		_RECORD_ID	head;                     //<���R�[�h�w�b�_�[>
		char		KettoNum[10];             //�����o�^�ԍ�
		char		HansyokuFNum[10];          //���n�ɐB�o�^�ԍ�
		char		HansyokuMNum[10];          //��n�ɐB�o�^�ԍ�
		char		BirthYear[4];             //���N
		char		SaleCode[6];              //��ÎҁE�s��R�[�h
		char		SaleHostName[40];         //��ÎҖ���
		char		SaleName[80];             //�s��̖���
		_YMD		FromDate;                 //�s��̊J�Ê���(�J�n��)
		_YMD		ToDate;                   //�s��̊J�Ê���(�I����)
		char		Barei[1];                 //������̋����n�̔N��
		char		Price[10];                //������i
		char		crlf[2];                  //���R�[�h��؂�
}				JV_HS_SALE;


//****** �Q�V�D�n���̈Ӗ��R�� ****************************************
typedef struct
{
		_RECORD_ID	head;                     //<���R�[�h�w�b�_�[>
		char		KettoNum[10];             //�����o�^�ԍ�
		char		Bamei[36];                //�n��
		char		Origin[64];               //�n���̈Ӗ��R��
		char		crlf[2];                  //���R�[�h��؂�
}				JV_HY_BAMEIORIGIN;

//****** �Q�W�D�o���ʒ��x�� ****************************************

//<�o���ʒ��x�� �����n���>
typedef struct
{
		char					KettoNum [10];            //�����o�^�ԍ�
		char					Bamei[36];                //�n��
		char					RuikeiHonsyoHeiti[9];     //���n�{�܋��݌v
		char					RuikeiHonsyoSyogai[9];    //��Q�{�܋��݌v
		char					RuikeiFukaHeichi[9];      //���n�t���܋��݌v
		char					RuikeiFukaSyogai[9];      //��Q�t���܋��݌v
		char					RuikeiSyutokuHeichi[9];   //���n�����܋��݌v
		char					RuikeiSyutokuSyogai[9];   //��Q�����܋��݌v
		_CHAKUKAISU3_INFO		ChakuSogo;                //��������
		_CHAKUKAISU3_INFO		ChakuChuo;                //�������v����
		_CHAKUKAISU3_INFO		ChakuKaisuBa[7];          //�n��ʒ���
		_CHAKUKAISU3_INFO		ChakuKaisuJyotai[12];     //�n���ԕʒ���
		_CHAKUKAISU3_INFO		ChakuKaisuSibaKyori[9];   //�ŋ����ʒ���
		_CHAKUKAISU3_INFO		ChakuKaisuDirtKyori[9];   //�_�[�g�����ʒ���
		_CHAKUKAISU3_INFO		ChakuKaisuJyoSiba[10];    //���n��ʎŒ���
		_CHAKUKAISU3_INFO		ChakuKaisuJyoDirt[10];    //���n��ʃ_�[�g����
		_CHAKUKAISU3_INFO		ChakuKaisuJyoSyogai[10];  //���n��ʏ�Q����
		char					Kyakusitu[4][3];          //�r���X��
		char					RaceCount[3];             //�o�^���[�X��
}				JV_CK_UMA;

//<�o���ʒ��x�� �{�N�E�݌v���я��>
typedef struct
{
		char					SetYear[4];               //�ݒ�N
		char					HonSyokinHeichi[10];      //���n�{�܋����v
		char					HonSyokinSyogai[10];      //��Q�{�܋����v
		char					FukaSyokinHeichi[10];     //���n�t���܋����v
		char					FukaSyokinSyogai[10];     //��Q�t���܋����v
		_CHAKUKAISU5_INFO		ChakuKaisuSiba;           //�Œ���
		_CHAKUKAISU5_INFO		ChakuKaisuDirt;           //�_�[�g����
		_CHAKUKAISU4_INFO		ChakuKaisuSyogai;         //��Q����
		_CHAKUKAISU4_INFO		ChakuKaisuSibaKyori[9];   //�ŋ����ʒ���
		_CHAKUKAISU4_INFO		ChakuKaisuDirtKyori[9];   //�_�[�g�����ʒ���
		_CHAKUKAISU4_INFO		ChakuKaisuJyoSiba[10];    //���n��ʎŒ���
		_CHAKUKAISU4_INFO		ChakuKaisuJyoDirt[10];    //���n��ʃ_�[�g����
		_CHAKUKAISU3_INFO		ChakuKaisuJyoSyogai[10];  //���n��ʏ�Q����
}				_CK_HON_RUIKEISEI_INFO;

//<�o���ʒ��x�� �R����>
typedef struct
{
		char					KisyuCode[5];             //�R��R�[�h
		char					KisyuName[34];            //�R�薼����
		_CK_HON_RUIKEISEI_INFO	HonRuikei[2];             //<�{�N�E�݌v���я��>
}				JV_CK_KISYU;

//<�o���ʒ��x�� �����t���>
typedef struct
{
		char					ChokyosiCode[5];          //�����t�R�[�h
		char					ChokyosiName[34];         //�����t������
		_CK_HON_RUIKEISEI_INFO	HonRuikei[2];             //<�{�N�E�݌v���я��>
}				JV_CK_CHOKYOSI;

//<�o���ʒ��x�� �n����>
typedef struct
{
		char					BanusiCode[6];            //�n��R�[�h
		char					BanusiName_Co[64];        //�n�喼�i�@�l�i�L�j
		char					BanusiName[64];           //�n�喼�i�@�l�i���j
		_SEI_RUIKEI_INFO		HonRuikei[2];             //<�{�N�E�݌v���я��>
}				JV_CK_BANUSI;

//<�o���ʒ��x�� ���Y�ҏ��>
typedef struct
{
		char					BreederCode[8];           //���Y�҃R�[�h
		char					BreederName_Co[72];       //���Y�Җ��i�@�l�i�L�j
		char					BreederName[72];          //���Y�Җ��i�@�l�i���j
		_SEI_RUIKEI_INFO		HonRuikei[2];             //<�{�N�E�݌v���я��>
}				JV_CK_BREEDER;

typedef struct
{
		_RECORD_ID				head;                     //<���R�[�h�w�b�_�[>
		_RACE_ID				id;                       //<�������ʏ��P>
		JV_CK_UMA				UmaChaku;                 //<�o���ʒ��x�� �����n���>
		JV_CK_KISYU				KisyuChaku;               //<�o���ʒ��x�� �R����>
		JV_CK_CHOKYOSI			ChokyoChaku;              //<�o���ʒ��x�� �����t���>
		JV_CK_BANUSI			BanusiChaku;              //<�o���ʒ��x�� �n����>
		JV_CK_BREEDER			BreederChaku;             //<�o���ʒ��x�� ���Y�ҏ��>
		char					crlf[2];                  //���R�[�h��؂�
}				JV_CK_CHAKU;

//****** �Q�X�D�n����� ****************************************
typedef struct
{
		_RECORD_ID	head;                     //<���R�[�h�w�b�_�[>
		char		HansyokuNum[10];           //�ɐB�o�^�ԍ�
		char		KeitoId[30];              //�n��ID
		char		KeitoName[36];            //�n����
		char		KeitoEx[6800];            //�n������
		char		crlf[2];                  //���R�[�h��؂�
}				JV_BT_KEITO;

//****** �R�O�D�R�[�X��� ****************************************
typedef struct
{
		_RECORD_ID	head;                     //<���R�[�h�w�b�_�[>
		char		JyoCD[2];                 //���n��R�[�h
		char		Kyori[4];                 //����
		char		TrackCD[2];               //�g���b�N�R�[�h
		_YMD		KaishuDate;               //�R�[�X���C�N����
		char		CourseEx[6800];           //�R�[�X����
		char		crlf[2];                  //���R�[�h��؂�
}				JV_CS_COURSE;

//****** �R�P�D�ΐ�^�f�[�^�}�C�j���O�\�z ****************************************
typedef struct
{
		_RECORD_ID	head;                     //<���R�[�h�w�b�_�[>
		_RACE_ID	id;                       //<�������ʏ��P>
		_HM		MakeHM;                       //�f�[�^�쐬����

		struct _TM_INFO                       //<�}�C�j���O�\�z>
                {
				char		Umaban[2];        //�n��
				char		TMScore[4];       //�\���X�R�A
		}			TMInfo[18];

		char		crlf[2];                  //���R�[�h��؂�
}				JV_TM_INFO;

//****** �R�Q�D�d����(WIN5) ****************************************
typedef struct
{
		_RECORD_ID	head;                     //<���R�[�h�w�b�_�[>
		_YMD		KaisaiDate;               //�J�ÔN����
		char		reserved1[2];             //�\��

		struct _WF_RACE_INFO
                {
                char JyoCD[2];                //���n��R�[�h
                char Kaiji[2];                //�J�É�[��N��]
                char Nichiji[2];              //�J�Ó���[N����]
                char RaceNum[2];              //���[�X�ԍ�
		}			WFRaceInfo[5];

		char 		reserved2[6];             //�\��
		char		Hatsubai_Hyo[11];         //�d���������[��

		struct _WF_YUKO_HYO_INFO
                {
                char Yuko_Hyo[11];            //�L���[��
		}			WFYukoHyoInfo[5];

		char		HenkanFlag[1];            //�Ԋ҃t���O
		char		FuseiritsuFlag[1];        //�s�����t���O
		char		TekichunashiFlag[1];      //�I�����t���O
		char		COShoki[15];              //�L�����[�I�[�o�[���z����
		char		COZanDaka[15];            //�L�����[�I�[�o�[���z�c��

		struct _WF_PAY_INFO
                {
                char Kumiban[10];             //�g��
                char Pay[9];                  //�d�������ߋ�
                char Tekichu_Hyo[10];         //�I���[��
		}			WFPayInfo[243];

		char		crlf[2];                  //���R�[�h��؂�
}				JV_WF_INFO;

//****** �R�R�D�����n���O��� ****************************************
typedef struct
{
		_RECORD_ID	head;                     //<���R�[�h�w�b�_�[>
		_RACE_ID	id;                       //<�������ʏ��P>
		char		KettoNum[10];             //�����o�^�ԍ�
		char		Bamei[36];                //�n��
		char		ShutsubaTohyoJun[3];      //�o�n���[��t����
		char		ShussoKubun[1];           //�o���敪
		char		JogaiJotaiKubun[1];       //���O��ԋ敪
		char		crlf[2];                  //���R�[�h��؂�
}				JV_JG_JOGAIBA;

//****** �R�S�D�E�b�h�`�b�v���� ****************************************
typedef struct
{
		_RECORD_ID	head;                     //<���R�[�h�w�b�_�[>
		char		TresenKubun[1];           //�g���Z���敪
		_YMD		ChokyoDate;               //�����N����
		char		ChokyoTime[4];            //��������
		char		KettoNum[10];             //�����o�^�ԍ�
		char		Course[1];                // �R�[�X
		char		BabaAround[1];            // �n�����
		char		reserved[1];              // �\��
		char		HaronTime10[4];           //10�n�����^�C�����v[2000M-0M]
		char		LapTime10[3];             //���b�v�^�C��[2000M-1800M]
		char		HaronTime9[4];            //9�n�����^�C�����v[1800M-0M]
		char		LapTime9[3];              //���b�v�^�C��[1800M-1600M]
		char		HaronTime8[4];            //8�n�����^�C�����v[1600M-0M]
		char		LapTime8[3];              //���b�v�^�C��[1600M-1400M]
		char		HaronTime7[4];            //7�n�����^�C�����v[1400M-0M]
		char		LapTime7[3];              //���b�v�^�C��[1400M-1200M]
		char		HaronTime6[4];            //6�n�����^�C�����v[1200M-0M]
		char		LapTime6[3];              //���b�v�^�C��[1200M-1000M]
		char		HaronTime5[4];            //5�n�����^�C�����v[1000M-0M]
		char		LapTime5[3];              //���b�v�^�C��[1000M-800M]
		char		HaronTime4[4];            //4�n�����^�C�����v[800M-0M]
		char		LapTime4[3];              //���b�v�^�C��[800M-600M]
		char		HaronTime3[4];            //3�n�����^�C�����v[600M-0M]
		char		LapTime3[3];              //���b�v�^�C��[600M-400M]
		char		HaronTime2[4];            //2�n�����^�C�����v[400M-0M]
		char		LapTime2[3];              //���b�v�^�C��[400M-200M]
		char		LapTime1[3];              //���b�v�^�C��[200M-0M]
		char		crlf[2];                  //���R�[�h��؂�
}				JV_WC_WOOD;


#endif	// __JV_DATA_STRUCT

