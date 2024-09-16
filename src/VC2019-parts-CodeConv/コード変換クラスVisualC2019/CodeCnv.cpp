// CodeCnv.cpp: CCodeCnv �N���X�̃C���v�������e�[�V����
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "Project1.h"
#include "CodeCnv.h"

//#include <string>
//using namespace std;

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

struct mudtCodeLine
{
		CString		strCodeNo;                   //�N
		CString		strCode;                     //��
		CString		strNames;                    //��
};

CString mFileName;
bool blnFlag;								////�f�[�^�Ǎ��m�F�t���O
const int MAX_LINE=500;
mudtCodeLine mArrData[MAX_LINE];			////�R�[�h�\�̍s��
long lngCt;

CCodeCnv::CCodeCnv()
{

}

CCodeCnv::~CCodeCnv()
{
}

//@(f)
//
//�@�\�@�@ : �f�[�^�̊i�[
//
//�������@ : ARG1 - �t�@�C����
//
//�Ԃ�l�@ : �Ȃ�
//
//�@�\���� : �w�肳�ꂽ�t�@�C���̃f�[�^����������Ɋi�[����
//
void CCodeCnv::FileName(CString strFile)
{
		mFileName = strFile;
		SetData();
}


//@(f)
//
//�@�\�@�@ : ���̂̎擾
//
//�������@ : ARG1 - �R�[�hNo.
//�@�@�@�@   ARG2 - �R�[�h
//
//�Ԃ�l�@ : ����
//
//�@�\���� : ��������Ɋi�[�����f�[�^���R�[�h�ɂ�茟�������̂��擾����
//
CString CCodeCnv::GetCodeName(CString strCodeNo, CString strCode, int intNo ) 
{
	int i;
	int j;
	int ct;
	CString strName;

	try{


		//�f�[�^���ǂݍ��߂Ă��Ȃ��ꍇ  
		if ( blnFlag == false){
			return "";
		}
    
		//���̕����񂩂�w��Ԗڂ̖��̂�Ԃ�
		for (i=0;i<=lngCt;i++){
			if (( mArrData[i].strCodeNo == strCodeNo) && (mArrData[i].strCode == strCode)) {
				ct = 1;
				for (j=1; j <= mArrData[i].strNames.GetLength();j++){
					if (mArrData[i].strNames.Mid(j-1, 1) == ",") {
						ct = ct + 1;
						if (ct > intNo) break;
					}else if (ct == intNo) {
						strName += mArrData[i].strNames.Mid(j-1, 1);
					}
				}
				break;
			}
		}
		return strName;
		
	}catch(...){    
		return "";
	}
}



//@(f)
//
//�@�\�@�@ : �f�[�^��1�s������
//
//�������@ : �Ȃ�
//
//�Ԃ�l�@ : �Ȃ�
//
//�@�\���� : CSV�f�[�^��1�s������؂��ď�������
//
int CCodeCnv::SetData()
{

	lngCt = 0;

	try{

		CString    strTemp;
		CStdioFile objFile;
		BOOL       bolEnd;
		CString    strResult;
		strResult = "";
		strTemp   = "";
		if (objFile.Open(mFileName, CFile::modeRead, NULL))
		{
			while (bolEnd = objFile.ReadString(strTemp), bolEnd)
			{
				strResult = strTemp + "\r\n";
				SetLine(strResult,lngCt);      
				lngCt = lngCt + 1;
			}
			objFile.Close();
		}

		blnFlag = true;
		return 0;

	}catch(...){
		blnFlag = false;
		return -1;

	}
}

//@(f)
//
//�@�\�@�@ : �z��Ɋi�[
//
//�������@ : ARG1 - ��s���̕�����
//�@�@�@�@ : ARG2 - ���݂̍s�ԍ�
//
//�Ԃ�l�@ : �Ȃ�
//
//�@�\���� : 1�s�����\���̂ɕϊ����Ĕz��Ɋi�[����
//
int CCodeCnv::SetLine(CString strLine, long lngCt)
{
	try{
		byte bytFieldCt;			//�t�B�[���h�i��j��
		CString strDelimiter;		//��؂�q
		long lngDelimiter;			//��؂�q�̈ʒu
		long lngBeforeDel;			//�O�̋�؂�q�̈ʒu
		CString strWord;			//�t�B�[���h1���̕�����
		mudtCodeLine udtWords;		//��s����strWord���i�[
    
		bytFieldCt = 0;
		lngDelimiter = 0;
		lngBeforeDel = 0;


		//��؂�q�̌���
		strDelimiter = ",";

		//���[�U��`�^mudtCodeLine�ɕϊ�
		while (bytFieldCt <= 2){
			if ( bytFieldCt < 2 )
				lngDelimiter = strLine.Find(strDelimiter,lngDelimiter + 1)+1;
			else
				lngDelimiter = strLine.GetLength() + 1;
       

			//�t�B�[���h��2�ȉ��̏ꍇ
			if (lngDelimiter == 0) {
				MessageBox(NULL,"CSV�t�@�C�����s���ł�",NULL,MB_OK);
				blnFlag = false;
				return -1;
			}

			strWord = strLine.Mid(lngBeforeDel , lngDelimiter - lngBeforeDel - 1);
        

			if (bytFieldCt==0)
				udtWords.strCodeNo = strWord;
			else if (bytFieldCt==1)
				udtWords.strCode = strWord;
			else if (bytFieldCt==2)
				udtWords.strNames = strWord;
			else
				return -1;

        
			bytFieldCt = bytFieldCt + 1;
			lngBeforeDel = lngDelimiter;
		}


		//���[�U��`�^mudtCodeLine��z��ɑ��
		mArrData[lngCt] = udtWords;

		return 0;

	}catch(...){
		return -1;
	}
}

