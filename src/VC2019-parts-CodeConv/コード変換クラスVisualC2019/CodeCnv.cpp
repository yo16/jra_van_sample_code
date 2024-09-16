// CodeCnv.cpp: CCodeCnv クラスのインプリメンテーション
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
		CString		strCodeNo;                   //年
		CString		strCode;                     //月
		CString		strNames;                    //日
};

CString mFileName;
bool blnFlag;								////データ読込確認フラグ
const int MAX_LINE=500;
mudtCodeLine mArrData[MAX_LINE];			////コード表の行数
long lngCt;

CCodeCnv::CCodeCnv()
{

}

CCodeCnv::~CCodeCnv()
{
}

//@(f)
//
//機能　　 : データの格納
//
//引き数　 : ARG1 - ファイル名
//
//返り値　 : なし
//
//機能説明 : 指定されたファイルのデータをメモリ上に格納する
//
void CCodeCnv::FileName(CString strFile)
{
		mFileName = strFile;
		SetData();
}


//@(f)
//
//機能　　 : 名称の取得
//
//引き数　 : ARG1 - コードNo.
//　　　　   ARG2 - コード
//
//返り値　 : 名称
//
//機能説明 : メモリ上に格納したデータをコードにより検索し名称を取得する
//
CString CCodeCnv::GetCodeName(CString strCodeNo, CString strCode, int intNo ) 
{
	int i;
	int j;
	int ct;
	CString strName;

	try{


		//データが読み込めていない場合  
		if ( blnFlag == false){
			return "";
		}
    
		//名称文字列から指定番目の名称を返す
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
//機能　　 : データを1行ずつ処理
//
//引き数　 : なし
//
//返り値　 : なし
//
//機能説明 : CSVデータを1行分ずつ区切って処理する
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
//機能　　 : 配列に格納
//
//引き数　 : ARG1 - 一行分の文字列
//　　　　 : ARG2 - 現在の行番号
//
//返り値　 : なし
//
//機能説明 : 1行分を構造体に変換して配列に格納する
//
int CCodeCnv::SetLine(CString strLine, long lngCt)
{
	try{
		byte bytFieldCt;			//フィールド（列）数
		CString strDelimiter;		//区切り子
		long lngDelimiter;			//区切り子の位置
		long lngBeforeDel;			//前の区切り子の位置
		CString strWord;			//フィールド1つ分の文字列
		mudtCodeLine udtWords;		//一行分のstrWordを格納
    
		bytFieldCt = 0;
		lngDelimiter = 0;
		lngBeforeDel = 0;


		//区切り子の決定
		strDelimiter = ",";

		//ユーザ定義型mudtCodeLineに変換
		while (bytFieldCt <= 2){
			if ( bytFieldCt < 2 )
				lngDelimiter = strLine.Find(strDelimiter,lngDelimiter + 1)+1;
			else
				lngDelimiter = strLine.GetLength() + 1;
       

			//フィールドが2以下の場合
			if (lngDelimiter == 0) {
				MessageBox(NULL,"CSVファイルが不正です",NULL,MB_OK);
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


		//ユーザ定義型mudtCodeLineを配列に代入
		mArrData[lngCt] = udtWords;

		return 0;

	}catch(...){
		return -1;
	}
}

