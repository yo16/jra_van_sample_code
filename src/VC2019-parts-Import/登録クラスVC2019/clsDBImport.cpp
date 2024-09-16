/*=======================================================================
  JRA-VAN Data Lab.プログラミングパーツ「データ登録パーツ　DB登録クラス」

	   作成: JRA-VAN ソフトウェア工房  2003年6月26日

========================================================================
   (C) Copyright JRA SYSTEM SERVICE CO.,LTD. 2007 All rights reserved
========================================================================*/

#include "StdAfx.h"
#include "clsDBImport.h"
#include "clsImportTK.h"
#include "clsImportRA.h"
#include "clsImportSE.h"
#include "clsImportHR.h"
#include "clsImportH1.h"
#include "clsImportH6.h"
#include "clsImportO1.h"
#include "clsimporto2.h"
#include "clsimporto3.h"
#include "clsimporto4.h"
#include "clsImportO5.h"
#include "clsImportO6.h"
#include "clsImportUM.h"
#include "clsImportKS.h"
#include "clsImportCH.h"
#include "clsImportBR.h"
#include "clsImportBN.h"
#include "clsImportRC.h"
#include "clsImportHN.h"
#include "clsImportSK.h"
#include "clsImportHC.h"
#include "clsImportWH.h"
#include "clsImportWE.h"
#include "clsImportAV.h"
#include "clsImportJC.h"
#include "clsImportTC.h"
#include "clsImportCC.h"
#include "clsImportDM.h"
#include "clsImportYS.h"
#include "clsImportHS.h"
#include "clsImportHY.h"
#include "clsImportBT.h"
#include "clsImportCS.h"



clsDBImport::clsDBImport(void)
{
 HRESULT hr;
 char currentpath[MAX_PATH];
 CString strDBline;
 try{

	GetCurrentDirectory(MAX_PATH,currentpath);
	strDBline = "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=";
	strDBline.Append(currentpath);
	strDBline.Append("\\Data.accdb");

	 hr = pCn.CreateInstance(__uuidof(Connection));
	 if(SUCCEEDED(hr)){
		hr = pCn->Open(_bstr_t(strDBline)
			, _bstr_t(L""), _bstr_t(L""), adModeUnknown);
		if(SUCCEEDED(hr) == FALSE){
		}
	 }else{
	 }
 }
	 
 catch(_com_error &e){
		MessageBox(NULL,e.Description(),NULL,NULL);
 }
 return;
}

clsDBImport::~clsDBImport(void)
{
	pCn->Close();
}

int clsDBImport::ClearData(LPCTSTR strTBLName)
{
	CString strDel;
	USES_CONVERSION;


	long option;

	option=0;

	
 try{
	 if(lstrlen(strTBLName)>0) {
		//指定したテーブルを削除する
		strDel = "DELETE * FROM ";
		strDel.Append(strTBLName);
		pCn->Execute(T2OLE(strDel),NULL,option);
	}else{
		//テーブルの内容を全て削除する
		strDel = "DELETE * FROM BANUSI";
		pCn->Execute(T2OLE(strDel),NULL,option);

		strDel = "DELETE * FROM BATAIJYU";
		pCn->Execute(T2OLE(strDel),NULL,option);
		strDel = "DELETE * FROM CHOKYO";
		pCn->Execute(T2OLE(strDel),NULL,option);
		strDel = "DELETE * FROM CHOKYO_SEISEKI";
		pCn->Execute(T2OLE(strDel),NULL,option);
		
		strDel = "DELETE * FROM HANRO";
		pCn->Execute(T2OLE(strDel),NULL,option);
		
		strDel = "DELETE * FROM HANSYOKU";
		pCn->Execute(T2OLE(strDel),NULL,option);
		
		strDel = "DELETE * FROM HARAI";
		pCn->Execute(T2OLE(strDel),NULL,option);
		
		strDel = "DELETE * FROM HYOSU";
		pCn->Execute(T2OLE(strDel),NULL,option);
		
		strDel = "DELETE * FROM HYOSU2";
		pCn->Execute(T2OLE(strDel),NULL,option);

		strDel = "DELETE * FROM HYOSU_SANRENTAN";
		pCn->Execute(T2OLE(strDel),NULL,option);

		strDel = "DELETE * FROM HYOSU_SANREN";
		pCn->Execute(T2OLE(strDel),NULL,option);
		
		strDel = "DELETE * FROM HYOSU_TANPUKU";
		pCn->Execute(T2OLE(strDel),NULL,option);
		
		strDel = "DELETE * FROM HYOSU_UMARENWIDE";
		pCn->Execute(T2OLE(strDel),NULL,option);
		
		strDel = "DELETE * FROM HYOSU_UMATAN";
		pCn->Execute(T2OLE(strDel),NULL,option);
		
		strDel = "DELETE * FROM HYOSU_WAKU";
		pCn->Execute(T2OLE(strDel),NULL,option);
		
		strDel = "DELETE * FROM KISYU";
		pCn->Execute(T2OLE(strDel),NULL,option);
		
		strDel = "DELETE * FROM KISYU_CHANGE";
		pCn->Execute(T2OLE(strDel),NULL,option);
		
		strDel = "DELETE * FROM KISYU_SEISEKI";
		pCn->Execute(T2OLE(strDel),NULL,option);
		
		strDel = "DELETE * FROM MINING";
		pCn->Execute(T2OLE(strDel),NULL,option);
		
		strDel = "DELETE * FROM ODDS_SANRENTAN";
		pCn->Execute(T2OLE(strDel),NULL,option);
		
		strDel = "DELETE * FROM ODDS_SANRENTAN_HEAD";
		pCn->Execute(T2OLE(strDel),NULL,option);

		strDel = "DELETE * FROM ODDS_SANREN";
		pCn->Execute(T2OLE(strDel),NULL,option);
		
		strDel = "DELETE * FROM ODDS_SANREN_HEAD";
		pCn->Execute(T2OLE(strDel),NULL,option);
		
		strDel = "DELETE * FROM ODDS_TANPUKU";
		pCn->Execute(T2OLE(strDel),NULL,option);
		
		strDel = "DELETE * FROM ODDS_TANPUKUWAKU_HEAD";
		pCn->Execute(T2OLE(strDel),NULL,option);
		
		strDel = "DELETE * FROM ODDS_UMAREN";
		pCn->Execute(T2OLE(strDel),NULL,option);
		
		strDel = "DELETE * FROM ODDS_UMAREN_HEAD";
		pCn->Execute(T2OLE(strDel),NULL,option);
				
		strDel = "DELETE * FROM ODDS_UMATAN";
		pCn->Execute(T2OLE(strDel),NULL,option);

		strDel = "DELETE * FROM ODDS_UMATAN_HEAD";
		pCn->Execute(T2OLE(strDel),NULL,option);

		strDel = "DELETE * FROM ODDS_WAKU";
		pCn->Execute(T2OLE(strDel),NULL,option);

		strDel = "DELETE * FROM ODDS_WIDE";
		pCn->Execute(T2OLE(strDel),NULL,option);
		
		strDel = "DELETE * FROM ODDS_WIDE_HEAD";
		pCn->Execute(T2OLE(strDel),NULL,option);
		
		strDel = "DELETE * FROM RACE";
		pCn->Execute(T2OLE(strDel),NULL,option);
		
		strDel = "DELETE * FROM RECORD";
		pCn->Execute(T2OLE(strDel),NULL,option);
		
		strDel = "DELETE * FROM SANKU";
		pCn->Execute(T2OLE(strDel),NULL,option);
		
		strDel = "DELETE * FROM SCHEDULE";
		pCn->Execute(T2OLE(strDel),NULL,option);
		
		strDel = "DELETE * FROM SEISAN";
		pCn->Execute(T2OLE(strDel),NULL,option);
		
		strDel = "DELETE * FROM TENKO_BABA";
		pCn->Execute(T2OLE(strDel),NULL,option);
		
		strDel = "DELETE * FROM TOKU";
		pCn->Execute(T2OLE(strDel),NULL,option);
		
		strDel = "DELETE * FROM TOKU_RACE";
		pCn->Execute(T2OLE(strDel),NULL,option);
		
		strDel = "DELETE * FROM TORIKESI_JYOGAI";
		pCn->Execute(T2OLE(strDel),NULL,option);
		
		strDel = "DELETE * FROM UMA";
		pCn->Execute(T2OLE(strDel),NULL,option);
		
		strDel = "DELETE * FROM UMA_RACE";
		pCn->Execute(T2OLE(strDel),NULL,option);

		strDel = "DELETE * FROM HASSOU_JIKOKU_CHANGE";
		pCn->Execute(T2OLE(strDel),NULL,option);

		strDel = "DELETE * FROM COURSE_CHANGE";
		pCn->Execute(T2OLE(strDel),NULL,option);

		strDel = "DELETE * FROM SALE";
		pCn->Execute(T2OLE(strDel),NULL,option);

		strDel = "DELETE * FROM BAMEIORIGIN";
		pCn->Execute(T2OLE(strDel),NULL,option);

		strDel = "DELETE * FROM KEITO";
		pCn->Execute(T2OLE(strDel),NULL,option);

		strDel = "DELETE * FROM COURSE";
		pCn->Execute(T2OLE(strDel),NULL,option);

	}
		
}
catch(_com_error &e)
{
	MessageBox(NULL,e.Description(),NULL,NULL);
	return -1;

}

//終了処理
return 0;
}


int clsDBImport::SetData(CString strBuff, long lngBuffSize)
{
		CString strRecIDCur;
		int retval;
		retval=0;
		//レコード種別IDを取得
		strRecIDCur = strBuff.Left(2);
		if(lstrcmp(strRecIDCur,"TK")==0){
				clsImportTK ImportObj;
				ImportObj.Init(pCn);
				retval=ImportObj.Add(strBuff, lngBuffSize);
		}else if(lstrcmp(strRecIDCur,"RA")==0){
				clsImportRA ImportObj;
				ImportObj.Init(pCn);
			 	retval=ImportObj.Add(strBuff, lngBuffSize);
		}else if(lstrcmp(strRecIDCur,"SE")==0){
				clsImportSE ImportObj;
				ImportObj.Init(pCn);
				retval=ImportObj.Add(strBuff, lngBuffSize);
		}else if(lstrcmp(strRecIDCur,"HR")==0){
				clsImportHR ImportObj;
				ImportObj.Init(pCn);
				retval=ImportObj.Add(strBuff, lngBuffSize);
		}else if(lstrcmp(strRecIDCur,"H1")==0){
				clsImportH1 ImportObj;
				ImportObj.Init(pCn);
				retval=ImportObj.Add(strBuff, lngBuffSize);
		}else if(lstrcmp(strRecIDCur,"H6")==0){
				clsImportH6 ImportObj;
				ImportObj.Init(pCn);
				retval=ImportObj.Add(strBuff, lngBuffSize);
		}else if(lstrcmp(strRecIDCur,"O1")==0){
				clsImportO1 ImportObj;
				ImportObj.Init(pCn);
				retval=ImportObj.Add(strBuff, lngBuffSize);
		}else if(lstrcmp(strRecIDCur,"O2")==0){
				clsImportO2 ImportObj;
				ImportObj.Init(pCn);
				retval=ImportObj.Add(strBuff, lngBuffSize);
		}else if(lstrcmp(strRecIDCur,"O3")==0){
				clsImportO3 ImportObj;
				ImportObj.Init(pCn);
				retval=ImportObj.Add(strBuff, lngBuffSize);
		}else if(lstrcmp(strRecIDCur,"O4")==0){
				clsImportO4 ImportObj;
				ImportObj.Init(pCn);
				retval=ImportObj.Add(strBuff, lngBuffSize);
		}else if(lstrcmp(strRecIDCur,"O5")==0){
				clsImportO5 ImportObj;
				ImportObj.Init(pCn);
				retval=ImportObj.Add(strBuff, lngBuffSize);
		}else if(lstrcmp(strRecIDCur,"O6")==0){
				clsImportO6 ImportObj;
				ImportObj.Init(pCn);
				retval=ImportObj.Add(strBuff, lngBuffSize);
		}else if(lstrcmp(strRecIDCur,"UM")==0){
				clsImportUM ImportObj;
				ImportObj.Init(pCn);
				retval=ImportObj.Add(strBuff, lngBuffSize);
		}else if(lstrcmp(strRecIDCur,"KS")==0){
				clsImportKS ImportObj;
				ImportObj.Init(pCn);
				retval=ImportObj.Add(strBuff, lngBuffSize);
		}else if(lstrcmp(strRecIDCur,"CH")==0){
				clsImportCH ImportObj;
				ImportObj.Init(pCn);
				retval=ImportObj.Add(strBuff, lngBuffSize);
		}else if(lstrcmp(strRecIDCur,"BR")==0){
				clsImportBR ImportObj;
				ImportObj.Init(pCn);
				retval=ImportObj.Add(strBuff, lngBuffSize);
		}else if(lstrcmp(strRecIDCur,"BN")==0){
				clsImportBN ImportObj;
				ImportObj.Init(pCn);
				retval=ImportObj.Add(strBuff, lngBuffSize);
		}else if(lstrcmp(strRecIDCur,"RC")==0){
				clsImportRC ImportObj;
				ImportObj.Init(pCn);
				retval=ImportObj.Add(strBuff, lngBuffSize);
		}else if(lstrcmp(strRecIDCur,"HN")==0){
				clsImportHN ImportObj;
				ImportObj.Init(pCn);
				retval=ImportObj.Add(strBuff, lngBuffSize);
		}else if(lstrcmp(strRecIDCur,"SK")==0){
				clsImportSK ImportObj;
				ImportObj.Init(pCn);
				retval=ImportObj.Add(strBuff, lngBuffSize);
		}else if(lstrcmp(strRecIDCur,"HC")==0){
				clsImportHC ImportObj;
				ImportObj.Init(pCn);
				retval=ImportObj.Add(strBuff, lngBuffSize);
		}else if(lstrcmp(strRecIDCur,"WH")==0){
				clsImportWH ImportObj;
				ImportObj.Init(pCn);
				retval=ImportObj.Add(strBuff, lngBuffSize);
		}else if(lstrcmp(strRecIDCur,"WE")==0){
				clsImportWE ImportObj;
				ImportObj.Init(pCn);
				retval=ImportObj.Add(strBuff, lngBuffSize);
		}else if(lstrcmp(strRecIDCur,"AV")==0){
				clsImportAV ImportObj;
				ImportObj.Init(pCn);
				retval=ImportObj.Add(strBuff, lngBuffSize);
		}else if(lstrcmp(strRecIDCur,"JC")==0){
				clsImportJC ImportObj;
				ImportObj.Init(pCn);
				retval=ImportObj.Add(strBuff, lngBuffSize);
		}else if(lstrcmp(strRecIDCur,"TC")==0){
				clsImportTC ImportObj;
				ImportObj.Init(pCn);
				retval=ImportObj.Add(strBuff, lngBuffSize);
		}else if(lstrcmp(strRecIDCur,"CC")==0){
				clsImportCC ImportObj;
				ImportObj.Init(pCn);
				retval=ImportObj.Add(strBuff, lngBuffSize);
		}else if(lstrcmp(strRecIDCur,"DM")==0){
				clsImportDM ImportObj;
				ImportObj.Init(pCn);
				retval=ImportObj.Add(strBuff, lngBuffSize);
		}else if(lstrcmp(strRecIDCur,"YS")==0){
				clsImportYS ImportObj;
				ImportObj.Init(pCn);
				retval=ImportObj.Add(strBuff, lngBuffSize);
		}else if(lstrcmp(strRecIDCur,"HS")==0){
				clsImportHS ImportObj;
				ImportObj.Init(pCn);
				retval=ImportObj.Add(strBuff, lngBuffSize);
		}else if(lstrcmp(strRecIDCur,"HY")==0){
				clsImportHY ImportObj;
				ImportObj.Init(pCn);
				retval=ImportObj.Add(strBuff, lngBuffSize);
		}else if(lstrcmp(strRecIDCur,"BT")==0){
				clsImportBT ImportObj;
				ImportObj.Init(pCn);
				retval=ImportObj.Add(strBuff, lngBuffSize);
		}else if(lstrcmp(strRecIDCur,"CS")==0){
				clsImportCS ImportObj;
				ImportObj.Init(pCn);
				retval=ImportObj.Add(strBuff, lngBuffSize);
		}
		
		return retval;
}



void clsDBImport::BeginTrans(void)
{
	pCn->BeginTrans();
}

void clsDBImport::CommitTrans(void)
{
	pCn->CommitTrans();
}

void clsDBImport::RollbackTrans(void)
{
	pCn->RollbackTrans();
}
