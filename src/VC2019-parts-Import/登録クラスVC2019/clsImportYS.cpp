/*=======================================================================
  JRA-VAN Data Lab.プログラミングパーツ「データ登録パーツ　YSレコード登録クラス」

	   作成: JRA-VAN ソフトウェア工房  2003年6月26日

========================================================================
   (C) Copyright JRA SYSTEM SERVICE CO.,LTD. 2007 All rights reserved
========================================================================*/
#include "StdAfx.h"
#include "clsImportYS.h"
#include <afxdisp.h>

clsImportYS::clsImportYS(void)
{
}

clsImportYS::~clsImportYS(void)
{
}
int clsImportYS::Add(CString strBuff, long lngBuffSize)
{
	CString strMakeDate;
    memcpy(&mBuf,strBuff.GetBuffer(lngBuffSize),lngBuffSize);
    strMakeDate.SetString(mBuf.head.MakeDate.Year,4);
	strMakeDate.Append(mBuf.head.MakeDate.Month,2);
	strMakeDate.Append(mBuf.head.MakeDate.Day,2);

//      INSERT処理
	if(InsertDB() != 0 ){
//			UPDATE処理（INSERTが失敗した場合）
		if(UpdateDB(strMakeDate)!=0){
//				System.Diagnostics.Debug.WriteLine("更新に失敗しました。" & Left(strBuf, 2))
		}
	}
	return 0;
}

int clsImportYS::InsertDB(void)
{
	_RecordsetPtr pRecordSet;
	CString strsql;
	strsql="SELECT * FROM SCHEDULE WHERE [Year]='";
	strsql.Append(mBuf.id.Year,4);
	strsql.Append("' AND [MonthDay]='");
	strsql.Append(mBuf.id.MonthDay,4);
	strsql.Append("' AND  [JyoCD]='");
	strsql.Append(mBuf.id.JyoCD,2);
	strsql.Append("' AND  [Kaiji]='");
	strsql.Append(mBuf.id.Kaiji,2);
	strsql.Append("' AND [Nichiji]='");
	strsql.Append(mBuf.id.Nichiji,2);
	strsql.Append("'");
	_bstr_t bstrQuery(strsql);

	_CommandPtr pCommand;      // Commandオブジェクト
	pCommand.CreateInstance(__uuidof(Command));
	pRecordSet.CreateInstance(__uuidof(Recordset));
	// SQLの設定
	pCommand->ActiveConnection = pCn;
	pCommand->CommandText = bstrQuery;
	pRecordSet->PutRefSource(pCommand);

	// レコードセットの取得
	_variant_t vNull;  // VARIANT型のNULLとして使う
	vNull.vt = VT_ERROR;
	vNull.scode = DISP_E_PARAMNOTFOUND;

	pRecordSet->Open(vNull, vNull, adOpenKeyset, adLockOptimistic, adCmdText);

	if (!pRecordSet->GetadoEOF()){
		pRecordSet->Close();
		pRecordSet = NULL;
		return -1;
	}
	USES_CONVERSION;

try{


	COleSafeArray vaFieldlist;
	vaFieldlist.CreateOneDim(VT_VARIANT,45);

	COleSafeArray vaValuelist;
	vaValuelist.CreateOneDim(VT_VARIANT,45);
	long lArrayIndex[1];

	lArrayIndex[0]=0;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("RecordSpec")));	
	strsql.SetString(mBuf.head.RecordSpec,2);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=1;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("DataKubun")));	
	strsql.SetString(mBuf.head.DataKubun,1);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=2;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("MakeDate")));	
	strsql.SetString(mBuf.head.MakeDate.Year,8);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=3;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Year")));	
	strsql.SetString(mBuf.id.Year,4);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=4;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("MonthDay")));	
	strsql.SetString(mBuf.id.MonthDay,4);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=5;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("JyoCD")));	
	strsql.SetString(mBuf.id.JyoCD,2);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=6;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Kaiji")));	
	strsql.SetString(mBuf.id.Kaiji,2);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=7;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Nichiji")));	
	strsql.SetString(mBuf.id.Nichiji,2);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=8;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("YoubiCD")));	
	strsql.SetString("X");
	strsql.SetString(mBuf.YoubiCD,1);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));




	int i;
	CString strsql_head;
	for(i=0;i<3;i++){
		strsql_head = "Jyusyo";
		strsql_head.AppendFormat("%d",i+1);


		lArrayIndex[0]=9+i*12;
		strsql=strsql_head;
		strsql.Append("TokuNum");
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.JyusyoInfo[i].TokuNum,4);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	

		lArrayIndex[0]=10+i*12;
		strsql=strsql_head;
		strsql.Append("Hondai");
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.JyusyoInfo[i].Hondai,60);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	

		lArrayIndex[0]=11+i*12;
		strsql=strsql_head;
		strsql.Append("Ryakusyo10");
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.JyusyoInfo[i].Ryakusyo10,20);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	


		lArrayIndex[0]=12+i*12;
		strsql=strsql_head;
		strsql.Append("Ryakusyo6");
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.JyusyoInfo[i].Ryakusyo6,12);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	


		lArrayIndex[0]=13+i*12;
		strsql=strsql_head;
		strsql.Append("Ryakusyo3");
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.JyusyoInfo[i].Ryakusyo3,6);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	

		lArrayIndex[0]=14+i*12;
		strsql=strsql_head;
		strsql.Append("Nkai");
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.JyusyoInfo[i].Nkai,3);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	

		lArrayIndex[0]=15+i*12;
		strsql=strsql_head;
		strsql.Append("GradeCD");
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.JyusyoInfo[i].GradeCD,1);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	

		lArrayIndex[0]=16+i*12;
		strsql=strsql_head;
		strsql.Append("SyubetuCD");
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.JyusyoInfo[i].SyubetuCD,2);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	

		lArrayIndex[0]=17+i*12;
		strsql=strsql_head;
		strsql.Append("KigoCD");
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.JyusyoInfo[i].KigoCD,3);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	

		lArrayIndex[0]=18+i*12;
		strsql=strsql_head;
		strsql.Append("JyuryoCD");
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.JyusyoInfo[i].JyuryoCD,1);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	

		lArrayIndex[0]=19+i*12;
		strsql=strsql_head;
		strsql.Append("Kyori");
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.JyusyoInfo[i].Kyori,4);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	

		lArrayIndex[0]=20+i*12;
		strsql=strsql_head;
		strsql.Append("TrackCD");
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.JyusyoInfo[i].TrackCD,2);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
	}
	


	pRecordSet->AddNew(vaFieldlist, vaValuelist);
	pRecordSet->Close();	

	
}
catch(_com_error &e){
	MessageBox(NULL,e.Description(),NULL,NULL);
	pRecordSet->Close();
	return -1;
}

	return 0;

}

int clsImportYS::UpdateDB(CString strMakeDate)
{

	_RecordsetPtr pRecordSet;
	_CommandPtr pCommand;      // Commandオブジェクト
	pCommand.CreateInstance(__uuidof(Command));
	pRecordSet.CreateInstance(__uuidof(Recordset));

	USES_CONVERSION;

	CString strsql;
	strsql="SELECT * FROM SCHEDULE WHERE [Year]='";
	strsql.Append(mBuf.id.Year,4);
	strsql.Append("' AND [MonthDay]='");
	strsql.Append(mBuf.id.MonthDay,4);
	strsql.Append("' AND  [JyoCD]='");
	strsql.Append(mBuf.id.JyoCD,2);
	strsql.Append("' AND  [Kaiji]='");
	strsql.Append(mBuf.id.Kaiji,2);
	strsql.Append("' AND [Nichiji]='");
	strsql.Append(mBuf.id.Nichiji,2);
	strsql.Append("' AND [MakeDate] <= '");
	strsql.Append(strMakeDate);
	strsql.Append("'");



	// SQLの設定
	pCommand->ActiveConnection = pCn;
	_bstr_t bstrQuery(strsql);
	pCommand->CommandText = bstrQuery;
	pRecordSet->PutRefSource(pCommand);

	// レコードセットの取得
	_variant_t vNull;  // VARIANT型のNULLとして使う
	vNull.vt = VT_ERROR;
	vNull.scode = DISP_E_PARAMNOTFOUND;
	pRecordSet->Open(vNull, vNull, adOpenForwardOnly, adLockOptimistic, adCmdText);

	COleSafeArray vaFieldlist;
	vaFieldlist.CreateOneDim(VT_VARIANT,45);

	COleSafeArray vaValuelist;
	vaValuelist.CreateOneDim(VT_VARIANT,45);
	long lArrayIndex[1];

	lArrayIndex[0]=0;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("RecordSpec")));	
	strsql.SetString(mBuf.head.RecordSpec,2);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=1;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("DataKubun")));	
	strsql.SetString(mBuf.head.DataKubun,1);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=2;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("MakeDate")));	
	strsql.SetString(mBuf.head.MakeDate.Year,8);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=3;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Year")));	
	strsql.SetString(mBuf.id.Year,4);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=4;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("MonthDay")));	
	strsql.SetString(mBuf.id.MonthDay,4);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=5;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("JyoCD")));	
	strsql.SetString(mBuf.id.JyoCD,2);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=6;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Kaiji")));	
	strsql.SetString(mBuf.id.Kaiji,2);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=7;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Nichiji")));	
	strsql.SetString(mBuf.id.Nichiji,2);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=8;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("YoubiCD")));	
	strsql.SetString("X");
	strsql.SetString(mBuf.YoubiCD,1);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));




	int i;
	CString strsql_head;
	for(i=0;i<3;i++){
		lArrayIndex[0]=9+i*12;

		strsql_head = "Jyusyo";
		strsql_head.AppendFormat("%d",i+1);
		strsql=strsql_head;
		strsql.Append("TokuNum");
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.JyusyoInfo[i].TokuNum,4);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	

		lArrayIndex[0]=10+i*12;
		strsql=strsql_head;
		strsql.Append("Hondai");
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.JyusyoInfo[i].Hondai,60);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	

		lArrayIndex[0]=11+i*12;
		strsql=strsql_head;
		strsql.Append("Ryakusyo10");
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.JyusyoInfo[i].Ryakusyo10,20);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	

		lArrayIndex[0]=12+i*12;
		strsql=strsql_head;
		strsql.Append("Ryakusyo6");
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.JyusyoInfo[i].Ryakusyo6,12);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	


		lArrayIndex[0]=13+i*12;
		strsql=strsql_head;
		strsql.Append("Ryakusyo3");
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.JyusyoInfo[i].Ryakusyo3,6);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	

		lArrayIndex[0]=14+i*12;
		strsql=strsql_head;
		strsql.Append("Nkai");
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.JyusyoInfo[i].Nkai,3);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	

		lArrayIndex[0]=15+i*12;
		strsql=strsql_head;
		strsql.Append("GradeCD");
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.JyusyoInfo[i].GradeCD,1);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	

		lArrayIndex[0]=16+i*12;
		strsql=strsql_head;
		strsql.Append("SyubetuCD");
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.JyusyoInfo[i].SyubetuCD,2);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	

		lArrayIndex[0]=17+i*12;
		strsql=strsql_head;
		strsql.Append("KigoCD");
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.JyusyoInfo[i].KigoCD,3);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	

		lArrayIndex[0]=18+i*12;
		strsql=strsql_head;
		strsql.Append("JyuryoCD");
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.JyusyoInfo[i].JyuryoCD,1);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	

		lArrayIndex[0]=19+i*12;
		strsql=strsql_head;
		strsql.Append("Kyori");
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.JyusyoInfo[i].Kyori,4);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	

		lArrayIndex[0]=20+i*12;
		strsql=strsql_head;
		strsql.Append("TrackCD");
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.JyusyoInfo[i].TrackCD,2);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
	}
	pRecordSet->Update(vaFieldlist, vaValuelist);
	pRecordSet->Close();	
	return 0;

}
int clsImportYS::Init(_ConnectionPtr m_Connect)
{
    pCn = m_Connect;

	 return 0;
}

