/*=======================================================================
  JRA-VAN Data Lab.プログラミングパーツ「データ登録パーツ　O1レコード登録クラス」

	   作成: JRA-VAN ソフトウェア工房  2003年6月26日

========================================================================
   (C) Copyright JRA SYSTEM SERVICE CO.,LTD. 2007 All rights reserved
========================================================================*/

#include "StdAfx.h"
#include "clsimporto1.h"

clsImportO1::clsImportO1(void)
{
}

clsImportO1::~clsImportO1(void)
{
}
int clsImportO1::Add(CString strBuff, long lngBuffSize)
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

int clsImportO1::InsertDB(void)
{

	int i;
	CString strsql;

	strsql="SELECT * FROM ODDS_TANPUKUWAKU_HEAD WHERE [Year]='";
	strsql.Append(mBuf.id.Year,4);
	strsql.Append("' AND [MonthDay]='");
	strsql.Append(mBuf.id.MonthDay,4);
	strsql.Append("' AND  [JyoCD]='");
	strsql.Append(mBuf.id.JyoCD,2);
	strsql.Append("' AND  [Kaiji]='");
	strsql.Append(mBuf.id.Kaiji,2);
	strsql.Append("' AND [Nichiji]='");
	strsql.Append(mBuf.id.Nichiji,2);
	strsql.Append("' AND [RaceNum]='");
	strsql.Append(mBuf.id.RaceNum,2);
	strsql.Append("'");

	_bstr_t bstrQuery(strsql);
    // Commandオブジェクト
	_CommandPtr pCommand;
	_RecordsetPtr pRecordSet;

	pCommand.CreateInstance(__uuidof(Command));
	pRecordSet.CreateInstance(__uuidof(Recordset));

	// SQLの設定
	pCommand->ActiveConnection = pCn;
	pCommand->CommandText = bstrQuery;
	pRecordSet->PutRefSource(pCommand);


	_variant_t vNull;  // VARIANT型のNULLとして使う
	vNull.vt = VT_ERROR;
	vNull.scode = DISP_E_PARAMNOTFOUND;

	// レコードセットの取得
	pRecordSet->Open(vNull, vNull, adOpenKeyset, adLockOptimistic, adCmdText);

	if (!pRecordSet->GetadoEOF()){
		pRecordSet->Close();
		pRecordSet = NULL;
		return -1;
	}

	USES_CONVERSION;

	COleSafeArray vaFieldlist;
	vaFieldlist.CreateOneDim(VT_VARIANT,19);

	COleSafeArray vaValuelist;
	vaValuelist.CreateOneDim(VT_VARIANT,19);
	long lArrayIndex[1];

	
try{

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
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("RaceNum")));	
	strsql.SetString(mBuf.id.RaceNum,2);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));


	lArrayIndex[0]=9;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("HappyoTime")));	
	strsql.SetString(mBuf.HappyoTime.Month,8);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=10;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("TorokuTosu")));	
	strsql.SetString(mBuf.TorokuTosu,2);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=11;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("SyussoTosu")));	
	strsql.SetString(mBuf.SyussoTosu,2);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=12;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("TansyoFlag")));	
	strsql.SetString(mBuf.TansyoFlag,1);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=13;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("FukusyoFlag")));	
	strsql.SetString(mBuf.FukusyoFlag,1);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=14;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("WakurenFlag")));	
	strsql.SetString(mBuf.WakurenFlag,1);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=15;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("FukuChakuBaraiKey")));	
	strsql.SetString(mBuf.FukuChakuBaraiKey,1);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=16;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("TotalHyosuTansyo")));	
	strsql.SetString(mBuf.TotalHyosuTansyo,11);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=17;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("TotalHyosuFukusyo")));	
	strsql.SetString(mBuf.TotalHyosuFukusyo,11);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=18;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("TotalHyosuWakuren")));	
	strsql.SetString(mBuf.TotalHyosuWakuren,11);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	pRecordSet->AddNew(vaFieldlist, vaValuelist);
	pRecordSet->Close();



	if(mBuf.TansyoFlag[0]!='0'){
		strsql="SELECT * FROM ODDS_TANPUKU WHERE [Year] IS NULL";
		bstrQuery=strsql;

		// SQLの設定
		pCommand->ActiveConnection = pCn;
		pCommand->CommandText = bstrQuery;
		pRecordSet->PutRefSource(pCommand);

		// レコードセットの取得
		pRecordSet->Open(vNull, vNull, adOpenKeyset, adLockOptimistic, adCmdText);

		vaFieldlist.CreateOneDim(VT_VARIANT,13);
		vaValuelist.CreateOneDim(VT_VARIANT,13);
		for(i=0;i<28;i++){
			if(strncmp(mBuf.OddsTansyoInfo[i].Umaban,"  ",2)!=0){

				lArrayIndex[0]=0;
				vaFieldlist.PutElement(lArrayIndex, &(_variant_t("MakeDate")));	
				strsql.SetString(mBuf.head.MakeDate.Year,8);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

				lArrayIndex[0]=1;
				vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Year")));	
				strsql.SetString(mBuf.id.Year,4);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

				lArrayIndex[0]=2;
				vaFieldlist.PutElement(lArrayIndex, &(_variant_t("MonthDay")));	
				strsql.SetString(mBuf.id.MonthDay,4);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

				lArrayIndex[0]=3;
				vaFieldlist.PutElement(lArrayIndex, &(_variant_t("JyoCD")));	
				strsql.SetString(mBuf.id.JyoCD,2);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

				lArrayIndex[0]=4;
				vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Kaiji")));	
				strsql.SetString(mBuf.id.Kaiji,2);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

				lArrayIndex[0]=5;
				vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Nichiji")));	
				strsql.SetString(mBuf.id.Nichiji,2);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

				lArrayIndex[0]=6;
				vaFieldlist.PutElement(lArrayIndex, &(_variant_t("RaceNum")));	
				strsql.SetString(mBuf.id.RaceNum,2);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

				lArrayIndex[0]=7;
				vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Umaban")));	
				strsql.SetString(mBuf.OddsTansyoInfo[i].Umaban,2);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

				lArrayIndex[0]=8;
				vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Umaban")));	
				strsql.SetString(mBuf.OddsTansyoInfo[i].Odds,4);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

				lArrayIndex[0]=9;
				vaFieldlist.PutElement(lArrayIndex, &(_variant_t("TanNinki")));	
				strsql.SetString(mBuf.OddsTansyoInfo[i].Ninki,2);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

				lArrayIndex[0]=10;
				vaFieldlist.PutElement(lArrayIndex, &(_variant_t("FukuOddsLow")));	
				strsql.SetString(mBuf.OddsFukusyoInfo[i].OddsLow,4);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

				lArrayIndex[0]=11;
				vaFieldlist.PutElement(lArrayIndex, &(_variant_t("FukuOddsHigh")));	
				strsql.SetString(mBuf.OddsFukusyoInfo[i].OddsHigh,4);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

				lArrayIndex[0]=12;
				vaFieldlist.PutElement(lArrayIndex, &(_variant_t("FukuNinki")));	
				strsql.SetString(mBuf.OddsFukusyoInfo[i].Ninki,2);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));


				pRecordSet->AddNew(vaFieldlist, vaValuelist);
			}
		}
		pRecordSet->Close();
	}
	if(mBuf.TansyoFlag[0]!='0'){
		strsql="SELECT * FROM ODDS_WAKU WHERE [Year] IS NULL";
		bstrQuery=strsql;

		// SQLの設定
		pCommand->ActiveConnection = pCn;
		pCommand->CommandText = bstrQuery;
		pRecordSet->PutRefSource(pCommand);

		// レコードセットの取得
		pRecordSet->Open(vNull, vNull, adOpenKeyset, adLockOptimistic, adCmdText);

		vaFieldlist.CreateOneDim(VT_VARIANT,10);
		vaValuelist.CreateOneDim(VT_VARIANT,10);

		for(i=0;i<36;i++){
			if(strncmp(mBuf.OddsWakurenInfo[i].Kumi,"  ",2)!=0){

				lArrayIndex[0]=0;
				vaFieldlist.PutElement(lArrayIndex, &(_variant_t("MakeDate")));	
				strsql.SetString(mBuf.head.MakeDate.Year,8);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

				lArrayIndex[0]=1;
				vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Year")));	
				strsql.SetString(mBuf.id.Year,4);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

				lArrayIndex[0]=2;
				vaFieldlist.PutElement(lArrayIndex, &(_variant_t("MonthDay")));	
				strsql.SetString(mBuf.id.MonthDay,4);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

				lArrayIndex[0]=3;
				vaFieldlist.PutElement(lArrayIndex, &(_variant_t("JyoCD")));	
				strsql.SetString(mBuf.id.JyoCD,2);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

				lArrayIndex[0]=4;
				vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Kaiji")));	
				strsql.SetString(mBuf.id.Kaiji,2);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

				lArrayIndex[0]=5;
				vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Nichiji")));	
				strsql.SetString(mBuf.id.Nichiji,2);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

				lArrayIndex[0]=6;
				vaFieldlist.PutElement(lArrayIndex, &(_variant_t("RaceNum")));	
				strsql.SetString(mBuf.id.RaceNum,2);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

				lArrayIndex[0]=7;
				vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Kumi")));	
				strsql.SetString(mBuf.OddsWakurenInfo[i].Kumi,2);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

				lArrayIndex[0]=8;
				vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Odds")));	
				strsql.SetString(mBuf.OddsWakurenInfo[i].Odds,4);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

				lArrayIndex[0]=9;
				vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Ninki")));	
				strsql.SetString(mBuf.OddsWakurenInfo[i].Ninki,2);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

				pRecordSet->AddNew(vaFieldlist, vaValuelist);

			}
		}
		pRecordSet->Close();
	}

	

}
	 catch(_com_error &e){
		MessageBox(NULL,e.Description(),NULL,NULL);
		pRecordSet->Close();
		  return -1;
	 }




	return 0;
}

int clsImportO1::UpdateDB(CString strMakeDate)
{
	int i;
	CString strsql;

	strsql="SELECT * FROM ODDS_TANPUKUWAKU_HEAD WHERE [Year]='";
	strsql.Append(mBuf.id.Year,4);
	strsql.Append("' AND [MonthDay]='");
	strsql.Append(mBuf.id.MonthDay,4);
	strsql.Append("' AND  [JyoCD]='");
	strsql.Append(mBuf.id.JyoCD,2);
	strsql.Append("' AND  [Kaiji]='");
	strsql.Append(mBuf.id.Kaiji,2);
	strsql.Append("' AND [Nichiji]='");
	strsql.Append(mBuf.id.Nichiji,2);
	strsql.Append("' AND [RaceNum]='");
	strsql.Append(mBuf.id.RaceNum,2);
	strsql.Append("' AND [MakeDate] <= '");
	strsql.Append(strMakeDate);
	strsql.Append("'");


	_bstr_t bstrQuery(strsql);
    // Commandオブジェクト
	_CommandPtr pCommand;
	_RecordsetPtr pRecordSet;

	pCommand.CreateInstance(__uuidof(Command));
	pRecordSet.CreateInstance(__uuidof(Recordset));

	// SQLの設定
	pCommand->ActiveConnection = pCn;
	pCommand->CommandText = bstrQuery;
	pRecordSet->PutRefSource(pCommand);


	_variant_t vNull;  // VARIANT型のNULLとして使う
	vNull.vt = VT_ERROR;
	vNull.scode = DISP_E_PARAMNOTFOUND;

	// レコードセットの取得
	pRecordSet->Open(vNull, vNull, adOpenKeyset, adLockOptimistic, adCmdText);

	USES_CONVERSION;

	COleSafeArray vaFieldlist;
	vaFieldlist.CreateOneDim(VT_VARIANT,19);

	COleSafeArray vaValuelist;
	vaValuelist.CreateOneDim(VT_VARIANT,19);
	long lArrayIndex[1];

	
try{

	strsql="INSERT INTO ODDS_TANPUKUWAKU_HEAD WHERE [Year] IS NULL";	
	
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
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("RaceNum")));	
	strsql.SetString(mBuf.id.RaceNum,2);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));


	lArrayIndex[0]=9;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("HappyoTime")));	
	strsql.SetString(mBuf.HappyoTime.Month,8);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=10;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("TorokuTosu")));	
	strsql.SetString(mBuf.TorokuTosu,2);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=11;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("SyussoTosu")));	
	strsql.SetString(mBuf.SyussoTosu,2);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=12;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("TansyoFlag")));	
	strsql.SetString(mBuf.TansyoFlag,1);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=13;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("FukusyoFlag")));	
	strsql.SetString(mBuf.FukusyoFlag,1);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=14;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("WakurenFlag")));	
	strsql.SetString(mBuf.WakurenFlag,1);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=15;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("FukuChakuBaraiKey")));	
	strsql.SetString(mBuf.FukuChakuBaraiKey,1);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=16;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("TotalHyosuTansyo")));	
	strsql.SetString(mBuf.TotalHyosuTansyo,11);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=17;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("TotalHyosuFukusyo")));	
	strsql.SetString(mBuf.TotalHyosuFukusyo,11);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=18;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("TotalHyosuWakuren")));	
	strsql.SetString(mBuf.TotalHyosuWakuren,11);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	pRecordSet->Update(vaFieldlist, vaValuelist);
	pRecordSet->Close();

	strsql = "DELETE FROM ODDS_TANPUKU";
	strsql.Append(" WHERE [Year]='");
	strsql.Append(mBuf.id.Year,4);
	strsql.Append("' AND [MonthDay] = '");
	strsql.Append(mBuf.id.MonthDay,4);
	strsql.Append("' AND [JyoCD] = '");
	strsql.Append(mBuf.id.JyoCD,2);
	strsql.Append("' AND [Kaiji] = '");
	strsql.Append(mBuf.id.Kaiji,2);
	strsql.Append("' AND [Nichiji] = '");
	strsql.Append(mBuf.id.Nichiji,2);
	strsql.Append("' AND [RaceNum]= '");
	strsql.Append(mBuf.id.RaceNum,2);
	strsql.Append("' AND [MakeDate]<= '");
	strsql.Append(strMakeDate);
	strsql.Append("'");

	bstrQuery=strsql;
	_variant_t vRecsAffected(0L);
	pCn->Execute(bstrQuery,&vRecsAffected,adOptionUnspecified);


	strsql = "DELETE FROM ODDS_WAKU";
	strsql.Append(" WHERE [Year]='");
	strsql.Append(mBuf.id.Year,4);
	strsql.Append("' AND [MonthDay] = '");
	strsql.Append(mBuf.id.MonthDay,4);
	strsql.Append("' AND [JyoCD] = '");
	strsql.Append(mBuf.id.JyoCD,2);
	strsql.Append("' AND [Kaiji] = '");
	strsql.Append(mBuf.id.Kaiji,2);
	strsql.Append("' AND [Nichiji] = '");
	strsql.Append(mBuf.id.Nichiji,2);
	strsql.Append("' AND [RaceNum]= '");
	strsql.Append(mBuf.id.RaceNum,2);
	strsql.Append("' AND [MakeDate]<= '");
	strsql.Append(strMakeDate);
	strsql.Append("'");

	bstrQuery=strsql;
	pCn->Execute(bstrQuery,&vRecsAffected,adOptionUnspecified);



	if(mBuf.TansyoFlag[0]!='0'){
				strsql="SELECT * FROM ODDS_TANPUKU WHERE [Year] IS NULL";
				bstrQuery=strsql;

				// SQLの設定
				pCommand->ActiveConnection = pCn;
				pCommand->CommandText = bstrQuery;
				pRecordSet->PutRefSource(pCommand);

				// レコードセットの取得
				pRecordSet->Open(vNull, vNull, adOpenKeyset, adLockOptimistic, adCmdText);

				vaFieldlist.CreateOneDim(VT_VARIANT,13);
				vaValuelist.CreateOneDim(VT_VARIANT,13);
		for(i=0;i<28;i++){
			if(strncmp(mBuf.OddsTansyoInfo[i].Umaban,"  ",2)!=0){

				lArrayIndex[0]=0;
				vaFieldlist.PutElement(lArrayIndex, &(_variant_t("MakeDate")));	
				strsql.SetString(mBuf.head.MakeDate.Year,8);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

				lArrayIndex[0]=1;
				vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Year")));	
				strsql.SetString(mBuf.id.Year,4);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

				lArrayIndex[0]=2;
				vaFieldlist.PutElement(lArrayIndex, &(_variant_t("MonthDay")));	
				strsql.SetString(mBuf.id.MonthDay,4);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

				lArrayIndex[0]=3;
				vaFieldlist.PutElement(lArrayIndex, &(_variant_t("JyoCD")));	
				strsql.SetString(mBuf.id.JyoCD,2);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

				lArrayIndex[0]=4;
				vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Kaiji")));	
				strsql.SetString(mBuf.id.Kaiji,2);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

				lArrayIndex[0]=5;
				vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Nichiji")));	
				strsql.SetString(mBuf.id.Nichiji,2);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

				lArrayIndex[0]=6;
				vaFieldlist.PutElement(lArrayIndex, &(_variant_t("RaceNum")));	
				strsql.SetString(mBuf.id.RaceNum,2);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

				lArrayIndex[0]=7;
				vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Umaban")));	
				strsql.SetString(mBuf.OddsTansyoInfo[i].Umaban,2);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

				lArrayIndex[0]=8;
				vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Umaban")));	
				strsql.SetString(mBuf.OddsTansyoInfo[i].Odds,4);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

				lArrayIndex[0]=9;
				vaFieldlist.PutElement(lArrayIndex, &(_variant_t("TanNinki")));	
				strsql.SetString(mBuf.OddsTansyoInfo[i].Ninki,2);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

				lArrayIndex[0]=10;
				vaFieldlist.PutElement(lArrayIndex, &(_variant_t("FukuOddsLow")));	
				strsql.SetString(mBuf.OddsFukusyoInfo[i].OddsLow,4);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

				lArrayIndex[0]=11;
				vaFieldlist.PutElement(lArrayIndex, &(_variant_t("FukuOddsHigh")));	
				strsql.SetString(mBuf.OddsFukusyoInfo[i].OddsHigh,4);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

				lArrayIndex[0]=12;
				vaFieldlist.PutElement(lArrayIndex, &(_variant_t("FukuNinki")));	
				strsql.SetString(mBuf.OddsFukusyoInfo[i].Ninki,2);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

				pRecordSet->AddNew(vaFieldlist, vaValuelist);


			}
		}
				pRecordSet->Close();
	}
	if(mBuf.TansyoFlag[0]!='0'){
				strsql="SELECT * FROM ODDS_WAKU WHERE [Year] IS NULL";
				bstrQuery=strsql;

				// SQLの設定
				pCommand->ActiveConnection = pCn;
				pCommand->CommandText = bstrQuery;
				pRecordSet->PutRefSource(pCommand);

				// レコードセットの取得
				pRecordSet->Open(vNull, vNull, adOpenKeyset, adLockOptimistic, adCmdText);

				vaFieldlist.CreateOneDim(VT_VARIANT,10);
				vaValuelist.CreateOneDim(VT_VARIANT,10);
		for(i=0;i<36;i++){
			if(strncmp(mBuf.OddsWakurenInfo[i].Kumi,"  ",2)!=0){

				lArrayIndex[0]=0;
				vaFieldlist.PutElement(lArrayIndex, &(_variant_t("MakeDate")));	
				strsql.SetString(mBuf.head.MakeDate.Year,8);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

				lArrayIndex[0]=1;
				vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Year")));	
				strsql.SetString(mBuf.id.Year,4);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

				lArrayIndex[0]=2;
				vaFieldlist.PutElement(lArrayIndex, &(_variant_t("MonthDay")));	
				strsql.SetString(mBuf.id.MonthDay,4);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

				lArrayIndex[0]=3;
				vaFieldlist.PutElement(lArrayIndex, &(_variant_t("JyoCD")));	
				strsql.SetString(mBuf.id.JyoCD,2);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

				lArrayIndex[0]=4;
				vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Kaiji")));	
				strsql.SetString(mBuf.id.Kaiji,2);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

				lArrayIndex[0]=5;
				vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Nichiji")));	
				strsql.SetString(mBuf.id.Nichiji,2);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

				lArrayIndex[0]=6;
				vaFieldlist.PutElement(lArrayIndex, &(_variant_t("RaceNum")));	
				strsql.SetString(mBuf.id.RaceNum,2);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

				lArrayIndex[0]=7;
				vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Kumi")));	
				strsql.SetString(mBuf.OddsWakurenInfo[i].Kumi,2);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

				lArrayIndex[0]=8;
				vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Odds")));	
				strsql.SetString(mBuf.OddsWakurenInfo[i].Odds,4);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

				lArrayIndex[0]=9;
				vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Ninki")));	
				strsql.SetString(mBuf.OddsWakurenInfo[i].Ninki,2);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

				pRecordSet->AddNew(vaFieldlist, vaValuelist);

			}
		}
				pRecordSet->Close();
	}

	

}
	 catch(_com_error &e){
		MessageBox(NULL,e.Description(),NULL,NULL);
		pRecordSet->Close();
		  return -1;
	 }




	return 0;
}
int clsImportO1::Init(_ConnectionPtr m_Connect)
{
    pCn = m_Connect;

	 return 0;
}

