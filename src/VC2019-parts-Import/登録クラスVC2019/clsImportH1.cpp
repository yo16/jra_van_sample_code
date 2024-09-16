/*=======================================================================
  JRA-VAN Data Lab.プログラミングパーツ「データ登録パーツ　H1レコード登録クラス」

	   作成: JRA-VAN ソフトウェア工房  2003年6月26日

========================================================================
   (C) Copyright JRA SYSTEM SERVICE CO.,LTD. 2007 All rights reserved
========================================================================*/

#include "StdAfx.h"
#include "clsimporth1.h"

clsImportH1::clsImportH1(void)
{
}

clsImportH1::~clsImportH1(void)
{
}
int clsImportH1::Add(CString strBuff, long lngBuffSize)
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

int clsImportH1::InsertDB(void)
{
	int i;
	CString strsql;

	strsql="SELECT * FROM HYOSU WHERE [Year]='";
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
	vaFieldlist.CreateOneDim(VT_VARIANT,77);

	COleSafeArray vaValuelist;
	vaValuelist.CreateOneDim(VT_VARIANT,77);
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
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("TorokuTosu")));	
	strsql.SetString(mBuf.TorokuTosu,2);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=10;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("SyussoTosu")));	
	strsql.SetString(mBuf.SyussoTosu,2);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	

	for(i=0;i<7;i++){
		lArrayIndex[0]=11+i;
		strsql.Format("HatubaiFlag%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
        strsql.SetString(mBuf.HatubaiFlag[i],1);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	

	}
	lArrayIndex[0]=18;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("FukuChakuBaraiKey")));	
	strsql.SetString(mBuf.FukuChakuBaraiKey,1);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	

	for(i=0;i<28;i++){
		lArrayIndex[0]=19+i;
		strsql.Format("HenkanUma%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.HenkanUma[i],1);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	

	}
	for(i=0;i<8;i++){
		lArrayIndex[0]=47+i;
		strsql.Format("HenkanWaku%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.HenkanWaku[i],1);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	

	}
	for(i=0;i<8;i++){
		lArrayIndex[0]=55+i;
		strsql.Format("HenkanWaku%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.HenkanDoWaku[i],1);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	

	}
	for(i=0;i<14;i++){
		lArrayIndex[0]=63+i;
		strsql.Format("HyoTotal%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.HyoTotal[i],11);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	

	}
	pRecordSet->AddNew(vaFieldlist, vaValuelist);
	pRecordSet->Close();	



		strsql="SELECT * FROM HYOSU_TANPUKU WHERE [Year] IS NULL";
		bstrQuery=strsql;

		// SQLの設定
		pCommand->ActiveConnection = pCn;
		pCommand->CommandText = bstrQuery;
		pRecordSet->PutRefSource(pCommand);

		// レコードセットの取得
		pRecordSet->Open(vNull, vNull, adOpenKeyset, adLockOptimistic, adCmdText);
		vaFieldlist.CreateOneDim(VT_VARIANT,12);
		vaValuelist.CreateOneDim(VT_VARIANT,12);
		for(i=0;i<28;i++){
			if(strncmp(mBuf.HyoTansyo[i].Umaban,"  ",2)!=0){

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
				strsql.SetString(mBuf.HyoTansyo[i].Umaban,2);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	

				lArrayIndex[0]=8;
				vaFieldlist.PutElement(lArrayIndex, &(_variant_t("TanHyo")));	
				strsql.SetString(mBuf.HyoTansyo[i].Hyo,11);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	

				lArrayIndex[0]=9;
				vaFieldlist.PutElement(lArrayIndex, &(_variant_t("TanNinki")));	
				strsql.SetString(mBuf.HyoTansyo[i].Ninki,2);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	

				lArrayIndex[0]=10;
				vaFieldlist.PutElement(lArrayIndex, &(_variant_t("FukuHyo")));	
				strsql.SetString(mBuf.HyoFukusyo[i].Hyo,11);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	

				lArrayIndex[0]=11;
				vaFieldlist.PutElement(lArrayIndex, &(_variant_t("FukuNinki")));	
				strsql.SetString(mBuf.HyoFukusyo[i].Ninki,2);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
				pRecordSet->AddNew(vaFieldlist, vaValuelist);
			}
		}
		pRecordSet->Close();	

		strsql="SELECT * FROM HYOSU_WAKU WHERE [Year] IS NULL";
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
			if(strncmp(mBuf.HyoWakuren[i].Umaban,"  ",2)!=0){

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
				strsql.SetString(mBuf.HyoWakuren[i].Umaban,2);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	

				lArrayIndex[0]=8;
				vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Hyo")));	
				strsql.SetString(mBuf.HyoWakuren[i].Hyo,11);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	

				lArrayIndex[0]=9;
				vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Ninki")));	
				strsql.SetString(mBuf.HyoWakuren[i].Ninki,2);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
				pRecordSet->AddNew(vaFieldlist, vaValuelist);
			}
		}
		pRecordSet->Close();	

		strsql="SELECT * FROM HYOSU_UMARENWIDE WHERE [Year] IS NULL";
		bstrQuery=strsql;
		// SQLの設定
		pCommand->ActiveConnection = pCn;
		pCommand->CommandText = bstrQuery;
		pRecordSet->PutRefSource(pCommand);

		// レコードセットの取得
		pRecordSet->Open(vNull, vNull, adOpenKeyset, adLockOptimistic, adCmdText);
		vaFieldlist.CreateOneDim(VT_VARIANT,12);
		vaValuelist.CreateOneDim(VT_VARIANT,12);
		for(i=0;i<153;i++){
			if(strncmp(mBuf.HyoUmaren[i].Kumi,"    ",4)!=0){

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
				strsql.SetString(mBuf.HyoUmaren[i].Kumi,4);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	

				lArrayIndex[0]=8;
				vaFieldlist.PutElement(lArrayIndex, &(_variant_t("UmarenHyo")));	
				strsql.SetString(mBuf.HyoUmaren[i].Hyo,11);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	

				lArrayIndex[0]=9;
				vaFieldlist.PutElement(lArrayIndex, &(_variant_t("UmarenNinki")));	
				strsql.SetString(mBuf.HyoUmaren[i].Ninki,3);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	

				lArrayIndex[0]=10;
				vaFieldlist.PutElement(lArrayIndex, &(_variant_t("WideHyo")));	
				strsql.SetString(mBuf.HyoWide[i].Hyo,11);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	

				lArrayIndex[0]=11;
				vaFieldlist.PutElement(lArrayIndex, &(_variant_t("WideNinki")));	
				strsql.SetString(mBuf.HyoWide[i].Ninki,3);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
				pRecordSet->AddNew(vaFieldlist, vaValuelist);


			}
		}
		pRecordSet->Close();	
		strsql="SELECT * FROM HYOSU_UMATAN WHERE [Year] IS NULL";
		bstrQuery=strsql;
		// SQLの設定
		pCommand->ActiveConnection = pCn;
		pCommand->CommandText = bstrQuery;
		pRecordSet->PutRefSource(pCommand);

		// レコードセットの取得
		pRecordSet->Open(vNull, vNull, adOpenKeyset, adLockOptimistic, adCmdText);
		vaFieldlist.CreateOneDim(VT_VARIANT,10);
		vaValuelist.CreateOneDim(VT_VARIANT,10);
		for(i=0;i<306;i++){
			if(strncmp(mBuf.HyoUmatan[i].Kumi,"    ",4)!=0){

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
				strsql.SetString(mBuf.HyoUmatan[i].Kumi,4);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	

				lArrayIndex[0]=8;
				vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Hyo")));	
				strsql.SetString(mBuf.HyoUmatan[i].Hyo,11);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	

				lArrayIndex[0]=9;
				vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Ninki")));	
				strsql.SetString(mBuf.HyoUmatan[i].Ninki,3);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	

				pRecordSet->AddNew(vaFieldlist, vaValuelist);

			}
		}
		pRecordSet->Close();	
		strsql="SELECT * FROM HYOSU_SANREN WHERE [Year] IS NULL";
		bstrQuery=strsql;
		// SQLの設定
		pCommand->ActiveConnection = pCn;
		pCommand->CommandText = bstrQuery;
		pRecordSet->PutRefSource(pCommand);

		// レコードセットの取得
		pRecordSet->Open(vNull, vNull, adOpenKeyset, adLockOptimistic, adCmdText);
		vaFieldlist.CreateOneDim(VT_VARIANT,10);
		vaValuelist.CreateOneDim(VT_VARIANT,10);
		for(i=0;i<816;i++){
			if(strncmp(mBuf.HyoSanrenpuku[i].Kumi,"      ",6)!=0){

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
				strsql.SetString(mBuf.HyoSanrenpuku[i].Kumi,6);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	

				lArrayIndex[0]=8;
				vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Hyo")));	
				strsql.SetString(mBuf.HyoSanrenpuku[i].Hyo,11);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	

				lArrayIndex[0]=9;
				vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Ninki")));	
				strsql.SetString(mBuf.HyoSanrenpuku[i].Ninki,3);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
				pRecordSet->AddNew(vaFieldlist, vaValuelist);

			}
		}
		pRecordSet->Close();	


	
}
	 catch(_com_error &e){
		MessageBox(NULL,e.Description(),NULL,NULL);
		pRecordSet->Close();	
		return -1;
	 }


	return 0;
}

int clsImportH1::UpdateDB(CString strMakeDate)
{

	int i;
	CString strsql;

	strsql="SELECT * FROM HYOSU WHERE [Year]='";
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
	vaFieldlist.CreateOneDim(VT_VARIANT,77);

	COleSafeArray vaValuelist;
	vaValuelist.CreateOneDim(VT_VARIANT,77);
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
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("TorokuTosu")));	
	strsql.SetString(mBuf.TorokuTosu,2);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=10;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("SyussoTosu")));	
	strsql.SetString(mBuf.SyussoTosu,2);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	

	for(i=0;i<7;i++){
		lArrayIndex[0]=11+i;
		strsql.Format("HatubaiFlag%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
        strsql.SetString(mBuf.HatubaiFlag[i],1);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	

	}
	lArrayIndex[0]=18;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("FukuChakuBaraiKey")));	
	strsql.SetString(mBuf.FukuChakuBaraiKey,1);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	

	for(i=0;i<28;i++){
		lArrayIndex[0]=19+i;
		strsql.Format("HenkanUma%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.HenkanUma[i],1);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	

	}
	for(i=0;i<8;i++){
		lArrayIndex[0]=47+i;
		strsql.Format("HenkanWaku%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.HenkanWaku[i],1);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	

	}
	for(i=0;i<8;i++){
		lArrayIndex[0]=55+i;
		strsql.Format("HenkanWaku%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.HenkanDoWaku[i],1);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	

	}
	for(i=0;i<14;i++){
		lArrayIndex[0]=63+i;
		strsql.Format("HyoTotal%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.HyoTotal[i],11);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	

	}
	pRecordSet->Update(vaFieldlist, vaValuelist);
	pRecordSet->Close();	

	if(mBuf.HatubaiFlag[0][1]!=(CHAR)'0'){
		for(i=0;i<28;i++){
			if(strncmp(mBuf.HyoTansyo[i].Umaban,"  ",2)!=0){
				strsql="SELECT * FROM HYOSU_TANPUKU WHERE [Year]='";
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
				strsql.Append("' AND [UmaBan]='");
				strsql.Append(mBuf.HyoTansyo[i].Umaban,2);
				strsql.Append("'");
				bstrQuery=strsql;

				// SQLの設定
				pCommand->ActiveConnection = pCn;
				pCommand->CommandText = bstrQuery;
				pRecordSet->PutRefSource(pCommand);

				// レコードセットの取得
				pRecordSet->Open(vNull, vNull, adOpenKeyset, adLockOptimistic, adCmdText);
				vaFieldlist.CreateOneDim(VT_VARIANT,12);
				vaValuelist.CreateOneDim(VT_VARIANT,12);

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
				strsql.SetString(mBuf.HyoTansyo[i].Umaban,2);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	

				lArrayIndex[0]=8;
				vaFieldlist.PutElement(lArrayIndex, &(_variant_t("TanHyo")));	
				strsql.SetString(mBuf.HyoTansyo[i].Hyo,11);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	

				lArrayIndex[0]=9;
				vaFieldlist.PutElement(lArrayIndex, &(_variant_t("TanNinki")));	
				strsql.SetString(mBuf.HyoTansyo[i].Ninki,2);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	

				lArrayIndex[0]=10;
				vaFieldlist.PutElement(lArrayIndex, &(_variant_t("FukuHyo")));	
				strsql.SetString(mBuf.HyoFukusyo[i].Hyo,11);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	

				lArrayIndex[0]=11;
				vaFieldlist.PutElement(lArrayIndex, &(_variant_t("FukuNinki")));	
				strsql.SetString(mBuf.HyoFukusyo[i].Ninki,2);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
				pRecordSet->Update(vaFieldlist, vaValuelist);
				pRecordSet->Close();	
			}
		}
	}

	if(mBuf.HatubaiFlag[2][1]!=(CHAR)'0'){
		for(i=0;i<36;i++){
			if(strncmp(mBuf.HyoWakuren[i].Umaban,"  ",2)!=0){
				strsql="SELECT * FROM HYOSU_WAKU WHERE [Year]='";
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
				strsql.Append("' AND [Kumi]='");
				strsql.Append(mBuf.HyoWakuren[i].Umaban,2);
				strsql.Append("'");
				bstrQuery=strsql;
				// SQLの設定
				pCommand->ActiveConnection = pCn;
				pCommand->CommandText = bstrQuery;
				pRecordSet->PutRefSource(pCommand);

				// レコードセットの取得
				pRecordSet->Open(vNull, vNull, adOpenKeyset, adLockOptimistic, adCmdText);
				vaFieldlist.CreateOneDim(VT_VARIANT,10);
				vaValuelist.CreateOneDim(VT_VARIANT,10);

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
				strsql.SetString(mBuf.HyoWakuren[i].Umaban,2);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	

				lArrayIndex[0]=8;
				vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Hyo")));	
				strsql.SetString(mBuf.HyoWakuren[i].Hyo,11);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	

				lArrayIndex[0]=9;
				vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Ninki")));	
				strsql.SetString(mBuf.HyoWakuren[i].Ninki,2);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
				pRecordSet->Update(vaFieldlist, vaValuelist);
				pRecordSet->Close();	
			}
		}
	}

	if(mBuf.HatubaiFlag[3][1]!=(CHAR)'0'){
		for(i=0;i<153;i++){
			if(strncmp(mBuf.HyoUmaren[i].Kumi,"    ",4)!=0){
				strsql="SELECT * FROM HYOSU_UMARENWIDE WHERE [Year]='";
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
				strsql.Append("' AND [Kumi]='");
				strsql.Append(mBuf.HyoUmaren[i].Kumi,4);
				strsql.Append("'");
				bstrQuery=strsql;

				// SQLの設定
				pCommand->ActiveConnection = pCn;
				pCommand->CommandText = bstrQuery;
				pRecordSet->PutRefSource(pCommand);

				// レコードセットの取得
				pRecordSet->Open(vNull, vNull, adOpenKeyset, adLockOptimistic, adCmdText);
				vaFieldlist.CreateOneDim(VT_VARIANT,12);
				vaValuelist.CreateOneDim(VT_VARIANT,12);

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
				strsql.SetString(mBuf.HyoUmaren[i].Kumi,4);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	

				lArrayIndex[0]=8;
				vaFieldlist.PutElement(lArrayIndex, &(_variant_t("UmarenHyo")));	
				strsql.SetString(mBuf.HyoUmaren[i].Hyo,11);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	

				lArrayIndex[0]=9;
				vaFieldlist.PutElement(lArrayIndex, &(_variant_t("UmarenNinki")));	
				strsql.SetString(mBuf.HyoUmaren[i].Ninki,3);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	

				lArrayIndex[0]=10;
				vaFieldlist.PutElement(lArrayIndex, &(_variant_t("WideHyo")));	
				strsql.SetString(mBuf.HyoWide[i].Hyo,11);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	

				lArrayIndex[0]=11;
				vaFieldlist.PutElement(lArrayIndex, &(_variant_t("WideNinki")));	
				strsql.SetString(mBuf.HyoWide[i].Ninki,3);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
				pRecordSet->Update(vaFieldlist, vaValuelist);
				pRecordSet->Close();	

			}
		}
	}
	if(mBuf.HatubaiFlag[5][1]!=(CHAR)'0'){
		for(i=0;i<306;i++){
			if(strncmp(mBuf.HyoUmatan[i].Kumi,"    ",4)!=0){
				strsql="SELECT * FROM HYOSU_UMATAN WHERE [Year]='";
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
				strsql.Append("' AND [Kumi]='");
				strsql.Append(mBuf.HyoUmatan[i].Kumi,4);
				strsql.Append("'");
				bstrQuery=strsql;

				// SQLの設定
				pCommand->ActiveConnection = pCn;
				pCommand->CommandText = bstrQuery;
				pRecordSet->PutRefSource(pCommand);

				// レコードセットの取得
				pRecordSet->Open(vNull, vNull, adOpenKeyset, adLockOptimistic, adCmdText);
				vaFieldlist.CreateOneDim(VT_VARIANT,10);
				vaValuelist.CreateOneDim(VT_VARIANT,10);

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
				strsql.SetString(mBuf.HyoUmatan[i].Kumi,4);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	

				lArrayIndex[0]=8;
				vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Hyo")));	
				strsql.SetString(mBuf.HyoUmatan[i].Hyo,11);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	

				lArrayIndex[0]=9;
				vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Ninki")));	
				strsql.SetString(mBuf.HyoUmatan[i].Ninki,3);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	

				pRecordSet->Update(vaFieldlist, vaValuelist);
				pRecordSet->Close();	

			}
		}
	}
	if(mBuf.HatubaiFlag[6][1]!=(CHAR)'0'){
		for(i=0;i<816;i++){
			if(strncmp(mBuf.HyoSanrenpuku[i].Kumi,"      ",6)!=0){

				strsql="SELECT * FROM HYOSU_SANREN WHERE [Year]='";
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
				strsql.Append("' AND [Kumi]='");
				strsql.Append(mBuf.HyoSanrenpuku[i].Kumi,6);
				strsql.Append("'");
				bstrQuery=strsql;
				// SQLの設定
				pCommand->ActiveConnection = pCn;
				pCommand->CommandText = bstrQuery;
				pRecordSet->PutRefSource(pCommand);

				// レコードセットの取得
				pRecordSet->Open(vNull, vNull, adOpenKeyset, adLockOptimistic, adCmdText);
				vaFieldlist.CreateOneDim(VT_VARIANT,10);
				vaValuelist.CreateOneDim(VT_VARIANT,10);

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
				strsql.SetString(mBuf.HyoSanrenpuku[i].Kumi,6);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	

				lArrayIndex[0]=8;
				vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Hyo")));	
				strsql.SetString(mBuf.HyoSanrenpuku[i].Hyo,11);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	

				lArrayIndex[0]=9;
				vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Ninki")));	
				strsql.SetString(mBuf.HyoSanrenpuku[i].Ninki,3);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
				pRecordSet->Update(vaFieldlist, vaValuelist);
				pRecordSet->Close();	
			}
		}
	}
}
	 catch(_com_error &e){
		MessageBox(NULL,e.Description(),NULL,NULL);
		pRecordSet->Close();
		return -1;
	 }


	return 0;

}

int clsImportH1::Init(_ConnectionPtr m_Connect)
{
    pCn = m_Connect;

	 return 0;
}

