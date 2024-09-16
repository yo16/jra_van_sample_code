/*=======================================================================
  JRA-VAN Data Lab.プログラミングパーツ「データ登録パーツ　SEレコード登録クラス」

	   作成: JRA-VAN ソフトウェア工房  2003年6月26日

========================================================================
   (C) Copyright JRA SYSTEM SERVICE CO.,LTD. 2007 All rights reserved
========================================================================*/
#include "StdAfx.h"
#include "clsimportse.h"

clsImportSE::clsImportSE(void)
{
}

clsImportSE::~clsImportSE(void)
{
}

int clsImportSE::Add(CString strBuff, long lngBuffSize)
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

int clsImportSE::InsertDB(void)
{

	CString strsql;

	strsql="SELECT * FROM UMA_RACE WHERE [Year]='";
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
	strsql.Append("' AND [Umaban]='");
	strsql.Append(mBuf.Umaban,2);
	strsql.Append("' AND [KettoNum]='");
	strsql.Append(mBuf.KettoNum,10);
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
	vaFieldlist.CreateOneDim(VT_VARIANT,73);

	COleSafeArray vaValuelist;
	vaValuelist.CreateOneDim(VT_VARIANT,73);
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
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Wakuban")));	
	strsql.SetString(mBuf.Wakuban,1);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=10;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Umaban")));	
	strsql.SetString(mBuf.Umaban,2);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=11;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("KettoNum")));	
	strsql.SetString(mBuf.KettoNum,10);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=12;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Bamei")));	
	strsql.SetString(mBuf.Bamei,36);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=13;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("UmaKigoCD")));	
	strsql.SetString(mBuf.UmaKigoCD,2);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=14;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("SexCD")));	
	strsql.SetString(mBuf.SexCD,1);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=15;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("HinsyuCD")));	
	strsql.SetString(mBuf.HinsyuCD,1);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=16;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("KeiroCD")));	
	strsql.SetString(mBuf.KeiroCD,2);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=17;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Barei")));	
	strsql.SetString(mBuf.Barei,2);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=18;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("TozaiCD")));	
	strsql.SetString(mBuf.TozaiCD,1);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=19;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("ChokyosiCode")));	
	strsql.SetString(mBuf.ChokyosiCode,5);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=20;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("ChokyosiRyakusyo")));	
	strsql.SetString(mBuf.ChokyosiRyakusyo,8);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=21;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("BanusiCode")));	
	strsql.SetString(mBuf.BanusiCode,6);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=22;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("BanusiName")));	
	strsql.SetString(mBuf.BanusiName,64);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=23;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Fukusyoku")));	
	strsql.SetString(mBuf.Fukusyoku,60);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=24;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("reserved1")));	
	strsql.SetString(mBuf.reserved1,60);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=25;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Futan")));	
	strsql.SetString(mBuf.Futan,3);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=26;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("FutanBefore")));	
	strsql.SetString(mBuf.FutanBefore,3);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=27;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Blinker")));	
	strsql.SetString(mBuf.Blinker,1);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=28;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("reserved2")));	
	strsql.SetString(mBuf.reserved2,1);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=29;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("KisyuCode")));	
	strsql.SetString(mBuf.KisyuCode,5);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=30;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("KisyuCodeBefore")));	
	strsql.SetString(mBuf.KisyuCodeBefore,5);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=31;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("KisyuRyakusyo")));	
	strsql.SetString(mBuf.KisyuRyakusyo,8);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=32;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("KisyuRyakusyoBefore")));	
	strsql.SetString(mBuf.KisyuRyakusyoBefore,8);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=33;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("MinaraiCD")));	
	strsql.SetString(mBuf.MinaraiCD,1);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=34;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("MinaraiCDBefore")));	
	strsql.SetString(mBuf.MinaraiCDBefore,1);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=35;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("BaTaijyu")));	
	strsql.SetString(mBuf.BaTaijyu,3);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=36;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("ZogenFugo")));	
	strsql.SetString(mBuf.ZogenFugo,1);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=37;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("ZogenSa")));	
	strsql.SetString(mBuf.ZogenSa,3);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=38;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("IJyoCD")));	
	strsql.SetString(mBuf.IJyoCD,1);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=39;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("NyusenJyuni")));	
	strsql.SetString(mBuf.NyusenJyuni,2);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=40;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("KakuteiJyuni")));	
	strsql.SetString(mBuf.KakuteiJyuni,2);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=41;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("DochakuKubun")));	
	strsql.SetString(mBuf.DochakuKubun,1);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=42;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("DochakuTosu")));	
	strsql.SetString(mBuf.DochakuTosu,1);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=43;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Time")));	
	strsql.SetString(mBuf.Time,4);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=44;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("ChakusaCD")));	
	strsql.SetString(mBuf.ChakusaCD,3);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=45;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("ChakusaCDP")));	
	strsql.SetString(mBuf.ChakusaCDP,3);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=46;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("ChakusaCDP")));	
	strsql.SetString(mBuf.ChakusaCDP,3);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=47;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Jyuni1c")));	
	strsql.SetString(mBuf.Jyuni1c,2);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=48;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Jyuni2c")));	
	strsql.SetString(mBuf.Jyuni2c,2);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=49;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Jyuni3c")));	
	strsql.SetString(mBuf.Jyuni3c,2);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=50;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Jyuni4c")));	
	strsql.SetString(mBuf.Jyuni4c,2);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=51;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Odds")));	
	strsql.SetString(mBuf.Odds,4);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=52;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Ninki")));	
	strsql.SetString(mBuf.Ninki,2);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=53;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Honsyokin")));	
	strsql.SetString(mBuf.Honsyokin,8);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=54;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Fukasyokin")));	
	strsql.SetString(mBuf.Fukasyokin,8);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=55;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("reserved3")));	
	strsql.SetString(mBuf.reserved3,3);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=56;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("reserved4")));	
	strsql.SetString(mBuf.reserved4,3);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=57;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("HaronTimeL4")));	
	strsql.SetString(mBuf.HaronTimeL4,3);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=58;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("HaronTimeL3")));	
	strsql.SetString(mBuf.HaronTimeL3,3);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));


	int i;
	for(i=0;i<3;i++){
		lArrayIndex[0]=59+i*2;
		strsql.Format("KettoNum%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.ChakuUmaInfo[0].KettoNum,10);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=60+i*2;
		strsql.Format("Bamei%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.ChakuUmaInfo[0].Bamei,36);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));
	}

	lArrayIndex[0]=65;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("TimeDiff")));	
	strsql.SetString(mBuf.TimeDiff,4);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=66;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("RecordUpKubun")));	
	strsql.SetString(mBuf.RecordUpKubun,1);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=67;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("DMKubun")));	
	strsql.SetString(mBuf.DMKubun,1);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=68;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("DMTime")));	
	strsql.SetString(mBuf.DMTime,5);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=69;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("DMGosaP")));	
	strsql.SetString(mBuf.DMGosaP,4);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=70;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("DMGosaM")));	
	strsql.SetString(mBuf.DMGosaM,4);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=71;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("DMJyuni")));	
	strsql.SetString(mBuf.DMJyuni,2);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=72;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("KyakusituKubun")));	
	strsql.SetString(mBuf.KyakusituKubun,1);
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

int clsImportSE::UpdateDB(CString strMakeDate)
{


	CString strsql;

	strsql="SELECT * FROM UMA_RACE WHERE [Year]='";
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
	strsql.Append("' AND [Umaban]='");
	strsql.Append(mBuf.Umaban,2);
	strsql.Append("' AND [KettoNum]='");
	strsql.Append(mBuf.KettoNum,10);
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
	vaFieldlist.CreateOneDim(VT_VARIANT,73);

	COleSafeArray vaValuelist;
	vaValuelist.CreateOneDim(VT_VARIANT,73);
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
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Wakuban")));	
	strsql.SetString(mBuf.Wakuban,1);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=10;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Umaban")));	
	strsql.SetString(mBuf.Umaban,2);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=11;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("KettoNum")));	
	strsql.SetString(mBuf.KettoNum,10);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=12;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Bamei")));	
	strsql.SetString(mBuf.Bamei,36);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=13;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("UmaKigoCD")));	
	strsql.SetString(mBuf.UmaKigoCD,2);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=14;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("SexCD")));	
	strsql.SetString(mBuf.SexCD,1);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=15;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("HinsyuCD")));	
	strsql.SetString(mBuf.HinsyuCD,1);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=16;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("KeiroCD")));	
	strsql.SetString(mBuf.KeiroCD,2);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=17;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Barei")));	
	strsql.SetString(mBuf.Barei,2);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=18;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("TozaiCD")));	
	strsql.SetString(mBuf.TozaiCD,1);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=19;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("ChokyosiCode")));	
	strsql.SetString(mBuf.ChokyosiCode,5);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=20;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("ChokyosiRyakusyo")));	
	strsql.SetString(mBuf.ChokyosiRyakusyo,8);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=21;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("BanusiCode")));	
	strsql.SetString(mBuf.BanusiCode,6);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=22;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("BanusiName")));	
	strsql.SetString(mBuf.BanusiName,64);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=23;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Fukusyoku")));	
	strsql.SetString(mBuf.Fukusyoku,60);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=24;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("reserved1")));	
	strsql.SetString(mBuf.reserved1,60);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=25;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Futan")));	
	strsql.SetString(mBuf.Futan,3);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=26;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("FutanBefore")));	
	strsql.SetString(mBuf.FutanBefore,3);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=27;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Blinker")));	
	strsql.SetString(mBuf.Blinker,1);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=28;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("reserved2")));	
	strsql.SetString(mBuf.reserved2,1);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=29;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("KisyuCode")));	
	strsql.SetString(mBuf.KisyuCode,5);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=30;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("KisyuCodeBefore")));	
	strsql.SetString(mBuf.KisyuCodeBefore,5);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=31;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("KisyuRyakusyo")));	
	strsql.SetString(mBuf.KisyuRyakusyo,8);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=32;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("KisyuRyakusyoBefore")));	
	strsql.SetString(mBuf.KisyuRyakusyoBefore,8);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=33;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("MinaraiCD")));	
	strsql.SetString(mBuf.MinaraiCD,1);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=34;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("MinaraiCDBefore")));	
	strsql.SetString(mBuf.MinaraiCDBefore,1);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=35;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("BaTaijyu")));	
	strsql.SetString(mBuf.BaTaijyu,3);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=36;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("ZogenFugo")));	
	strsql.SetString(mBuf.ZogenFugo,1);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=37;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("ZogenSa")));	
	strsql.SetString(mBuf.ZogenSa,3);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=38;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("IJyoCD")));	
	strsql.SetString(mBuf.IJyoCD,1);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=39;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("NyusenJyuni")));	
	strsql.SetString(mBuf.NyusenJyuni,2);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=40;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("KakuteiJyuni")));	
	strsql.SetString(mBuf.KakuteiJyuni,2);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=41;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("DochakuKubun")));	
	strsql.SetString(mBuf.DochakuKubun,1);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=42;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("DochakuTosu")));	
	strsql.SetString(mBuf.DochakuTosu,1);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=43;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Time")));	
	strsql.SetString(mBuf.Time,4);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=44;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("ChakusaCD")));	
	strsql.SetString(mBuf.ChakusaCD,3);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=45;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("ChakusaCDP")));	
	strsql.SetString(mBuf.ChakusaCDP,3);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=46;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("ChakusaCDP")));	
	strsql.SetString(mBuf.ChakusaCDP,3);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=47;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Jyuni1c")));	
	strsql.SetString(mBuf.Jyuni1c,2);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=48;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Jyuni2c")));	
	strsql.SetString(mBuf.Jyuni2c,2);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=49;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Jyuni3c")));	
	strsql.SetString(mBuf.Jyuni3c,2);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=50;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Jyuni4c")));	
	strsql.SetString(mBuf.Jyuni4c,2);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=51;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Odds")));	
	strsql.SetString(mBuf.Odds,4);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=52;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Ninki")));	
	strsql.SetString(mBuf.Ninki,2);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=53;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Honsyokin")));	
	strsql.SetString(mBuf.Honsyokin,8);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=54;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Fukasyokin")));	
	strsql.SetString(mBuf.Fukasyokin,8);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=55;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("reserved3")));	
	strsql.SetString(mBuf.reserved3,3);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=56;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("reserved4")));	
	strsql.SetString(mBuf.reserved4,3);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=57;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("HaronTimeL4")));	
	strsql.SetString(mBuf.HaronTimeL4,3);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=58;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("HaronTimeL3")));	
	strsql.SetString(mBuf.HaronTimeL3,3);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));


	int i;
	for(i=0;i<3;i++){
		lArrayIndex[0]=59+i*2;
		strsql.Format("KettoNum%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.ChakuUmaInfo[0].KettoNum,10);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=60+i*2;
		strsql.Format("Bamei%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.ChakuUmaInfo[0].Bamei,36);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));
	}

	lArrayIndex[0]=65;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("TimeDiff")));	
	strsql.SetString(mBuf.TimeDiff,4);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=66;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("RecordUpKubun")));	
	strsql.SetString(mBuf.RecordUpKubun,1);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=67;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("DMKubun")));	
	strsql.SetString(mBuf.DMKubun,1);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=68;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("DMTime")));	
	strsql.SetString(mBuf.DMTime,5);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=69;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("DMGosaP")));	
	strsql.SetString(mBuf.DMGosaP,4);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=70;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("DMGosaM")));	
	strsql.SetString(mBuf.DMGosaM,4);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=71;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("DMJyuni")));	
	strsql.SetString(mBuf.DMJyuni,2);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=72;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("KyakusituKubun")));	
	strsql.SetString(mBuf.KyakusituKubun,1);
	pRecordSet->Update(vaFieldlist, vaValuelist);
	pRecordSet->Close();	

}
catch(_com_error &e){
	MessageBox(NULL,e.Description(),NULL,NULL);
	pRecordSet->Close();
	return -1;

}

	return 0;
}

int clsImportSE::Init(_ConnectionPtr m_Connect)
{ 
    pCn = m_Connect;

	return 0;
}
