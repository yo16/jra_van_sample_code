/*=======================================================================
  JRA-VAN Data Lab.プログラミングパーツ「データ登録パーツ　UMレコード登録クラス」

	   作成: JRA-VAN ソフトウェア工房  2003年6月26日

========================================================================
   (C) Copyright JRA SYSTEM SERVICE CO.,LTD. 2007 All rights reserved
========================================================================*/
#include "StdAfx.h"
#include "clsimportum.h"

clsImportUM::clsImportUM(void)
{
}

clsImportUM::~clsImportUM(void)
{
}
int clsImportUM::Add(CString strBuff, long lngBuffSize)
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

	return 0;
}

int clsImportUM::InsertDB(void)
{


	int i;
	CString strsql;
	strsql="SELECT * FROM UMA WHERE [KettoNum] = '";
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
	vaFieldlist.CreateOneDim(VT_VARIANT,227);

	COleSafeArray vaValuelist;
	vaValuelist.CreateOneDim(VT_VARIANT,227);
	long lArrayIndex[1];

try{
	lArrayIndex[0]=0;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("RecordSpec")));	
	strsql.SetString(mBuf.head.RecordSpec,2);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(strsql)));

	lArrayIndex[0]=1;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("DataKubun")));	
	strsql.SetString(mBuf.head.DataKubun,1);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(strsql)));

	lArrayIndex[0]=2;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("MakeDate")));	
	strsql.SetString(mBuf.head.MakeDate.Year,8);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(strsql)));

	lArrayIndex[0]=3;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("KettoNum")));	
	strsql.SetString(mBuf.KettoNum,10);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(strsql)));

	lArrayIndex[0]=4;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("DelKubun")));	
	strsql.SetString(mBuf.DelKubun,1);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(strsql)));

	lArrayIndex[0]=5;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("RegDate")));	
	strsql.SetString(mBuf.RegDate.Year,8);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(strsql)));

	lArrayIndex[0]=6;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("DelDate")));	
	strsql.SetString(mBuf.DelDate.Year,8);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(strsql)));

	lArrayIndex[0]=7;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("BirthDate")));	
	strsql.SetString(mBuf.BirthDate.Year,8);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(strsql)));

	lArrayIndex[0]=8;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Bamei")));	
	strsql.SetString(mBuf.Bamei,36);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(strsql)));

	lArrayIndex[0]=9;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("BameiKana")));	
	strsql.SetString(mBuf.BameiKana,36);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(strsql)));

	lArrayIndex[0]=10;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("BameiEng")));	
	strsql.SetString(mBuf.BameiEng,60);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(strsql)));

	lArrayIndex[0]=11;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("ZaikyuFlag")));	
	strsql.SetString(mBuf.ZaikyuFlag,1);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(strsql)));

	lArrayIndex[0]=12;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Reserved")));	
	strsql.SetString(mBuf.Reserved,19);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(strsql)));

	lArrayIndex[0]=13;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("UmaKigoCD")));	
	strsql.SetString(mBuf.UmaKigoCD,2);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(strsql)));

	lArrayIndex[0]=14;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("SexCD")));	
	strsql.SetString(mBuf.SexCD,1);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(strsql)));

	lArrayIndex[0]=15;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("HinsyuCD")));	
	strsql.SetString(mBuf.HinsyuCD,1);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(strsql)));

	lArrayIndex[0]=16;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("KeiroCD")));	
	strsql.SetString(mBuf.KeiroCD,2);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(strsql)));

	for(i=0;i<14;i++){
		lArrayIndex[0]=17+2*i;
		strsql.Format("Ketto3InfoHansyokuNum%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.Ketto3Info[i].HansyokuNum,10);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(strsql)));

		strsql.Format("Ketto3InfoBamei%d",i+1);
		lArrayIndex[0]=18+2*i;
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.Ketto3Info[i].Bamei,36);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(strsql)));
	}
	lArrayIndex[0]=45;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("TozaiCD")));	
	strsql.SetString(mBuf.TozaiCD,1);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(strsql)));

	lArrayIndex[0]=46;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("ChokyosiCode")));	
	strsql.SetString(mBuf.ChokyosiCode,5);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(strsql)));

	lArrayIndex[0]=47;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("ChokyosiRyakusyo")));	
	strsql.SetString(mBuf.ChokyosiRyakusyo,8);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(strsql)));

	lArrayIndex[0]=48;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Syotai")));	
	strsql.SetString(mBuf.Syotai,20);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(strsql)));

	lArrayIndex[0]=49;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("BreederCode")));
	strsql.SetString(mBuf.BreederCode,8);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(strsql)));

	lArrayIndex[0]=50;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("BreederName")));	
	strsql.SetString(mBuf.BreederName,72);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(strsql)));

	lArrayIndex[0]=51;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("SanchiName")));	
	strsql.SetString(mBuf.SanchiName,20);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(strsql)));

	lArrayIndex[0]=52;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("BanusiCode")));	
	strsql.SetString(mBuf.BanusiCode,6);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(strsql)));

	lArrayIndex[0]=53;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("BanusiName")));	
	strsql.SetString(mBuf.BanusiName,64);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(strsql)));

	lArrayIndex[0]=54;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("RuikeiHonsyoHeiti")));	
	strsql.SetString(mBuf.RuikeiHonsyoHeiti,9);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(strsql)));

	lArrayIndex[0]=55;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("RuikeiHonsyoSyogai")));	
	strsql.SetString(mBuf.RuikeiHonsyoSyogai,9);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(strsql)));

	lArrayIndex[0]=56;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("RuikeiFukaHeichi")));	
	strsql.SetString(mBuf.RuikeiFukaHeichi,9);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(strsql)));

	lArrayIndex[0]=57;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("RuikeiFukaSyogai")));	
	strsql.SetString(mBuf.RuikeiFukaSyogai,9);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(strsql)));

	lArrayIndex[0]=58;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("RuikeiSyutokuHeichi")));	
	strsql.SetString(mBuf.RuikeiSyutokuHeichi,9);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(strsql)));

	lArrayIndex[0]=59;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("RuikeiSyutokuSyogai")));	
	strsql.SetString(mBuf.RuikeiSyutokuSyogai,9);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(strsql)));

	for(i=0;i<6;i++){
		lArrayIndex[0]=60+i;
		strsql.Format("SogoChakukaisu%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.ChakuSogo.Chakukaisu[i],3);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(strsql)));
	}
	for(i=0;i<6;i++){
		lArrayIndex[0]=66+i;
		strsql.Format("ChuoChakukaisu%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.ChakuChuo.Chakukaisu[i],3);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(strsql)));
	}    

	int j;
	for(i=0;i<7;i++){
		for(j=0;j<6;j++){	
			lArrayIndex[0]=72+i*6+j;
			strsql.Format("Ba%dChakukaisu%d",i+1,j+1);
			vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		    strsql.SetString(mBuf.ChakuKaisuBa[i].Chakukaisu[j],3);
			vaValuelist.PutElement(lArrayIndex, &(_variant_t(strsql)));
		}
	}    
	for(i=0;i<12;i++){
		for(j=0;j<6;j++){
			lArrayIndex[0]=114+i*6+j;
			strsql.Format("Jyotai%dChakukaisu%d",i+1,j+1);
			vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		    strsql.SetString(mBuf.ChakuKaisuJyotai[i].Chakukaisu[j],3);
			vaValuelist.PutElement(lArrayIndex, &(_variant_t(strsql)));
		}
	}    
	for(i=0;i<6;i++){
		for(j=0;j<6;j++){
			strsql.Format("Kyori%dChakukaisu%d",i+1,j+1);
			lArrayIndex[0]=186+i*6+j;
			vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		    strsql.SetString(mBuf.ChakuKaisuKyori[i].Chakukaisu[j],3);
			vaValuelist.PutElement(lArrayIndex, &(_variant_t(strsql)));
		}
	}    
	for(i=0;i<4;i++){
		strsql.Format("Kyakusitu%d",i+1);
			lArrayIndex[0]=222+i;
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
        strsql.SetString(mBuf.Kyakusitu[i],3);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(strsql)));
	}
	lArrayIndex[0]=226;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("RaceCount")));	
	strsql.SetString(mBuf.RaceCount,3);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(strsql)));
	
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

int clsImportUM::UpdateDB(CString strMakeDate)
{
	int i;
	CString strsql;
	strsql="SELECT * FROM UMA WHERE [KettoNum] = '";
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

	if (!pRecordSet->GetadoEOF()){
		pRecordSet->Close();
		pRecordSet = NULL;
		return -1;
	}

	USES_CONVERSION;

	COleSafeArray vaFieldlist;
	vaFieldlist.CreateOneDim(VT_VARIANT,227);

	COleSafeArray vaValuelist;
	vaValuelist.CreateOneDim(VT_VARIANT,227);
	long lArrayIndex[1];
	
	try{	
	lArrayIndex[0]=0;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("RecordSpec")));	
	strsql.SetString(mBuf.head.RecordSpec,2);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(strsql)));

	lArrayIndex[0]=1;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("DataKubun")));	
	strsql.SetString(mBuf.head.DataKubun,1);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(strsql)));

	lArrayIndex[0]=2;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("MakeDate")));	
	strsql.SetString(mBuf.head.MakeDate.Year,8);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(strsql)));

	lArrayIndex[0]=3;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("KettoNum")));	
	strsql.SetString(mBuf.KettoNum,10);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(strsql)));

	lArrayIndex[0]=4;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("DelKubun")));	
	strsql.SetString(mBuf.DelKubun,1);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(strsql)));

	lArrayIndex[0]=5;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("RegDate")));	
	strsql.SetString(mBuf.RegDate.Year,8);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(strsql)));

	lArrayIndex[0]=6;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("DelDate")));	
	strsql.SetString(mBuf.DelDate.Year,8);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(strsql)));

	lArrayIndex[0]=7;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("BirthDate")));	
	strsql.SetString(mBuf.BirthDate.Year,8);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(strsql)));

	lArrayIndex[0]=8;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Bamei")));	
	strsql.SetString(mBuf.Bamei,36);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(strsql)));

	lArrayIndex[0]=9;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("BameiKana")));	
	strsql.SetString(mBuf.BameiKana,36);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(strsql)));

	lArrayIndex[0]=10;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("BameiEng")));	
	strsql.SetString(mBuf.BameiEng,60);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(strsql)));

	lArrayIndex[0]=11;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("ZaikyuFlag")));	
	strsql.SetString(mBuf.ZaikyuFlag,1);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(strsql)));

	lArrayIndex[0]=12;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Reserved")));	
	strsql.SetString(mBuf.Reserved,19);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(strsql)));

	lArrayIndex[0]=13;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("UmaKigoCD")));	
	strsql.SetString(mBuf.UmaKigoCD,2);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(strsql)));

	lArrayIndex[0]=14;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("SexCD")));	
	strsql.SetString(mBuf.SexCD,1);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(strsql)));

	lArrayIndex[0]=15;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("HinsyuCD")));	
	strsql.SetString(mBuf.HinsyuCD,1);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(strsql)));

	lArrayIndex[0]=16;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("KeiroCD")));	
	strsql.SetString(mBuf.KeiroCD,2);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(strsql)));

	for(i=0;i<14;i++){
		lArrayIndex[0]=17+2*i;
		strsql.Format("Ketto3InfoHansyokuNum%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.Ketto3Info[i].HansyokuNum,10);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(strsql)));

		strsql.Format("Ketto3InfoBamei%d",i+1);
		lArrayIndex[0]=18+2*i;
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.Ketto3Info[i].Bamei,36);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(strsql)));
	}
	lArrayIndex[0]=45;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("TozaiCD")));	
	strsql.SetString(mBuf.TozaiCD,1);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(strsql)));

	lArrayIndex[0]=46;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("ChokyosiCode")));	
	strsql.SetString(mBuf.ChokyosiCode,5);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(strsql)));

	lArrayIndex[0]=47;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("ChokyosiRyakusyo")));	
	strsql.SetString(mBuf.ChokyosiRyakusyo,8);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(strsql)));

	lArrayIndex[0]=48;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Syotai")));	
	strsql.SetString(mBuf.Syotai,20);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(strsql)));

	lArrayIndex[0]=49;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("BreederCode")));	
	strsql.SetString(mBuf.BreederCode,8);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(strsql)));

	lArrayIndex[0]=50;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("BreederName")));	
	strsql.SetString(mBuf.BreederName,72);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(strsql)));

	lArrayIndex[0]=51;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("SanchiName")));	
	strsql.SetString(mBuf.SanchiName,20);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(strsql)));

	lArrayIndex[0]=52;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("BanusiCode")));	
	strsql.SetString(mBuf.BanusiCode,6);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(strsql)));

	lArrayIndex[0]=53;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("BanusiName")));	
	strsql.SetString(mBuf.BanusiName,64);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(strsql)));

	lArrayIndex[0]=54;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("RuikeiHonsyoHeiti")));	
	strsql.SetString(mBuf.RuikeiHonsyoHeiti,9);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(strsql)));

	lArrayIndex[0]=55;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("RuikeiHonsyoSyogai")));	
	strsql.SetString(mBuf.RuikeiHonsyoSyogai,9);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(strsql)));

	lArrayIndex[0]=56;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("RuikeiFukaHeichi")));	
	strsql.SetString(mBuf.RuikeiFukaHeichi,9);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(strsql)));

	lArrayIndex[0]=57;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("RuikeiFukaSyogai")));	
	strsql.SetString(mBuf.RuikeiFukaSyogai,9);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(strsql)));

	lArrayIndex[0]=58;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("RuikeiSyutokuHeichi")));	
	strsql.SetString(mBuf.RuikeiSyutokuHeichi,9);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(strsql)));

	lArrayIndex[0]=59;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("RuikeiSyutokuSyogai")));	
	strsql.SetString(mBuf.RuikeiSyutokuSyogai,9);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(strsql)));

	for(i=0;i<6;i++){
		lArrayIndex[0]=60+i;
		strsql.Format("SogoChakukaisu%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.ChakuSogo.Chakukaisu[i],3);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(strsql)));
	}
	for(i=0;i<6;i++){
		lArrayIndex[0]=66+i;
		strsql.Format("ChuoChakukaisu%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.ChakuChuo.Chakukaisu[i],3);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(strsql)));
	}    

	int j;
	for(i=0;i<7;i++){
		for(j=0;j<6;j++){	
			lArrayIndex[0]=72+i*6+j;
			strsql.Format("Ba%dChakukaisu%d",i+1,j+1);
			vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		    strsql.SetString(mBuf.ChakuKaisuBa[i].Chakukaisu[j],3);
			vaValuelist.PutElement(lArrayIndex, &(_variant_t(strsql)));
		}
	}    
	for(i=0;i<12;i++){
		for(j=0;j<6;j++){
			lArrayIndex[0]=114+i*6+j;
			strsql.Format("Jyotai%dChakukaisu%d",i+1,j+1);
			vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		    strsql.SetString(mBuf.ChakuKaisuJyotai[i].Chakukaisu[j],3);
			vaValuelist.PutElement(lArrayIndex, &(_variant_t(strsql)));
		}
	}    
	for(i=0;i<6;i++){
		for(j=0;j<6;j++){
			strsql.Format("Kyori%dChakukaisu%d",i+1,j+1);
			lArrayIndex[0]=186+i*6+j;
			vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		    strsql.SetString(mBuf.ChakuKaisuKyori[i].Chakukaisu[j],3);
			vaValuelist.PutElement(lArrayIndex, &(_variant_t(strsql)));
		}
	}    
	for(i=0;i<4;i++){
		strsql.Format("Kyakusitu%d",i+1);
			lArrayIndex[0]=222+i;
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
        strsql.SetString(mBuf.Kyakusitu[i],3);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(strsql)));
	}
	lArrayIndex[0]=226;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("RaceCount")));	
	strsql.SetString(mBuf.RaceCount,3);
	
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
int clsImportUM::Init(_ConnectionPtr m_Connect)
{
    pCn = m_Connect;

	 return 0;
}


