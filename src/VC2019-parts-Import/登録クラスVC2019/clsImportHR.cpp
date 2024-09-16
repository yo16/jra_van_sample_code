/*=======================================================================
  JRA-VAN Data Lab.プログラミングパーツ「データ登録パーツ　HRレコード登録クラス」

	   作成: JRA-VAN ソフトウェア工房  2003年6月26日

========================================================================
   (C) Copyright JRA SYSTEM SERVICE CO.,LTD. 2007 All rights reserved
========================================================================*/

#include "StdAfx.h"
#include "clsimporthr.h"

clsImportHR::clsImportHR(void)
{
}

clsImportHR::~clsImportHR(void)
{
}

int clsImportHR::Add(CString strBuff, long lngBuffSize)
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

int clsImportHR::InsertDB(void)
{

	int i;
	CString strsql;

	strsql="SELECT * FROM HARAI WHERE [Year]='";
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
	vaFieldlist.CreateOneDim(VT_VARIANT,198);

	COleSafeArray vaValuelist;
	vaValuelist.CreateOneDim(VT_VARIANT,198);
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

	for(i=0;i<9;i++){
		lArrayIndex[0]=11+i;
		strsql.Format("FuseirituFlag%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.FuseirituFlag[i],1);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	}
	for(i=0;i<9;i++){
		lArrayIndex[0]=20+i;
		strsql.Format("TokubaraiFlag%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.TokubaraiFlag[i],1);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	}
	for(i=0;i<9;i++){
		lArrayIndex[0]=29+i;
		strsql.Format("HenkanFlag%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.HenkanFlag[i],1);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	}
	for(i=0;i<28;i++){
		lArrayIndex[0]=38+i;
		strsql.Format("HenkanUma%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.HenkanUma[i],1);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	}
	for(i=0;i<8;i++){
		lArrayIndex[0]=66+i;
		strsql.Format("HenkanWaku%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.HenkanWaku[i],1);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));
	}
	for(i=0;i<8;i++){
		lArrayIndex[0]=74+i;
		strsql.Format("HenkanDoWaku%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.HenkanDoWaku[i],1);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	}
	for(i=0;i<3;i++){
		lArrayIndex[0]=82+3*i;
		strsql.Format("PayTansyoUmaban%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.PayTansyo[i].Umaban,2);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=83+3*i;
		strsql.Format("PayTansyoPay%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.PayTansyo[i].Pay,9);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=84+3*i;
		strsql.Format("PayTansyoNinki%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.PayTansyo[i].Ninki,2);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	}
	for(i=0;i<5;i++){
		lArrayIndex[0]=91+i*3;
		strsql.Format("PayFukusyoUmaban%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.PayFukusyo[i].Umaban,2);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=92+i*3;
		strsql.Format("PayFukusyoPay%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.PayFukusyo[i].Pay,9);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=93+i*3;
		strsql.Format("PayFukusyoNinki%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.PayFukusyo[i].Ninki,2);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	}
	for(i=0;i<3;i++){
		lArrayIndex[0]=106+3*i;
		strsql.Format("PayWakurenKumi%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.PayWakuren[i].Umaban,2);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=107+3*i;
		strsql.Format("PayWakurenPay%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.PayWakuren[i].Pay,9);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=108+3*i;
		strsql.Format("PayWakurenNinki%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.PayWakuren[i].Ninki,2);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	}
	for(i=0;i<3;i++){
		lArrayIndex[0]=115+i*3;
		strsql.Format("PayUmarenKumi%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.PayUmaren[i].Kumi,4);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=116+i*3;
		strsql.Format("PayUmarenPay%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.PayUmaren[i].Pay,9);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=117+i*3;
		strsql.Format("PayUmarenNinki%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.PayUmaren[i].Ninki,3);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));
	}
	for(i=0;i<7;i++){
		lArrayIndex[0]=124+i*3;
		strsql.Format("PayWideKumi%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.PayWide[i].Kumi,4);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=125+i*3;
		strsql.Format("PayWidePay%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.PayWide[i].Pay,9);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=126+i*3;
		strsql.Format("PayWidePay%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.PayWide[i].Ninki,3);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	}
	for(i=0;i<3;i++){
		lArrayIndex[0]=145+i*3;
		strsql.Format("PayReserved1Kumi%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.PayReserved1[i].Kumi,4);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=146+i*3;
		strsql.Format("PayReserved1Pay%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.PayReserved1[i].Pay,9);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=147+i*3;
		strsql.Format("PayReserved1Ninki%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.PayReserved1[i].Ninki,3);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));
	}
	
	
	for(i=0;i<6;i++){
		lArrayIndex[0]=154+i*3;
		strsql.Format("PayUmatanKumi%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.PayUmatan[i].Kumi,4);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=155+i*3;
		strsql.Format("PayUmatanPay%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.PayUmatan[i].Pay,9);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=156+i*3;
		strsql.Format("PayUmatanNinki%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.PayUmatan[i].Ninki,3);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));
	}
	for(i=0;i<3;i++){
		lArrayIndex[0]=171+i*3;
		strsql.Format("PaySanrenpukuKumi%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.PaySanrenpuku[i].Kumi,6);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=172+i*3;
		strsql.Format("PaySanrenpukuPay%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.PaySanrenpuku[i].Pay,9);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=173+i*3;
		strsql.Format("PaySanrenpukuNinki%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.PaySanrenpuku[i].Ninki,3);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));
	}
	for(i=0;i<6;i++){
		lArrayIndex[0]=180+i*3;
		strsql.Format("PaySanrentanKumi%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.PaySanrentan[i].Kumi,6);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=181+i*3;
		strsql.Format("PaySanrentanPay%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.PaySanrentan[i].Pay,9);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=182+i*3;
		strsql.Format("PaySanrentanNinki%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.PaySanrentan[i].Ninki,4);
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



int clsImportHR::UpdateDB(CString strMakeDate)
{
	int i;
	CString strsql;

	strsql="SELECT * FROM HARAI WHERE [Year]='";
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
	vaFieldlist.CreateOneDim(VT_VARIANT,198);

	COleSafeArray vaValuelist;
	vaValuelist.CreateOneDim(VT_VARIANT,198);
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

	for(i=0;i<9;i++){
		lArrayIndex[0]=11+i;
		strsql.Format("FuseirituFlag%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.FuseirituFlag[i],1);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	}
	for(i=0;i<9;i++){
		lArrayIndex[0]=20+i;
		strsql.Format("TokubaraiFlag%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.TokubaraiFlag[i],1);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	}
	for(i=0;i<9;i++){
		lArrayIndex[0]=29+i;
		strsql.Format("HenkanFlag%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.HenkanFlag[i],1);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	}
	for(i=0;i<28;i++){
		lArrayIndex[0]=38+i;
		strsql.Format("HenkanUma%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.HenkanUma[i],1);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	}
	for(i=0;i<8;i++){
		lArrayIndex[0]=66+i;
		strsql.Format("HenkanWaku%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.HenkanWaku[i],1);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));
	}
	for(i=0;i<8;i++){
		lArrayIndex[0]=74+i;
		strsql.Format("HenkanDoWaku%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.HenkanDoWaku[i],1);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	}
	for(i=0;i<3;i++){
		lArrayIndex[0]=82+3*i;
		strsql.Format("PayTansyoUmaban%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.PayTansyo[i].Umaban,2);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=83+3*i;
		strsql.Format("PayTansyoPay%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.PayTansyo[i].Pay,9);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=84+3*i;
		strsql.Format("PayTansyoNinki%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.PayTansyo[i].Ninki,2);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	}
	for(i=0;i<5;i++){
		lArrayIndex[0]=91+i*3;
		strsql.Format("PayFukusyoUmaban%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.PayFukusyo[i].Umaban,2);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=92+i*3;
		strsql.Format("PayFukusyoPay%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.PayFukusyo[i].Pay,9);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=93+i*3;
		strsql.Format("PayFukusyoNinki%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.PayFukusyo[i].Ninki,2);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	}
	for(i=0;i<3;i++){
		lArrayIndex[0]=106+3*i;
		strsql.Format("PayWakurenKumi%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.PayWakuren[i].Umaban,2);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=107+3*i;
		strsql.Format("PayWakurenPay%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.PayWakuren[i].Pay,9);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=108+3*i;
		strsql.Format("PayWakurenNinki%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.PayWakuren[i].Ninki,2);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	}
	for(i=0;i<3;i++){
		lArrayIndex[0]=115+i*3;
		strsql.Format("PayUmarenKumi%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.PayUmaren[i].Kumi,4);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=116+i*3;
		strsql.Format("PayUmarenPay%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.PayUmaren[i].Pay,9);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=117+i*3;
		strsql.Format("PayUmarenNinki%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.PayUmaren[i].Ninki,3);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));
	}
	for(i=0;i<7;i++){
		lArrayIndex[0]=124+i*3;
		strsql.Format("PayWideKumi%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.PayWide[i].Kumi,4);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=125+i*3;
		strsql.Format("PayWidePay%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.PayWide[i].Pay,9);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=126+i*3;
		strsql.Format("PayWidePay%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.PayWide[i].Ninki,3);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	}
	for(i=0;i<3;i++){
		lArrayIndex[0]=145+i*3;
		strsql.Format("PayReserved1Kumi%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.PayReserved1[i].Kumi,4);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=146+i*3;
		strsql.Format("PayReserved1Pay%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.PayReserved1[i].Pay,9);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=147+i*3;
		strsql.Format("PayReserved1Ninki%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.PayReserved1[i].Ninki,3);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));
	}
	
	
	for(i=0;i<6;i++){
		lArrayIndex[0]=154+i*3;
		strsql.Format("PayUmatanKumi%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.PayUmatan[i].Kumi,4);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=155+i*3;
		strsql.Format("PayUmatanPay%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.PayUmatan[i].Pay,9);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=156+i*3;
		strsql.Format("PayUmatanNinki%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.PayUmatan[i].Ninki,3);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));
	}
	for(i=0;i<3;i++){
		lArrayIndex[0]=171+i*3;
		strsql.Format("PaySanrenpukuKumi%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.PaySanrenpuku[i].Kumi,6);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=172+i*3;
		strsql.Format("PaySanrenpukuPay%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.PaySanrenpuku[i].Pay,9);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=173+i*3;
		strsql.Format("PaySanrenpukuNinki%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.PaySanrenpuku[i].Ninki,3);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));
	}
	for(i=0;i<6;i++){
		lArrayIndex[0]=180+i*3;
		strsql.Format("PaySanrentanKumi%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.PaySanrentan[i].Kumi,6);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=181+i*3;
		strsql.Format("PaySanrentanPay%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.PaySanrentan[i].Pay,9);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=182+i*3;
		strsql.Format("PaySanrentanNinki%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.PaySanrentan[i].Ninki,4);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));
	}
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

int clsImportHR::Init(_ConnectionPtr m_Connect)
{
    pCn = m_Connect;

	 return 0;
}
