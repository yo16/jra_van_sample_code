/*=======================================================================
  JRA-VAN Data Lab.プログラミングパーツ「データ登録パーツ　BRレコード登録クラス」

	   作成: JRA-VAN ソフトウェア工房  2003年6月26日

========================================================================
   (C) Copyright JRA SYSTEM SERVICE CO.,LTD. 2007 All rights reserved
========================================================================*/
#include "StdAfx.h"
#include "clsimportBR.h"
#include <afxdisp.h>

clsImportBR::clsImportBR(void)
{
}

clsImportBR::~clsImportBR(void)
{
}
int clsImportBR::Add(CString strBuff, long lngBuffSize)
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

int clsImportBR::InsertDB(void)
{
	CString strsql;
	strsql="SELECT * FROM SEISAN WHERE [BreederCode] = '";
	strsql.Append(mBuf.BreederCode,8);
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

	
try{


	COleSafeArray vaFieldlist;
	vaFieldlist.CreateOneDim(VT_VARIANT,27);

	COleSafeArray vaValuelist;
	vaValuelist.CreateOneDim(VT_VARIANT,27);
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
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("BreederCode")));	
	strsql.SetString(mBuf.BreederCode,8);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=4;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("BreederName_Co")));	
	strsql.SetString(mBuf.BreederName_Co,72);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=5;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("BreederName")));	
	strsql.SetString(mBuf.BreederName,72);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=6;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("BreederNameKana")));	
	strsql.SetString(mBuf.BreederNameKana,72);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));


	lArrayIndex[0]=7;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("BreederNameEng")));	
	strsql.SetString(mBuf.BreederNameEng,168);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=8;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Address")));	
	strsql.SetString(mBuf.Address,20);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));


	lArrayIndex[0]=9;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("H_SetYear")));
	strsql.SetString(mBuf.HonRuikei[0].SetYear,4);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=10;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("H_HonSyokinTotal")));
	strsql.SetString(mBuf.HonRuikei[0].HonSyokinTotal,10);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=11;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("H_FukaSyokin")));
	strsql.SetString(mBuf.HonRuikei[0].FukaSyokin,10);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	int i;
	for(i=0;i<6;i++){
		strsql = "H_ChakuKaisu";
		strsql.AppendFormat("%d",i+1);
		lArrayIndex[0]=12+i;
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));
		strsql.SetString(mBuf.HonRuikei[0].ChakuKaisu[i],6);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));
	}

	lArrayIndex[0]=18;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("R_SetYear")));
	strsql.SetString(mBuf.HonRuikei[1].SetYear,4);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=19;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("R_HonSyokinTotal")));
	strsql.SetString(mBuf.HonRuikei[1].HonSyokinTotal,10);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=20;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("R_FukaSyokin")));
	strsql.SetString(mBuf.HonRuikei[1].FukaSyokin,10);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	for(i=0;i<6;i++){
		strsql = "R_ChakuKaisu";
		strsql.AppendFormat("%d",i+1);
		lArrayIndex[0]=21+i;
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));
		strsql.SetString(mBuf.HonRuikei[1].ChakuKaisu[i],6);
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

int clsImportBR::UpdateDB(CString strMakeDate)
{
	_RecordsetPtr pRecordSet;

	CString strsql;
	strsql="SELECT * FROM SEISAN WHERE [BreederCode] = '";
	strsql.Append(mBuf.BreederCode,8);
	strsql.Append("' AND [MakeDate] <= '");
	strsql.Append(strMakeDate);
	strsql.Append("'");

	USES_CONVERSION;
	// Commandオブジェクト
	_CommandPtr pCommand;
	pCommand.CreateInstance(__uuidof(Command));
	pRecordSet.CreateInstance(__uuidof(Recordset));

	// SQLの設定
	pCommand->ActiveConnection = pCn;
	_bstr_t bstrQuery(strsql);
	pCommand->CommandText = bstrQuery;
	pRecordSet->PutRefSource(pCommand);

	_variant_t vNull;  // VARIANT型のNULLとして使う
	vNull.vt = VT_ERROR;
	vNull.scode = DISP_E_PARAMNOTFOUND;

	// レコードセットの取得
	pRecordSet->Open(vNull, vNull, adOpenKeyset, adLockOptimistic, adCmdText);

	COleSafeArray vaFieldlist;
	vaFieldlist.CreateOneDim(VT_VARIANT,27);

	COleSafeArray vaValuelist;
	vaValuelist.CreateOneDim(VT_VARIANT,27);
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
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("BreederCode")));	
	strsql.SetString(mBuf.BreederCode,8);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=4;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("BreederName_Co")));	
	strsql.SetString(mBuf.BreederName_Co,72);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=5;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("BreederName")));	
	strsql.SetString(mBuf.BreederName,72);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=6;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("BreederNameKana")));	
	strsql.SetString(mBuf.BreederNameKana,72);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));


	lArrayIndex[0]=7;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("BreederNameEng")));	
	strsql.SetString(mBuf.BreederNameEng,168);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=8;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Address")));	
	strsql.SetString(mBuf.Address,20);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));


	lArrayIndex[0]=9;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("H_SetYear")));
	strsql.SetString(mBuf.HonRuikei[0].SetYear,4);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=10;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("H_HonSyokinTotal")));
	strsql.SetString(mBuf.HonRuikei[0].HonSyokinTotal,10);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=11;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("H_FukaSyokin")));
	strsql.SetString(mBuf.HonRuikei[0].FukaSyokin,10);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	int i;
	for(i=0;i<6;i++){
		strsql = "H_ChakuKaisu";
		strsql.AppendFormat("%d",i+1);
		lArrayIndex[0]=12+i;
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));
		strsql.SetString(mBuf.HonRuikei[0].ChakuKaisu[i],6);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));
	}

	lArrayIndex[0]=18;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("R_SetYear")));
	strsql.SetString(mBuf.HonRuikei[1].SetYear,4);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=19;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("R_HonSyokinTotal")));
	strsql.SetString(mBuf.HonRuikei[1].HonSyokinTotal,10);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=20;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("R_FukaSyokin")));
	strsql.SetString(mBuf.HonRuikei[1].FukaSyokin,10);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	for(i=0;i<6;i++){
		strsql = "R_ChakuKaisu";
		strsql.AppendFormat("%d",i+1);
		lArrayIndex[0]=21+i;
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));
		strsql.SetString(mBuf.HonRuikei[1].ChakuKaisu[i],6);
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


int clsImportBR::Init(_ConnectionPtr m_Connect)
{
    pCn = m_Connect;

	 return 0;
}

