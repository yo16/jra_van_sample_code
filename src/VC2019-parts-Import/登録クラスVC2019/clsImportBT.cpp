/*=======================================================================
  JRA-VAN Data Lab.プログラミングパーツ「データ登録パーツ　BTレコード登録クラス」

	   作成: JRA-VAN ソフトウェア工房  2009年6月23日

========================================================================
   (C) Copyright JRA SYSTEM SERVICE CO.,LTD. 2007 All rights reserved
========================================================================*/
#include "StdAfx.h"
#include "clsImportBT.h"
#include <afxdisp.h>

clsImportBT::clsImportBT(void)
{
}

clsImportBT::~clsImportBT(void)
{
}
int clsImportBT::Add(CString strBuff, long lngBuffSize)
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

int clsImportBT::InsertDB(void)
{
	_RecordsetPtr pRecordSet;
	CString strsql;
	strsql="SELECT * FROM KEITO WHERE [HansyokuNum] = '";
	strsql.Append(mBuf.HansyokuNum,10);
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

	COleSafeArray vaFieldlist;
	vaFieldlist.CreateOneDim(VT_VARIANT,7);

	COleSafeArray vaValuelist;
	vaValuelist.CreateOneDim(VT_VARIANT,7);
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
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("HansyokuNum")));	
	strsql.SetString(mBuf.HansyokuNum,10);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=4;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("KeitoId")));	
	strsql.SetString(mBuf.KeitoId,30);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=5;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("KeitoName")));	
	strsql.SetString(mBuf.KeitoName,36);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=6;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("KeitoEx")));	
	strsql.SetString(mBuf.KeitoEx,6800);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

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

int clsImportBT::UpdateDB(CString strMakeDate)
{
	CString strsql;
	strsql="SELECT * FROM KEITO WHERE [HansyokuNum] = '";
	strsql.Append(mBuf.HansyokuNum,10);
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
	vaFieldlist.CreateOneDim(VT_VARIANT,7);

	COleSafeArray vaValuelist;
	vaValuelist.CreateOneDim(VT_VARIANT,7);
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
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("HansyokuNum")));	
	strsql.SetString(mBuf.HansyokuNum,10);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=4;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("KeitoId")));	
	strsql.SetString(mBuf.KeitoId,30);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=5;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("KeitoName")));	
	strsql.SetString(mBuf.KeitoName,36);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=6;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("KeitoEx")));	
	strsql.SetString(mBuf.KeitoEx,6800);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

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
int clsImportBT::Init(_ConnectionPtr m_Connect)
{
    pCn = m_Connect;

	 return 0;
}

