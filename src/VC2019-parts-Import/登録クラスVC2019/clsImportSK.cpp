/*=======================================================================
  JRA-VAN Data Lab.プログラミングパーツ「データ登録パーツ　SEレコード登録クラス」

	   作成: JRA-VAN ソフトウェア工房  2003年6月26日

========================================================================
   (C) Copyright JRA SYSTEM SERVICE CO.,LTD. 2007 All rights reserved
========================================================================*/
#include "StdAfx.h"
#include "clsImportSK.h"

clsImportSK::clsImportSK(void)
{
}

clsImportSK::~clsImportSK(void)
{
}
int clsImportSK::Add(CString strBuff, long lngBuffSize)
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

int clsImportSK::InsertDB(void)
{
	CString strsql;
	strsql="SELECT * FROM SANKU WHERE [KettoNum] = '";
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
	vaFieldlist.CreateOneDim(VT_VARIANT,26);

	COleSafeArray vaValuelist;
	vaValuelist.CreateOneDim(VT_VARIANT,26);
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
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("KettoNum")));	
	strsql.SetString(mBuf.KettoNum,10);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=4;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("BirthDate")));	
	strsql.SetString(mBuf.BirthDate.Year,8);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=5;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("SexCD")));	
	strsql.SetString(mBuf.SexCD,1);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=6;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("HinsyuCD")));	
	strsql.SetString(mBuf.HinsyuCD,1);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=7;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("KeiroCD")));	
	strsql.SetString(mBuf.KeiroCD,2);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=8;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("SankuMochiKubun")));	
	strsql.SetString(mBuf.SankuMochiKubun,1);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=9;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("ImportYear")));	
	strsql.SetString(mBuf.ImportYear,4);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=10;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("BreederCode")));	
	strsql.SetString(mBuf.BreederCode,8);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=11;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("SanchiName")));	
	strsql.SetString(mBuf.SanchiName,20);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=12;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("FNum")));	
	strsql.SetString(mBuf.HansyokuNum[0],10);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=13;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("MNum")));	
	strsql.SetString(mBuf.HansyokuNum[1],10);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=14;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("FFNum")));	
	strsql.SetString(mBuf.HansyokuNum[2],10);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=15;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("FMNum")));	
	strsql.SetString(mBuf.HansyokuNum[3],10);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=16;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("MFNum")));
	strsql.SetString(mBuf.HansyokuNum[4],10);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=17;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("MMNum")));
	strsql.SetString(mBuf.HansyokuNum[5],10);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=18;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("FFFNum")));
	strsql.SetString(mBuf.HansyokuNum[6],10);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=19;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("FFMNum")));
	strsql.SetString(mBuf.HansyokuNum[7],10);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=20;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("FMFNum")));
	strsql.SetString(mBuf.HansyokuNum[8],10);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=21;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("FMMNum")));
	strsql.SetString(mBuf.HansyokuNum[9],10);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=22;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("MFFNum")));
	strsql.SetString(mBuf.HansyokuNum[10],10);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=23;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("MFMNum")));	
	strsql.SetString(mBuf.HansyokuNum[11],10);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=24;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("MMFNum")));	
	strsql.SetString(mBuf.HansyokuNum[12],10);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=25;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("MMMNum")));	
	strsql.SetString(mBuf.HansyokuNum[13],10);
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

int clsImportSK::UpdateDB(CString strMakeDate)
{
	CString strsql;
	strsql="SELECT * FROM SANKU WHERE [KettoNum] = '";
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
	vaFieldlist.CreateOneDim(VT_VARIANT,26);

	COleSafeArray vaValuelist;
	vaValuelist.CreateOneDim(VT_VARIANT,26);
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
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("KettoNum")));	
	strsql.SetString(mBuf.KettoNum,10);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=4;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("BirthDate")));	
	strsql.SetString(mBuf.BirthDate.Year,8);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=5;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("SexCD")));	
	strsql.SetString(mBuf.SexCD,1);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=6;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("HinsyuCD")));	
	strsql.SetString(mBuf.HinsyuCD,1);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=7;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("KeiroCD")));	
	strsql.SetString(mBuf.KeiroCD,2);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=8;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("SankuMochiKubun")));	
	strsql.SetString(mBuf.SankuMochiKubun,1);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=9;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("ImportYear")));	
	strsql.SetString(mBuf.ImportYear,4);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=10;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("BreederCode")));	
	strsql.SetString(mBuf.BreederCode,8);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=11;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("SanchiName")));	
	strsql.SetString(mBuf.SanchiName,20);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=12;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("FNum")));	
	strsql.SetString(mBuf.HansyokuNum[0],10);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=13;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("MNum")));
	strsql.SetString(mBuf.HansyokuNum[1],10);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=14;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("FFNum")));	
	strsql.SetString(mBuf.HansyokuNum[2],10);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=15;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("FMNum")));	
	strsql.SetString(mBuf.HansyokuNum[3],10);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=16;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("MFNum")));
	strsql.SetString(mBuf.HansyokuNum[4],10);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=17;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("MMNum")));	
	strsql.SetString(mBuf.HansyokuNum[5],10);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=18;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("FFFNum")));
	strsql.SetString(mBuf.HansyokuNum[6],10);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=19;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("FFFNum")));	
	strsql.SetString(mBuf.HansyokuNum[7],10);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=20;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("FMFNum")));
	strsql.SetString(mBuf.HansyokuNum[8],10);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=21;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("FMFNum")));
	strsql.SetString(mBuf.HansyokuNum[9],10);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=22;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("MFFNum")));	
	strsql.SetString(mBuf.HansyokuNum[10],10);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=23;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("MFMNum")));
	strsql.SetString(mBuf.HansyokuNum[11],10);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=24;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("MMFNum")));
	strsql.SetString(mBuf.HansyokuNum[12],10);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=25;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("MMMNum")));	
	strsql.SetString(mBuf.HansyokuNum[13],10);
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
int clsImportSK::Init(_ConnectionPtr m_Connect)
{
    pCn = m_Connect;

	 return 0;
}

