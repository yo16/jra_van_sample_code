/*=======================================================================
  JRA-VAN Data Lab.プログラミングパーツ「データ登録パーツ　AVレコード登録クラス」

	   作成: JRA-VAN ソフトウェア工房  2003年6月26日

========================================================================
   (C) Copyright JRA SYSTEM SERVICE CO.,LTD. 2007 All rights reserved
========================================================================*/

#include "StdAfx.h"
#include "clsImportAV.h"
#include <afxdisp.h>

clsImportAV::clsImportAV(void)
{
}

clsImportAV::~clsImportAV(void)
{
}
int clsImportAV::Add(CString strBuff, long lngBuffSize)
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

int clsImportAV::InsertDB(void)
{
	CString strsql;

	USES_CONVERSION;
	strsql="SELECT * FROM TORIKESI_JYOGAI WHERE [Year]='";
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
	strsql.Append("'");
	_bstr_t bstrQuery(strsql);

	// SQLの設定
	_CommandPtr pCommand;      // Commandオブジェクト
	_RecordsetPtr pRecordSet;
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

	COleSafeArray vaFieldlist;
	vaFieldlist.CreateOneDim(VT_VARIANT,13);

	COleSafeArray vaValuelist;
	vaValuelist.CreateOneDim(VT_VARIANT,13);
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
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Umaban")));
		strsql.SetString(mBuf.Umaban,2);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=11;
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Bamei")));
		strsql.SetString(mBuf.Bamei,36);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=12;
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t("JiyuKubun")));
		strsql.SetString(mBuf.JiyuKubun,3);
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

int clsImportAV::UpdateDB(CString strMakeDate)
{
	CString strsql;
	
	USES_CONVERSION;
	
	strsql="SELECT * FROM TORIKESI_JYOGAI WHERE [Year]='";
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
	strsql.Append("' AND [MakeDate] <= '");
	strsql.Append(strMakeDate);
	strsql.Append("'");

	_bstr_t bstrQuery(strsql);



	// SQLの設定
	_CommandPtr pCommand;      // Commandオブジェクト
	_RecordsetPtr pRecordSet;

	pCommand.CreateInstance(__uuidof(Command));
	pRecordSet.CreateInstance(__uuidof(Recordset));

	pCommand->ActiveConnection = pCn;
	pCommand->CommandText = bstrQuery;
	pRecordSet->PutRefSource(pCommand);

	// レコードセットの取得
	_variant_t vNull;  // VARIANT型のNULLとして使う
	vNull.vt = VT_ERROR;
	vNull.scode = DISP_E_PARAMNOTFOUND;
	pRecordSet->Open(vNull, vNull, adOpenKeyset, adLockOptimistic, adCmdText);

	COleSafeArray vaFieldlist;
	vaFieldlist.CreateOneDim(VT_VARIANT,13);

	COleSafeArray vaValuelist;
	vaValuelist.CreateOneDim(VT_VARIANT,13);
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
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Umaban")));
		strsql.SetString(mBuf.Umaban,2);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=11;
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Bamei")));
		strsql.SetString(mBuf.Bamei,36);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=12;
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t("JiyuKubun")));
		strsql.SetString(mBuf.JiyuKubun,3);
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
int clsImportAV::Init(_ConnectionPtr m_Connect)
{
	pCn=m_Connect;
	return 0;
}

