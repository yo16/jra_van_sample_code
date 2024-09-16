/*=======================================================================
  JRA-VAN Data Lab.プログラミングパーツ「データ登録パーツ　DMレコード登録クラス」

	   作成: JRA-VAN ソフトウェア工房  2003年6月26日

========================================================================
   (C) Copyright JRA SYSTEM SERVICE CO.,LTD. 2007 All rights reserved
========================================================================*/


#include <afxdisp.h>
#include "StdAfx.h"
#include "clsImportDM.h"

clsImportDM::clsImportDM(void)
{
}

clsImportDM::~clsImportDM(void)
{
}
int clsImportDM::Add(CString strBuff, long lngBuffSize)
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

int clsImportDM::InsertDB(void)
{


	CString strsql;
	strsql="SELECT * FROM MINING WHERE [Year]='";
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

	// SQLの設定
	_RecordsetPtr pRecordSet;
	_CommandPtr pCommand;      // Commandオブジェクト
	pCommand.CreateInstance(__uuidof(Command));
	pRecordSet.CreateInstance(__uuidof(Recordset));

	pCommand->ActiveConnection = pCn;

	pCommand->CommandText = bstrQuery;
	pRecordSet->PutRefSource(pCommand);

	// レコードセットの取得
	_variant_t vNull;  // VARIANT型のNULLとして使う
	vNull.vt = VT_ERROR;
	vNull.scode = DISP_E_PARAMNOTFOUND;
	pRecordSet->Open(vNull, vNull, adOpenForwardOnly, adLockOptimistic, adCmdText);


	if (!pRecordSet->GetadoEOF()){
		pRecordSet->Close();
		pRecordSet = NULL;
		return -1;
	}

	COleSafeArray vaFieldlist;
	vaFieldlist.CreateOneDim(VT_VARIANT,82);

	COleSafeArray vaValuelist;
	vaValuelist.CreateOneDim(VT_VARIANT,82);
	long lArrayIndex[1];

	USES_CONVERSION;

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
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("MakeHM")));	
	strsql.SetString(mBuf.MakeHM.Hour,4);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	int i;
	for(i=0;i<18;i++){
		lArrayIndex[0]=10+4*i;
		strsql.Format("Umaban%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.DMInfo[i].Umaban,2);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));
			
		lArrayIndex[0]=11+4*i;
		strsql.Format("DMTime%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.DMInfo[i].DMTime,5);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));
			
		lArrayIndex[0]=12+4*i;
		strsql.Format("DMGosaP%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.DMInfo[i].DMGosaP,4);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));
			
		lArrayIndex[0]=13+4*i;
		strsql.Format("DMGosaM%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.DMInfo[i].DMGosaM,4);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));
	}

	pRecordSet->AddNew(vaFieldlist, vaValuelist);
	pRecordSet->Close();	
	pRecordSet=NULL;
}
catch(_com_error &e){
	MessageBox(NULL,e.Description(),NULL,NULL);
	pRecordSet->Close();
	return -1;
}

	return 0;

}

int clsImportDM::UpdateDB(CString strMakeDate)
{

	_RecordsetPtr pRecordSet;
	_CommandPtr pCommand;      // Commandオブジェクト
	pCommand.CreateInstance(__uuidof(Command));
	pRecordSet.CreateInstance(__uuidof(Recordset));

	CString strsql;
	strsql="SELECT * FROM MINING WHERE [Year]='";
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
	// SQLの設定
	pCommand->ActiveConnection = pCn;
	pCommand->CommandText = bstrQuery;
	pRecordSet->PutRefSource(pCommand);

	// レコードセットの取得
	_variant_t vNull;  // VARIANT型のNULLとして使う
	vNull.vt = VT_ERROR;
	vNull.scode = DISP_E_PARAMNOTFOUND;
	pRecordSet->Open(vNull, vNull, adOpenForwardOnly, adLockOptimistic, adCmdText);

	COleSafeArray vaFieldlist;
	vaFieldlist.CreateOneDim(VT_VARIANT,82);

	COleSafeArray vaValuelist;
	vaValuelist.CreateOneDim(VT_VARIANT,82);
	long lArrayIndex[1];
	USES_CONVERSION;

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
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("MakeHM")));	
	strsql.SetString(mBuf.MakeHM.Hour,4);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	int i;
	for(i=0;i<18;i++){
		lArrayIndex[0]=10+4*i;
		strsql.Format("Umaban%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.DMInfo[i].Umaban,2);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));
			
		lArrayIndex[0]=11+4*i;
		strsql.Format("DMTime%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.DMInfo[i].DMTime,5);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));
			
		lArrayIndex[0]=12+4*i;
		strsql.Format("DMGosaP%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.DMInfo[i].DMGosaP,4);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));
			
		lArrayIndex[0]=13+4*i;
		strsql.Format("DMGosaM%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.DMInfo[i].DMGosaM,4);
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

	pRecordSet=NULL;
	return 0;
}


int clsImportDM::Init(_ConnectionPtr m_Connect)
{
    pCn = m_Connect;

	 return 0;
}

