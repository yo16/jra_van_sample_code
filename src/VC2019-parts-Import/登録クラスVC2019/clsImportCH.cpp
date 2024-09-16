/*=======================================================================
  JRA-VAN Data Lab.プログラミングパーツ「データ登録パーツ　CHレコード登録クラス」

	   作成: JRA-VAN ソフトウェア工房  2003年6月26日

========================================================================
   (C) Copyright JRA SYSTEM SERVICE CO.,LTD. 2007 All rights reserved
========================================================================*/


#include "StdAfx.h"
#include "clsImportCH.h"
#include <afxdisp.h>


clsImportCH::clsImportCH(void)
{
}

clsImportCH::~clsImportCH(void)
{
}
int clsImportCH::Add(CString strBuff, long lngBuffSize)
{
	CString strMakeDate;
    memcpy(&mBuf,strBuff.GetBuffer(lngBuffSize),lngBuffSize);
    strMakeDate.SetString(mBuf.head.MakeDate.Year,4);
	strMakeDate.Append(mBuf.head.MakeDate.Month,2);
	strMakeDate.Append(mBuf.head.MakeDate.Day,2);

	// INSERT処理
	if(InsertDB() != 0 ){
		// UPDATE処理（INSERTが失敗した場合）
		if(UpdateDB(strMakeDate)!=0){
			// System.Diagnostics.Debug.WriteLine("更新に失敗しました。" & Left(strBuf, 2))
		}
	}
	return 0;
}

int clsImportCH::InsertDB(void)
{
	_RecordsetPtr pRecordSet;

	int i,j,k;
	CString strsql;
	strsql="SELECT * FROM CHOKYO WHERE [ChokyosiCode] = '";
	strsql.Append(mBuf.ChokyosiCode,5);
	strsql.Append("'");

	_bstr_t bstrQuery(strsql);

	// Commandオブジェクト
	_CommandPtr pCommand;
	pCommand.CreateInstance(__uuidof(Command));
	pRecordSet.CreateInstance(__uuidof(Recordset));

	pCommand->ActiveConnection = pCn;
	pCommand->CommandText = bstrQuery;
	pRecordSet->PutRefSource(pCommand);

	_variant_t vNull;  // VARIANT型のNULLとして使う
	vNull.vt = VT_ERROR;
	vNull.scode = DISP_E_PARAMNOTFOUND;
	pRecordSet->Open(vNull, vNull, adOpenKeyset, adLockOptimistic, adCmdText);
	if (!pRecordSet->GetadoEOF()){
		pRecordSet->Close();
		pRecordSet = NULL;
		return -1;
	}

	try{
		USES_CONVERSION;

		COleSafeArray vaFieldlist;
		vaFieldlist.CreateOneDim(VT_VARIANT,42);

		COleSafeArray vaValuelist;
		vaValuelist.CreateOneDim(VT_VARIANT,42);
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
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t("ChokyosiCode")));	
		strsql.SetString(mBuf.ChokyosiCode,5);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=4;
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t("DelKubun")));	
		strsql.SetString(mBuf.DelKubun,1);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=5;
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t("IssueDate")));	
		strsql.SetString(mBuf.IssueDate.Year,8);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=6;
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t("DelDate")));	
		strsql.SetString(mBuf.DelDate.Year,8);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=7;
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t("BirthDate")));	
		strsql.SetString(mBuf.BirthDate.Year,8);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=8;
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t("ChokyosiName")));	
		strsql.SetString(mBuf.ChokyosiName,34);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=9;
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t("ChokyosiNameKana")));	
		strsql.SetString(mBuf.ChokyosiNameKana,30);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=10;
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t("ChokyosiRyakusyo")));	
		strsql.SetString(mBuf.ChokyosiRyakusyo,8);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=11;
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t("ChokyosiNameEng")));
		strsql.SetString(mBuf.ChokyosiNameEng,80);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=12;
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t("SexCD")));	
		strsql.SetString(mBuf.SexCD,1);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=13;
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t("TozaiCD")));	
		strsql.SetString(mBuf.TozaiCD,1);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=14;
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Syotai")));	
		strsql.SetString(mBuf.Syotai,20);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		CString strHead;
		for(i=0;i<3;i++){

			strHead = "SaikinJyusyo";
			strHead.AppendFormat("%d",i+1);
			
			lArrayIndex[0]=15+9*i;
			strsql = strHead;
			strsql.Append("SaikinJyusyoid");
			vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
			strsql.SetString(mBuf.SaikinJyusyo[i].SaikinJyusyoid.Year,16);
			vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

			lArrayIndex[0]=16+9*i;
			strsql = strHead;
			strsql.Append("Hondai");
			vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
			strsql.SetString(mBuf.SaikinJyusyo[i].Hondai,60);
			vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

			lArrayIndex[0]=17+9*i;
			strsql = strHead;
			strsql.Append("Ryakusyo10");
			vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
			strsql.SetString(mBuf.SaikinJyusyo[i].Ryakusyo10,20);
			vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

			lArrayIndex[0]=18+9*i;
			strsql = strHead;
			strsql.Append("Ryakusyo6");
			vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
			strsql.SetString(mBuf.SaikinJyusyo[i].Ryakusyo6,12);
			vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

			lArrayIndex[0]=19+9*i;
			strsql = strHead;
			strsql.Append("Ryakusyo3");
			vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
			strsql.SetString(mBuf.SaikinJyusyo[i].Ryakusyo3,6);
			vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

			lArrayIndex[0]=20+9*i;
			strsql = strHead;
			strsql.Append("GradeCD");
			vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
			strsql.SetString(mBuf.SaikinJyusyo[i].GradeCD,1);
			vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

			lArrayIndex[0]=21+9*i;
			strsql = strHead;
			strsql.Append("SyussoTosu");
			vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
			strsql.SetString(mBuf.SaikinJyusyo[i].SyussoTosu,2);
			vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

			lArrayIndex[0]=22+9*i;
			strsql = strHead;
			strsql.Append("KettoNum");
			vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
			strsql.SetString(mBuf.SaikinJyusyo[i].KettoNum,10);
			vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

			lArrayIndex[0]=23+9*i;
			strsql = strHead;
			strsql.Append("Bamei");
			vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
			strsql.SetString(mBuf.SaikinJyusyo[i].Bamei,36);
			vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		}
		pRecordSet->AddNew(vaFieldlist, vaValuelist);
		pRecordSet->Close();	

		bstrQuery ="SELECT * FROM CHOKYO_SEISEKI WHERE [ChokyosiCode] IS NULL";
		vaFieldlist.CreateOneDim(VT_VARIANT,176);
		vaValuelist.CreateOneDim(VT_VARIANT,176);

		pCommand.CreateInstance(__uuidof(Command));
		pCommand->ActiveConnection = pCn;
		pCommand->CommandText = bstrQuery;		
		pRecordSet->PutRefSource(pCommand);

		// レコードセットの取得
		pRecordSet->Open(vNull, vNull, adOpenKeyset, adLockOptimistic, adCmdText);
		for(i=0;i<3;i++){

			lArrayIndex[0]=0;
			vaFieldlist.PutElement(lArrayIndex, &(_variant_t("MakeDate")));	
			strsql.SetString(mBuf.head.MakeDate.Year,8);
			vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

			lArrayIndex[0]=1;
			vaFieldlist.PutElement(lArrayIndex, &(_variant_t("ChokyosiCode")));			
			strsql.SetString(mBuf.ChokyosiCode,5);
			vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

			lArrayIndex[0]=2;
			vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Num")));			
			strsql.Format("%1d",i);
			vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

			lArrayIndex[0]=3;
			vaFieldlist.PutElement(lArrayIndex, &(_variant_t("SetYear")));			
			strsql.SetString(mBuf.HonZenRuikei[i].SetYear,4);
			vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

			lArrayIndex[0]=4;
			vaFieldlist.PutElement(lArrayIndex, &(_variant_t("HonSyokinHeichi")));			
			strsql.SetString(mBuf.HonZenRuikei[i].HonSyokinHeichi,10);
			vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

			lArrayIndex[0]=5;
			vaFieldlist.PutElement(lArrayIndex, &(_variant_t("HonSyokinSyogai")));			
			strsql.SetString(mBuf.HonZenRuikei[i].HonSyokinSyogai,10);
			vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

			lArrayIndex[0]=6;
			vaFieldlist.PutElement(lArrayIndex, &(_variant_t("FukaSyokinHeichi")));			
			strsql.SetString(mBuf.HonZenRuikei[i].FukaSyokinHeichi,10);
			vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

			lArrayIndex[0]=7;
			vaFieldlist.PutElement(lArrayIndex, &(_variant_t("FukaSyokinSyogai")));			
			strsql.SetString(mBuf.HonZenRuikei[i].FukaSyokinSyogai,10);
			vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

			for(j=0;j<6;j++){
				lArrayIndex[0]=8+j;
				strsql="HeichiChakukaisu";
				strsql.AppendFormat("%d",j+1);
				vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));			
				strsql.SetString(mBuf.HonZenRuikei[i].ChakuKaisuHeichi.Chakukaisu[j],6);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

			}
			for(j=0;j<6;j++){
				lArrayIndex[0]=14+j;
				strsql="SyogaiChakukaisu";
				strsql.AppendFormat("%d",j+1);
				vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));		
				strsql.SetString(mBuf.HonZenRuikei[i].ChakuKaisuSyogai.Chakukaisu[j],6);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));
			}
			for(j=0;j<20;j++){
				for(k=0;k<6;k++){
					lArrayIndex[0]=20+6*j+k;
					strsql="Jyo";
					strsql.AppendFormat("%dChakukaisu",j+1);
					strsql.AppendFormat("%d",k+1);
					vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));		
					strsql.SetString(mBuf.HonZenRuikei[i].ChakuKaisuJyo[j].Chakukaisu[k],6);
					vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));
				}
			}
			for(j=0;j<6;j++){
				for(k=0;k<6;k++){
					lArrayIndex[0]=140+6*j+k;
					strsql="Kyori";
					strsql.AppendFormat("%dChakukaisu",j+1);
					strsql.AppendFormat("%d",k+1);
					vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));		
					strsql.SetString(mBuf.HonZenRuikei[i].ChakuKaisuKyori[j].Chakukaisu[k],6);
					vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));
				}
			}
			pRecordSet->AddNew(vaFieldlist, vaValuelist);
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

int clsImportCH::UpdateDB(CString strMakeDate)
{
	_RecordsetPtr pRecordSet;

	int i,j,k;
	CString strsql;
	strsql="SELECT * FROM CHOKYO WHERE [ChokyosiCode] = '";
	strsql.Append(mBuf.ChokyosiCode,5);
	strsql.Append("' AND [MakeDate] <= '");
	strsql.Append(strMakeDate);
	strsql.Append("'");

	_bstr_t bstrQuery(strsql);

	// Commandオブジェクト
	_CommandPtr pCommand;
	pCommand.CreateInstance(__uuidof(Command));
	pRecordSet.CreateInstance(__uuidof(Recordset));

	pCommand->ActiveConnection = pCn;
	pCommand->CommandText = bstrQuery;
	pRecordSet->PutRefSource(pCommand);

	_variant_t vNull;  // VARIANT型のNULLとして使う
	vNull.vt = VT_ERROR;
	vNull.scode = DISP_E_PARAMNOTFOUND;
	pRecordSet->Open(vNull, vNull, adOpenKeyset, adLockOptimistic, adCmdText);
	
	try{
		USES_CONVERSION;

		COleSafeArray vaFieldlist;
		vaFieldlist.CreateOneDim(VT_VARIANT,42);

		COleSafeArray vaValuelist;
		vaValuelist.CreateOneDim(VT_VARIANT,42);
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
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t("ChokyosiCode")));	
		strsql.SetString(mBuf.ChokyosiCode,5);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=4;
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t("DelKubun")));	
		strsql.SetString(mBuf.DelKubun,1);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=5;
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t("IssueDate")));	
		strsql.SetString(mBuf.IssueDate.Year,8);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=6;
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t("DelDate")));	
		strsql.SetString(mBuf.DelDate.Year,8);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=7;
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t("BirthDate")));	
		strsql.SetString(mBuf.BirthDate.Year,8);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=8;
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t("ChokyosiName")));	
		strsql.SetString(mBuf.ChokyosiName,34);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=9;
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t("ChokyosiNameKana")));	
		strsql.SetString(mBuf.ChokyosiNameKana,30);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=10;
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t("ChokyosiRyakusyo")));	
		strsql.SetString(mBuf.ChokyosiRyakusyo,8);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=11;
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t("ChokyosiNameEng")));	
		strsql.SetString(mBuf.ChokyosiNameEng,80);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=12;
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t("SexCD")));	
		strsql.SetString(mBuf.SexCD,1);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=13;
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t("TozaiCD")));	
		strsql.SetString(mBuf.TozaiCD,1);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=14;
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Syotai")));	
		strsql.SetString(mBuf.Syotai,20);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		CString strHead;
		for(i=0;i<3;i++){
			strHead = "SaikinJyusyo";
			strHead.AppendFormat("%d",i+1);

			lArrayIndex[0]=15+9*i;
			strsql = strHead;
			strsql.Append("SaikinJyusyoid");
			vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
			strsql.SetString(mBuf.SaikinJyusyo[i].SaikinJyusyoid.Year,16);
			vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

			lArrayIndex[0]=16+9*i;
			strsql = strHead;
			strsql.Append("Hondai");
			vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
			strsql.SetString(mBuf.SaikinJyusyo[i].Hondai,60);
			vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

			lArrayIndex[0]=17+9*i;
			strsql = strHead;
			strsql.Append("Ryakusyo10");
			vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
			strsql.SetString(mBuf.SaikinJyusyo[i].Ryakusyo10,20);
			vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

			lArrayIndex[0]=18+9*i;
			strsql = strHead;
			strsql.Append("Ryakusyo6");
			vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
			strsql.SetString(mBuf.SaikinJyusyo[i].Ryakusyo6,12);
			vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

			lArrayIndex[0]=19+9*i;
			strsql = strHead;
			strsql.Append("Ryakusyo3");
			vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
			strsql.SetString(mBuf.SaikinJyusyo[i].Ryakusyo3,6);
			vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

			lArrayIndex[0]=20+9*i;
			strsql = strHead;
			strsql.Append("GradeCD");
			vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
			strsql.SetString(mBuf.SaikinJyusyo[i].GradeCD,1);
			vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

			lArrayIndex[0]=21+9*i;
			strsql = strHead;
			strsql.Append("SyussoTosu");
			vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
			strsql.SetString(mBuf.SaikinJyusyo[i].SyussoTosu,2);
			vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

			lArrayIndex[0]=22+9*i;
			strsql = strHead;
			strsql.Append("KettoNum");
			vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
			strsql.SetString(mBuf.SaikinJyusyo[i].KettoNum,10);
			vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

			lArrayIndex[0]=23+9*i;
			strsql = strHead;
			strsql.Append("Bamei");
			vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
			strsql.SetString(mBuf.SaikinJyusyo[i].Bamei,36);
			vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		}
		pRecordSet->Update(vaFieldlist, vaValuelist);
		pRecordSet->Close();	

		for(i=0;i<3;i++){

			strsql="SELECT * FROM CHOKYO_SEISEKI WHERE [ChokyosiCode] = '";
				strsql.Append(mBuf.ChokyosiCode,5);
				strsql.Append("'");
				strsql.Append(" AND [MakeDate] <= '");
				strsql.Append(strMakeDate);
				strsql.Append("'");
				strsql.Append(" AND [Num] = '");
				strsql.AppendFormat("%d",i);
				strsql.Append("'");
			
			_bstr_t bstrQuery(strsql);

			pCommand.CreateInstance(__uuidof(Command));
			pRecordSet.CreateInstance(__uuidof(Recordset));

			// SQLの設定
			pCommand->ActiveConnection = pCn;
			pCommand->CommandText = bstrQuery;
			pRecordSet->PutRefSource(pCommand);

			// レコードセットの取得
			pRecordSet->Open(vNull, vNull, adOpenKeyset, adLockOptimistic, adCmdText);
			if (pRecordSet->GetadoEOF()){
				pRecordSet->Close();
				continue;
			}

			vaFieldlist.CreateOneDim(VT_VARIANT,176);
			vaValuelist.CreateOneDim(VT_VARIANT,176);

			lArrayIndex[0]=0;
			vaFieldlist.PutElement(lArrayIndex, &(_variant_t("MakeDate")));	
			strsql.SetString(mBuf.head.MakeDate.Year,8);
			vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

			lArrayIndex[0]=1;
			vaFieldlist.PutElement(lArrayIndex, &(_variant_t("ChokyosiCode")));			
			strsql.SetString(mBuf.ChokyosiCode,5);
			vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

			lArrayIndex[0]=2;
			vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Num")));			
			strsql.Format("%1d",i);
			vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

			lArrayIndex[0]=3;
			vaFieldlist.PutElement(lArrayIndex, &(_variant_t("SetYear")));			
			strsql.SetString(mBuf.HonZenRuikei[i].SetYear,4);
			vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

			lArrayIndex[0]=4;
			vaFieldlist.PutElement(lArrayIndex, &(_variant_t("HonSyokinHeichi")));			
			strsql.SetString(mBuf.HonZenRuikei[i].HonSyokinHeichi,10);
			vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

			lArrayIndex[0]=5;
			vaFieldlist.PutElement(lArrayIndex, &(_variant_t("HonSyokinSyogai")));			
			strsql.SetString(mBuf.HonZenRuikei[i].HonSyokinSyogai,10);
			vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

			lArrayIndex[0]=6;
			vaFieldlist.PutElement(lArrayIndex, &(_variant_t("FukaSyokinHeichi")));			
			strsql.SetString(mBuf.HonZenRuikei[i].FukaSyokinHeichi,10);
			vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

			lArrayIndex[0]=7;
			vaFieldlist.PutElement(lArrayIndex, &(_variant_t("FukaSyokinSyogai")));			
			strsql.SetString(mBuf.HonZenRuikei[i].FukaSyokinSyogai,10);
			vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

			for(j=0;j<6;j++){
				lArrayIndex[0]=8+j;
				strsql="HeichiChakukaisu";
				strsql.AppendFormat("%d",j+1);
				vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));			
				strsql.SetString(mBuf.HonZenRuikei[i].ChakuKaisuHeichi.Chakukaisu[j],6);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

			}
			for(j=0;j<6;j++){
				lArrayIndex[0]=14+j;
				strsql="SyogaiChakukaisu";
				strsql.AppendFormat("%d",j+1);
				vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));		
				strsql.SetString(mBuf.HonZenRuikei[i].ChakuKaisuSyogai.Chakukaisu[j],6);
				vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));
			}
			for(j=0;j<20;j++){
				for(k=0;k<6;k++){
					lArrayIndex[0]=20+6*j+k;
					strsql="Jyo";
					strsql.AppendFormat("%dChakukaisu",j+1);
					strsql.AppendFormat("%d",k+1);
					vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));		
					strsql.SetString(mBuf.HonZenRuikei[i].ChakuKaisuJyo[j].Chakukaisu[k],6);
					vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));
				}
			}
			for(j=0;j<6;j++){
				for(k=0;k<6;k++){
					lArrayIndex[0]=140+6*j+k;
					strsql="Kyori";
					strsql.AppendFormat("%dChakukaisu",j+1);
					strsql.AppendFormat("%d",k+1);
					vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));		
					strsql.SetString(mBuf.HonZenRuikei[i].ChakuKaisuKyori[j].Chakukaisu[k],6);
					vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));
				}
			}
			pRecordSet->Update(vaFieldlist, vaValuelist);
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
int clsImportCH::Init(_ConnectionPtr m_Connect)
{
	pCn = m_Connect;
	return 0;
}

