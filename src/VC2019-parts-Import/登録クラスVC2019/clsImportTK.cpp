/*=======================================================================
  JRA-VAN Data Lab.プログラミングパーツ「データ登録パーツ　TKレコード登録クラス」

	   作成: JRA-VAN ソフトウェア工房  2003年6月26日

========================================================================
   (C) Copyright JRA SYSTEM SERVICE CO.,LTD. 2007 All rights reserved
========================================================================*/
#include "StdAfx.h"
#include "clsimporttk.h"
#include "JVData_Structure.h"



clsImportTK::clsImportTK(void)
{
}

clsImportTK::~clsImportTK(void)
{
//	mRS1->Close();
//	mRS1=0;
}


int clsImportTK::Add(CString strBuff, long lngBuffSize)
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

int clsImportTK::Init(_ConnectionPtr m_Connect)
{
    pCn = m_Connect;

	return 0;
}

int clsImportTK::InsertDB(void)
{



	CString strsql;

	strsql="SELECT * FROM TOKU_RACE WHERE [Year]='";
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
	vaFieldlist.CreateOneDim(VT_VARIANT,36);

	COleSafeArray vaValuelist;
	vaValuelist.CreateOneDim(VT_VARIANT,36);
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
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t("YoubiCD")));	
		strsql.SetString(mBuf.RaceInfo.YoubiCD,1);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=10;
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t("TokuNum")));	
		strsql.SetString(mBuf.RaceInfo.TokuNum,4);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=11;
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Hondai")));	
		strsql.SetString(mBuf.RaceInfo.Hondai,60);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=12;
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Fukudai")));	
		strsql.SetString(mBuf.RaceInfo.Fukudai,60);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=13;
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Kakko")));	
		strsql.SetString(mBuf.RaceInfo.Kakko,60);	
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=14;
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t("HondaiEng")));	
		strsql.SetString(mBuf.RaceInfo.HondaiEng,120);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=15;
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t("FukudaiEng")));	
		strsql.SetString(mBuf.RaceInfo.FukudaiEng,120);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=16;
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t("KakkoEng")));	
		strsql.SetString(mBuf.RaceInfo.KakkoEng,120);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=17;
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Ryakusyo10")));	
		strsql.SetString(mBuf.RaceInfo.Ryakusyo10,20);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=18;
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Ryakusyo6")));	
		strsql.SetString(mBuf.RaceInfo.Ryakusyo6,12);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=19;
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Ryakusyo3")));	
		strsql.SetString(mBuf.RaceInfo.Ryakusyo3,6);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=20;
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Kubun")));	
		strsql.SetString(mBuf.RaceInfo.Kubun,1);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=21;
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Nkai")));	
		strsql.SetString(mBuf.RaceInfo.Nkai,3);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=22;
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t("GradeCD")));	
		strsql.SetString(mBuf.GradeCD,1);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=23;
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t("SyubetuCD")));	
		strsql.SetString(mBuf.JyokenInfo.SyubetuCD,2);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=24;
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t("KigoCD")));	
		strsql.SetString(mBuf.JyokenInfo.KigoCD,3);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=25;
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t("JyuryoCD")));	
		strsql.SetString(mBuf.JyokenInfo.JyuryoCD,1);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		int i;
		for(i=0;i<5;i++){
			lArrayIndex[0]=26+i;
			strsql.Format("JyokenCD%d",i+1);
			vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
			strsql.SetString(mBuf.JyokenInfo.JyokenCD[i],3);
			vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));
		}
		lArrayIndex[0]=31;
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Kyori")));	
		strsql.SetString(mBuf.Kyori,4);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=32;
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t("TrackCD")));	
		strsql.SetString(mBuf.TrackCD,2);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=33;
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t("CourseKubunCD")));	
		strsql.SetString(mBuf.CourseKubunCD,2);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=34;
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t("HandiDate")));	
		strsql.SetString(mBuf.HandiDate.Year,8);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=35;
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t("TorokuTosu")));	
		strsql.SetString(mBuf.TorokuTosu,3);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));
			
		pRecordSet->AddNew(vaFieldlist, vaValuelist);
		pRecordSet->Close();	



		char tosu[4];
		strncpy_s(tosu,mBuf.TorokuTosu,3);
		int LoopTosuMax = atoi(tosu);

		strsql="SELECT * FROM TOKU WHERE [Year] IS NULL";
		bstrQuery=strsql;

		// SQLの設定
		pCommand->ActiveConnection = pCn;
		pCommand->CommandText = bstrQuery;
		pRecordSet->PutRefSource(pCommand);

		// レコードセットの取得
		pRecordSet->Open(vNull, vNull, adOpenKeyset, adLockOptimistic, adCmdText);

		vaFieldlist.CreateOneDim(VT_VARIANT,17);
		vaValuelist.CreateOneDim(VT_VARIANT,17);
		for(i=0;i<LoopTosuMax-1;i++){
			

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
			vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Num")));	
			strsql.SetString(mBuf.TokuUmaInfo[i].Num,3);
			vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

			lArrayIndex[0]=8;
			vaFieldlist.PutElement(lArrayIndex, &(_variant_t("KettoNum")));	
			strsql.SetString(mBuf.TokuUmaInfo[i].KettoNum,10);
			vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

			lArrayIndex[0]=9;
			vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Bamei")));	
			strsql.SetString(mBuf.TokuUmaInfo[i].Bamei,36);
			vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

			lArrayIndex[0]=10;
			vaFieldlist.PutElement(lArrayIndex, &(_variant_t("UmaKigoCD")));	
			strsql.SetString(mBuf.TokuUmaInfo[i].UmaKigoCD,2);
			vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

			lArrayIndex[0]=11;
			vaFieldlist.PutElement(lArrayIndex, &(_variant_t("SexCD")));	
			strsql.SetString(mBuf.TokuUmaInfo[i].SexCD,1);
			vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

			lArrayIndex[0]=12;
			vaFieldlist.PutElement(lArrayIndex, &(_variant_t("TozaiCD")));	
			strsql.SetString(mBuf.TokuUmaInfo[i].TozaiCD,1);
			vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

			lArrayIndex[0]=13;
			vaFieldlist.PutElement(lArrayIndex, &(_variant_t("ChokyosiCode")));	
			strsql.SetString(mBuf.TokuUmaInfo[i].ChokyosiCode,5);
			vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

			lArrayIndex[0]=14;
			vaFieldlist.PutElement(lArrayIndex, &(_variant_t("ChokyosiRyakusyo")));	
			strsql.SetString(mBuf.TokuUmaInfo[i].ChokyosiRyakusyo,8);
			vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

			lArrayIndex[0]=15;
			vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Futan")));	
			strsql.SetString(mBuf.TokuUmaInfo[i].Futan,3);
			vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

			lArrayIndex[0]=16;
			vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Koryu")));	
			strsql.SetString(mBuf.TokuUmaInfo[i].Koryu,1);
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

int clsImportTK::UpdateDB(CString strMakeDate)
{
	CString strsql;
	strsql="SELECT * FROM TOKU_RACE WHERE [Year]='";
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
	vaFieldlist.CreateOneDim(VT_VARIANT,36);

	COleSafeArray vaValuelist;
	vaValuelist.CreateOneDim(VT_VARIANT,36);
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
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t("YoubiCD")));	
		strsql.SetString(mBuf.RaceInfo.YoubiCD,1);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=10;
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t("TokuNum")));	
		strsql.SetString(mBuf.RaceInfo.TokuNum,4);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=11;
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Hondai")));	
		strsql.SetString(mBuf.RaceInfo.Hondai,60);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=12;
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Fukudai")));	
		strsql.SetString(mBuf.RaceInfo.Fukudai,60);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=13;
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Kakko")));	
		strsql.SetString(mBuf.RaceInfo.Kakko,60);	
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=14;
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t("HondaiEng")));	
		strsql.SetString(mBuf.RaceInfo.HondaiEng,120);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=15;
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t("FukudaiEng")));	
		strsql.SetString(mBuf.RaceInfo.FukudaiEng,120);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=16;
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t("KakkoEng")));	
		strsql.SetString(mBuf.RaceInfo.KakkoEng,120);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=17;
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Ryakusyo10")));	
		strsql.SetString(mBuf.RaceInfo.Ryakusyo10,20);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=18;
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Ryakusyo6")));	
		strsql.SetString(mBuf.RaceInfo.Ryakusyo6,12);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=19;
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Ryakusyo3")));	
		strsql.SetString(mBuf.RaceInfo.Ryakusyo3,6);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=20;
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Kubun")));	
		strsql.SetString(mBuf.RaceInfo.Kubun,1);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=21;
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Nkai")));	
		strsql.SetString(mBuf.RaceInfo.Nkai,3);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=22;
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t("GradeCD")));	
		strsql.SetString(mBuf.GradeCD,1);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=23;
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t("SyubetuCD")));	
		strsql.SetString(mBuf.JyokenInfo.SyubetuCD,2);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=24;
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t("KigoCD")));	
		strsql.SetString(mBuf.JyokenInfo.KigoCD,3);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=25;
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t("JyuryoCD")));	
		strsql.SetString(mBuf.JyokenInfo.JyuryoCD,1);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		int i;
		for(i=0;i<5;i++){
			lArrayIndex[0]=26+i;
			strsql.Format("JyokenCD%d",i+1);
			vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
			strsql.SetString(mBuf.JyokenInfo.JyokenCD[i],3);
			vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));
		}
		lArrayIndex[0]=31;
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Kyori")));	
		strsql.SetString(mBuf.Kyori,4);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=32;
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t("TrackCD")));	
		strsql.SetString(mBuf.TrackCD,2);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=33;
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t("CourseKubunCD")));	
		strsql.SetString(mBuf.CourseKubunCD,2);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=34;
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t("HandiDate")));	
		strsql.SetString(mBuf.HandiDate.Year,8);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		lArrayIndex[0]=35;
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t("TorokuTosu")));	
		strsql.SetString(mBuf.TorokuTosu,3);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));
			
		pRecordSet->Update(vaFieldlist, vaValuelist);
		pRecordSet->Close();	



		strsql = "DELETE FROM TOKU";
		strsql.Append(" WHERE [Year]='");
		strsql.Append(mBuf.id.Year,4);
		strsql.Append("' AND [MonthDay] = '");
		strsql.Append(mBuf.id.MonthDay,4);
		strsql.Append("' AND [JyoCD] = '");
		strsql.Append(mBuf.id.JyoCD,2);
		strsql.Append("' AND [Kaiji] = '");
		strsql.Append(mBuf.id.Kaiji,2);
		strsql.Append("' AND [Nichiji] = '");
		strsql.Append(mBuf.id.Nichiji,2);
		strsql.Append("' AND [RaceNum]= '");
		strsql.Append(mBuf.id.RaceNum,2);
		strsql.Append("' AND [MakeDate]<= '");
		strsql.Append(strMakeDate);
		strsql.Append("'");

		bstrQuery=strsql;
		_variant_t vRecsAffected(0L);
		pCn->Execute(bstrQuery,&vRecsAffected,adOptionUnspecified);


		char tosu[4];
		strncpy_s(tosu,mBuf.TorokuTosu,3);
		int LoopTosuMax = atoi(tosu);

		strsql="SELECT * FROM TOKU WHERE [Year] IS NULL";
		bstrQuery=strsql;

		// SQLの設定
		pCommand->ActiveConnection = pCn;
		pCommand->CommandText = bstrQuery;
		pRecordSet->PutRefSource(pCommand);

		// レコードセットの取得
		pRecordSet->Open(vNull, vNull, adOpenKeyset, adLockOptimistic, adCmdText);

		vaFieldlist.CreateOneDim(VT_VARIANT,17);
		vaValuelist.CreateOneDim(VT_VARIANT,17);
		for(i=0;i<LoopTosuMax-1;i++){
			

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
			vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Num")));	
			strsql.SetString(mBuf.TokuUmaInfo[i].Num,3);
			vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

			lArrayIndex[0]=8;
			vaFieldlist.PutElement(lArrayIndex, &(_variant_t("KettoNum")));	
			strsql.SetString(mBuf.TokuUmaInfo[i].KettoNum,10);
			vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

			lArrayIndex[0]=9;
			vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Bamei")));	
			strsql.SetString(mBuf.TokuUmaInfo[i].Bamei,36);
			vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

			lArrayIndex[0]=10;
			vaFieldlist.PutElement(lArrayIndex, &(_variant_t("UmaKigoCD")));	
			strsql.SetString(mBuf.TokuUmaInfo[i].UmaKigoCD,2);
			vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

			lArrayIndex[0]=11;
			vaFieldlist.PutElement(lArrayIndex, &(_variant_t("SexCD")));	
			strsql.SetString(mBuf.TokuUmaInfo[i].SexCD,1);
			vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

			lArrayIndex[0]=12;
			vaFieldlist.PutElement(lArrayIndex, &(_variant_t("TozaiCD")));	
			strsql.SetString(mBuf.TokuUmaInfo[i].TozaiCD,1);
			vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

			lArrayIndex[0]=13;
			vaFieldlist.PutElement(lArrayIndex, &(_variant_t("ChokyosiCode")));	
			strsql.SetString(mBuf.TokuUmaInfo[i].ChokyosiCode,5);
			vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

			lArrayIndex[0]=14;
			vaFieldlist.PutElement(lArrayIndex, &(_variant_t("ChokyosiRyakusyo")));	
			strsql.SetString(mBuf.TokuUmaInfo[i].ChokyosiRyakusyo,8);
			vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

			lArrayIndex[0]=15;
			vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Futan")));	
			strsql.SetString(mBuf.TokuUmaInfo[i].Futan,3);
			vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

			lArrayIndex[0]=16;
			vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Koryu")));	
			strsql.SetString(mBuf.TokuUmaInfo[i].Koryu,1);
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
