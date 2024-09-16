/*=======================================================================
  JRA-VAN Data Lab.プログラミングパーツ「データ登録パーツ　RAレコード登録クラス」

	   作成: JRA-VAN ソフトウェア工房  2003年6月26日

========================================================================
   (C) Copyright JRA SYSTEM SERVICE CO.,LTD. 2007 All rights reserved
========================================================================*/

#include "StdAfx.h"
#include "clsimportra.h"

clsImportRA::clsImportRA(void)
{

}

clsImportRA::~clsImportRA(void)
{
//	pCn->Close();
}
int clsImportRA::Init(_ConnectionPtr m_Connect)
{
    pCn = m_Connect;
	return 0;
}

int clsImportRA::InsertDB(void)
{



	CString strsql;

	strsql="SELECT * FROM RACE WHERE [Year]='";
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
	USES_CONVERSION;


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


	COleSafeArray vaFieldlist;
	vaFieldlist.CreateOneDim(VT_VARIANT,110);

	COleSafeArray vaValuelist;
	vaValuelist.CreateOneDim(VT_VARIANT,110);
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
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("GradeCDBefore")));	
	strsql.SetString(mBuf.GradeCDBefore,1);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=24;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("SyubetuCD")));	
	strsql.SetString(mBuf.JyokenInfo.SyubetuCD,2);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=25;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("KigoCD")));	
	strsql.SetString(mBuf.JyokenInfo.KigoCD,3);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=26;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("JyuryoCD")));	
	strsql.SetString(mBuf.JyokenInfo.JyuryoCD,1);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	int i;
	for(i=0;i<5;i++){
		lArrayIndex[0]=27+i;
		strsql.Format("JyokenCD%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.JyokenInfo.JyokenCD[i],3);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));
	}
	lArrayIndex[0]=32;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("JyokenName")));	
	strsql.SetString(mBuf.JyokenName,60);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=33;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Kyori")));	
	strsql.SetString(mBuf.Kyori,4);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=34;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("KyoriBefore")));	
	strsql.SetString(mBuf.KyoriBefore,4);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=35;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("TrackCD")));	
	strsql.SetString(mBuf.TrackCD,2);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=36;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("TrackCDBefore")));	
	strsql.SetString(mBuf.TrackCDBefore,2);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=37;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("CourseKubunCD")));	
	strsql.SetString(mBuf.CourseKubunCD,2);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=38;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("CourseKubunCDBefore")));	
	strsql.SetString(mBuf.CourseKubunCDBefore,2);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	for(i=0;i<7;i++){
		lArrayIndex[0]=39+i;
		strsql.Format("Honsyokin%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.Honsyokin[i],8);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));
	}
	for(i=0;i<5;i++){
		lArrayIndex[0]=46+i;
		strsql.Format("HonsyokinBefore%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.HonsyokinBefore[i],8);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));
	}
	for(i=0;i<5;i++){
		lArrayIndex[0]=51+i;
		strsql.Format("Fukasyokin%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.Fukasyokin[i],8);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));
	}
	for(i=0;i<3;i++){
		lArrayIndex[0]=56+i;
		strsql.Format("FukasyokinBefore%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.FukasyokinBefore[i],8);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));
	}
	lArrayIndex[0]=59;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("HassoTime")));	
	strsql.SetString(mBuf.HassoTime,4);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=60;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("HassoTimeBefore")));	
	strsql.SetString(mBuf.HassoTimeBefore,4);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=61;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("TorokuTosu")));	
	strsql.SetString(mBuf.TorokuTosu,2);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=62;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("SyussoTosu")));	
	strsql.SetString(mBuf.SyussoTosu,2);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=63;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("NyusenTosu")));	
	strsql.SetString(mBuf.NyusenTosu,2);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=64;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("TenkoCD")));	
	strsql.SetString(mBuf.TenkoBaba.TenkoCD,1);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=65;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("SibaBabaCD")));	
	strsql.SetString(mBuf.TenkoBaba.SibaBabaCD,1);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=66;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("DirtBabaCD")));	
	strsql.SetString(mBuf.TenkoBaba.DirtBabaCD,1);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	for(i=0;i<25;i++){
		lArrayIndex[0]=67+i;
		strsql.Format("LapTime%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.LapTime[i],3);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));
	}

	lArrayIndex[0]=92;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("SyogaiMileTime")));	
	strsql.SetString(mBuf.SyogaiMileTime,4);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=93;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("HaronTimeS3")));	
	strsql.SetString(mBuf.HaronTimeS3,3);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=94;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("HaronTimeS4")));	
	strsql.SetString(mBuf.HaronTimeS4,3);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=95;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("HaronTimeL3")));	
	strsql.SetString(mBuf.HaronTimeL3,3);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=96;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("HaronTimeL4")));	
	strsql.SetString(mBuf.HaronTimeL4,3);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	for(i=0;i<4;i++){	
		lArrayIndex[0]=97+3*i;
		strsql.Format("Corner%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.CornerInfo[i].Corner,1);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		strsql.Format("Syukaisu%d",i+1);
		lArrayIndex[0]=98+3*i;
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.CornerInfo[i].Syukaisu,1);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		strsql.Format("Jyuni%d",i+1);
		lArrayIndex[0]=99+3*i;
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.CornerInfo[i].Jyuni,70);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));
	}

	lArrayIndex[0]=109;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("RecordUpKubun")));	
	strsql.SetString(mBuf.RecordUpKubun,1);
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

int clsImportRA::UpdateDB(CString strMakeDate)
{

		CString strsql;

	strsql="SELECT * FROM RACE WHERE [Year]='";
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
	USES_CONVERSION;


	_variant_t vNull;  // VARIANT型のNULLとして使う
	vNull.vt = VT_ERROR;
	vNull.scode = DISP_E_PARAMNOTFOUND;

	// レコードセットの取得
	pRecordSet->Open(vNull, vNull, adOpenKeyset, adLockOptimistic, adCmdText);


	COleSafeArray vaFieldlist;
	vaFieldlist.CreateOneDim(VT_VARIANT,110);

	COleSafeArray vaValuelist;
	vaValuelist.CreateOneDim(VT_VARIANT,110);
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
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("GradeCDBefore")));	
	strsql.SetString(mBuf.GradeCDBefore,1);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=24;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("SyubetuCD")));	
	strsql.SetString(mBuf.JyokenInfo.SyubetuCD,2);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=25;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("KigoCD")));	
	strsql.SetString(mBuf.JyokenInfo.KigoCD,3);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=26;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("JyuryoCD")));	
	strsql.SetString(mBuf.JyokenInfo.JyuryoCD,1);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	int i;
	for(i=0;i<5;i++){
		lArrayIndex[0]=27+i;
		strsql.Format("JyokenCD%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.JyokenInfo.JyokenCD[i],3);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));
	}
	lArrayIndex[0]=32;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("JyokenName")));	
	strsql.SetString(mBuf.JyokenName,60);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=33;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("Kyori")));	
	strsql.SetString(mBuf.Kyori,4);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=34;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("KyoriBefore")));	
	strsql.SetString(mBuf.KyoriBefore,4);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=35;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("TrackCD")));	
	strsql.SetString(mBuf.TrackCD,2);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=36;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("TrackCDBefore")));	
	strsql.SetString(mBuf.TrackCDBefore,2);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=37;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("CourseKubunCD")));	
	strsql.SetString(mBuf.CourseKubunCD,2);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=38;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("CourseKubunCDBefore")));	
	strsql.SetString(mBuf.CourseKubunCDBefore,2);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	for(i=0;i<7;i++){
		lArrayIndex[0]=39+i;
		strsql.Format("Honsyokin%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.Honsyokin[i],8);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));
	}
	for(i=0;i<5;i++){
		lArrayIndex[0]=46+i;
		strsql.Format("HonsyokinBefore%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.HonsyokinBefore[i],8);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));
	}
	for(i=0;i<5;i++){
		lArrayIndex[0]=51+i;
		strsql.Format("Fukasyokin%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.Fukasyokin[i],8);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));
	}
	for(i=0;i<3;i++){
		lArrayIndex[0]=56+i;
		strsql.Format("FukasyokinBefore%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.FukasyokinBefore[i],8);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));
	}
	lArrayIndex[0]=59;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("HassoTime")));	
	strsql.SetString(mBuf.HassoTime,4);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=60;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("HassoTimeBefore")));	
	strsql.SetString(mBuf.HassoTimeBefore,4);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=61;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("TorokuTosu")));	
	strsql.SetString(mBuf.TorokuTosu,2);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=62;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("SyussoTosu")));	
	strsql.SetString(mBuf.SyussoTosu,2);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=63;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("NyusenTosu")));	
	strsql.SetString(mBuf.NyusenTosu,2);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=64;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("TenkoCD")));	
	strsql.SetString(mBuf.TenkoBaba.TenkoCD,1);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=65;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("SibaBabaCD")));	
	strsql.SetString(mBuf.TenkoBaba.SibaBabaCD,1);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=66;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("DirtBabaCD")));	
	strsql.SetString(mBuf.TenkoBaba.DirtBabaCD,1);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	for(i=0;i<25;i++){
		lArrayIndex[0]=67+i;
		strsql.Format("LapTime%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.LapTime[i],3);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));
	}

	lArrayIndex[0]=92;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("SyogaiMileTime")));	
	strsql.SetString(mBuf.SyogaiMileTime,4);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=93;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("HaronTimeS3")));	
	strsql.SetString(mBuf.HaronTimeS3,3);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=94;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("HaronTimeS4")));	
	strsql.SetString(mBuf.HaronTimeS4,3);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=95;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("HaronTimeL3")));	
	strsql.SetString(mBuf.HaronTimeL3,3);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=96;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("HaronTimeL4")));	
	strsql.SetString(mBuf.HaronTimeL4,3);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	for(i=0;i<4;i++){	
		lArrayIndex[0]=97+3*i;
		strsql.Format("Corner%d",i+1);
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.CornerInfo[i].Corner,1);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		strsql.Format("Syukaisu%d",i+1);
		lArrayIndex[0]=98+3*i;
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.CornerInfo[i].Syukaisu,1);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

		strsql.Format("Jyuni%d",i+1);
		lArrayIndex[0]=99+3*i;
		vaFieldlist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));	
		strsql.SetString(mBuf.CornerInfo[i].Jyuni,70);
		vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));
	}

	lArrayIndex[0]=109;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("RecordUpKubun")));	
	strsql.SetString(mBuf.RecordUpKubun,1);
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

	


	return 0;
}

int clsImportRA::Add(CString strBuff, long lngBuffSize)
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