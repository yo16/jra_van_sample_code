#include "StdAfx.h"
#include "Shlwapi.h"
#include <afxdao.h>
#include <windows.h>

#include "clsDBBuilder.h"
#pragma warning(disable : 4995)
clsDBBuilder::clsDBBuilder(void)
{
}

clsDBBuilder::~clsDBBuilder(void)
{
}

int clsDBBuilder::CreateDB(CString strFilePath)
{
	

	CString strsql;
	HRESULT hr = S_OK;
    try
    {
		CString strcnn = "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=" + strFilePath;
		ADOX::_CatalogPtr m_pCatalog = NULL;
		hr = m_pCatalog.CreateInstance(__uuidof(ADOX::Catalog));
		if (SUCCEEDED(hr)) {
			m_pCatalog->Create(_bstr_t(strcnn));
		}
		else {
			return -1;
		}
		
    }
    catch(_com_error &e)
    {
        // エラー処理
        AfxMessageBox(e.Description());
        return -1;
    }

	
	
	try{
		hr = pCn.CreateInstance(__uuidof(Connection));
		if(SUCCEEDED(hr)){
			strsql= "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=" ;
			_bstr_t commandline(strsql+strFilePath);
			hr = pCn->Open(commandline, _bstr_t(L""), _bstr_t(L""), adModeUnknown);
		}else{
		}
	}
	catch(_com_error &e){
		(void)e;
		return -1;
	}
	
	pCn->BeginTrans();

	try{
		strsql = "CREATE TABLE TOKU_RACE (";
		strsql.Append("RecordSpec TEXT(2),");
		strsql.Append("DataKubun TEXT(1),");
		strsql.Append("MakeDate TEXT(8),");
		strsql.Append("[Year] TEXT(4),");
		strsql.Append("MonthDay TEXT(4),");
		strsql.Append("JyoCD TEXT(2),");
		strsql.Append("Kaiji TEXT(2),");
		strsql.Append("Nichiji TEXT(2),");
		strsql.Append("RaceNum TEXT(2),");
		strsql.Append("YoubiCD TEXT(1),");
		strsql.Append("TokuNum TEXT(4),");
		strsql.Append("Hondai TEXT(60),");
		strsql.Append("Fukudai TEXT(60),");
		strsql.Append("Kakko TEXT(60),");
		strsql.Append("HondaiEng TEXT(120),");
		strsql.Append("FukudaiEng TEXT(120),");
		strsql.Append("KakkoEng TEXT(120),");
		strsql.Append("Ryakusyo10 TEXT(20),");
		strsql.Append("Ryakusyo6 TEXT(12),");
		strsql.Append("Ryakusyo3 TEXT(6),");
		strsql.Append("Kubun TEXT(1),");
		strsql.Append("Nkai TEXT(3),");
		strsql.Append("GradeCD TEXT(1),");
		strsql.Append("SyubetuCD TEXT(2),");
		strsql.Append("KigoCD TEXT(3),");
		strsql.Append("JyuryoCD TEXT(1),");
		strsql.Append("JyokenCD1 TEXT(3),");
		strsql.Append("JyokenCD2 TEXT(3),");
		strsql.Append("JyokenCD3 TEXT(3),");
		strsql.Append("JyokenCD4 TEXT(3),");
		strsql.Append("JyokenCD5 TEXT(3),");
		strsql.Append("Kyori TEXT(4),");
		strsql.Append("TrackCD TEXT(2),");
		strsql.Append("CourseKubunCD TEXT(2),");
		strsql.Append("HandiDate TEXT(8),");
		strsql.Append("TorokuTosu TEXT(3),");
		strsql.Append("CONSTRAINT TOKU_RACE PRIMARY KEY ([Year],MonthDay,JyoCD,Kaiji,Nichiji,RaceNum));");

		_bstr_t bstrQuery(strsql);
		_variant_t vRecsAffected(0L);
		pCn->Execute(bstrQuery,&vRecsAffected,adOptionUnspecified);
		
		
		//特別登録馬
		strsql = "CREATE TABLE TOKU (";
		strsql.Append("MakeDate TEXT(8),");
		strsql.Append("[Year]     TEXT(4),");
		strsql.Append("MonthDay TEXT(4),");
		strsql.Append("JyoCD    TEXT(2),");
		strsql.Append("Kaiji    TEXT(2),");
		strsql.Append("Nichiji  TEXT(2),");
		strsql.Append("RaceNum  TEXT(2),");
		strsql.Append("Num      TEXT(3),");
		strsql.Append("KettoNum TEXT(10),");
		strsql.Append("Bamei    TEXT(36),");
		strsql.Append("UmaKigoCD        TEXT(2),");
		strsql.Append("SexCD    TEXT(1),");
		strsql.Append("TozaiCD  TEXT(1),");
		strsql.Append("ChokyosiCode     TEXT(5),");
		strsql.Append("ChokyosiRyakusyo TEXT(8),");
		strsql.Append("Futan    TEXT(3),");
		strsql.Append("Koryu    TEXT(1),");
		strsql.Append("CONSTRAINT TOKU PRIMARY KEY ([Year],MonthDay,JyoCD,Kaiji,Nichiji,RaceNum,Num));");
bstrQuery = strsql; pCn->Execute(bstrQuery,&vRecsAffected,adOptionUnspecified);;
		
		
		//レース詳細
		strsql = "CREATE TABLE RACE (";
		strsql.Append("RecordSpec       TEXT(2),");
		strsql.Append("DataKubun        TEXT(1),");
		strsql.Append("MakeDate TEXT(8),");
		strsql.Append("[Year]     TEXT(4),");
		strsql.Append("MonthDay TEXT(4),");
		strsql.Append("JyoCD    TEXT(2),");
		strsql.Append("Kaiji    TEXT(2),");
		strsql.Append("Nichiji  TEXT(2),");
		strsql.Append("RaceNum  TEXT(2),");
		strsql.Append("YoubiCD  TEXT(1),");
		strsql.Append("TokuNum  TEXT(4),");
		strsql.Append("Hondai   TEXT(60),");
		strsql.Append("Fukudai  TEXT(60),");
		strsql.Append("Kakko    TEXT(60),");
		strsql.Append("HondaiEng        TEXT(120),");
		strsql.Append("FukudaiEng       TEXT(120),");
		strsql.Append("KakkoEng TEXT(120),");
		strsql.Append("Ryakusyo10       TEXT(20),");
		strsql.Append("Ryakusyo6        TEXT(12),");
		strsql.Append("Ryakusyo3        TEXT(6),");
		strsql.Append("Kubun    TEXT(1),");
		strsql.Append("Nkai     TEXT(3),");
		strsql.Append("GradeCD  TEXT(1),");
		strsql.Append("GradeCDBefore    TEXT(1),");
		strsql.Append("SyubetuCD        TEXT(2),");
		strsql.Append("KigoCD   TEXT(3),");
		strsql.Append("JyuryoCD TEXT(1),");
		strsql.Append("JyokenCD1        TEXT(3),");
		strsql.Append("JyokenCD2        TEXT(3),");
		strsql.Append("JyokenCD3        TEXT(3),");
		strsql.Append("JyokenCD4        TEXT(3),");
		strsql.Append("JyokenCD5        TEXT(3),");
		strsql.Append("JyokenName       TEXT(60),");
		strsql.Append("Kyori    TEXT(4),");
		strsql.Append("KyoriBefore      TEXT(4),");
		strsql.Append("TrackCD  TEXT(2),");
		strsql.Append("TrackCDBefore    TEXT(2),");
		strsql.Append("CourseKubunCD    TEXT(2),");
		strsql.Append("CourseKubunCDBefore      TEXT(2),");
		strsql.Append("Honsyokin1       TEXT(8),");
		strsql.Append("Honsyokin2       TEXT(8),");
		strsql.Append("Honsyokin3       TEXT(8),");
		strsql.Append("Honsyokin4       TEXT(8),");
		strsql.Append("Honsyokin5       TEXT(8),");
		strsql.Append("Honsyokin6       TEXT(8),");
		strsql.Append("Honsyokin7       TEXT(8),");
		strsql.Append("HonsyokinBefore1 TEXT(8),");
		strsql.Append("HonsyokinBefore2 TEXT(8),");
		strsql.Append("HonsyokinBefore3 TEXT(8),");
		strsql.Append("HonsyokinBefore4 TEXT(8),");
		strsql.Append("HonsyokinBefore5 TEXT(8),");
		strsql.Append("Fukasyokin1      TEXT(8),");
		strsql.Append("Fukasyokin2      TEXT(8),");
		strsql.Append("Fukasyokin3      TEXT(8),");
		strsql.Append("Fukasyokin4      TEXT(8),");
		strsql.Append("Fukasyokin5      TEXT(8),");
		strsql.Append("FukasyokinBefore1        TEXT(8),");
		strsql.Append("FukasyokinBefore2        TEXT(8),");
		strsql.Append("FukasyokinBefore3        TEXT(8),");
		strsql.Append("HassoTime        TEXT(4),");
		strsql.Append("HassoTimeBefore  TEXT(4),");
		strsql.Append("TorokuTosu       TEXT(2),");
		strsql.Append("SyussoTosu       TEXT(2),");
		strsql.Append("NyusenTosu       TEXT(2),");
		strsql.Append("TenkoCD  TEXT(1),");
		strsql.Append("SibaBabaCD       TEXT(1),");
		strsql.Append("DirtBabaCD       TEXT(1),");
		strsql.Append("LapTime1 TEXT(3),");
		strsql.Append("LapTime2 TEXT(3),");
		strsql.Append("LapTime3 TEXT(3),");
		strsql.Append("LapTime4 TEXT(3),");
		strsql.Append("LapTime5 TEXT(3),");
		strsql.Append("LapTime6 TEXT(3),");
		strsql.Append("LapTime7 TEXT(3),");
		strsql.Append("LapTime8 TEXT(3),");
		strsql.Append("LapTime9 TEXT(3),");
		strsql.Append("LapTime10        TEXT(3),");
		strsql.Append("LapTime11        TEXT(3),");
		strsql.Append("LapTime12        TEXT(3),");
		strsql.Append("LapTime13        TEXT(3),");
		strsql.Append("LapTime14        TEXT(3),");
		strsql.Append("LapTime15        TEXT(3),");
		strsql.Append("LapTime16        TEXT(3),");
		strsql.Append("LapTime17        TEXT(3),");
		strsql.Append("LapTime18        TEXT(3),");
		strsql.Append("LapTime19        TEXT(3),");
		strsql.Append("LapTime20        TEXT(3),");
		strsql.Append("LapTime21        TEXT(3),");
		strsql.Append("LapTime22        TEXT(3),");
		strsql.Append("LapTime23        TEXT(3),");
		strsql.Append("LapTime24        TEXT(3),");
		strsql.Append("LapTime25        TEXT(3),");
		strsql.Append("SyogaiMileTime   TEXT(4),");
		strsql.Append("HaronTimeS3      TEXT(3),");
		strsql.Append("HaronTimeS4      TEXT(3),");
		strsql.Append("HaronTimeL3      TEXT(3),");
		strsql.Append("HaronTimeL4      TEXT(3),");
		strsql.Append("Corner1  TEXT(1),");
		strsql.Append("Syukaisu1        TEXT(1),");
		strsql.Append("Jyuni1   TEXT(70),");
		strsql.Append("Corner2  TEXT(1),");
		strsql.Append("Syukaisu2        TEXT(1),");
		strsql.Append("Jyuni2   TEXT(70),");
		strsql.Append("Corner3  TEXT(1),");
		strsql.Append("Syukaisu3        TEXT(1),");
		strsql.Append("Jyuni3   TEXT(70),");
		strsql.Append("Corner4  TEXT(1),");
		strsql.Append("Syukaisu4        TEXT(1),");
		strsql.Append("Jyuni4   TEXT(70),");
		strsql.Append("RecordUpKubun    TEXT(1),");
		strsql.Append("CONSTRAINT RACE PRIMARY KEY ([Year],MonthDay,JyoCD,Kaiji,Nichiji,RaceNum));");
		
bstrQuery = strsql; pCn->Execute(bstrQuery,&vRecsAffected,adOptionUnspecified);
		
		
		//馬毎レース情報
		strsql = "CREATE TABLE UMA_RACE (";
		strsql.Append("RecordSpec       TEXT(2),");
		strsql.Append("DataKubun        TEXT(1),");
		strsql.Append("MakeDate     TEXT(8),");
		strsql.Append("[Year]       TEXT (4),");
		strsql.Append("MonthDay   TEXT (4),");
		strsql.Append("JyoCD      TEXT (2),");
		strsql.Append("Kaiji      TEXT (2),");
		strsql.Append("Nichiji    TEXT (2),");
		strsql.Append("RaceNum    TEXT (2),");
		strsql.Append("Wakuban  TEXT(1),");
		strsql.Append("Umaban   TEXT(2),");
		strsql.Append("KettoNum TEXT(10),");
		strsql.Append("Bamei    TEXT(36),");
		strsql.Append("UmaKigoCD        TEXT(2),");
		strsql.Append("SexCD    TEXT(1),");
		strsql.Append("HinsyuCD TEXT(1),");
		strsql.Append("KeiroCD  TEXT(2),");
		strsql.Append("Barei    TEXT(2),");
		strsql.Append("TozaiCD  TEXT(1),");
		strsql.Append("ChokyosiCode     TEXT(5),");
		strsql.Append("ChokyosiRyakusyo TEXT(8),");
		strsql.Append("BanusiCode       TEXT(6),");
		strsql.Append("BanusiName       TEXT(64),");
		strsql.Append("Fukusyoku        TEXT(60),");
		strsql.Append("reserved1        TEXT(60),");
		strsql.Append("Futan    TEXT(3),");
		strsql.Append("FutanBefore      TEXT(3),");
		strsql.Append("Blinker  TEXT(1),");
		strsql.Append("reserved2        TEXT(1),");
		strsql.Append("KisyuCode        TEXT(5),");
		strsql.Append("KisyuCodeBefore  TEXT(5),");
		strsql.Append("KisyuRyakusyo    TEXT(8),");
		strsql.Append("KisyuRyakusyoBefore      TEXT(8),");
		strsql.Append("MinaraiCD        TEXT(1),");
		strsql.Append("MinaraiCDBefore  TEXT(1),");
		strsql.Append("BaTaijyu TEXT(3),");
		strsql.Append("ZogenFugo        TEXT(1),");
		strsql.Append("ZogenSa  TEXT(3),");
		strsql.Append("IJyoCD   TEXT(1),");
		strsql.Append("NyusenJyuni      TEXT(2),");
		strsql.Append("KakuteiJyuni     TEXT(2),");
		strsql.Append("DochakuKubun     TEXT(1),");
		strsql.Append("DochakuTosu      TEXT(1),");
		strsql.Append("[Time]     TEXT(4),");
		strsql.Append("ChakusaCD        TEXT(3),");
		strsql.Append("ChakusaCDP       TEXT(3),");
		strsql.Append("ChakusaCDPP      TEXT(3),");
		strsql.Append("Jyuni1c  TEXT(2),");
		strsql.Append("Jyuni2c  TEXT(2),");
		strsql.Append("Jyuni3c  TEXT(2),");
		strsql.Append("Jyuni4c  TEXT(2),");
		strsql.Append("Odds     TEXT(4),");
		strsql.Append("Ninki    TEXT(2),");
		strsql.Append("Honsyokin        TEXT(8),");
		strsql.Append("Fukasyokin       TEXT(8),");
		strsql.Append("reserved3        TEXT(3),");
		strsql.Append("reserved4        TEXT(3),");
		strsql.Append("HaronTimeL4      TEXT(3),");
		strsql.Append("HaronTimeL3      TEXT(3),");
		strsql.Append("KettoNum1        TEXT(10),");
		strsql.Append("Bamei1   TEXT(36),");
		strsql.Append("KettoNum2        TEXT(10),");
		strsql.Append("Bamei2   TEXT(36),");
		strsql.Append("KettoNum3        TEXT(10),");
		strsql.Append("Bamei3   TEXT(36),");
		strsql.Append("TimeDiff TEXT(4),");
		strsql.Append("RecordUpKubun    TEXT(1),");
		strsql.Append("DMKubun  TEXT(1),");
		strsql.Append("DMTime   TEXT(5),");
		strsql.Append("DMGosaP  TEXT(4),");
		strsql.Append("DMGosaM  TEXT(4),");
		strsql.Append("DMJyuni  TEXT(2),");
		strsql.Append("KyakusituKubun   TEXT(1),");
		strsql.Append("CONSTRAINT UMA_RACE PRIMARY KEY ([Year],MonthDay,JyoCD,Kaiji,Nichiji,RaceNum,Umaban,KettoNum));");
		
bstrQuery = strsql; pCn->Execute(bstrQuery,&vRecsAffected,adOptionUnspecified);;
		
		
		///払戻
		strsql = "CREATE TABLE HARAI (";
		strsql.Append("RecordSpec TEXT (2),");
		strsql.Append("DataKubun  TEXT (1),");
		strsql.Append("MakeDate   TEXT (8),");
		strsql.Append("[Year]       TEXT (4),");
		strsql.Append("MonthDay   TEXT (4),");
		strsql.Append("JyoCD      TEXT (2),");;
		strsql.Append("Kaiji      TEXT (2),");;
		strsql.Append("Nichiji    TEXT (2),");;
		strsql.Append("RaceNum    TEXT (2),");;
		strsql.Append("TorokuTosu TEXT (2),");;
		strsql.Append("SyussoTosu TEXT (2),");;
		strsql.Append("FuseirituFlag1   TEXT (1),");
		strsql.Append("FuseirituFlag2   TEXT (1),");
		strsql.Append("FuseirituFlag3   TEXT (1),");
		strsql.Append("FuseirituFlag4   TEXT (1),");
		strsql.Append("FuseirituFlag5   TEXT (1),");
		strsql.Append("FuseirituFlag6   TEXT (1),");
		strsql.Append("FuseirituFlag7   TEXT (1),");
		strsql.Append("FuseirituFlag8   TEXT (1),");
		strsql.Append("FuseirituFlag9   TEXT (1),");
		strsql.Append("TokubaraiFlag1   TEXT (1),");
		strsql.Append("TokubaraiFlag2   TEXT (1),");
		strsql.Append("TokubaraiFlag3   TEXT (1),");
		strsql.Append("TokubaraiFlag4   TEXT (1),");
		strsql.Append("TokubaraiFlag5   TEXT (1),");
		strsql.Append("TokubaraiFlag6   TEXT (1),");
		strsql.Append("TokubaraiFlag7   TEXT (1),");
		strsql.Append("TokubaraiFlag8   TEXT (1),");
		strsql.Append("TokubaraiFlag9   TEXT (1),");
		strsql.Append("HenkanFlag1   TEXT (1),");
		strsql.Append("HenkanFlag2   TEXT (1),");
		strsql.Append("HenkanFlag3   TEXT (1),");
		strsql.Append("HenkanFlag4   TEXT (1),");
		strsql.Append("HenkanFlag5   TEXT (1),");
		strsql.Append("HenkanFlag6   TEXT (1),");
		strsql.Append("HenkanFlag7   TEXT (1),");
		strsql.Append("HenkanFlag8   TEXT (1),");
		strsql.Append("HenkanFlag9   TEXT (1),");
		strsql.Append("HenkanUma1   TEXT (1),");
		strsql.Append("HenkanUma2   TEXT (1),");
		strsql.Append("HenkanUma3   TEXT (1),");
		strsql.Append("HenkanUma4   TEXT (1),");
		strsql.Append("HenkanUma5   TEXT (1),");
		strsql.Append("HenkanUma6   TEXT (1),");
		strsql.Append("HenkanUma7   TEXT (1),");
		strsql.Append("HenkanUma8   TEXT (1),");
		strsql.Append("HenkanUma9   TEXT (1),");
		strsql.Append("HenkanUma10   TEXT (1),");
		strsql.Append("HenkanUma11   TEXT (1),");
		strsql.Append("HenkanUma12   TEXT (1),");
		strsql.Append("HenkanUma13   TEXT (1),");
		strsql.Append("HenkanUma14   TEXT (1),");
		strsql.Append("HenkanUma15   TEXT (1),");
		strsql.Append("HenkanUma16   TEXT (1),");
		strsql.Append("HenkanUma17   TEXT (1),");
		strsql.Append("HenkanUma18   TEXT (1),");
		strsql.Append("HenkanUma19   TEXT (1),");
		strsql.Append("HenkanUma20   TEXT (1),");
		strsql.Append("HenkanUma21   TEXT (1),");
		strsql.Append("HenkanUma22   TEXT (1),");
		strsql.Append("HenkanUma23   TEXT (1),");
		strsql.Append("HenkanUma24   TEXT (1),");
		strsql.Append("HenkanUma25   TEXT (1),");
		strsql.Append("HenkanUma26   TEXT (1),");
		strsql.Append("HenkanUma27   TEXT (1),");
		strsql.Append("HenkanUma28   TEXT (1),");
		strsql.Append("HenkanWaku1   TEXT (1),");
		strsql.Append("HenkanWaku2   TEXT (1),");
		strsql.Append("HenkanWaku3   TEXT (1),");
		strsql.Append("HenkanWaku4   TEXT (1),");
		strsql.Append("HenkanWaku5   TEXT (1),");
		strsql.Append("HenkanWaku6   TEXT (1),");
		strsql.Append("HenkanWaku7   TEXT (1),");
		strsql.Append("HenkanWaku8   TEXT (1),");
		strsql.Append("HenkanDoWaku1   TEXT (1),");
		strsql.Append("HenkanDoWaku2   TEXT (1),");
		strsql.Append("HenkanDoWaku3   TEXT (1),");
		strsql.Append("HenkanDoWaku4   TEXT (1),");
		strsql.Append("HenkanDoWaku5   TEXT (1),");
		strsql.Append("HenkanDoWaku6   TEXT (1),");
		strsql.Append("HenkanDoWaku7   TEXT (1),");
		strsql.Append("HenkanDoWaku8   TEXT (1),");
		strsql.Append("PayTansyoUmaban1 TEXT (2),");
		strsql.Append("PayTansyoPay1    TEXT (9),");
		strsql.Append("PayTansyoNinki1  TEXT (2),");
		strsql.Append("PayTansyoUmaban2 TEXT (2),");
		strsql.Append("PayTansyoPay2    TEXT (9),");
		strsql.Append("PayTansyoNinki2  TEXT (2),");
		strsql.Append("PayTansyoUmaban3 TEXT (2),");
		strsql.Append("PayTansyoPay3    TEXT (9),");
		strsql.Append("PayTansyoNinki3  TEXT (2),");
		strsql.Append("PayFukusyoUmaban1 TEXT (2),");
		strsql.Append("PayFukusyoPay1    TEXT (9),");
		strsql.Append("PayFukusyoNinki1  TEXT (2),");
		strsql.Append("PayFukusyoUmaban2 TEXT (2),");
		strsql.Append("PayFukusyoPay2    TEXT (9),");
		strsql.Append("PayFukusyoNinki2  TEXT (2),");
		strsql.Append("PayFukusyoUmaban3 TEXT (2),");
		strsql.Append("PayFukusyoPay3    TEXT (9),");
		strsql.Append("PayFukusyoNinki3  TEXT (2),");
		strsql.Append("PayFukusyoUmaban4 TEXT (2),");
		strsql.Append("PayFukusyoPay4    TEXT (9),");
		strsql.Append("PayFukusyoNinki4  TEXT (2),");
		strsql.Append("PayFukusyoUmaban5 TEXT (2),");
		strsql.Append("PayFukusyoPay5    TEXT (9),");
		strsql.Append("PayFukusyoNinki5  TEXT (2),");
		strsql.Append("PayWakurenKumi1 TEXT (2),");
		strsql.Append("PayWakurenPay1    TEXT (9),");
		strsql.Append("PayWakurenNinki1  TEXT (2),");
		strsql.Append("PayWakurenKumi2 TEXT (2),");
		strsql.Append("PayWakurenPay2    TEXT (9),");
		strsql.Append("PayWakurenNinki2  TEXT (2),");
		strsql.Append("PayWakurenKumi3 TEXT (2),");
		strsql.Append("PayWakurenPay3    TEXT (9),");
		strsql.Append("PayWakurenNinki3  TEXT (2),");
		strsql.Append("PayUmarenKumi1 TEXT (4),");
		strsql.Append("PayUmarenPay1    TEXT (9),");
		strsql.Append("PayUmarenNinki1  TEXT (3),");
		strsql.Append("PayUmarenKumi2 TEXT (4),");
		strsql.Append("PayUmarenPay2    TEXT (9),");
		strsql.Append("PayUmarenNinki2  TEXT (3),");
		strsql.Append("PayUmarenKumi3 TEXT (4),");
		strsql.Append("PayUmarenPay3    TEXT (9),");
		strsql.Append("PayUmarenNinki3  TEXT (3),");
		strsql.Append("PayWideKumi1 TEXT (4),");
		strsql.Append("PayWidePay1    TEXT (9),");
		strsql.Append("PayWideNinki1  TEXT (3),");
		strsql.Append("PayWideKumi2 TEXT (4),");
		strsql.Append("PayWidePay2    TEXT (9),");
		strsql.Append("PayWideNinki2  TEXT (3),");
		strsql.Append("PayWideKumi3 TEXT (4),");
		strsql.Append("PayWidePay3    TEXT (9),");
		strsql.Append("PayWideNinki3  TEXT (3),");
		strsql.Append("PayWideKumi4 TEXT (4),");
		strsql.Append("PayWidePay4    TEXT (9),");
		strsql.Append("PayWideNinki4  TEXT (3),");
		strsql.Append("PayWideKumi5 TEXT (4),");
		strsql.Append("PayWidePay5    TEXT (9),");
		strsql.Append("PayWideNinki5  TEXT (3),");
		strsql.Append("PayWideKumi6 TEXT (4),");
		strsql.Append("PayWidePay6    TEXT (9),");
		strsql.Append("PayWideNinki6  TEXT (3),");
		strsql.Append("PayWideKumi7 TEXT (4),");
		strsql.Append("PayWidePay7    TEXT (9),");
		strsql.Append("PayWideNinki7  TEXT (3),");
		strsql.Append("PayReserved1Kumi1 TEXT (4),");
		strsql.Append("PayReserved1Pay1    TEXT (9),");
		strsql.Append("PayReserved1Ninki1  TEXT (3),");
		strsql.Append("PayReserved1Kumi2 TEXT (4),");
		strsql.Append("PayReserved1Pay2    TEXT (9),");
		strsql.Append("PayReserved1Ninki2  TEXT (3),");
		strsql.Append("PayReserved1Kumi3 TEXT (4),");
		strsql.Append("PayReserved1Pay3    TEXT (9),");
		strsql.Append("PayReserved1Ninki3  TEXT (3),");
		strsql.Append("PayUmatanKumi1  TEXT (4),");
		strsql.Append("PayUmatanPay1   TEXT (9),");
		strsql.Append("PayUmatanNinki1 TEXT (3),");
		strsql.Append("PayUmatanKumi2  TEXT (4),");
		strsql.Append("PayUmatanPay2   TEXT (9),");
		strsql.Append("PayUmatanNinki2 TEXT (3),");
		strsql.Append("PayUmatanKumi3  TEXT (4),");
		strsql.Append("PayUmatanPay3   TEXT (9),");
		strsql.Append("PayUmatanNinki3 TEXT (3),");
		strsql.Append("PayUmatanKumi4  TEXT (4),");
		strsql.Append("PayUmatanPay4   TEXT (9),");
		strsql.Append("PayUmatanNinki4 TEXT (3),");
		strsql.Append("PayUmatanKumi5  TEXT (4),");
		strsql.Append("PayUmatanPay5   TEXT (9),");
		strsql.Append("PayUmatanNinki5 TEXT (3),");
		strsql.Append("PayUmatanKumi6  TEXT (4),");
		strsql.Append("PayUmatanPay6   TEXT (9),");
		strsql.Append("PayUmatanNinki6 TEXT (3),");
		strsql.Append("PaySanrenpukuKumi1  TEXT (6),");
		strsql.Append("PaySanrenpukuPay1   TEXT (9),");
		strsql.Append("PaySanrenpukuNinki1 TEXT (3),");
		strsql.Append("PaySanrenpukuKumi2  TEXT (6),");
		strsql.Append("PaySanrenpukuPay2   TEXT (9),");
		strsql.Append("PaySanrenpukuNinki2 TEXT (3),");
		strsql.Append("PaySanrenpukuKumi3  TEXT (6),");
		strsql.Append("PaySanrenpukuPay3   TEXT (9),");
		strsql.Append("PaySanrenpukuNinki3 TEXT (3),");
		strsql.Append("PaySanrentanKumi1   TEXT (6),");
		strsql.Append("PaySanrentanPay1    TEXT (9),");
		strsql.Append("PaySanrentanNinki1  TEXT (4),");
		strsql.Append("PaySanrentanKumi2   TEXT (6),");
		strsql.Append("PaySanrentanPay2    TEXT (9),");
		strsql.Append("PaySanrentanNinki2  TEXT (4),");
		strsql.Append("PaySanrentanKumi3   TEXT (6),");
		strsql.Append("PaySanrentanPay3    TEXT (9),");
		strsql.Append("PaySanrentanNinki3  TEXT (4),");
		strsql.Append("PaySanrentanKumi4   TEXT (6),");
		strsql.Append("PaySanrentanPay4    TEXT (9),");
		strsql.Append("PaySanrentanNinki4  TEXT (4),");
		strsql.Append("PaySanrentanKumi5   TEXT (6),");
		strsql.Append("PaySanrentanPay5    TEXT (9),");
		strsql.Append("PaySanrentanNinki5  TEXT (4),");
		strsql.Append("PaySanrentanKumi6   TEXT (6),");
		strsql.Append("PaySanrentanPay6    TEXT (9),");
		strsql.Append("PaySanrentanNinki6  TEXT (4),");
		strsql.Append("CONSTRAINT HARAI PRIMARY KEY ([Year],MonthDay,JyoCD,Kaiji,Nichiji,RaceNum));");
		
bstrQuery = strsql; pCn->Execute(bstrQuery,&vRecsAffected,adOptionUnspecified);
		
		
		//票数
		strsql = "CREATE TABLE HYOSU (";
		strsql.Append("RecordSpec       TEXT(2),");
		strsql.Append("DataKubun        TEXT(1),");
		strsql.Append("MakeDate TEXT(8),");
		strsql.Append("[Year]     TEXT(4),");
		strsql.Append("MonthDay TEXT(4),");
		strsql.Append("JyoCD    TEXT(2),");
		strsql.Append("Kaiji    TEXT(2),");
		strsql.Append("Nichiji  TEXT(2),");
		strsql.Append("RaceNum  TEXT(2),");
		strsql.Append("TorokuTosu       TEXT(2),");
		strsql.Append("SyussoTosu       TEXT(2),");
		strsql.Append("HatubaiFlag1     TEXT(1),");
		strsql.Append("HatubaiFlag2     TEXT(1),");
		strsql.Append("HatubaiFlag3     TEXT(1),");
		strsql.Append("HatubaiFlag4     TEXT(1),");
		strsql.Append("HatubaiFlag5     TEXT(1),");
		strsql.Append("HatubaiFlag6     TEXT(1),");
		strsql.Append("HatubaiFlag7     TEXT(1),");
		strsql.Append("FukuChakuBaraiKey        TEXT(1),");
		strsql.Append("HenkanUma1       TEXT(1),");
		strsql.Append("HenkanUma2       TEXT(1),");
		strsql.Append("HenkanUma3       TEXT(1),");
		strsql.Append("HenkanUma4       TEXT(1),");
		strsql.Append("HenkanUma5       TEXT(1),");
		strsql.Append("HenkanUma6       TEXT(1),");
		strsql.Append("HenkanUma7       TEXT(1),");
		strsql.Append("HenkanUma8       TEXT(1),");
		strsql.Append("HenkanUma9       TEXT(1),");
		strsql.Append("HenkanUma10      TEXT(1),");
		strsql.Append("HenkanUma11      TEXT(1),");
		strsql.Append("HenkanUma12      TEXT(1),");
		strsql.Append("HenkanUma13      TEXT(1),");
		strsql.Append("HenkanUma14      TEXT(1),");
		strsql.Append("HenkanUma15      TEXT(1),");
		strsql.Append("HenkanUma16      TEXT(1),");
		strsql.Append("HenkanUma17      TEXT(1),");
		strsql.Append("HenkanUma18      TEXT(1),");
		strsql.Append("HenkanUma19      TEXT(1),");
		strsql.Append("HenkanUma20      TEXT(1),");
		strsql.Append("HenkanUma21      TEXT(1),");
		strsql.Append("HenkanUma22      TEXT(1),");
		strsql.Append("HenkanUma23      TEXT(1),");
		strsql.Append("HenkanUma24      TEXT(1),");
		strsql.Append("HenkanUma25      TEXT(1),");
		strsql.Append("HenkanUma26      TEXT(1),");
		strsql.Append("HenkanUma27      TEXT(1),");
		strsql.Append("HenkanUma28      TEXT(1),");
		strsql.Append("HenkanWaku1      TEXT(1),");
		strsql.Append("HenkanWaku2      TEXT(1),");
		strsql.Append("HenkanWaku3      TEXT(1),");
		strsql.Append("HenkanWaku4      TEXT(1),");
		strsql.Append("HenkanWaku5      TEXT(1),");
		strsql.Append("HenkanWaku6      TEXT(1),");
		strsql.Append("HenkanWaku7      TEXT(1),");
		strsql.Append("HenkanWaku8      TEXT(1),");
		strsql.Append("HenkanDoWaku1    TEXT(1),");
		strsql.Append("HenkanDoWaku2    TEXT(1),");
		strsql.Append("HenkanDoWaku3    TEXT(1),");
		strsql.Append("HenkanDoWaku4    TEXT(1),");
		strsql.Append("HenkanDoWaku5    TEXT(1),");
		strsql.Append("HenkanDoWaku6    TEXT(1),");
		strsql.Append("HenkanDoWaku7    TEXT(1),");
		strsql.Append("HenkanDoWaku8    TEXT(1),");
		strsql.Append("HyoTotal1        TEXT(11),");
		strsql.Append("HyoTotal2        TEXT(11),");
		strsql.Append("HyoTotal3        TEXT(11),");
		strsql.Append("HyoTotal4        TEXT(11),");
		strsql.Append("HyoTotal5        TEXT(11),");
		strsql.Append("HyoTotal6        TEXT(11),");
		strsql.Append("HyoTotal7        TEXT(11),");
		strsql.Append("HyoTotal8        TEXT(11),");
		strsql.Append("HyoTotal9        TEXT(11),");
		strsql.Append("HyoTotal10       TEXT(11),");
		strsql.Append("HyoTotal11       TEXT(11),");
		strsql.Append("HyoTotal12       TEXT(11),");
		strsql.Append("HyoTotal13       TEXT(11),");
		strsql.Append("HyoTotal14       TEXT(11),");
		strsql.Append("CONSTRAINT HYOSU PRIMARY KEY ([Year],MonthDay,JyoCD,Kaiji,Nichiji,RaceNum));");
		
bstrQuery = strsql; pCn->Execute(bstrQuery,&vRecsAffected,adOptionUnspecified);
		
		
		//票数２(３連単)
		strsql = "CREATE TABLE HYOSU2 (";
		strsql.Append("RecordSpec       TEXT(2),");
		strsql.Append("DataKubun        TEXT(1),");
		strsql.Append("MakeDate TEXT(8),");
		strsql.Append("[Year]     TEXT(4),");
		strsql.Append("MonthDay TEXT(4),");
		strsql.Append("JyoCD    TEXT(2),");
		strsql.Append("Kaiji    TEXT(2),");
		strsql.Append("Nichiji  TEXT(2),");
		strsql.Append("RaceNum  TEXT(2),");
		strsql.Append("TorokuTosu       TEXT(2),");
		strsql.Append("SyussoTosu       TEXT(2),");
		strsql.Append("HatubaiFlag1     TEXT(1),");
		strsql.Append("HenkanUma1       TEXT(1),");
		strsql.Append("HenkanUma2       TEXT(1),");
		strsql.Append("HenkanUma3       TEXT(1),");
		strsql.Append("HenkanUma4       TEXT(1),");
		strsql.Append("HenkanUma5       TEXT(1),");
		strsql.Append("HenkanUma6       TEXT(1),");
		strsql.Append("HenkanUma7       TEXT(1),");
		strsql.Append("HenkanUma8       TEXT(1),");
		strsql.Append("HenkanUma9       TEXT(1),");
		strsql.Append("HenkanUma10      TEXT(1),");
		strsql.Append("HenkanUma11      TEXT(1),");
		strsql.Append("HenkanUma12      TEXT(1),");
		strsql.Append("HenkanUma13      TEXT(1),");
		strsql.Append("HenkanUma14      TEXT(1),");
		strsql.Append("HenkanUma15      TEXT(1),");
		strsql.Append("HenkanUma16      TEXT(1),");
		strsql.Append("HenkanUma17      TEXT(1),");
		strsql.Append("HenkanUma18      TEXT(1),");
		strsql.Append("HyoTotal1        TEXT(11),");
		strsql.Append("HyoTotal2        TEXT(11),");
		strsql.Append("CONSTRAINT HYOSU2 PRIMARY KEY ([Year],MonthDay,JyoCD,Kaiji,Nichiji,RaceNum));");
		
bstrQuery = strsql; pCn->Execute(bstrQuery,&vRecsAffected,adOptionUnspecified);


		//票数_単複
		strsql = "CREATE TABLE HYOSU_TANPUKU (";
		strsql.Append("MakeDate TEXT(8),");
		strsql.Append("[Year]     TEXT(4),");
		strsql.Append("MonthDay TEXT(4),");
		strsql.Append("JyoCD    TEXT(2),");
		strsql.Append("Kaiji    TEXT(2),");
		strsql.Append("Nichiji  TEXT(2),");
		strsql.Append("RaceNum  TEXT(2),");
		strsql.Append("Umaban   TEXT(2),");
		strsql.Append("TanHyo   TEXT(11),");
		strsql.Append("TanNinki TEXT(2),");
		strsql.Append("FukuHyo  TEXT(11),");
		strsql.Append("FukuNinki        TEXT(2),");
		strsql.Append("CONSTRAINT HYOSU_TANPUKU PRIMARY KEY ([Year],MonthDay,JyoCD,Kaiji,Nichiji,RaceNum,Umaban));");
		
bstrQuery = strsql; pCn->Execute(bstrQuery,&vRecsAffected,adOptionUnspecified);
		
		
		//票数_枠連
		strsql = "CREATE TABLE HYOSU_WAKU (";
		strsql.Append("MakeDate TEXT(8),");
		strsql.Append("[Year]     TEXT(4),");
		strsql.Append("MonthDay TEXT(4),");
		strsql.Append("JyoCD    TEXT(2),");
		strsql.Append("Kaiji    TEXT(2),");
		strsql.Append("Nichiji  TEXT(2),");
		strsql.Append("RaceNum  TEXT(2),");
		strsql.Append("Kumi     TEXT(2),");
		strsql.Append("Hyo      TEXT(11),");
		strsql.Append("Ninki    TEXT(2),");
		strsql.Append("CONSTRAINT HYOSU_WAKU PRIMARY KEY ([Year],MonthDay,JyoCD,Kaiji,Nichiji,RaceNum,Kumi));");
		
bstrQuery = strsql; pCn->Execute(bstrQuery,&vRecsAffected,adOptionUnspecified);
		
		
		//票数_馬連ワイド
		strsql = "CREATE TABLE HYOSU_UMARENWIDE (";
		strsql.Append("MakeDate TEXT(8),");
		strsql.Append("[Year]     TEXT(4),");
		strsql.Append("MonthDay TEXT(4),");
		strsql.Append("JyoCD    TEXT(2),");
		strsql.Append("Kaiji    TEXT(2),");
		strsql.Append("Nichiji  TEXT(2),");
		strsql.Append("RaceNum  TEXT(2),");
		strsql.Append("Kumi     TEXT(4),");
		strsql.Append("UmarenHyo        TEXT(11),");
		strsql.Append("UmarenNinki      TEXT(3),");
		strsql.Append("WideHyo  TEXT(11),");
		strsql.Append("WideNinki        TEXT(3),");
		strsql.Append("CONSTRAINT HYOSU_UMARENWIDE PRIMARY KEY ([Year],MonthDay,JyoCD,Kaiji,Nichiji,RaceNum,Kumi));");
		
bstrQuery = strsql; pCn->Execute(bstrQuery,&vRecsAffected,adOptionUnspecified);
		
		
		//票数_馬単
		strsql = "CREATE TABLE HYOSU_UMATAN (";
		strsql.Append("MakeDate TEXT(8),");
		strsql.Append("[Year]     TEXT(4),");
		strsql.Append("MonthDay TEXT(4),");
		strsql.Append("JyoCD    TEXT(2),");
		strsql.Append("Kaiji    TEXT(2),");
		strsql.Append("Nichiji  TEXT(2),");
		strsql.Append("RaceNum  TEXT(2),");
		strsql.Append("Kumi     TEXT(4),");
		strsql.Append("Hyo      TEXT(11),");
		strsql.Append("Ninki    TEXT(3),");
		strsql.Append("CONSTRAINT HYOSU_UMATAN PRIMARY KEY ([Year],MonthDay,JyoCD,Kaiji,Nichiji,RaceNum,Kumi));");
		
bstrQuery = strsql; pCn->Execute(bstrQuery,&vRecsAffected,adOptionUnspecified);
		
		
		//票数_３連複
		strsql = "CREATE TABLE HYOSU_SANREN (";
		strsql.Append("MakeDate TEXT(8),");
		strsql.Append("[Year]     TEXT(4),");
		strsql.Append("MonthDay TEXT(4),");
		strsql.Append("JyoCD    TEXT(2),");
		strsql.Append("Kaiji    TEXT(2),");
		strsql.Append("Nichiji  TEXT(2),");
		strsql.Append("RaceNum  TEXT(2),");
		strsql.Append("Kumi     TEXT(6),");
		strsql.Append("Hyo      TEXT(11),");
		strsql.Append("Ninki    TEXT(3),");
		strsql.Append("CONSTRAINT HYOSU_SANREN PRIMARY KEY ([Year],MonthDay,JyoCD,Kaiji,Nichiji,RaceNum,Kumi));");

bstrQuery = strsql; pCn->Execute(bstrQuery,&vRecsAffected,adOptionUnspecified);

		//票数_３連単
		strsql = "CREATE TABLE HYOSU_SANRENTAN (";
		strsql.Append("MakeDate TEXT(8),");
		strsql.Append("[Year]     TEXT(4),");
		strsql.Append("MonthDay TEXT(4),");
		strsql.Append("JyoCD    TEXT(2),");
		strsql.Append("Kaiji    TEXT(2),");
		strsql.Append("Nichiji  TEXT(2),");
		strsql.Append("RaceNum  TEXT(2),");
		strsql.Append("Kumi     TEXT(6),");
		strsql.Append("Hyo      TEXT(11),");
		strsql.Append("Ninki    TEXT(4),");
		strsql.Append("CONSTRAINT HYOSU_SANRENTAN PRIMARY KEY ([Year],MonthDay,JyoCD,Kaiji,Nichiji,RaceNum,Kumi));");
		
bstrQuery = strsql; pCn->Execute(bstrQuery,&vRecsAffected,adOptionUnspecified);
		
		
		//オッズ_単複枠_ヘッダ
		strsql = "CREATE TABLE ODDS_TANPUKUWAKU_HEAD (";
		strsql.Append("RecordSpec       TEXT(2),");
		strsql.Append("DataKubun        TEXT(1),");
		strsql.Append("MakeDate TEXT(8),");
		strsql.Append("[Year]     TEXT(4),");
		strsql.Append("MonthDay TEXT(4),");
		strsql.Append("JyoCD    TEXT(2),");
		strsql.Append("Kaiji    TEXT(2),");
		strsql.Append("Nichiji  TEXT(2),");
		strsql.Append("RaceNum  TEXT(2),");
		strsql.Append("HappyoTime       TEXT(8),");
		strsql.Append("TorokuTosu       TEXT(2),");
		strsql.Append("SyussoTosu       TEXT(2),");
		strsql.Append("TansyoFlag       TEXT(1),");
		strsql.Append("FukusyoFlag      TEXT(1),");
		strsql.Append("WakurenFlag      TEXT(1),");
		strsql.Append("FukuChakuBaraiKey        TEXT(1),");
		strsql.Append("TotalHyosuTansyo     TEXT(11),");
		strsql.Append("TotalHyosuFukusyo        TEXT(11),");
		strsql.Append("TotalHyosuWakuren        TEXT(11),");
		strsql.Append("CONSTRAINT ODDS_TANPUKUWAKU_HEAD PRIMARY KEY ([Year],MonthDay,JyoCD,Kaiji,Nichiji,RaceNum));");
		
bstrQuery = strsql; pCn->Execute(bstrQuery,&vRecsAffected,adOptionUnspecified);
		
		
		//オッズ_単複
		strsql = "CREATE TABLE ODDS_TANPUKU (";
		strsql.Append("MakeDate TEXT(8),");
		strsql.Append("[Year]     TEXT(4),");
		strsql.Append("MonthDay TEXT(4),");
		strsql.Append("JyoCD    TEXT(2),");
		strsql.Append("Kaiji    TEXT(2),");
		strsql.Append("Nichiji  TEXT(2),");
		strsql.Append("RaceNum  TEXT(2),");
		strsql.Append("Umaban   TEXT(2),");
		strsql.Append("TanOdds  TEXT(4),");
		strsql.Append("TanNinki TEXT(2),");
		strsql.Append("FukuOddsLow      TEXT(4),");
		strsql.Append("FukuOddsHigh     TEXT(4),");
		strsql.Append("FukuNinki        TEXT(2),");
		strsql.Append("CONSTRAINT ODDS_TANPUKU PRIMARY KEY ([Year],MonthDay,JyoCD,Kaiji,Nichiji,RaceNum,Umaban));");
		
bstrQuery = strsql; pCn->Execute(bstrQuery,&vRecsAffected,adOptionUnspecified);
		
		
		//オッズ_枠連
		strsql = "CREATE TABLE ODDS_WAKU (";
		strsql.Append("MakeDate TEXT(8),");
		strsql.Append("[Year]     TEXT(4),");
		strsql.Append("MonthDay TEXT(4),");
		strsql.Append("JyoCD    TEXT(2),");
		strsql.Append("Kaiji    TEXT(2),");
		strsql.Append("Nichiji  TEXT(2),");
		strsql.Append("RaceNum  TEXT(2),");
		strsql.Append("Kumi     TEXT(2),");
		strsql.Append("Odds     TEXT(5),");
		strsql.Append("Ninki    TEXT(2),");
		strsql.Append("CONSTRAINT ODDS_WAKU PRIMARY KEY ([Year],MonthDay,JyoCD,Kaiji,Nichiji,RaceNum,Kumi));");
		
bstrQuery = strsql; pCn->Execute(bstrQuery,&vRecsAffected,adOptionUnspecified);
		
		
		//オッズ_馬連_ヘッダ
		strsql = "CREATE TABLE ODDS_UMAREN_HEAD (";
		strsql.Append("RecordSpec       TEXT(2),");
		strsql.Append("DataKubun        TEXT(1),");
		strsql.Append("MakeDate TEXT(8),");
		strsql.Append("[Year]     TEXT(4),");
		strsql.Append("MonthDay TEXT(4),");
		strsql.Append("JyoCD    TEXT(2),");
		strsql.Append("Kaiji    TEXT(2),");
		strsql.Append("Nichiji  TEXT(2),");
		strsql.Append("RaceNum  TEXT(2),");
		strsql.Append("HappyoTime       TEXT(8),");
		strsql.Append("TorokuTosu       TEXT(2),");
		strsql.Append("SyussoTosu       TEXT(2),");
		strsql.Append("UmarenFlag       TEXT(1),");
		strsql.Append("TotalHyosuUmaren TEXT(11),");
		strsql.Append("CONSTRAINT ODDS_UMAREN_HEAD PRIMARY KEY ([Year],MonthDay,JyoCD,Kaiji,Nichiji,RaceNum));");
		
bstrQuery = strsql; pCn->Execute(bstrQuery,&vRecsAffected,adOptionUnspecified);
		
		
		//オッズ_馬連
		strsql = "CREATE TABLE ODDS_UMAREN (";
		strsql.Append("MakeDate TEXT(8),");
		strsql.Append("[Year]     TEXT(4),");
		strsql.Append("MonthDay TEXT(4),");
		strsql.Append("JyoCD    TEXT(2),");
		strsql.Append("Kaiji    TEXT(2),");
		strsql.Append("Nichiji  TEXT(2),");
		strsql.Append("RaceNum  TEXT(2),");
		strsql.Append("Kumi     TEXT(4),");
		strsql.Append("Odds     TEXT(6),");
		strsql.Append("Ninki    TEXT(3),");
		strsql.Append("CONSTRAINT ODDS_UMAREN PRIMARY KEY ([Year],MonthDay,JyoCD,Kaiji,Nichiji,RaceNum,Kumi));");
		
bstrQuery = strsql; pCn->Execute(bstrQuery,&vRecsAffected,adOptionUnspecified);
		
		
		//オッズ_ワイド_ヘッダ
		strsql = "CREATE TABLE ODDS_WIDE_HEAD (";
		strsql.Append("RecordSpec       TEXT(2),");
		strsql.Append("DataKubun        TEXT(1),");
		strsql.Append("MakeDate TEXT(8),");
		strsql.Append("[Year]     TEXT(4),");
		strsql.Append("MonthDay TEXT(4),");
		strsql.Append("JyoCD    TEXT(2),");
		strsql.Append("Kaiji    TEXT(2),");
		strsql.Append("Nichiji  TEXT(2),");
		strsql.Append("RaceNum  TEXT(2),");
		strsql.Append("HappyoTime       TEXT(8),");
		strsql.Append("TorokuTosu       TEXT(2),");
		strsql.Append("SyussoTosu       TEXT(2),");
		strsql.Append("WideFlag TEXT(1),");
		strsql.Append("TotalHyosuWide   TEXT(11),");
		strsql.Append("CONSTRAINT ODDS_WIDE_HEAD PRIMARY KEY ([Year],MonthDay,JyoCD,Kaiji,Nichiji,RaceNum));");
		
bstrQuery = strsql; pCn->Execute(bstrQuery,&vRecsAffected,adOptionUnspecified);
		
		
		//オッズ_ワイド
		strsql = "CREATE TABLE ODDS_WIDE (";
		strsql.Append("MakeDate TEXT(8),");
		strsql.Append("[Year]     TEXT(4),");
		strsql.Append("MonthDay TEXT(4),");
		strsql.Append("JyoCD    TEXT(2),");
		strsql.Append("Kaiji    TEXT(2),");
		strsql.Append("Nichiji  TEXT(2),");
		strsql.Append("RaceNum  TEXT(2),");
		strsql.Append("Kumi     TEXT(4),");
		strsql.Append("OddsLow  TEXT(5),");
		strsql.Append("OddsHigh TEXT(5),");
		strsql.Append("Ninki    TEXT(3),");
		strsql.Append("CONSTRAINT ODDS_WIDE PRIMARY KEY ([Year],MonthDay,JyoCD,Kaiji,Nichiji,RaceNum,Kumi));");
		
bstrQuery = strsql; pCn->Execute(bstrQuery,&vRecsAffected,adOptionUnspecified);
		
		
		//オッズ_馬単_ヘッダ
		strsql = "CREATE TABLE ODDS_UMATAN_HEAD (";
		strsql.Append("RecordSpec       TEXT(2),");
		strsql.Append("DataKubun        TEXT(1),");
		strsql.Append("MakeDate TEXT(8),");
		strsql.Append("[Year]     TEXT(4),");
		strsql.Append("MonthDay TEXT(4),");
		strsql.Append("JyoCD    TEXT(2),");
		strsql.Append("Kaiji    TEXT(2),");
		strsql.Append("Nichiji  TEXT(2),");
		strsql.Append("RaceNum  TEXT(2),");
		strsql.Append("HappyoTime       TEXT(8),");
		strsql.Append("TorokuTosu       TEXT(2),");
		strsql.Append("SyussoTosu       TEXT(2),");
		strsql.Append("UmatanFlag       TEXT(1),");
		strsql.Append("TotalHyosuUmatan TEXT(11),");
		strsql.Append("CONSTRAINT ODDS_UMATAN_HEAD PRIMARY KEY ([Year],MonthDay,JyoCD,Kaiji,Nichiji,RaceNum));");
		
bstrQuery = strsql; pCn->Execute(bstrQuery,&vRecsAffected,adOptionUnspecified);
		
		
		//オッズ_馬単
		strsql = "CREATE TABLE ODDS_UMATAN (";
		strsql.Append("MakeDate TEXT(8),");
		strsql.Append("[Year]     TEXT(4),");
		strsql.Append("MonthDay TEXT(4),");
		strsql.Append("JyoCD    TEXT(2),");
		strsql.Append("Kaiji    TEXT(2),");
		strsql.Append("Nichiji  TEXT(2),");
		strsql.Append("RaceNum  TEXT(2),");
		strsql.Append("Kumi     TEXT(4),");
		strsql.Append("Odds     TEXT(6),");
		strsql.Append("Ninki    TEXT(3),");
		strsql.Append("CONSTRAINT ODDS_UMATAN PRIMARY KEY ([Year],MonthDay,JyoCD,Kaiji,Nichiji,RaceNum,Kumi));");
		
bstrQuery = strsql; pCn->Execute(bstrQuery,&vRecsAffected,adOptionUnspecified);
		
		
		//オッズ_3連複_ヘッダ
		strsql = "CREATE TABLE ODDS_SANREN_HEAD (";
		strsql.Append("RecordSpec       TEXT(2),");
		strsql.Append("DataKubun        TEXT(1),");
		strsql.Append("MakeDate TEXT(8),");
		strsql.Append("[Year]     TEXT(4),");
		strsql.Append("MonthDay TEXT(4),");
		strsql.Append("JyoCD    TEXT(2),");
		strsql.Append("Kaiji    TEXT(2),");
		strsql.Append("Nichiji  TEXT(2),");
		strsql.Append("RaceNum  TEXT(2),");
		strsql.Append("HappyoTime       TEXT(8),");
		strsql.Append("TorokuTosu       TEXT(2),");
		strsql.Append("SyussoTosu       TEXT(2),");
		strsql.Append("SanrenFlag       TEXT(1),");
		strsql.Append("TotalHyosuSanren TEXT(11),");
		strsql.Append("CONSTRAINT ODDS_SANREN_HEAD PRIMARY KEY ([Year],MonthDay,JyoCD,Kaiji,Nichiji,RaceNum));");
		
bstrQuery = strsql; pCn->Execute(bstrQuery,&vRecsAffected,adOptionUnspecified);
		
		
		//オッズ_3連複
		strsql = "CREATE TABLE ODDS_SANREN (";
		strsql.Append("MakeDate TEXT(8),");
		strsql.Append("[Year]     TEXT(4),");
		strsql.Append("MonthDay TEXT(4),");
		strsql.Append("JyoCD    TEXT(2),");
		strsql.Append("Kaiji    TEXT(2),");
		strsql.Append("Nichiji  TEXT(2),");
		strsql.Append("RaceNum  TEXT(2),");
		strsql.Append("Kumi     TEXT(6),");
		strsql.Append("Odds     TEXT(6),");
		strsql.Append("Ninki    TEXT(3),");
		strsql.Append("CONSTRAINT ODDS_SANREN PRIMARY KEY ([Year],MonthDay,JyoCD,Kaiji,Nichiji,RaceNum,Kumi));");
		
bstrQuery = strsql; pCn->Execute(bstrQuery,&vRecsAffected,adOptionUnspecified);
		
		
		//オッズ_3連単_ヘッダ
		strsql = "CREATE TABLE ODDS_SANRENTAN_HEAD (";
		strsql.Append("RecordSpec       TEXT(2),");
		strsql.Append("DataKubun        TEXT(1),");
		strsql.Append("MakeDate TEXT(8),");
		strsql.Append("[Year]     TEXT(4),");
		strsql.Append("MonthDay TEXT(4),");
		strsql.Append("JyoCD    TEXT(2),");
		strsql.Append("Kaiji    TEXT(2),");
		strsql.Append("Nichiji  TEXT(2),");
		strsql.Append("RaceNum  TEXT(2),");
		strsql.Append("HappyoTime       TEXT(8),");
		strsql.Append("TorokuTosu       TEXT(2),");
		strsql.Append("SyussoTosu       TEXT(2),");
		strsql.Append("SanrentanFlag       TEXT(1),");
		strsql.Append("TotalHyosuSanrentan TEXT(11),");
		strsql.Append("CONSTRAINT ODDS_SANRENTAN_HEAD PRIMARY KEY ([Year],MonthDay,JyoCD,Kaiji,Nichiji,RaceNum));");
		
bstrQuery = strsql; pCn->Execute(bstrQuery,&vRecsAffected,adOptionUnspecified);
		
		
		//オッズ_3連単
		strsql = "CREATE TABLE ODDS_SANRENTAN (";
		strsql.Append("MakeDate TEXT(8),");
		strsql.Append("[Year]     TEXT(4),");
		strsql.Append("MonthDay TEXT(4),");
		strsql.Append("JyoCD    TEXT(2),");
		strsql.Append("Kaiji    TEXT(2),");
		strsql.Append("Nichiji  TEXT(2),");
		strsql.Append("RaceNum  TEXT(2),");
		strsql.Append("Kumi     TEXT(6),");
		strsql.Append("Odds     TEXT(7),");
		strsql.Append("Ninki    TEXT(4),");
		strsql.Append("CONSTRAINT ODDS_SANRENTAN PRIMARY KEY ([Year],MonthDay,JyoCD,Kaiji,Nichiji,RaceNum,Kumi));");
		
bstrQuery = strsql; pCn->Execute(bstrQuery,&vRecsAffected,adOptionUnspecified);


		//競走馬マスタ
		strsql = "CREATE TABLE UMA (";
		strsql.Append("RecordSpec       TEXT(2),");
		strsql.Append("DataKubun        TEXT(1),");
		strsql.Append("MakeDate TEXT(8),");
		strsql.Append("KettoNum TEXT(10) CONSTRAINT UMA PRIMARY KEY UNIQUE,");
		strsql.Append("DelKubun TEXT(1),");
		strsql.Append("RegDate  TEXT(8),");
		strsql.Append("DelDate  TEXT(8),");
		strsql.Append("BirthDate        TEXT(8),");
		strsql.Append("Bamei    TEXT(36),");
		strsql.Append("BameiKana        TEXT(36),");
		strsql.Append("BameiEng TEXT(60),");
		strsql.Append("ZaikyuFlag TEXT(1),");
		strsql.Append("Reserved TEXT(19),");
		strsql.Append("UmaKigoCD        TEXT(2),");
		strsql.Append("SexCD    TEXT(1),");
		strsql.Append("HinsyuCD TEXT(1),");
		strsql.Append("KeiroCD  TEXT(2),");
		strsql.Append("Ketto3InfoHansyokuNum1   TEXT(10),");
		strsql.Append("Ketto3InfoBamei1 TEXT(36),");
		strsql.Append("Ketto3InfoHansyokuNum2   TEXT(10),");
		strsql.Append("Ketto3InfoBamei2 TEXT(36),");
		strsql.Append("Ketto3InfoHansyokuNum3   TEXT(10),");
		strsql.Append("Ketto3InfoBamei3 TEXT(36),");
		strsql.Append("Ketto3InfoHansyokuNum4   TEXT(10),");
		strsql.Append("Ketto3InfoBamei4 TEXT(36),");
		strsql.Append("Ketto3InfoHansyokuNum5   TEXT(10),");
		strsql.Append("Ketto3InfoBamei5 TEXT(36),");
		strsql.Append("Ketto3InfoHansyokuNum6   TEXT(10),");
		strsql.Append("Ketto3InfoBamei6 TEXT(36),");
		strsql.Append("Ketto3InfoHansyokuNum7   TEXT(10),");
		strsql.Append("Ketto3InfoBamei7 TEXT(36),");
		strsql.Append("Ketto3InfoHansyokuNum8   TEXT(10),");
		strsql.Append("Ketto3InfoBamei8 TEXT(36),");
		strsql.Append("Ketto3InfoHansyokuNum9   TEXT(10),");
		strsql.Append("Ketto3InfoBamei9 TEXT(36),");
		strsql.Append("Ketto3InfoHansyokuNum10  TEXT(10),");
		strsql.Append("Ketto3InfoBamei10        TEXT(36),");
		strsql.Append("Ketto3InfoHansyokuNum11  TEXT(10),");
		strsql.Append("Ketto3InfoBamei11        TEXT(36),");
		strsql.Append("Ketto3InfoHansyokuNum12  TEXT(10),");
		strsql.Append("Ketto3InfoBamei12        TEXT(36),");
		strsql.Append("Ketto3InfoHansyokuNum13  TEXT(10),");
		strsql.Append("Ketto3InfoBamei13        TEXT(36),");
		strsql.Append("Ketto3InfoHansyokuNum14  TEXT(10),");
		strsql.Append("Ketto3InfoBamei14        TEXT(36),");
		strsql.Append("TozaiCD  TEXT(1),");
		strsql.Append("ChokyosiCode     TEXT(5),");
		strsql.Append("ChokyosiRyakusyo TEXT(8),");
		strsql.Append("Syotai   TEXT(20),");
		strsql.Append("BreederCode      TEXT(8),");
		strsql.Append("BreederName      TEXT(72),");
		strsql.Append("SanchiName       TEXT(20),");
		strsql.Append("BanusiCode       TEXT(6),");
		strsql.Append("BanusiName       TEXT(64),");
		strsql.Append("RuikeiHonsyoHeiti        TEXT(9),");
		strsql.Append("RuikeiHonsyoSyogai       TEXT(9),");
		strsql.Append("RuikeiFukaHeichi TEXT(9),");
		strsql.Append("RuikeiFukaSyogai TEXT(9),");
		strsql.Append("RuikeiSyutokuHeichi      TEXT(9),");
		strsql.Append("RuikeiSyutokuSyogai      TEXT(9),");
		strsql.Append("SogoChakukaisu1  TEXT(3),");
		strsql.Append("SogoChakukaisu2  TEXT(3),");
		strsql.Append("SogoChakukaisu3  TEXT(3),");
		strsql.Append("SogoChakukaisu4  TEXT(3),");
		strsql.Append("SogoChakukaisu5  TEXT(3),");
		strsql.Append("SogoChakukaisu6  TEXT(3),");
		strsql.Append("ChuoChakukaisu1  TEXT(3),");
		strsql.Append("ChuoChakukaisu2  TEXT(3),");
		strsql.Append("ChuoChakukaisu3  TEXT(3),");
		strsql.Append("ChuoChakukaisu4  TEXT(3),");
		strsql.Append("ChuoChakukaisu5  TEXT(3),");
		strsql.Append("ChuoChakukaisu6  TEXT(3),");
		strsql.Append("Ba1Chakukaisu1   TEXT(3),");
		strsql.Append("Ba1Chakukaisu2   TEXT(3),");
		strsql.Append("Ba1Chakukaisu3   TEXT(3),");
		strsql.Append("Ba1Chakukaisu4   TEXT(3),");
		strsql.Append("Ba1Chakukaisu5   TEXT(3),");
		strsql.Append("Ba1Chakukaisu6   TEXT(3),");
		strsql.Append("Ba2Chakukaisu1   TEXT(3),");
		strsql.Append("Ba2Chakukaisu2   TEXT(3),");
		strsql.Append("Ba2Chakukaisu3   TEXT(3),");
		strsql.Append("Ba2Chakukaisu4   TEXT(3),");
		strsql.Append("Ba2Chakukaisu5   TEXT(3),");
		strsql.Append("Ba2Chakukaisu6   TEXT(3),");
		strsql.Append("Ba3Chakukaisu1   TEXT(3),");
		strsql.Append("Ba3Chakukaisu2   TEXT(3),");
		strsql.Append("Ba3Chakukaisu3   TEXT(3),");
		strsql.Append("Ba3Chakukaisu4   TEXT(3),");
		strsql.Append("Ba3Chakukaisu5   TEXT(3),");
		strsql.Append("Ba3Chakukaisu6   TEXT(3),");
		strsql.Append("Ba4Chakukaisu1   TEXT(3),");
		strsql.Append("Ba4Chakukaisu2   TEXT(3),");
		strsql.Append("Ba4Chakukaisu3   TEXT(3),");
		strsql.Append("Ba4Chakukaisu4   TEXT(3),");
		strsql.Append("Ba4Chakukaisu5   TEXT(3),");
		strsql.Append("Ba4Chakukaisu6   TEXT(3),");
		strsql.Append("Ba5Chakukaisu1   TEXT(3),");
		strsql.Append("Ba5Chakukaisu2   TEXT(3),");
		strsql.Append("Ba5Chakukaisu3   TEXT(3),");
		strsql.Append("Ba5Chakukaisu4   TEXT(3),");
		strsql.Append("Ba5Chakukaisu5   TEXT(3),");
		strsql.Append("Ba5Chakukaisu6   TEXT(3),");
		strsql.Append("Ba6Chakukaisu1   TEXT(3),");
		strsql.Append("Ba6Chakukaisu2   TEXT(3),");
		strsql.Append("Ba6Chakukaisu3   TEXT(3),");
		strsql.Append("Ba6Chakukaisu4   TEXT(3),");
		strsql.Append("Ba6Chakukaisu5   TEXT(3),");
		strsql.Append("Ba6Chakukaisu6   TEXT(3),");
		strsql.Append("Ba7Chakukaisu1   TEXT(3),");
		strsql.Append("Ba7Chakukaisu2   TEXT(3),");
		strsql.Append("Ba7Chakukaisu3   TEXT(3),");
		strsql.Append("Ba7Chakukaisu4   TEXT(3),");
		strsql.Append("Ba7Chakukaisu5   TEXT(3),");
		strsql.Append("Ba7Chakukaisu6   TEXT(3),");
		strsql.Append("Jyotai1Chakukaisu1       TEXT(3),");
		strsql.Append("Jyotai1Chakukaisu2       TEXT(3),");
		strsql.Append("Jyotai1Chakukaisu3       TEXT(3),");
		strsql.Append("Jyotai1Chakukaisu4       TEXT(3),");
		strsql.Append("Jyotai1Chakukaisu5       TEXT(3),");
		strsql.Append("Jyotai1Chakukaisu6       TEXT(3),");
		strsql.Append("Jyotai2Chakukaisu1       TEXT(3),");
		strsql.Append("Jyotai2Chakukaisu2       TEXT(3),");
		strsql.Append("Jyotai2Chakukaisu3       TEXT(3),");
		strsql.Append("Jyotai2Chakukaisu4       TEXT(3),");
		strsql.Append("Jyotai2Chakukaisu5       TEXT(3),");
		strsql.Append("Jyotai2Chakukaisu6       TEXT(3),");
		strsql.Append("Jyotai3Chakukaisu1       TEXT(3),");
		strsql.Append("Jyotai3Chakukaisu2       TEXT(3),");
		strsql.Append("Jyotai3Chakukaisu3       TEXT(3),");
		strsql.Append("Jyotai3Chakukaisu4       TEXT(3),");
		strsql.Append("Jyotai3Chakukaisu5       TEXT(3),");
		strsql.Append("Jyotai3Chakukaisu6       TEXT(3),");
		strsql.Append("Jyotai4Chakukaisu1       TEXT(3),");
		strsql.Append("Jyotai4Chakukaisu2       TEXT(3),");
		strsql.Append("Jyotai4Chakukaisu3       TEXT(3),");
		strsql.Append("Jyotai4Chakukaisu4       TEXT(3),");
		strsql.Append("Jyotai4Chakukaisu5       TEXT(3),");
		strsql.Append("Jyotai4Chakukaisu6       TEXT(3),");
		strsql.Append("Jyotai5Chakukaisu1       TEXT(3),");
		strsql.Append("Jyotai5Chakukaisu2       TEXT(3),");
		strsql.Append("Jyotai5Chakukaisu3       TEXT(3),");
		strsql.Append("Jyotai5Chakukaisu4       TEXT(3),");
		strsql.Append("Jyotai5Chakukaisu5       TEXT(3),");
		strsql.Append("Jyotai5Chakukaisu6       TEXT(3),");
		strsql.Append("Jyotai6Chakukaisu1       TEXT(3),");
		strsql.Append("Jyotai6Chakukaisu2       TEXT(3),");
		strsql.Append("Jyotai6Chakukaisu3       TEXT(3),");
		strsql.Append("Jyotai6Chakukaisu4       TEXT(3),");
		strsql.Append("Jyotai6Chakukaisu5       TEXT(3),");
		strsql.Append("Jyotai6Chakukaisu6       TEXT(3),");
		strsql.Append("Jyotai7Chakukaisu1       TEXT(3),");
		strsql.Append("Jyotai7Chakukaisu2       TEXT(3),");
		strsql.Append("Jyotai7Chakukaisu3       TEXT(3),");
		strsql.Append("Jyotai7Chakukaisu4       TEXT(3),");
		strsql.Append("Jyotai7Chakukaisu5       TEXT(3),");
		strsql.Append("Jyotai7Chakukaisu6       TEXT(3),");
		strsql.Append("Jyotai8Chakukaisu1       TEXT(3),");
		strsql.Append("Jyotai8Chakukaisu2       TEXT(3),");
		strsql.Append("Jyotai8Chakukaisu3       TEXT(3),");
		strsql.Append("Jyotai8Chakukaisu4       TEXT(3),");
		strsql.Append("Jyotai8Chakukaisu5       TEXT(3),");
		strsql.Append("Jyotai8Chakukaisu6       TEXT(3),");
		strsql.Append("Jyotai9Chakukaisu1       TEXT(3),");
		strsql.Append("Jyotai9Chakukaisu2       TEXT(3),");
		strsql.Append("Jyotai9Chakukaisu3       TEXT(3),");
		strsql.Append("Jyotai9Chakukaisu4       TEXT(3),");
		strsql.Append("Jyotai9Chakukaisu5       TEXT(3),");
		strsql.Append("Jyotai9Chakukaisu6       TEXT(3),");
		strsql.Append("Jyotai10Chakukaisu1      TEXT(3),");
		strsql.Append("Jyotai10Chakukaisu2      TEXT(3),");
		strsql.Append("Jyotai10Chakukaisu3      TEXT(3),");
		strsql.Append("Jyotai10Chakukaisu4      TEXT(3),");
		strsql.Append("Jyotai10Chakukaisu5      TEXT(3),");
		strsql.Append("Jyotai10Chakukaisu6      TEXT(3),");
		strsql.Append("Jyotai11Chakukaisu1      TEXT(3),");
		strsql.Append("Jyotai11Chakukaisu2      TEXT(3),");
		strsql.Append("Jyotai11Chakukaisu3      TEXT(3),");
		strsql.Append("Jyotai11Chakukaisu4      TEXT(3),");
		strsql.Append("Jyotai11Chakukaisu5      TEXT(3),");
		strsql.Append("Jyotai11Chakukaisu6      TEXT(3),");
		strsql.Append("Jyotai12Chakukaisu1      TEXT(3),");
		strsql.Append("Jyotai12Chakukaisu2      TEXT(3),");
		strsql.Append("Jyotai12Chakukaisu3      TEXT(3),");
		strsql.Append("Jyotai12Chakukaisu4      TEXT(3),");
		strsql.Append("Jyotai12Chakukaisu5      TEXT(3),");
		strsql.Append("Jyotai12Chakukaisu6      TEXT(3),");
		strsql.Append("Kyori1Chakukaisu1        TEXT(3),");
		strsql.Append("Kyori1Chakukaisu2        TEXT(3),");
		strsql.Append("Kyori1Chakukaisu3        TEXT(3),");
		strsql.Append("Kyori1Chakukaisu4        TEXT(3),");
		strsql.Append("Kyori1Chakukaisu5        TEXT(3),");
		strsql.Append("Kyori1Chakukaisu6        TEXT(3),");
		strsql.Append("Kyori2Chakukaisu1        TEXT(3),");
		strsql.Append("Kyori2Chakukaisu2        TEXT(3),");
		strsql.Append("Kyori2Chakukaisu3        TEXT(3),");
		strsql.Append("Kyori2Chakukaisu4        TEXT(3),");
		strsql.Append("Kyori2Chakukaisu5        TEXT(3),");
		strsql.Append("Kyori2Chakukaisu6        TEXT(3),");
		strsql.Append("Kyori3Chakukaisu1        TEXT(3),");
		strsql.Append("Kyori3Chakukaisu2        TEXT(3),");
		strsql.Append("Kyori3Chakukaisu3        TEXT(3),");
		strsql.Append("Kyori3Chakukaisu4        TEXT(3),");
		strsql.Append("Kyori3Chakukaisu5        TEXT(3),");
		strsql.Append("Kyori3Chakukaisu6        TEXT(3),");
		strsql.Append("Kyori4Chakukaisu1        TEXT(3),");
		strsql.Append("Kyori4Chakukaisu2        TEXT(3),");
		strsql.Append("Kyori4Chakukaisu3        TEXT(3),");
		strsql.Append("Kyori4Chakukaisu4        TEXT(3),");
		strsql.Append("Kyori4Chakukaisu5        TEXT(3),");
		strsql.Append("Kyori4Chakukaisu6        TEXT(3),");
		strsql.Append("Kyori5Chakukaisu1        TEXT(3),");
		strsql.Append("Kyori5Chakukaisu2        TEXT(3),");
		strsql.Append("Kyori5Chakukaisu3        TEXT(3),");
		strsql.Append("Kyori5Chakukaisu4        TEXT(3),");
		strsql.Append("Kyori5Chakukaisu5        TEXT(3),");
		strsql.Append("Kyori5Chakukaisu6        TEXT(3),");
		strsql.Append("Kyori6Chakukaisu1        TEXT(3),");
		strsql.Append("Kyori6Chakukaisu2        TEXT(3),");
		strsql.Append("Kyori6Chakukaisu3        TEXT(3),");
		strsql.Append("Kyori6Chakukaisu4        TEXT(3),");
		strsql.Append("Kyori6Chakukaisu5        TEXT(3),");
		strsql.Append("Kyori6Chakukaisu6        TEXT(3),");
		strsql.Append("Kyakusitu1       TEXT(3),");
		strsql.Append("Kyakusitu2       TEXT(3),");
		strsql.Append("Kyakusitu3       TEXT(3),");
		strsql.Append("Kyakusitu4       TEXT(3),");
		strsql.Append("RaceCount        TEXT(3));");
		
bstrQuery = strsql; pCn->Execute(bstrQuery,&vRecsAffected,adOptionUnspecified);
		
		
		//騎手マスタ
		strsql = "CREATE TABLE KISYU (";
		strsql.Append("RecordSpec       TEXT(2),");
		strsql.Append("DataKubun        TEXT(1),");
		strsql.Append("MakeDate TEXT(8),");
		strsql.Append("KisyuCode        TEXT(5) CONSTRAINT KISYU PRIMARY KEY UNIQUE,");
		strsql.Append("DelKubun TEXT(1),");
		strsql.Append("IssueDate        TEXT(8),");
		strsql.Append("DelDate  TEXT(8),");
		strsql.Append("BirthDate        TEXT(8),");
		strsql.Append("KisyuName        TEXT(34),");
		strsql.Append("reserved TEXT(34),");
		strsql.Append("KisyuNameKana    TEXT(30),");
		strsql.Append("KisyuRyakusyo    TEXT(8),");
		strsql.Append("KisyuNameEng     TEXT(80),");
		strsql.Append("SexCD    TEXT(1),");
		strsql.Append("SikakuCD TEXT(1),");
		strsql.Append("MinaraiCD        TEXT(1),");
		strsql.Append("TozaiCD  TEXT(1),");
		strsql.Append("Syotai   TEXT(20),");
		strsql.Append("ChokyosiCode     TEXT(5),");
		strsql.Append("ChokyosiRyakusyo TEXT(8),");
		strsql.Append("HatuKiJyo1Hatukijyoid    TEXT(16),");
		strsql.Append("HatuKiJyo1SyussoTosu     TEXT(2),");
		strsql.Append("HatuKiJyo1KettoNum       TEXT(10),");
		strsql.Append("HatuKiJyo1Bamei  TEXT(36),");
		strsql.Append("HatuKiJyo1KakuteiJyuni   TEXT(2),");
		strsql.Append("HatuKiJyo1IJyoCD TEXT(1),");
		strsql.Append("HatuKiJyo2Hatukijyoid    TEXT(16),");
		strsql.Append("HatuKiJyo2SyussoTosu     TEXT(2),");
		strsql.Append("HatuKiJyo2KettoNum       TEXT(10),");
		strsql.Append("HatuKiJyo2Bamei  TEXT(36),");
		strsql.Append("HatuKiJyo2KakuteiJyuni   TEXT(2),");
		strsql.Append("HatuKiJyo2IJyoCD TEXT(1),");
		strsql.Append("HatuSyori1Hatusyoriid    TEXT(16),");
		strsql.Append("HatuSyori1SyussoTosu     TEXT(2),");
		strsql.Append("HatuSyori1KettoNum       TEXT(10),");
		strsql.Append("HatuSyori1Bamei  TEXT(36),");
		strsql.Append("HatuSyori2Hatusyoriid    TEXT(16),");
		strsql.Append("HatuSyori2SyussoTosu     TEXT(2),");
		strsql.Append("HatuSyori2KettoNum       TEXT(10),");
		strsql.Append("HatuSyori2Bamei  TEXT(36),");
		strsql.Append("SaikinJyusyo1SaikinJyusyoid      TEXT(16),");
		strsql.Append("SaikinJyusyo1Hondai      TEXT(60),");
		strsql.Append("SaikinJyusyo1Ryakusyo10  TEXT(20),");
		strsql.Append("SaikinJyusyo1Ryakusyo6   TEXT(12),");
		strsql.Append("SaikinJyusyo1Ryakusyo3   TEXT(6),");
		strsql.Append("SaikinJyusyo1GradeCD     TEXT(1),");
		strsql.Append("SaikinJyusyo1SyussoTosu  TEXT(2),");
		strsql.Append("SaikinJyusyo1KettoNum    TEXT(10),");
		strsql.Append("SaikinJyusyo1Bamei       TEXT(36),");
		strsql.Append("SaikinJyusyo2SaikinJyusyoid      TEXT(16),");
		strsql.Append("SaikinJyusyo2Hondai      TEXT(60),");
		strsql.Append("SaikinJyusyo2Ryakusyo10  TEXT(20),");
		strsql.Append("SaikinJyusyo2Ryakusyo6   TEXT(12),");
		strsql.Append("SaikinJyusyo2Ryakusyo3   TEXT(6),");
		strsql.Append("SaikinJyusyo2GradeCD     TEXT(1),");
		strsql.Append("SaikinJyusyo2SyussoTosu  TEXT(2),");
		strsql.Append("SaikinJyusyo2KettoNum    TEXT(10),");
		strsql.Append("SaikinJyusyo2Bamei       TEXT(36),");
		strsql.Append("SaikinJyusyo3SaikinJyusyoid      TEXT(16),");
		strsql.Append("SaikinJyusyo3Hondai      TEXT(60),");
		strsql.Append("SaikinJyusyo3Ryakusyo10  TEXT(20),");
		strsql.Append("SaikinJyusyo3Ryakusyo6   TEXT(12),");
		strsql.Append("SaikinJyusyo3Ryakusyo3   TEXT(6),");
		strsql.Append("SaikinJyusyo3GradeCD     TEXT(1),");
		strsql.Append("SaikinJyusyo3SyussoTosu  TEXT(2),");
		strsql.Append("SaikinJyusyo3KettoNum    TEXT(10),");
		strsql.Append("SaikinJyusyo3Bamei       TEXT(36));");
		
bstrQuery = strsql; pCn->Execute(bstrQuery,&vRecsAffected,adOptionUnspecified);
		
		
		//騎手マスタ_成績
		strsql = "CREATE TABLE KISYU_SEISEKI (";
		strsql.Append("MakeDate TEXT(8),");
		strsql.Append("KisyuCode        TEXT(5),");
		strsql.Append("Num      TEXT(1),");
		strsql.Append("SetYear  TEXT(4),");
		strsql.Append("HonSyokinHeichi  TEXT(10),");
		strsql.Append("HonSyokinSyogai  TEXT(10),");
		strsql.Append("FukaSyokinHeichi TEXT(10),");
		strsql.Append("FukaSyokinSyogai TEXT(10),");
		strsql.Append("HeichiChakukaisu1        TEXT(6),");
		strsql.Append("HeichiChakukaisu2        TEXT(6),");
		strsql.Append("HeichiChakukaisu3        TEXT(6),");
		strsql.Append("HeichiChakukaisu4        TEXT(6),");
		strsql.Append("HeichiChakukaisu5        TEXT(6),");
		strsql.Append("HeichiChakukaisu6        TEXT(6),");
		strsql.Append("SyogaiChakukaisu1        TEXT(6),");
		strsql.Append("SyogaiChakukaisu2        TEXT(6),");
		strsql.Append("SyogaiChakukaisu3        TEXT(6),");
		strsql.Append("SyogaiChakukaisu4        TEXT(6),");
		strsql.Append("SyogaiChakukaisu5        TEXT(6),");
		strsql.Append("SyogaiChakukaisu6        TEXT(6),");
		strsql.Append("Jyo1Chakukaisu1  TEXT(6),");
		strsql.Append("Jyo1Chakukaisu2  TEXT(6),");
		strsql.Append("Jyo1Chakukaisu3  TEXT(6),");
		strsql.Append("Jyo1Chakukaisu4  TEXT(6),");
		strsql.Append("Jyo1Chakukaisu5  TEXT(6),");
		strsql.Append("Jyo1Chakukaisu6  TEXT(6),");
		strsql.Append("Jyo2Chakukaisu1  TEXT(6),");
		strsql.Append("Jyo2Chakukaisu2  TEXT(6),");
		strsql.Append("Jyo2Chakukaisu3  TEXT(6),");
		strsql.Append("Jyo2Chakukaisu4  TEXT(6),");
		strsql.Append("Jyo2Chakukaisu5  TEXT(6),");
		strsql.Append("Jyo2Chakukaisu6  TEXT(6),");
		strsql.Append("Jyo3Chakukaisu1  TEXT(6),");
		strsql.Append("Jyo3Chakukaisu2  TEXT(6),");
		strsql.Append("Jyo3Chakukaisu3  TEXT(6),");
		strsql.Append("Jyo3Chakukaisu4  TEXT(6),");
		strsql.Append("Jyo3Chakukaisu5  TEXT(6),");
		strsql.Append("Jyo3Chakukaisu6  TEXT(6),");
		strsql.Append("Jyo4Chakukaisu1  TEXT(6),");
		strsql.Append("Jyo4Chakukaisu2  TEXT(6),");
		strsql.Append("Jyo4Chakukaisu3  TEXT(6),");
		strsql.Append("Jyo4Chakukaisu4  TEXT(6),");
		strsql.Append("Jyo4Chakukaisu5  TEXT(6),");
		strsql.Append("Jyo4Chakukaisu6  TEXT(6),");
		strsql.Append("Jyo5Chakukaisu1  TEXT(6),");
		strsql.Append("Jyo5Chakukaisu2  TEXT(6),");
		strsql.Append("Jyo5Chakukaisu3  TEXT(6),");
		strsql.Append("Jyo5Chakukaisu4  TEXT(6),");
		strsql.Append("Jyo5Chakukaisu5  TEXT(6),");
		strsql.Append("Jyo5Chakukaisu6  TEXT(6),");
		strsql.Append("Jyo6Chakukaisu1  TEXT(6),");
		strsql.Append("Jyo6Chakukaisu2  TEXT(6),");
		strsql.Append("Jyo6Chakukaisu3  TEXT(6),");
		strsql.Append("Jyo6Chakukaisu4  TEXT(6),");
		strsql.Append("Jyo6Chakukaisu5  TEXT(6),");
		strsql.Append("Jyo6Chakukaisu6  TEXT(6),");
		strsql.Append("Jyo7Chakukaisu1  TEXT(6),");
		strsql.Append("Jyo7Chakukaisu2  TEXT(6),");
		strsql.Append("Jyo7Chakukaisu3  TEXT(6),");
		strsql.Append("Jyo7Chakukaisu4  TEXT(6),");
		strsql.Append("Jyo7Chakukaisu5  TEXT(6),");
		strsql.Append("Jyo7Chakukaisu6  TEXT(6),");
		strsql.Append("Jyo8Chakukaisu1  TEXT(6),");
		strsql.Append("Jyo8Chakukaisu2  TEXT(6),");
		strsql.Append("Jyo8Chakukaisu3  TEXT(6),");
		strsql.Append("Jyo8Chakukaisu4  TEXT(6),");
		strsql.Append("Jyo8Chakukaisu5  TEXT(6),");
		strsql.Append("Jyo8Chakukaisu6  TEXT(6),");
		strsql.Append("Jyo9Chakukaisu1  TEXT(6),");
		strsql.Append("Jyo9Chakukaisu2  TEXT(6),");
		strsql.Append("Jyo9Chakukaisu3  TEXT(6),");
		strsql.Append("Jyo9Chakukaisu4  TEXT(6),");
		strsql.Append("Jyo9Chakukaisu5  TEXT(6),");
		strsql.Append("Jyo9Chakukaisu6  TEXT(6),");
		strsql.Append("Jyo10Chakukaisu1 TEXT(6),");
		strsql.Append("Jyo10Chakukaisu2 TEXT(6),");
		strsql.Append("Jyo10Chakukaisu3 TEXT(6),");
		strsql.Append("Jyo10Chakukaisu4 TEXT(6),");
		strsql.Append("Jyo10Chakukaisu5 TEXT(6),");
		strsql.Append("Jyo10Chakukaisu6 TEXT(6),");
		strsql.Append("Jyo11Chakukaisu1 TEXT(6),");
		strsql.Append("Jyo11Chakukaisu2 TEXT(6),");
		strsql.Append("Jyo11Chakukaisu3 TEXT(6),");
		strsql.Append("Jyo11Chakukaisu4 TEXT(6),");
		strsql.Append("Jyo11Chakukaisu5 TEXT(6),");
		strsql.Append("Jyo11Chakukaisu6 TEXT(6),");
		strsql.Append("Jyo12Chakukaisu1 TEXT(6),");
		strsql.Append("Jyo12Chakukaisu2 TEXT(6),");
		strsql.Append("Jyo12Chakukaisu3 TEXT(6),");
		strsql.Append("Jyo12Chakukaisu4 TEXT(6),");
		strsql.Append("Jyo12Chakukaisu5 TEXT(6),");
		strsql.Append("Jyo12Chakukaisu6 TEXT(6),");
		strsql.Append("Jyo13Chakukaisu1 TEXT(6),");
		strsql.Append("Jyo13Chakukaisu2 TEXT(6),");
		strsql.Append("Jyo13Chakukaisu3 TEXT(6),");
		strsql.Append("Jyo13Chakukaisu4 TEXT(6),");
		strsql.Append("Jyo13Chakukaisu5 TEXT(6),");
		strsql.Append("Jyo13Chakukaisu6 TEXT(6),");
		strsql.Append("Jyo14Chakukaisu1 TEXT(6),");
		strsql.Append("Jyo14Chakukaisu2 TEXT(6),");
		strsql.Append("Jyo14Chakukaisu3 TEXT(6),");
		strsql.Append("Jyo14Chakukaisu4 TEXT(6),");
		strsql.Append("Jyo14Chakukaisu5 TEXT(6),");
		strsql.Append("Jyo14Chakukaisu6 TEXT(6),");
		strsql.Append("Jyo15Chakukaisu1 TEXT(6),");
		strsql.Append("Jyo15Chakukaisu2 TEXT(6),");
		strsql.Append("Jyo15Chakukaisu3 TEXT(6),");
		strsql.Append("Jyo15Chakukaisu4 TEXT(6),");
		strsql.Append("Jyo15Chakukaisu5 TEXT(6),");
		strsql.Append("Jyo15Chakukaisu6 TEXT(6),");
		strsql.Append("Jyo16Chakukaisu1 TEXT(6),");
		strsql.Append("Jyo16Chakukaisu2 TEXT(6),");
		strsql.Append("Jyo16Chakukaisu3 TEXT(6),");
		strsql.Append("Jyo16Chakukaisu4 TEXT(6),");
		strsql.Append("Jyo16Chakukaisu5 TEXT(6),");
		strsql.Append("Jyo16Chakukaisu6 TEXT(6),");
		strsql.Append("Jyo17Chakukaisu1 TEXT(6),");
		strsql.Append("Jyo17Chakukaisu2 TEXT(6),");
		strsql.Append("Jyo17Chakukaisu3 TEXT(6),");
		strsql.Append("Jyo17Chakukaisu4 TEXT(6),");
		strsql.Append("Jyo17Chakukaisu5 TEXT(6),");
		strsql.Append("Jyo17Chakukaisu6 TEXT(6),");
		strsql.Append("Jyo18Chakukaisu1 TEXT(6),");
		strsql.Append("Jyo18Chakukaisu2 TEXT(6),");
		strsql.Append("Jyo18Chakukaisu3 TEXT(6),");
		strsql.Append("Jyo18Chakukaisu4 TEXT(6),");
		strsql.Append("Jyo18Chakukaisu5 TEXT(6),");
		strsql.Append("Jyo18Chakukaisu6 TEXT(6),");
		strsql.Append("Jyo19Chakukaisu1 TEXT(6),");
		strsql.Append("Jyo19Chakukaisu2 TEXT(6),");
		strsql.Append("Jyo19Chakukaisu3 TEXT(6),");
		strsql.Append("Jyo19Chakukaisu4 TEXT(6),");
		strsql.Append("Jyo19Chakukaisu5 TEXT(6),");
		strsql.Append("Jyo19Chakukaisu6 TEXT(6),");
		strsql.Append("Jyo20Chakukaisu1 TEXT(6),");
		strsql.Append("Jyo20Chakukaisu2 TEXT(6),");
		strsql.Append("Jyo20Chakukaisu3 TEXT(6),");
		strsql.Append("Jyo20Chakukaisu4 TEXT(6),");
		strsql.Append("Jyo20Chakukaisu5 TEXT(6),");
		strsql.Append("Jyo20Chakukaisu6 TEXT(6),");
		strsql.Append("Kyori1Chakukaisu1        TEXT(6),");
		strsql.Append("Kyori1Chakukaisu2        TEXT(6),");
		strsql.Append("Kyori1Chakukaisu3        TEXT(6),");
		strsql.Append("Kyori1Chakukaisu4        TEXT(6),");
		strsql.Append("Kyori1Chakukaisu5        TEXT(6),");
		strsql.Append("Kyori1Chakukaisu6        TEXT(6),");
		strsql.Append("Kyori2Chakukaisu1        TEXT(6),");
		strsql.Append("Kyori2Chakukaisu2        TEXT(6),");
		strsql.Append("Kyori2Chakukaisu3        TEXT(6),");
		strsql.Append("Kyori2Chakukaisu4        TEXT(6),");
		strsql.Append("Kyori2Chakukaisu5        TEXT(6),");
		strsql.Append("Kyori2Chakukaisu6        TEXT(6),");
		strsql.Append("Kyori3Chakukaisu1        TEXT(6),");
		strsql.Append("Kyori3Chakukaisu2        TEXT(6),");
		strsql.Append("Kyori3Chakukaisu3        TEXT(6),");
		strsql.Append("Kyori3Chakukaisu4        TEXT(6),");
		strsql.Append("Kyori3Chakukaisu5        TEXT(6),");
		strsql.Append("Kyori3Chakukaisu6        TEXT(6),");
		strsql.Append("Kyori4Chakukaisu1        TEXT(6),");
		strsql.Append("Kyori4Chakukaisu2        TEXT(6),");
		strsql.Append("Kyori4Chakukaisu3        TEXT(6),");
		strsql.Append("Kyori4Chakukaisu4        TEXT(6),");
		strsql.Append("Kyori4Chakukaisu5        TEXT(6),");
		strsql.Append("Kyori4Chakukaisu6        TEXT(6),");
		strsql.Append("Kyori5Chakukaisu1        TEXT(6),");
		strsql.Append("Kyori5Chakukaisu2        TEXT(6),");
		strsql.Append("Kyori5Chakukaisu3        TEXT(6),");
		strsql.Append("Kyori5Chakukaisu4        TEXT(6),");
		strsql.Append("Kyori5Chakukaisu5        TEXT(6),");
		strsql.Append("Kyori5Chakukaisu6        TEXT(6),");
		strsql.Append("Kyori6Chakukaisu1        TEXT(6),");
		strsql.Append("Kyori6Chakukaisu2        TEXT(6),");
		strsql.Append("Kyori6Chakukaisu3        TEXT(6),");
		strsql.Append("Kyori6Chakukaisu4        TEXT(6),");
		strsql.Append("Kyori6Chakukaisu5        TEXT(6),");
		strsql.Append("Kyori6Chakukaisu6        TEXT(6),");
		strsql.Append("CONSTRAINT KISYU_SEISEKI PRIMARY KEY (KisyuCode,Num));");
		
bstrQuery = strsql; pCn->Execute(bstrQuery,&vRecsAffected,adOptionUnspecified);
		
		
		//調教師マスタ
		strsql = "CREATE TABLE CHOKYO (";
		strsql.Append("RecordSpec  TEXT (2),");
		strsql.Append("DataKubun   TEXT (1),");
		strsql.Append("MakeDate    TEXT (8),");
		strsql.Append("ChokyosiCode TEXT (5) CONSTRAINT CHOKYO PRIMARY KEY UNIQUE,");
		strsql.Append("DelKubun    TEXT (1),");
		strsql.Append("IssueDate   TEXT (8),");
		strsql.Append("DelDate     TEXT (8),");
		strsql.Append("BirthDate   TEXT (8),");
		strsql.Append("ChokyosiName TEXT (34),");
		strsql.Append("ChokyosiNameKana TEXT (30),");
		strsql.Append("ChokyosiRyakusyo TEXT (8),");
		strsql.Append("ChokyosiNameEng  TEXT (80),");
		strsql.Append("SexCD            TEXT (1),");
		strsql.Append("TozaiCD          TEXT (1),");
		strsql.Append("Syotai           TEXT (20),");
		strsql.Append("SaikinJyusyo1SaikinJyusyoid TEXT (16),");
		strsql.Append("SaikinJyusyo1Hondai         TEXT (60),");
		strsql.Append("SaikinJyusyo1Ryakusyo10     TEXT (20),");
		strsql.Append("SaikinJyusyo1Ryakusyo6      TEXT (12),");
		strsql.Append("SaikinJyusyo1Ryakusyo3      TEXT (6),");
		strsql.Append("SaikinJyusyo1GradeCD        TEXT (1),");
		strsql.Append("SaikinJyusyo1SyussoTosu     TEXT (2),");
		strsql.Append("SaikinJyusyo1KettoNum       TEXT (10),");
		strsql.Append("SaikinJyusyo1Bamei          TEXT (36),");
		strsql.Append("SaikinJyusyo2SaikinJyusyoid TEXT (16),");
		strsql.Append("SaikinJyusyo2Hondai         TEXT (60),");
		strsql.Append("SaikinJyusyo2Ryakusyo10     TEXT (20),");
		strsql.Append("SaikinJyusyo2Ryakusyo6      TEXT (12),");
		strsql.Append("SaikinJyusyo2Ryakusyo3      TEXT (6),");
		strsql.Append("SaikinJyusyo2GradeCD        TEXT (1),");
		strsql.Append("SaikinJyusyo2SyussoTosu     TEXT (2),");
		strsql.Append("SaikinJyusyo2KettoNum       TEXT (10),");
		strsql.Append("SaikinJyusyo2Bamei          TEXT (36),");
		strsql.Append("SaikinJyusyo3SaikinJyusyoid TEXT (16),");
		strsql.Append("SaikinJyusyo3Hondai         TEXT (60),");
		strsql.Append("SaikinJyusyo3Ryakusyo10     TEXT (20),");
		strsql.Append("SaikinJyusyo3Ryakusyo6      TEXT (12),");
		strsql.Append("SaikinJyusyo3Ryakusyo3      TEXT (6),");
		strsql.Append("SaikinJyusyo3GradeCD        TEXT (1),");
		strsql.Append("SaikinJyusyo3SyussoTosu     TEXT (2),");
		strsql.Append("SaikinJyusyo3KettoNum       TEXT (10),");
		strsql.Append("SaikinJyusyo3Bamei          TEXT (36));");
		
bstrQuery = strsql; pCn->Execute(bstrQuery,&vRecsAffected,adOptionUnspecified);
		
		
		//調教師マスタ_成績
		strsql = "CREATE TABLE CHOKYO_SEISEKI (";
		strsql.Append("MakeDate     TEXT (8),");
		strsql.Append("ChokyosiCode TEXT (5),");
		strsql.Append("Num          TEXT (1),");
		strsql.Append("SetYear      TEXT (4),");
		strsql.Append("HonSyokinHeichi      TEXT (10),");
		strsql.Append("HonSyokinSyogai      TEXT (10),");
		strsql.Append("FukaSyokinHeichi     TEXT (10),");
		strsql.Append("FukaSyokinSyogai     TEXT (10),");
		strsql.Append("HeichiChakukaisu1    TEXT (6),");
		strsql.Append("HeichiChakukaisu2    TEXT (6),");
		strsql.Append("HeichiChakukaisu3    TEXT (6),");
		strsql.Append("HeichiChakukaisu4    TEXT (6),");
		strsql.Append("HeichiChakukaisu5    TEXT (6),");
		strsql.Append("HeichiChakukaisu6    TEXT (6),");
		strsql.Append("SyogaiChakukaisu1    TEXT (6),");
		strsql.Append("SyogaiChakukaisu2    TEXT (6),");
		strsql.Append("SyogaiChakukaisu3    TEXT (6),");
		strsql.Append("SyogaiChakukaisu4    TEXT (6),");
		strsql.Append("SyogaiChakukaisu5    TEXT (6),");
		strsql.Append("SyogaiChakukaisu6    TEXT (6),");
		strsql.Append("Jyo1Chakukaisu1      TEXT (6),");
		strsql.Append("Jyo1Chakukaisu2      TEXT (6),");
		strsql.Append("Jyo1Chakukaisu3      TEXT (6),");
		strsql.Append("Jyo1Chakukaisu4      TEXT (6),");
		strsql.Append("Jyo1Chakukaisu5      TEXT (6),");
		strsql.Append("Jyo1Chakukaisu6      TEXT (6),");
		strsql.Append("Jyo2Chakukaisu1      TEXT (6),");
		strsql.Append("Jyo2Chakukaisu2      TEXT (6),");
		strsql.Append("Jyo2Chakukaisu3      TEXT (6),");
		strsql.Append("Jyo2Chakukaisu4      TEXT (6),");
		strsql.Append("Jyo2Chakukaisu5      TEXT (6),");
		strsql.Append("Jyo2Chakukaisu6      TEXT (6),");
		strsql.Append("Jyo3Chakukaisu1      TEXT (6),");
		strsql.Append("Jyo3Chakukaisu2      TEXT (6),");
		strsql.Append("Jyo3Chakukaisu3      TEXT (6),");
		strsql.Append("Jyo3Chakukaisu4      TEXT (6),");
		strsql.Append("Jyo3Chakukaisu5      TEXT (6),");
		strsql.Append("Jyo3Chakukaisu6      TEXT (6),");
		strsql.Append("Jyo4Chakukaisu1      TEXT (6),");
		strsql.Append("Jyo4Chakukaisu2      TEXT (6),");
		strsql.Append("Jyo4Chakukaisu3      TEXT (6),");
		strsql.Append("Jyo4Chakukaisu4      TEXT (6),");
		strsql.Append("Jyo4Chakukaisu5      TEXT (6),");
		strsql.Append("Jyo4Chakukaisu6      TEXT (6),");
		strsql.Append("Jyo5Chakukaisu1      TEXT (6),");
		strsql.Append("Jyo5Chakukaisu2      TEXT (6),");
		strsql.Append("Jyo5Chakukaisu3      TEXT (6),");
		strsql.Append("Jyo5Chakukaisu4      TEXT (6),");
		strsql.Append("Jyo5Chakukaisu5      TEXT (6),");
		strsql.Append("Jyo5Chakukaisu6      TEXT (6),");
		strsql.Append("Jyo6Chakukaisu1      TEXT (6),");
		strsql.Append("Jyo6Chakukaisu2      TEXT (6),");
		strsql.Append("Jyo6Chakukaisu3      TEXT (6),");
		strsql.Append("Jyo6Chakukaisu4      TEXT (6),");
		strsql.Append("Jyo6Chakukaisu5      TEXT (6),");
		strsql.Append("Jyo6Chakukaisu6      TEXT (6),");
		strsql.Append("Jyo7Chakukaisu1      TEXT (6),");
		strsql.Append("Jyo7Chakukaisu2      TEXT (6),");
		strsql.Append("Jyo7Chakukaisu3      TEXT (6),");
		strsql.Append("Jyo7Chakukaisu4      TEXT (6),");
		strsql.Append("Jyo7Chakukaisu5      TEXT (6),");
		strsql.Append("Jyo7Chakukaisu6      TEXT (6),");
		strsql.Append("Jyo8Chakukaisu1      TEXT (6),");
		strsql.Append("Jyo8Chakukaisu2      TEXT (6),");
		strsql.Append("Jyo8Chakukaisu3      TEXT (6),");
		strsql.Append("Jyo8Chakukaisu4      TEXT (6),");
		strsql.Append("Jyo8Chakukaisu5      TEXT (6),");
		strsql.Append("Jyo8Chakukaisu6      TEXT (6),");
		strsql.Append("Jyo9Chakukaisu1      TEXT (6),");
		strsql.Append("Jyo9Chakukaisu2      TEXT (6),");
		strsql.Append("Jyo9Chakukaisu3      TEXT (6),");
		strsql.Append("Jyo9Chakukaisu4      TEXT (6),");
		strsql.Append("Jyo9Chakukaisu5      TEXT (6),");
		strsql.Append("Jyo9Chakukaisu6      TEXT (6),");
		strsql.Append("Jyo10Chakukaisu1      TEXT (6),");
		strsql.Append("Jyo10Chakukaisu2      TEXT (6),");
		strsql.Append("Jyo10Chakukaisu3      TEXT (6),");
		strsql.Append("Jyo10Chakukaisu4      TEXT (6),");
		strsql.Append("Jyo10Chakukaisu5      TEXT (6),");
		strsql.Append("Jyo10Chakukaisu6      TEXT (6),");
		strsql.Append("Jyo11Chakukaisu1      TEXT (6),");
		strsql.Append("Jyo11Chakukaisu2      TEXT (6),");
		strsql.Append("Jyo11Chakukaisu3      TEXT (6),");
		strsql.Append("Jyo11Chakukaisu4      TEXT (6),");
		strsql.Append("Jyo11Chakukaisu5      TEXT (6),");
		strsql.Append("Jyo11Chakukaisu6      TEXT (6),");
		strsql.Append("Jyo12Chakukaisu1      TEXT (6),");
		strsql.Append("Jyo12Chakukaisu2      TEXT (6),");
		strsql.Append("Jyo12Chakukaisu3      TEXT (6),");
		strsql.Append("Jyo12Chakukaisu4      TEXT (6),");
		strsql.Append("Jyo12Chakukaisu5      TEXT (6),");
		strsql.Append("Jyo12Chakukaisu6      TEXT (6),");
		strsql.Append("Jyo13Chakukaisu1      TEXT (6),");
		strsql.Append("Jyo13Chakukaisu2      TEXT (6),");
		strsql.Append("Jyo13Chakukaisu3      TEXT (6),");
		strsql.Append("Jyo13Chakukaisu4      TEXT (6),");
		strsql.Append("Jyo13Chakukaisu5      TEXT (6),");
		strsql.Append("Jyo13Chakukaisu6      TEXT (6),");
		strsql.Append("Jyo14Chakukaisu1      TEXT (6),");
		strsql.Append("Jyo14Chakukaisu2      TEXT (6),");
		strsql.Append("Jyo14Chakukaisu3      TEXT (6),");
		strsql.Append("Jyo14Chakukaisu4      TEXT (6),");
		strsql.Append("Jyo14Chakukaisu5      TEXT (6),");
		strsql.Append("Jyo14Chakukaisu6      TEXT (6),");
		strsql.Append("Jyo15Chakukaisu1      TEXT (6),");
		strsql.Append("Jyo15Chakukaisu2      TEXT (6),");
		strsql.Append("Jyo15Chakukaisu3      TEXT (6),");
		strsql.Append("Jyo15Chakukaisu4      TEXT (6),");
		strsql.Append("Jyo15Chakukaisu5      TEXT (6),");
		strsql.Append("Jyo15Chakukaisu6      TEXT (6),");
		strsql.Append("Jyo16Chakukaisu1      TEXT (6),");
		strsql.Append("Jyo16Chakukaisu2      TEXT (6),");
		strsql.Append("Jyo16Chakukaisu3      TEXT (6),");
		strsql.Append("Jyo16Chakukaisu4      TEXT (6),");
		strsql.Append("Jyo16Chakukaisu5      TEXT (6),");
		strsql.Append("Jyo16Chakukaisu6      TEXT (6),");
		strsql.Append("Jyo17Chakukaisu1      TEXT (6),");
		strsql.Append("Jyo17Chakukaisu2      TEXT (6),");
		strsql.Append("Jyo17Chakukaisu3      TEXT (6),");
		strsql.Append("Jyo17Chakukaisu4      TEXT (6),");
		strsql.Append("Jyo17Chakukaisu5      TEXT (6),");
		strsql.Append("Jyo17Chakukaisu6      TEXT (6),");
		strsql.Append("Jyo18Chakukaisu1      TEXT (6),");
		strsql.Append("Jyo18Chakukaisu2      TEXT (6),");
		strsql.Append("Jyo18Chakukaisu3      TEXT (6),");
		strsql.Append("Jyo18Chakukaisu4      TEXT (6),");
		strsql.Append("Jyo18Chakukaisu5      TEXT (6),");
		strsql.Append("Jyo18Chakukaisu6      TEXT (6),");
		strsql.Append("Jyo19Chakukaisu1      TEXT (6),");
		strsql.Append("Jyo19Chakukaisu2      TEXT (6),");
		strsql.Append("Jyo19Chakukaisu3      TEXT (6),");
		strsql.Append("Jyo19Chakukaisu4      TEXT (6),");
		strsql.Append("Jyo19Chakukaisu5      TEXT (6),");
		strsql.Append("Jyo19Chakukaisu6      TEXT (6),");
		strsql.Append("Jyo20Chakukaisu1      TEXT (6),");
		strsql.Append("Jyo20Chakukaisu2      TEXT (6),");
		strsql.Append("Jyo20Chakukaisu3      TEXT (6),");
		strsql.Append("Jyo20Chakukaisu4      TEXT (6),");
		strsql.Append("Jyo20Chakukaisu5      TEXT (6),");
		strsql.Append("Jyo20Chakukaisu6      TEXT (6),");
		strsql.Append("Kyori1Chakukaisu1      TEXT (6),");
		strsql.Append("Kyori1Chakukaisu2      TEXT (6),");
		strsql.Append("Kyori1Chakukaisu3      TEXT (6),");
		strsql.Append("Kyori1Chakukaisu4      TEXT (6),");
		strsql.Append("Kyori1Chakukaisu5      TEXT (6),");
		strsql.Append("Kyori1Chakukaisu6      TEXT (6),");
		strsql.Append("Kyori2Chakukaisu1      TEXT (6),");
		strsql.Append("Kyori2Chakukaisu2      TEXT (6),");
		strsql.Append("Kyori2Chakukaisu3      TEXT (6),");
		strsql.Append("Kyori2Chakukaisu4      TEXT (6),");
		strsql.Append("Kyori2Chakukaisu5      TEXT (6),");
		strsql.Append("Kyori2Chakukaisu6      TEXT (6),");
		strsql.Append("Kyori3Chakukaisu1      TEXT (6),");
		strsql.Append("Kyori3Chakukaisu2      TEXT (6),");
		strsql.Append("Kyori3Chakukaisu3      TEXT (6),");
		strsql.Append("Kyori3Chakukaisu4      TEXT (6),");
		strsql.Append("Kyori3Chakukaisu5      TEXT (6),");
		strsql.Append("Kyori3Chakukaisu6      TEXT (6),");
		strsql.Append("Kyori4Chakukaisu1      TEXT (6),");
		strsql.Append("Kyori4Chakukaisu2      TEXT (6),");
		strsql.Append("Kyori4Chakukaisu3      TEXT (6),");
		strsql.Append("Kyori4Chakukaisu4      TEXT (6),");
		strsql.Append("Kyori4Chakukaisu5      TEXT (6),");
		strsql.Append("Kyori4Chakukaisu6      TEXT (6),");
		strsql.Append("Kyori5Chakukaisu1      TEXT (6),");
		strsql.Append("Kyori5Chakukaisu2      TEXT (6),");
		strsql.Append("Kyori5Chakukaisu3      TEXT (6),");
		strsql.Append("Kyori5Chakukaisu4      TEXT (6),");
		strsql.Append("Kyori5Chakukaisu5      TEXT (6),");
		strsql.Append("Kyori5Chakukaisu6      TEXT (6),");
		strsql.Append("Kyori6Chakukaisu1      TEXT (6),");
		strsql.Append("Kyori6Chakukaisu2      TEXT (6),");
		strsql.Append("Kyori6Chakukaisu3      TEXT (6),");
		strsql.Append("Kyori6Chakukaisu4      TEXT (6),");
		strsql.Append("Kyori6Chakukaisu5      TEXT (6),");
		strsql.Append("Kyori6Chakukaisu6      TEXT (6),");
		strsql.Append("CONSTRAINT CHOKYO_SEISEKI PRIMARY KEY (ChokyosiCode,Num));");
		
bstrQuery = strsql; pCn->Execute(bstrQuery,&vRecsAffected,adOptionUnspecified);
		
		
		//生産者マスタ
		strsql = "CREATE TABLE SEISAN (";
		strsql.Append("RecordSpec       TEXT(2),");
		strsql.Append("DataKubun        TEXT(1),");
		strsql.Append("MakeDate TEXT(8),");
		strsql.Append("BreederCode      TEXT(8) CONSTRAINT SEISAN PRIMARY KEY,");
		strsql.Append("BreederName_Co   TEXT(72),");
		strsql.Append("BreederName      TEXT(72),");
		strsql.Append("BreederNameKana  TEXT(72),");
		strsql.Append("BreederNameEng   TEXT(168),");
		strsql.Append("Address  TEXT(20),");
		strsql.Append("H_SetYear        TEXT(4),");
		strsql.Append("H_HonSyokinTotal TEXT(10),");
		strsql.Append("H_FukaSyokin     TEXT(10),");
		strsql.Append("H_ChakuKaisu1    TEXT(6),");
		strsql.Append("H_ChakuKaisu2    TEXT(6),");
		strsql.Append("H_ChakuKaisu3    TEXT(6),");
		strsql.Append("H_ChakuKaisu4    TEXT(6),");
		strsql.Append("H_ChakuKaisu5    TEXT(6),");
		strsql.Append("H_ChakuKaisu6    TEXT(6),");
		strsql.Append("R_SetYear        TEXT(4),");
		strsql.Append("R_HonSyokinTotal TEXT(10),");
		strsql.Append("R_FukaSyokin     TEXT(10),");
		strsql.Append("R_ChakuKaisu1    TEXT(6),");
		strsql.Append("R_ChakuKaisu2    TEXT(6),");
		strsql.Append("R_ChakuKaisu3    TEXT(6),");
		strsql.Append("R_ChakuKaisu4    TEXT(6),");
		strsql.Append("R_ChakuKaisu5    TEXT(6),");
		strsql.Append("R_ChakuKaisu6    TEXT(6));");
		
bstrQuery = strsql; pCn->Execute(bstrQuery,&vRecsAffected,adOptionUnspecified);
		
		
		//馬主マスタ
		strsql = "CREATE TABLE BANUSI (";
		strsql.Append("RecordSpec   TEXT (2),");
		strsql.Append("DataKubun    TEXT (1),");
		strsql.Append("MakeDate     TEXT (8),");
		strsql.Append("BanusiCode   TEXT (6) CONSTRAINT BANUSI PRIMARY KEY UNIQUE, ");
		strsql.Append("BanusiName   TEXT (64),");
		strsql.Append("BanusiName_Co  TEXT(64),");
		strsql.Append("BanusiNameKana TEXT (50),");
		strsql.Append("BanusiNameEng  TEXT (100),");
		strsql.Append("Fukusyoku      TEXT (60),");
		strsql.Append("H_SetYear      TEXT (4),");
		strsql.Append("H_HonSyokinTotal TEXT (10),");
		strsql.Append("H_FukaSyokin   TEXT (10),");
		strsql.Append("H_ChakuKaisu1  TEXT (6),");
		strsql.Append("H_ChakuKaisu2  TEXT (6),");
		strsql.Append("H_ChakuKaisu3  TEXT (6),");
		strsql.Append("H_ChakuKaisu4  TEXT (6),");
		strsql.Append("H_ChakuKaisu5  TEXT (6),");
		strsql.Append("H_ChakuKaisu6  TEXT (6),");
		strsql.Append("R_SetYear      TEXT (4),");
		strsql.Append("R_HonSyokinTotal TEXT (10),");
		strsql.Append("R_FukaSyokin   TEXT (10),");
		strsql.Append("R_ChakuKaisu1  TEXT (6),");
		strsql.Append("R_ChakuKaisu2  TEXT (6),");
		strsql.Append("R_ChakuKaisu3  TEXT (6),");
		strsql.Append("R_ChakuKaisu4  TEXT (6),");
		strsql.Append("R_ChakuKaisu5  TEXT (6),");
		strsql.Append("R_ChakuKaisu6  TEXT (6));");
		
		
bstrQuery = strsql; pCn->Execute(bstrQuery,&vRecsAffected,adOptionUnspecified);
		
		
		//繁殖馬マスタ
		strsql = "CREATE TABLE HANSYOKU (";
		strsql.Append("RecordSpec   TEXT (2),");
		strsql.Append("DataKubun    TEXT (1),");
		strsql.Append("MakeDate     TEXT (8),");
		strsql.Append("HansyokuNum  TEXT (10) CONSTRAINT HANSYOKU PRIMARY KEY UNIQUE,");
		strsql.Append("reserved     TEXT (8),");
		strsql.Append("KettoNum     TEXT (10),");
		strsql.Append("DelKubun     TEXT (1),");		// 現在は予備として使用
		strsql.Append("Bamei        TEXT (36),");
		strsql.Append("BameiKana    TEXT (40),");
		strsql.Append("BameiEng     TEXT (80),");
		strsql.Append("BirthYear    TEXT (4),");
		strsql.Append("SexCD        TEXT (1),");
		strsql.Append("HinsyuCD     TEXT (1),");
		strsql.Append("KeiroCD      TEXT (2),");
		strsql.Append("HansyokuMochiKubun   TEXT (1),");
		strsql.Append("ImportYear   TEXT (4),");
		strsql.Append("SanchiName   TEXT (20),");
		strsql.Append("HansyokuFNum TEXT (10),");
		strsql.Append("HansyokuMNum TEXT (10));");
		
bstrQuery = strsql; pCn->Execute(bstrQuery,&vRecsAffected,adOptionUnspecified);
		
		
		//産駒マスタ
		strsql = "CREATE TABLE SANKU (";
		strsql.Append("RecordSpec       TEXT(2),");
		strsql.Append("DataKubun        TEXT(1),");
		strsql.Append("MakeDate TEXT(8),");
		strsql.Append("KettoNum TEXT(10) CONSTRAINT SANKU PRIMARY KEY,");
		strsql.Append("BirthDate        TEXT(8),");
		strsql.Append("SexCD    TEXT(1),");
		strsql.Append("HinsyuCD TEXT(1),");
		strsql.Append("KeiroCD  TEXT(2),");
		strsql.Append("SankuMochiKubun  TEXT(1),");
		strsql.Append("ImportYear       TEXT(4),");
		strsql.Append("BreederCode      TEXT(8),");
		strsql.Append("SanchiName       TEXT(20),");
		strsql.Append("FNum     TEXT(10),");
		strsql.Append("MNum     TEXT(10),");
		strsql.Append("FFNum    TEXT(10),");
		strsql.Append("FMNum    TEXT(10),");
		strsql.Append("MFNum    TEXT(10),");
		strsql.Append("MMNum    TEXT(10),");
		strsql.Append("FFFNum   TEXT(10),");
		strsql.Append("FFMNum   TEXT(10),");
		strsql.Append("FMFNum   TEXT(10),");
		strsql.Append("FMMNum   TEXT(10),");
		strsql.Append("MFFNum   TEXT(10),");
		strsql.Append("MFMNum   TEXT(10),");
		strsql.Append("MMFNum   TEXT(10),");
		strsql.Append("MMMNum   TEXT(10));");
		
bstrQuery = strsql; pCn->Execute(bstrQuery,&vRecsAffected,adOptionUnspecified);
		
		
		//レコードマスタ
		strsql = "CREATE TABLE RECORD (";
		strsql.Append("RecordSpec       TEXT(2),");
		strsql.Append("DataKubun        TEXT(1),");
		strsql.Append("MakeDate TEXT(8),");
		strsql.Append("RecInfoKubun     TEXT(1),");
		strsql.Append("[Year]     TEXT(4),");
		strsql.Append("MonthDay TEXT(4),");
		strsql.Append("JyoCD    TEXT(2),");
		strsql.Append("Kaiji    TEXT(2),");
		strsql.Append("Nichiji  TEXT(2),");
		strsql.Append("RaceNum  TEXT(2),");
		strsql.Append("TokuNum  TEXT(4),");
		strsql.Append("Hondai   TEXT(60),");
		strsql.Append("GradeCD  TEXT(1),");
		strsql.Append("SyubetuCD_TrackCD        TEXT(4),");
		strsql.Append("Kyori    TEXT(4),");
		strsql.Append("RecKubun TEXT(1),");
		strsql.Append("RecTime  TEXT(4),");
		strsql.Append("TenkoCD  TEXT(1),");
		strsql.Append("SibaBabaCD       TEXT(1),");
		strsql.Append("DirtBabaCD       TEXT(1),");
		strsql.Append("RecUmaKettoNum1  TEXT(10),");
		strsql.Append("RecUmaBamei1     TEXT(36),");
		strsql.Append("RecUmaUmaKigoCD1 TEXT(2),");
		strsql.Append("RecUmaSexCD1     TEXT(1),");
		strsql.Append("RecUmaChokyosiCode1      TEXT(5),");
		strsql.Append("RecUmaChokyosiName1      TEXT(34),");
		strsql.Append("RecUmaFutan1     TEXT(3),");
		strsql.Append("RecUmaKisyuCode1 TEXT(5),");
		strsql.Append("RecUmaKisyuName1 TEXT(34),");
		strsql.Append("RecUmaKettoNum2  TEXT(10),");
		strsql.Append("RecUmaBamei2     TEXT(36),");
		strsql.Append("RecUmaUmaKigoCD2 TEXT(2),");
		strsql.Append("RecUmaSexCD2     TEXT(1),");
		strsql.Append("RecUmaChokyosiCode2      TEXT(5),");
		strsql.Append("RecUmaChokyosiName2      TEXT(34),");
		strsql.Append("RecUmaFutan2     TEXT(3),");
		strsql.Append("RecUmaKisyuCode2 TEXT(5),");
		strsql.Append("RecUmaKisyuName2 TEXT(34),");
		strsql.Append("RecUmaKettoNum3  TEXT(10),");
		strsql.Append("RecUmaBamei3     TEXT(36),");
		strsql.Append("RecUmaUmaKigoCD3 TEXT(2),");
		strsql.Append("RecUmaSexCD3     TEXT(1),");
		strsql.Append("RecUmaChokyosiCode3      TEXT(5),");
		strsql.Append("RecUmaChokyosiName3      TEXT(34),");
		strsql.Append("RecUmaFutan3     TEXT(3),");
		strsql.Append("RecUmaKisyuCode3 TEXT(5),");
		strsql.Append("RecUmaKisyuName3 TEXT(34),");
		strsql.Append("CONSTRAINT RECORD PRIMARY KEY (RecInfoKubun,[Year],MonthDay,JyoCD,Kaiji,Nichiji,RaceNum,TokuNum,SyubetuCD_TrackCD,Kyori));");
		
bstrQuery = strsql; pCn->Execute(bstrQuery,&vRecsAffected,adOptionUnspecified);
		
		
		//坂路調教
		strsql = "CREATE TABLE HANRO (";
		strsql.Append("RecordSpec   TEXT (2),");
		strsql.Append("DataKubun    TEXT (1),");
		strsql.Append("MakeDate     TEXT (8),");
		strsql.Append("TresenKubun  TEXT (1),");
		strsql.Append("ChokyoDate   TEXT (8),");
		strsql.Append("ChokyoTime   TEXT (4),");
		strsql.Append("KettoNum     TEXT (10),");
		strsql.Append("HaronTime4   TEXT (4),");
		strsql.Append("LapTime4     TEXT (3),");
		strsql.Append("HaronTime3   TEXT (4),");
		strsql.Append("LapTime3     TEXT (3),");
		strsql.Append("HaronTime2   TEXT (4),");
		strsql.Append("LapTime2     TEXT (3),");
		strsql.Append("LapTime1     TEXT (3),");
		strsql.Append("CONSTRAINT HANRO PRIMARY KEY (TresenKubun,ChokyoDate,ChokyoTime,KettoNum));");
		
bstrQuery = strsql; pCn->Execute(bstrQuery,&vRecsAffected,adOptionUnspecified);
		
		
		//馬体重
		strsql = "CREATE TABLE BATAIJYU (";
		strsql.Append("RecordSpec TEXT (2),");
		strsql.Append("DataKubun  TEXT (1),");
		strsql.Append("MakeDate   TEXT (8),");
		strsql.Append("[Year]       TEXT (4),");
		strsql.Append("MonthDay   TEXT (4),");
		strsql.Append("JyoCD      TEXT (2),");
		strsql.Append("Kaiji      TEXT (2),");
		strsql.Append("Nichiji    TEXT (2),");
		strsql.Append("RaceNum    TEXT (2),");
		strsql.Append("HappyoTime TEXT (8),");
		strsql.Append("Umaban1    TEXT (2),");
		strsql.Append("Bamei1     TEXT (36),");
		strsql.Append("BaTaijyu1  TEXT (3),");
		strsql.Append("ZogenFugo1 TEXT (1),");
		strsql.Append("ZogenSa1   TEXT (3),");
		strsql.Append("Umaban2    TEXT (2),");
		strsql.Append("Bamei2     TEXT (36),");
		strsql.Append("BaTaijyu2  TEXT (3),");
		strsql.Append("ZogenFugo2 TEXT (1),");
		strsql.Append("ZogenSa2   TEXT (3),");
		strsql.Append("Umaban3    TEXT (2),");
		strsql.Append("Bamei3     TEXT (36),");
		strsql.Append("BaTaijyu3  TEXT (3),");
		strsql.Append("ZogenFugo3 TEXT (1),");
		strsql.Append("ZogenSa3   TEXT (3),");
		strsql.Append("Umaban4    TEXT (2),");
		strsql.Append("Bamei4     TEXT (36),");
		strsql.Append("BaTaijyu4  TEXT (3),");
		strsql.Append("ZogenFugo4 TEXT (1),");
		strsql.Append("ZogenSa4   TEXT (3),");
		strsql.Append("Umaban5    TEXT (2),");
		strsql.Append("Bamei5     TEXT (36),");
		strsql.Append("BaTaijyu5  TEXT (3),");
		strsql.Append("ZogenFugo5 TEXT (1),");
		strsql.Append("ZogenSa5   TEXT (3),");
		strsql.Append("Umaban6    TEXT (2),");
		strsql.Append("Bamei6     TEXT (36),");
		strsql.Append("BaTaijyu6  TEXT (3),");
		strsql.Append("ZogenFugo6 TEXT (1),");
		strsql.Append("ZogenSa6   TEXT (3),");
		strsql.Append("Umaban7    TEXT (2),");
		strsql.Append("Bamei7     TEXT (36),");
		strsql.Append("BaTaijyu7  TEXT (3),");
		strsql.Append("ZogenFugo7 TEXT (1),");
		strsql.Append("ZogenSa7   TEXT (3),");
		strsql.Append("Umaban8    TEXT (2),");
		strsql.Append("Bamei8     TEXT (36),");
		strsql.Append("BaTaijyu8  TEXT (3),");
		strsql.Append("ZogenFugo8 TEXT (1),");
		strsql.Append("ZogenSa8   TEXT (3),");
		strsql.Append("Umaban9    TEXT (2),");
		strsql.Append("Bamei9     TEXT (36),");
		strsql.Append("BaTaijyu9  TEXT (3),");
		strsql.Append("ZogenFugo9 TEXT (1),");
		strsql.Append("ZogenSa9   TEXT (3),");
		strsql.Append("Umaban10    TEXT (2),");
		strsql.Append("Bamei10     TEXT (36),");
		strsql.Append("BaTaijyu10  TEXT (3),");
		strsql.Append("ZogenFugo10 TEXT (1),");
		strsql.Append("ZogenSa10   TEXT (3),");
		strsql.Append("Umaban11    TEXT (2),");
		strsql.Append("Bamei11     TEXT (36),");
		strsql.Append("BaTaijyu11  TEXT (3),");
		strsql.Append("ZogenFugo11 TEXT (1),");
		strsql.Append("ZogenSa11   TEXT (3),");
		strsql.Append("Umaban12    TEXT (2),");
		strsql.Append("Bamei12     TEXT (36),");
		strsql.Append("BaTaijyu12  TEXT (3),");
		strsql.Append("ZogenFugo12 TEXT (1),");
		strsql.Append("ZogenSa12   TEXT (3),");
		strsql.Append("Umaban13    TEXT (2),");
		strsql.Append("Bamei13     TEXT (36),");
		strsql.Append("BaTaijyu13  TEXT (3),");
		strsql.Append("ZogenFugo13 TEXT (1),");
		strsql.Append("ZogenSa13   TEXT (3),");
		strsql.Append("Umaban14    TEXT (2),");
		strsql.Append("Bamei14     TEXT (36),");
		strsql.Append("BaTaijyu14  TEXT (3),");
		strsql.Append("ZogenFugo14 TEXT (1),");
		strsql.Append("ZogenSa14   TEXT (3),");
		strsql.Append("Umaban15    TEXT (2),");
		strsql.Append("Bamei15     TEXT (36),");
		strsql.Append("BaTaijyu15  TEXT (3),");
		strsql.Append("ZogenFugo15 TEXT (1),");
		strsql.Append("ZogenSa15   TEXT (3),");
		strsql.Append("Umaban16    TEXT (2),");
		strsql.Append("Bamei16     TEXT (36),");
		strsql.Append("BaTaijyu16  TEXT (3),");
		strsql.Append("ZogenFugo16 TEXT (1),");
		strsql.Append("ZogenSa16   TEXT (3),");
		strsql.Append("Umaban17    TEXT (2),");
		strsql.Append("Bamei17     TEXT (36),");
		strsql.Append("BaTaijyu17  TEXT (3),");
		strsql.Append("ZogenFugo17 TEXT (1),");
		strsql.Append("ZogenSa17   TEXT (3),");
		strsql.Append("Umaban18    TEXT (2),");
		strsql.Append("Bamei18     TEXT (36),");
		strsql.Append("BaTaijyu18  TEXT (3),");
		strsql.Append("ZogenFugo18 TEXT (1),");
		strsql.Append("ZogenSa18   TEXT (3),");
		strsql.Append("CONSTRAINT BATAIJYU PRIMARY KEY ([Year],MonthDay,JyoCD,Kaiji,Nichiji,RaceNum));");
		
bstrQuery = strsql; pCn->Execute(bstrQuery,&vRecsAffected,adOptionUnspecified);
		
		
		//天候馬場状態
		strsql = "CREATE TABLE TENKO_BABA (";
		strsql.Append("RecordSpec       TEXT(2),");
		strsql.Append("DataKubun        TEXT(1),");
		strsql.Append("MakeDate TEXT(8),");
		strsql.Append("[Year]     TEXT(4),");
		strsql.Append("MonthDay TEXT(4),");
		strsql.Append("JyoCD    TEXT(2),");
		strsql.Append("Kaiji    TEXT(2),");
		strsql.Append("Nichiji  TEXT(2),");
		strsql.Append("HappyoTime       TEXT(8),");
		strsql.Append("HenkoID  TEXT(1),");
		strsql.Append("AtoTenkoCD       TEXT(1),");
		strsql.Append("AtoSibaBabaCD    TEXT(1),");
		strsql.Append("AtoDirtBabaCD    TEXT(1),");
		strsql.Append("MaeTenkoCD       TEXT(1),");
		strsql.Append("MaeSibaBabaCD    TEXT(1),");
		strsql.Append("MaeDirtBabaCD    TEXT(1),");
		strsql.Append("CONSTRAINT TENKO_BABA PRIMARY KEY ([Year],MonthDay,JyoCD,Kaiji,Nichiji,HappyoTime,HenkoID));");
		
bstrQuery = strsql; pCn->Execute(bstrQuery,&vRecsAffected,adOptionUnspecified);
		
		
		//出走取消・競走除外
		strsql = "CREATE TABLE TORIKESI_JYOGAI (";
		strsql.Append("RecordSpec       TEXT(2),");
		strsql.Append("DataKubun        TEXT(1),");
		strsql.Append("MakeDate TEXT(8),");
		strsql.Append("[Year]     TEXT(4),");
		strsql.Append("MonthDay TEXT(4),");
		strsql.Append("JyoCD    TEXT(2),");
		strsql.Append("Kaiji    TEXT(2),");
		strsql.Append("Nichiji  TEXT(2),");
		strsql.Append("RaceNum  TEXT(2),");
		strsql.Append("HappyoTime       TEXT(8),");
		strsql.Append("Umaban   TEXT(2),");
		strsql.Append("Bamei    TEXT(36),");
		strsql.Append("JiyuKubun        TEXT(3),");
		strsql.Append("CONSTRAINT TORIKESI_JYOGAI PRIMARY KEY ([Year],MonthDay,JyoCD,Kaiji,Nichiji,RaceNum,Umaban));");
		
bstrQuery = strsql; pCn->Execute(bstrQuery,&vRecsAffected,adOptionUnspecified);
		
		
		//騎手変更
		strsql = "CREATE TABLE KISYU_CHANGE (";
		strsql.Append("RecordSpec       TEXT(2),");
		strsql.Append("DataKubun        TEXT(1),");
		strsql.Append("MakeDate TEXT(8),");
		strsql.Append("[Year]     TEXT(4),");
		strsql.Append("MonthDay TEXT(4),");
		strsql.Append("JyoCD    TEXT(2),");
		strsql.Append("Kaiji    TEXT(2),");
		strsql.Append("Nichiji  TEXT(2),");
		strsql.Append("RaceNum  TEXT(2),");
		strsql.Append("HappyoTime       TEXT(8),");
		strsql.Append("Umaban   TEXT(2),");
		strsql.Append("Bamei    TEXT(36),");
		strsql.Append("AtoFutan TEXT(3),");
		strsql.Append("AtoKisyuCode     TEXT(5),");
		strsql.Append("AtoKisyuName     TEXT(34),");
		strsql.Append("AtoMinaraiCD     TEXT(1),");
		strsql.Append("MaeFutan TEXT(3),");
		strsql.Append("MaeKisyuCode     TEXT(5),");
		strsql.Append("MaeKisyuName     TEXT(34),");
		strsql.Append("MaeMinaraiCD     TEXT(1),");
		strsql.Append("CONSTRAINT KISYU_CHANGE PRIMARY KEY ([Year],MonthDay,JyoCD,Kaiji,Nichiji,RaceNum,HappyoTime,Umaban));");
		
bstrQuery = strsql; pCn->Execute(bstrQuery,&vRecsAffected,adOptionUnspecified);
		
		
		//発走時刻変更
		strsql = "CREATE TABLE HASSOU_JIKOKU_CHANGE (";
		strsql.Append("RecordSpec       TEXT(2),");
		strsql.Append("DataKubun        TEXT(1),");
		strsql.Append("MakeDate TEXT(8),");
		strsql.Append("[Year]     TEXT(4),");
		strsql.Append("MonthDay TEXT(4),");
		strsql.Append("JyoCD    TEXT(2),");
		strsql.Append("Kaiji    TEXT(2),");
		strsql.Append("Nichiji  TEXT(2),");
		strsql.Append("RaceNum  TEXT(2),");
		strsql.Append("HappyoTime       TEXT(8),");
		strsql.Append("AtoJi     TEXT(2),");
		strsql.Append("AtoFun    TEXT(2),");
		strsql.Append("MaeJi     TEXT(2),");
		strsql.Append("MaeFun    TEXT(2),");

		strsql.Append("CONSTRAINT HASSOU_JIKOKU_CHANGE PRIMARY KEY ([Year],MonthDay,JyoCD,Kaiji,Nichiji,RaceNum,HappyoTime));");
		
bstrQuery = strsql; pCn->Execute(bstrQuery,&vRecsAffected,adOptionUnspecified);


		//コース変更
		strsql = "CREATE TABLE COURSE_CHANGE (";
		strsql.Append("RecordSpec       TEXT(2),");
		strsql.Append("DataKubun        TEXT(1),");
		strsql.Append("MakeDate			TEXT(8),");
		strsql.Append("[Year]			TEXT(4),");
		strsql.Append("MonthDay			TEXT(4),");
		strsql.Append("JyoCD			TEXT(2),");
		strsql.Append("Kaiji			TEXT(2),");
		strsql.Append("Nichiji			TEXT(2),");
		strsql.Append("RaceNum			TEXT(2),");
		strsql.Append("HappyoTime		TEXT(8),");
		strsql.Append("AtoKyori			TEXT(4),");
		strsql.Append("AtoTruckCD		TEXT(2),");
		strsql.Append("MaeKyori			TEXT(4),");
		strsql.Append("MaeTruckCD		TEXT(2),");
		strsql.Append("JiyuCD			TEXT(1),");
		strsql.Append("CONSTRAINT COURSE_CHANGE PRIMARY KEY ([Year],MonthDay,JyoCD,Kaiji,Nichiji,RaceNum,HappyoTime));");
		
bstrQuery = strsql; pCn->Execute(bstrQuery,&vRecsAffected,adOptionUnspecified);


		//データマイニング予想
		strsql = "CREATE TABLE MINING (";
		strsql.Append("RecordSpec       TEXT(2),");
		strsql.Append("DataKubun        TEXT(1),");
		strsql.Append("MakeDate TEXT(8),");
		strsql.Append("[Year]     TEXT(4),");
		strsql.Append("MonthDay TEXT(4),");
		strsql.Append("JyoCD    TEXT(2),");
		strsql.Append("Kaiji    TEXT(2),");
		strsql.Append("Nichiji  TEXT(2),");
		strsql.Append("RaceNum  TEXT(2),");
		strsql.Append("MakeHM   TEXT(4),");
		strsql.Append("Umaban1  TEXT(2),");
		strsql.Append("DMTime1  TEXT(5),");
		strsql.Append("DMGosaP1 TEXT(4),");
		strsql.Append("DMGosaM1 TEXT(4),");
		strsql.Append("Umaban2  TEXT(2),");
		strsql.Append("DMTime2  TEXT(5),");
		strsql.Append("DMGosaP2 TEXT(4),");
		strsql.Append("DMGosaM2 TEXT(4),");
		strsql.Append("Umaban3  TEXT(2),");
		strsql.Append("DMTime3  TEXT(5),");
		strsql.Append("DMGosaP3 TEXT(4),");
		strsql.Append("DMGosaM3 TEXT(4),");
		strsql.Append("Umaban4  TEXT(2),");
		strsql.Append("DMTime4  TEXT(5),");
		strsql.Append("DMGosaP4 TEXT(4),");
		strsql.Append("DMGosaM4 TEXT(4),");
		strsql.Append("Umaban5  TEXT(2),");
		strsql.Append("DMTime5  TEXT(5),");
		strsql.Append("DMGosaP5 TEXT(4),");
		strsql.Append("DMGosaM5 TEXT(4),");
		strsql.Append("Umaban6  TEXT(2),");
		strsql.Append("DMTime6  TEXT(5),");
		strsql.Append("DMGosaP6 TEXT(4),");
		strsql.Append("DMGosaM6 TEXT(4),");
		strsql.Append("Umaban7  TEXT(2),");
		strsql.Append("DMTime7  TEXT(5),");
		strsql.Append("DMGosaP7 TEXT(4),");
		strsql.Append("DMGosaM7 TEXT(4),");
		strsql.Append("Umaban8  TEXT(2),");
		strsql.Append("DMTime8  TEXT(5),");
		strsql.Append("DMGosaP8 TEXT(4),");
		strsql.Append("DMGosaM8 TEXT(4),");
		strsql.Append("Umaban9  TEXT(2),");
		strsql.Append("DMTime9  TEXT(5),");
		strsql.Append("DMGosaP9 TEXT(4),");
		strsql.Append("DMGosaM9 TEXT(4),");
		strsql.Append("Umaban10 TEXT(2),");
		strsql.Append("DMTime10 TEXT(5),");
		strsql.Append("DMGosaP10        TEXT(4),");
		strsql.Append("DMGosaM10        TEXT(4),");
		strsql.Append("Umaban11 TEXT(2),");
		strsql.Append("DMTime11 TEXT(5),");
		strsql.Append("DMGosaP11        TEXT(4),");
		strsql.Append("DMGosaM11        TEXT(4),");
		strsql.Append("Umaban12 TEXT(2),");
		strsql.Append("DMTime12 TEXT(5),");
		strsql.Append("DMGosaP12        TEXT(4),");
		strsql.Append("DMGosaM12        TEXT(4),");
		strsql.Append("Umaban13 TEXT(2),");
		strsql.Append("DMTime13 TEXT(5),");
		strsql.Append("DMGosaP13        TEXT(4),");
		strsql.Append("DMGosaM13        TEXT(4),");
		strsql.Append("Umaban14 TEXT(2),");
		strsql.Append("DMTime14 TEXT(5),");
		strsql.Append("DMGosaP14        TEXT(4),");
		strsql.Append("DMGosaM14        TEXT(4),");
		strsql.Append("Umaban15 TEXT(2),");
		strsql.Append("DMTime15 TEXT(5),");
		strsql.Append("DMGosaP15        TEXT(4),");
		strsql.Append("DMGosaM15        TEXT(4),");
		strsql.Append("Umaban16 TEXT(2),");
		strsql.Append("DMTime16 TEXT(5),");
		strsql.Append("DMGosaP16        TEXT(4),");
		strsql.Append("DMGosaM16        TEXT(4),");
		strsql.Append("Umaban17 TEXT(2),");
		strsql.Append("DMTime17 TEXT(5),");
		strsql.Append("DMGosaP17        TEXT(4),");
		strsql.Append("DMGosaM17        TEXT(4),");
		strsql.Append("Umaban18 TEXT(2),");
		strsql.Append("DMTime18 TEXT(5),");
		strsql.Append("DMGosaP18        TEXT(4),");
		strsql.Append("DMGosaM18        TEXT(4),");
		strsql.Append("CONSTRAINT MINING PRIMARY KEY ([Year],MonthDay,JyoCD,Kaiji,Nichiji,RaceNum));");
		
bstrQuery = strsql; pCn->Execute(bstrQuery,&vRecsAffected,adOptionUnspecified);
		
		
		//年間スケジュール
		strsql = "CREATE TABLE SCHEDULE (";
		strsql.Append("RecordSpec TEXT (2),");
		strsql.Append("DataKubun  TEXT (1),");
		strsql.Append("MakeDate   TEXT (8),");
		strsql.Append("[Year]       TEXT (4),");
		strsql.Append("MonthDay   TEXT (4),");
		strsql.Append("JyoCD      TEXT (2),");
		strsql.Append("Kaiji      TEXT (2),");
		strsql.Append("Nichiji    TEXT (2),");
		strsql.Append("YoubiCD  TEXT(1),");
		strsql.Append("Jyusyo1TokuNum   TEXT(4),");
		strsql.Append("Jyusyo1Hondai    TEXT(60),");
		strsql.Append("Jyusyo1Ryakusyo10        TEXT(20),");
		strsql.Append("Jyusyo1Ryakusyo6 TEXT(12),");
		strsql.Append("Jyusyo1Ryakusyo3 TEXT(6),");
		strsql.Append("Jyusyo1Nkai      TEXT(3),");
		strsql.Append("Jyusyo1GradeCD   TEXT(1),");
		strsql.Append("Jyusyo1SyubetuCD TEXT(2),");
		strsql.Append("Jyusyo1KigoCD    TEXT(3),");
		strsql.Append("Jyusyo1JyuryoCD  TEXT(1),");
		strsql.Append("Jyusyo1Kyori     TEXT(4),");
		strsql.Append("Jyusyo1TrackCD   TEXT(2),");
		strsql.Append("Jyusyo2TokuNum   TEXT(4),");
		strsql.Append("Jyusyo2Hondai    TEXT(60),");
		strsql.Append("Jyusyo2Ryakusyo10        TEXT(20),");
		strsql.Append("Jyusyo2Ryakusyo6 TEXT(12),");
		strsql.Append("Jyusyo2Ryakusyo3 TEXT(6),");
		strsql.Append("Jyusyo2Nkai      TEXT(3),");
		strsql.Append("Jyusyo2GradeCD   TEXT(1),");
		strsql.Append("Jyusyo2SyubetuCD TEXT(2),");
		strsql.Append("Jyusyo2KigoCD    TEXT(3),");
		strsql.Append("Jyusyo2JyuryoCD  TEXT(1),");
		strsql.Append("Jyusyo2Kyori     TEXT(4),");
		strsql.Append("Jyusyo2TrackCD   TEXT(2),");
		strsql.Append("Jyusyo3TokuNum   TEXT(4),");
		strsql.Append("Jyusyo3Hondai    TEXT(60),");
		strsql.Append("Jyusyo3Ryakusyo10        TEXT(20),");
		strsql.Append("Jyusyo3Ryakusyo6 TEXT(12),");
		strsql.Append("Jyusyo3Ryakusyo3 TEXT(6),");
		strsql.Append("Jyusyo3Nkai      TEXT(3),");
		strsql.Append("Jyusyo3GradeCD   TEXT(1),");
		strsql.Append("Jyusyo3SyubetuCD TEXT(2),");
		strsql.Append("Jyusyo3KigoCD    TEXT(3),");
		strsql.Append("Jyusyo3JyuryoCD  TEXT(1),");
		strsql.Append("Jyusyo3Kyori     TEXT(4),");
		strsql.Append("Jyusyo3TrackCD   TEXT(2),");
		strsql.Append("CONSTRAINT SCHEDULE PRIMARY KEY ([Year],MonthDay,JyoCD,Kaiji,Nichiji));");
		
		bstrQuery = strsql; 
		pCn->Execute(bstrQuery,&vRecsAffected,adOptionUnspecified);

		//競走馬市場取引価格
		strsql = "CREATE TABLE SALE (";
		strsql.Append("RecordSpec		TEXT(2),");
		strsql.Append("DataKubun		TEXT(1),");
		strsql.Append("MakeDate			TEXT(8),");
		strsql.Append("KettoNum			TEXT(10),");
		strsql.Append("HansyokuFNum		TEXT(10),");
		strsql.Append("HansyokuMNum		TEXT(10),");
		strsql.Append("BirthYear		TEXT(4),");
		strsql.Append("SaleCode			TEXT(6),");
		strsql.Append("SaleHostName     TEXT(40),");
		strsql.Append("SaleName			TEXT(80),");
		strsql.Append("FromDate			TEXT(8),");
		strsql.Append("ToDate			TEXT(8),");
		strsql.Append("Barei			TEXT(1),");
		strsql.Append("Price			TEXT(10),");
		strsql.Append("CONSTRAINT SALE PRIMARY KEY (KettoNum,SaleCode,FromDate));");
		
		bstrQuery = strsql; 
		pCn->Execute(bstrQuery,&vRecsAffected,adOptionUnspecified);

        //馬名の意味由来
		strsql = "CREATE TABLE BAMEIORIGIN (";
		strsql.Append("RecordSpec		TEXT(2),");
		strsql.Append("DataKubun		TEXT(1),");
		strsql.Append("MakeDate			TEXT(8),");
		strsql.Append("KettoNum			TEXT(10) CONSTRAINT BAMEIORIGIN PRIMARY KEY UNIQUE,");
		strsql.Append("Bamei			TEXT(36),");
		strsql.Append("Origin			TEXT(64));");

		bstrQuery = strsql; 
		pCn->Execute(bstrQuery,&vRecsAffected,adOptionUnspecified);
		
		//系統情報
		strsql = "CREATE TABLE KEITO (";
		strsql.Append("RecordSpec       TEXT(2),");
		strsql.Append("DataKubun        TEXT(1),");
		strsql.Append("MakeDate         TEXT(8),");
		strsql.Append("HansyokuNum      TEXT(10) CONSTRAINT KEITO PRIMARY KEY,");
		strsql.Append("KeitoId          TEXT(30),");
		strsql.Append("KeitoName        TEXT(36),");
		strsql.Append("KeitoEx          MEMO);");
		
		bstrQuery = strsql;
		pCn->Execute(bstrQuery,&vRecsAffected,adOptionUnspecified);

		//コース情報
		strsql = "CREATE TABLE COURSE (";
		strsql.Append("RecordSpec       TEXT(2),");
		strsql.Append("DataKubun        TEXT(1),");
		strsql.Append("MakeDate         TEXT(8),");
		strsql.Append("JyoCD            TEXT(2),");
		strsql.Append("Kyori            TEXT(4),");
		strsql.Append("TrackCD          TEXT(2),");
		strsql.Append("KaishuDate       TEXT(8),");
		strsql.Append("CourseEx         MEMO,");
		strsql.Append("CONSTRAINT COURSE PRIMARY KEY (JyoCD,Kyori,TrackCD,KaishuDate));");

		bstrQuery = strsql;
		pCn->Execute(bstrQuery,&vRecsAffected,adOptionUnspecified);
		pCn->CommitTrans();
}
		catch(_com_error &e){
			(void)e;
			pCn->RollbackTrans();
		return -1;
	}
	return 0;
}

int clsDBBuilder::CompactDB(CString strFilePath)
{
	DAO::_DBEngine* pEngine = NULL;
	int i = 0;
	CString strFilePath_tmp = "";

	do {
		CString str;
		str.Format(_T("%d"), i);
		strFilePath_tmp = strFilePath + str;
		i++;
	} while (PathFileExists(strFilePath_tmp));

	HRESULT hr = CoCreateInstance(__uuidof(DAO::DBEngine), NULL, CLSCTX_ALL, IID_IDispatch, (LPVOID*)&pEngine);
	if (SUCCEEDED(hr) && pEngine) {
		pEngine->CompactDatabase(_bstr_t(strFilePath), _bstr_t(strFilePath_tmp));
		pEngine->Release();
		pEngine = NULL;
		DeleteFile(strFilePath);
		MoveFile(strFilePath_tmp, strFilePath);
	} else {
		return -1;
	}

	return 0;
}

int clsDBBuilder::KillDB(CString strFilePath)
{

	if(GetFileAttributes(strFilePath) != (DWORD)-1){
		DeleteFile(strFilePath); 
	}else{
		MessageBox(NULL,"指定されたファイルは存在しません。",NULL,NULL);
		return -1;
	}
	return 0;
}
