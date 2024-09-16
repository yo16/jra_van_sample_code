/*=======================================================================
  JRA-VAN Data Lab.�v���O���~���O�p�[�c�u�f�[�^�o�^�p�[�c�@TC���R�[�h�o�^�N���X�v

	   �쐬: JRA-VAN �\�t�g�E�F�A�H�[

========================================================================
   (C) Copyright JRA SYSTEM SERVICE CO.,LTD. 2007 All rights reserved
========================================================================*/

#include "StdAfx.h"
#include "clsImportTC.h"

clsImportTC::clsImportTC(void)
{
}

clsImportTC::~clsImportTC(void)
{
}
int clsImportTC::Add(CString strBuff, long lngBuffSize)
{
	CString strMakeDate;
    memcpy(&mBuf,strBuff.GetBuffer(lngBuffSize),lngBuffSize);
    strMakeDate.SetString(mBuf.head.MakeDate.Year,4);
	strMakeDate.Append(mBuf.head.MakeDate.Month,2);
	strMakeDate.Append(mBuf.head.MakeDate.Day,2);

//      INSERT����
	if(InsertDB() != 0 ){
//			UPDATE�����iINSERT�����s�����ꍇ�j
		if(UpdateDB(strMakeDate)!=0){
//				System.Diagnostics.Debug.WriteLine("�X�V�Ɏ��s���܂����B" & Left(strBuf, 2))
		}
	}
	return 0;
}

int clsImportTC::InsertDB(void)
{
	CString strsql;
	strsql="SELECT * FROM HASSOU_JIKOKU_CHANGE WHERE [Year]='";
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
	strsql.Append("' AND [HappyoTime]='");
	strsql.Append(mBuf.HappyoTime.Month,8);
	strsql.Append("'");
	_bstr_t bstrQuery(strsql);
    // Command�I�u�W�F�N�g
	_CommandPtr pCommand;
	_RecordsetPtr pRecordSet;

	pCommand.CreateInstance(__uuidof(Command));
	pRecordSet.CreateInstance(__uuidof(Recordset));

	// SQL�̐ݒ�
	pCommand->ActiveConnection = pCn;
	pCommand->CommandText = bstrQuery;
	pRecordSet->PutRefSource(pCommand);


	_variant_t vNull;  // VARIANT�^��NULL�Ƃ��Ďg��
	vNull.vt = VT_ERROR;
	vNull.scode = DISP_E_PARAMNOTFOUND;

	// ���R�[�h�Z�b�g�̎擾
	pRecordSet->Open(vNull, vNull, adOpenKeyset, adLockOptimistic, adCmdText);

	if (!pRecordSet->GetadoEOF()){
		pRecordSet->Close();
		pRecordSet = NULL;
		return -1;
	}

	USES_CONVERSION;

	COleSafeArray vaFieldlist;
	vaFieldlist.CreateOneDim(VT_VARIANT,14);

	COleSafeArray vaValuelist;
	vaValuelist.CreateOneDim(VT_VARIANT,14);
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
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("AtoJi")));	
	strsql.SetString(mBuf.TCInfoAfter.Ji,2);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=11;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("AtoFun")));	
	strsql.SetString(mBuf.TCInfoAfter.Fun,2);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=12;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("MaeJi")));	
	strsql.SetString(mBuf.TCInfoBefore.Ji,2);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=13;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("MaeFun")));	
	strsql.SetString(mBuf.TCInfoBefore.Fun,2);
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

int clsImportTC::UpdateDB(CString strMakeDate)
{
	CString strsql;
	strsql="SELECT * FROM HASSOU_JIKOKU_CHANGE WHERE [Year]='";
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
	strsql.Append("' AND [HappyoTime]='");
	strsql.Append(mBuf.HappyoTime.Month,8);
	strsql.Append("' AND [MakeDate] <= '");
	strsql.Append(strMakeDate);
	strsql.Append("'");

	_bstr_t bstrQuery(strsql);
    // Command�I�u�W�F�N�g
	_CommandPtr pCommand;
	_RecordsetPtr pRecordSet;

	pCommand.CreateInstance(__uuidof(Command));
	pRecordSet.CreateInstance(__uuidof(Recordset));

	// SQL�̐ݒ�
	pCommand->ActiveConnection = pCn;
	pCommand->CommandText = bstrQuery;
	pRecordSet->PutRefSource(pCommand);


	_variant_t vNull;  // VARIANT�^��NULL�Ƃ��Ďg��
	vNull.vt = VT_ERROR;
	vNull.scode = DISP_E_PARAMNOTFOUND;

	// ���R�[�h�Z�b�g�̎擾
	pRecordSet->Open(vNull, vNull, adOpenKeyset, adLockOptimistic, adCmdText);
	USES_CONVERSION;

	COleSafeArray vaFieldlist;
	vaFieldlist.CreateOneDim(VT_VARIANT,14);

	COleSafeArray vaValuelist;
	vaValuelist.CreateOneDim(VT_VARIANT,14);
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
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("AtoJi")));	
	strsql.SetString(mBuf.TCInfoAfter.Ji,2);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=11;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("AtoFun")));	
	strsql.SetString(mBuf.TCInfoAfter.Fun,2);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=12;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("MaeJi")));	
	strsql.SetString(mBuf.TCInfoBefore.Ji,2);
	vaValuelist.PutElement(lArrayIndex, &(_variant_t(T2OLE(strsql))));

	lArrayIndex[0]=13;
	vaFieldlist.PutElement(lArrayIndex, &(_variant_t("MaeFun")));	
	strsql.SetString(mBuf.TCInfoBefore.Fun,2);
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
int clsImportTC::Init(_ConnectionPtr m_Connect)
{
    pCn = m_Connect;

	 return 0;
}

