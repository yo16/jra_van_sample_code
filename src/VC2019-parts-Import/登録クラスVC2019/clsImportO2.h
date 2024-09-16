#pragma once
#include "JVData_Structure.h"


class clsImportO2
{
public:
	clsImportO2(void);
	~clsImportO2(void);
	int Add(CString strBuff, long lngBuffSize);
private:
	int InsertDB(void);
	int UpdateDB(CString strMakeDate);
	JV_O2_ODDS_UMAREN mBuf;
	_ConnectionPtr pCn;
public:
	int Init(_ConnectionPtr m_Connect);
};
