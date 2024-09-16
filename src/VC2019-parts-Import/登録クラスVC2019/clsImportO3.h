#pragma once
#include "JVData_Structure.h"


class clsImportO3
{
public:
	clsImportO3(void);
	~clsImportO3(void);
	int Add(CString strBuff, long lngBuffSize);
private:
	int InsertDB(void);
	int UpdateDB(CString strMakeDate);
	JV_O3_ODDS_WIDE mBuf;
	_ConnectionPtr pCn;
public:
	int Init(_ConnectionPtr m_Connect);
};
