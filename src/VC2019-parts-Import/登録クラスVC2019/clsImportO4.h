#pragma once
#include "JVData_Structure.h"


class clsImportO4
{
public:
	clsImportO4(void);
	~clsImportO4(void);
	int Add(CString strBuff, long lngBuffSize);
private:
	int InsertDB(void);
	int UpdateDB(CString strMakeDate);
	JV_O4_ODDS_UMATAN mBuf;
	_ConnectionPtr pCn;
public:
	int Init(_ConnectionPtr m_Connect);
};
