#pragma once

#pragma once
#include "JVData_Structure.h"


class clsImportO6
{
public:
	clsImportO6(void);
	~clsImportO6(void);
	int Add(CString strBuff, long lngBuffSize);
private:
	int InsertDB(void);
	int UpdateDB(CString strMakeDate);
	JV_O6_ODDS_SANRENTAN mBuf;
	_ConnectionPtr pCn;
public:
	int Init(_ConnectionPtr m_Connect);
};
