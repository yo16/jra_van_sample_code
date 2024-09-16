#pragma once

#pragma once
#include "JVData_Structure.h"


class clsImportO5
{
public:
	clsImportO5(void);
	~clsImportO5(void);
	int Add(CString strBuff, long lngBuffSize);
private:
	int InsertDB(void);
	int UpdateDB(CString strMakeDate);
	JV_O5_ODDS_SANREN mBuf;
	_ConnectionPtr pCn;
public:
	int Init(_ConnectionPtr m_Connect);
};
