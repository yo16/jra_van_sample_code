#pragma once

#include "JVData_Structure.h"

class clsImportRA
{
public:
	clsImportRA(void);
	~clsImportRA(void);
public:
	int Init(_ConnectionPtr m_Connect);
	int Add(CString strBuff, long lngBuffSize);
	JV_RA_RACE mBuf;
private:
	int InsertDB(void);
	int UpdateDB(CString strMakeDate);
	_ConnectionPtr pCn;
};
