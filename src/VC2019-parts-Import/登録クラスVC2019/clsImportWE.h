#pragma once
#include "JVData_Structure.h"


class clsImportWE
{
public:
	clsImportWE(void);
	~clsImportWE(void);
	int Add(CString strBuff, long lngBuffSize);
private:
	int InsertDB(void);
	int UpdateDB(CString strMakeDate);
	JV_WE_WEATHER mBuf;
	_ConnectionPtr pCn;
public:
	int Init(_ConnectionPtr m_Connect);
};
