#pragma once
#include "JVData_Structure.h"

class clsImportSE
{
public:
	clsImportSE(void);
	~clsImportSE(void);
	int Add(CString strBuff, long lngBuffSize);
	int Init(_ConnectionPtr m_Connect);
private:
	int InsertDB(void);
	int UpdateDB(CString strMakeDate);
	_ConnectionPtr pCn;
	JV_SE_RACE_UMA mBuf;
};
