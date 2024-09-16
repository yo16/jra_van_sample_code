#pragma once
#include "JVData_Structure.h"


class clsImportHC
{
public:
	clsImportHC(void);
	~clsImportHC(void);
	int Add(CString strBuff, long lngBuffSize);
private:
	int InsertDB(void);
	int UpdateDB(CString strMakeDate);
	JV_HC_HANRO mBuf;
	_ConnectionPtr pCn;
public:
	int Init(_ConnectionPtr m_Connect);
};
