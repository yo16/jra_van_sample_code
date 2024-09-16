#pragma once
#include "JVData_Structure.h"


class clsImportSK
{
public:
	clsImportSK(void);
	~clsImportSK(void);
	int Add(CString strBuff, long lngBuffSize);
private:
	int InsertDB(void);
	int UpdateDB(CString strMakeDate);
	JV_SK_SANKU mBuf;
	_ConnectionPtr pCn;
public:
	int Init(_ConnectionPtr m_Connect);
};
