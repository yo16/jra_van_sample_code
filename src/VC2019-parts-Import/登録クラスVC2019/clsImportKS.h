#pragma once
#include "JVData_Structure.h"


class clsImportKS
{
public:
	clsImportKS(void);
	~clsImportKS(void);
	int Add(CString strBuff, long lngBuffSize);
private:
	int InsertDB(void);
	int UpdateDB(CString strMakeDate);
	JV_KS_KISYU mBuf;
	_ConnectionPtr pCn;
public:
	int Init(_ConnectionPtr m_Connect);
};
