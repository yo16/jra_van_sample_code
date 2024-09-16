#pragma once
#include "JVData_Structure.h"


class clsImportYS
{
public:
	clsImportYS(void);
	~clsImportYS(void);
	int Add(CString strBuff, long lngBuffSize);
private:
	int InsertDB(void);
	int UpdateDB(CString strMakeDate);
	JV_YS_SCHEDULE mBuf;
	_ConnectionPtr pCn;
public:
	int Init(_ConnectionPtr m_Connect);
};
