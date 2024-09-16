#pragma once
#include "JVData_Structure.h"


class clsImportUM
{
public:
	clsImportUM(void);
	~clsImportUM(void);
	int Add(CString strBuff, long lngBuffSize);
private:
	int InsertDB(void);
	int UpdateDB(CString strMakeDate);
	JV_UM_UMA mBuf;
	_ConnectionPtr pCn;
public:
	int Init(_ConnectionPtr m_Connect);
};
