#pragma once
#include "JVData_Structure.h"


class clsImportHS
{
public:
	clsImportHS(void);
	~clsImportHS(void);
	int Add(CString strBuff, long lngBuffSize);
private:
	int InsertDB(void);
	int UpdateDB(CString strMakeDate);
	JV_HS_SALE mBuf;
	_ConnectionPtr pCn;
public:
	int Init(_ConnectionPtr m_Connect);
};
