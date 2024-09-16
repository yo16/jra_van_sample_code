#pragma once
#include "JVData_Structure.h"


class clsImportH6
{
public:
	clsImportH6(void);
	~clsImportH6(void);
	int Add(CString strBuff, long lngBuffSize);
private:
	int InsertDB(void);
	int UpdateDB(CString strMakeDate);
	JV_H6_HYOSU_SANRENTAN mBuf;
	_ConnectionPtr pCn;
public:
	int Init(_ConnectionPtr m_Connect);
};
