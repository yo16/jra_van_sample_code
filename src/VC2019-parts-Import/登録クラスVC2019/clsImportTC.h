#pragma once
#include "JVData_Structure.h"


class clsImportTC
{
public:
	clsImportTC(void);
	~clsImportTC(void);
	int Add(CString strBuff, long lngBuffSize);
private:
	int InsertDB(void);
	int UpdateDB(CString strMakeDate);
	JV_TC_INFO mBuf;
	_ConnectionPtr pCn;
public:
	int Init(_ConnectionPtr m_Connect);
};
