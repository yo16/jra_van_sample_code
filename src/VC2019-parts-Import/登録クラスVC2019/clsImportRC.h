#pragma once
#include "JVData_Structure.h"


class clsImportRC
{
public:
	clsImportRC(void);
	~clsImportRC(void);
	int Add(CString strBuff, long lngBuffSize);
private:
	int InsertDB(void);
	int UpdateDB(CString strMakeDate);
	JV_RC_RECORD mBuf;
	_ConnectionPtr pCn;
public:
	int Init(_ConnectionPtr m_Connect);
};
