#pragma once
#include "JVData_Structure.h"


class clsImportCS
{
public:
	clsImportCS(void);
	~clsImportCS(void);
	int Add(CString strBuff, long lngBuffSize);
private:
	int InsertDB(void);
	int UpdateDB(CString strMakeDate);
	JV_CS_COURSE mBuf;
	_ConnectionPtr pCn;
public:
	int Init(_ConnectionPtr m_Connect);
};
