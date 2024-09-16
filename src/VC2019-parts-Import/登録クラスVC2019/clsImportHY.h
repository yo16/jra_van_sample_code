#pragma once
#include "JVData_Structure.h"


class clsImportHY
{
public:
	clsImportHY(void);
	~clsImportHY(void);
	int Add(CString strBuff, long lngBuffSize);
private:
	int InsertDB(void);
	int UpdateDB(CString strMakeDate);
	JV_HY_BAMEIORIGIN mBuf;
	_ConnectionPtr pCn;
public:
	int Init(_ConnectionPtr m_Connect);
};
