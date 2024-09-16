#pragma once
#include "JVData_Structure.h"


class clsImportCC
{
public:
	clsImportCC(void);
	~clsImportCC(void);
	int Add(CString strBuff, long lngBuffSize);
private:
	int InsertDB(void);
	int UpdateDB(CString strMakeDate);
	JV_CC_INFO mBuf;
	_ConnectionPtr pCn;
public:
	int Init(_ConnectionPtr m_Connect);
};
