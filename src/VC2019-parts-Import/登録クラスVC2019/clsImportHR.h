#pragma once
#include "JVData_Structure.h"

class clsImportHR
{
public:
	clsImportHR(void);
	~clsImportHR(void);
	int Add(CString strBuff, long lngBuffSize);
private:
	int InsertDB(void);
	int UpdateDB(CString strMakeDate);
	JV_HR_PAY mBuf;
	_ConnectionPtr pCn;
public:
	int Init(_ConnectionPtr m_Connect);
};
