#pragma once
#include "JVData_Structure.h"


class clsImportJC
{
public:
	clsImportJC(void);
	~clsImportJC(void);
	int Add(CString strBuff, long lngBuffSize);
private:
	int InsertDB(void);
	int UpdateDB(CString strMakeDate);
	JV_JC_INFO mBuf;
	_ConnectionPtr pCn;
public:
	int Init(_ConnectionPtr m_Connect);
};
