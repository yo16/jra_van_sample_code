#pragma once
#include "JVData_Structure.h"


class clsImportAV
{
public:
	clsImportAV(void);
	~clsImportAV(void);
	int Add(CString strBuff, long lngBuffSize);
private:
	int InsertDB(void);
	int UpdateDB(CString strMakeDate);
	JV_AV_INFO mBuf;
	_ConnectionPtr pCn;
public:
	int Init(_ConnectionPtr m_Connect);
};
