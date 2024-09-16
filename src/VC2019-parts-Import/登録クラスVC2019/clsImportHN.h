#pragma once
#include "JVData_Structure.h"


class clsImportHN
{
public:
	clsImportHN(void);
	~clsImportHN(void);
	int Add(CString strBuff, long lngBuffSize);
private:
	int InsertDB(void);
	int UpdateDB(CString strMakeDate);
	JV_HN_HANSYOKU mBuf;
	_ConnectionPtr pCn;
public:
	int Init(_ConnectionPtr m_Connect);
};
