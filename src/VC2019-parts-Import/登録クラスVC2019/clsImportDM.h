#pragma once
#include "JVData_Structure.h"


class clsImportDM
{
public:
	clsImportDM(void);
	~clsImportDM(void);
	int Add(CString strBuff, long lngBuffSize);
private:
	int InsertDB(void);
	int UpdateDB(CString strMakeDate);
	JV_DM_INFO mBuf;
	_ConnectionPtr pCn;
public:
	int Init(_ConnectionPtr m_Connect);
};
