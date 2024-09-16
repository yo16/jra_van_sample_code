#pragma once
#include "JVData_Structure.h"


class clsImportWH
{
public:
	clsImportWH(void);
	~clsImportWH(void);
	int Add(CString strBuff, long lngBuffSize);
private:
	int InsertDB(void);
	int UpdateDB(CString strMakeDate);
	JV_WH_BATAIJYU mBuf;
	_ConnectionPtr pCn;
public:
	int Init(_ConnectionPtr m_Connect);
};
