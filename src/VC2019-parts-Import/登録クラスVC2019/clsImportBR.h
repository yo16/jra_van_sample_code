#pragma once
#include "JVData_Structure.h"


class clsImportBR
{
public:
	clsImportBR(void);
	~clsImportBR(void);
	int Add(CString strBuff, long lngBuffSize);
private:
	int InsertDB(void);
	int UpdateDB(CString strMakeDate);
	JV_BR_BREEDER mBuf;
	_ConnectionPtr pCn;
public:
	int Init(_ConnectionPtr m_Connect);
};
