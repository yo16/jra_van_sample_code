#pragma once
#include "JVData_Structure.h"


class clsImportBN
{
public:
	clsImportBN(void);
	~clsImportBN(void);
	int Add(CString strBuff, long lngBuffSize);
private:
	int InsertDB(void);
	int UpdateDB(CString strMakeDate);
	JV_BN_BANUSI mBuf;
	_ConnectionPtr pCn;
public:
	int Init(_ConnectionPtr m_Connect);
};
