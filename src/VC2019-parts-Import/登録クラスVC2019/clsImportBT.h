#pragma once
#include "JVData_Structure.h"


class clsImportBT
{
public:
	clsImportBT(void);
	~clsImportBT(void);
	int Add(CString strBuff, long lngBuffSize);
private:
	int InsertDB(void);
	int UpdateDB(CString strMakeDate);
	JV_BT_KEITO mBuf;
	_ConnectionPtr pCn;
public:
	int Init(_ConnectionPtr m_Connect);
};
