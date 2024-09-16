#pragma once
#include "JVData_Structure.h"


class clsImportCH
{
public:
	clsImportCH(void);
	~clsImportCH(void);
	int Add(CString strBuff, long lngBuffSize);
private:
	int InsertDB(void);
	int UpdateDB(CString strMakeDate);
	JV_CH_CHOKYOSI mBuf;
	_ConnectionPtr pCn;
public:
	int Init(_ConnectionPtr m_Connect);
};
