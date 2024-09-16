#pragma once
#include "JVData_Structure.h"


class clsImportO1
{
public:
	clsImportO1(void);
	~clsImportO1(void);
	int Add(CString strBuff, long lngBuffSize);
private:
	int InsertDB(void);
	int UpdateDB(CString strMakeDate);
	JV_O1_ODDS_TANFUKUWAKU mBuf;
	_ConnectionPtr pCn;
public:
	int Init(_ConnectionPtr m_Connect);
};
