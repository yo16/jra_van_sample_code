#pragma once
#include "JVData_Structure.h"


class clsImportH1
{
public:
	clsImportH1(void);
	~clsImportH1(void);
	int Add(CString strBuff, long lngBuffSize);
private:
	int InsertDB(void);
	int UpdateDB(CString strMakeDate);
	JV_H1_HYOSU_ZENKAKE mBuf;
	_ConnectionPtr pCn;
public:
	int Init(_ConnectionPtr m_Connect);
};
