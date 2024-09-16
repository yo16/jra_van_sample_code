#pragma once
#include "JVData_Structure.h"

class clsImportTK
{
public:
	clsImportTK(void);
	~clsImportTK(void);
	int Init(_ConnectionPtr m_Connect);
	int Add(CString strBuff, long lngBuffSize);
	JV_TK_TOKUUMA mBuf;
private:
	int InsertDB(void);
	int UpdateDB(CString strMakeDate);
	_RecordsetPtr mRS1;  // Recordset�I�u�W�F�N�g
	_RecordsetPtr mRS2;  // Recordset�I�u�W�F�N�g
	_ConnectionPtr pCn;

};

//This Class extracts empid, fname and lastname  
