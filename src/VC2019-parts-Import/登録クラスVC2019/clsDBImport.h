#pragma once

class clsDBImport
{
public:
	clsDBImport(void);
	~clsDBImport(void);
	int ClearData(LPCTSTR   strTBLName);
	int SetData(CString strBuff, long lngBuffSize);
private:
	_ConnectionPtr pCn;
public:
	void BeginTrans(void);
	void CommitTrans(void);
	void RollbackTrans(void);
};
