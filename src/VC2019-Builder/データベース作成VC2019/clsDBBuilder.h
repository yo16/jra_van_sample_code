#pragma once

class clsDBBuilder
{
public:
	clsDBBuilder(void);
	~clsDBBuilder(void);
private:
	_ConnectionPtr pCn;
public:
	int CreateDB(CString strFilePath);
	int CompactDB(CString strFilePath);
	int KillDB(CString strFilePath);
};
