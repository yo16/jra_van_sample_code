// CodeCnv.h: CCodeCnv クラスのインターフェイス
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_CODECNV_H__4715F9C3_38C5_47B1_B93F_6A809A867A38__INCLUDED_)
#define AFX_CODECNV_H__4715F9C3_38C5_47B1_B93F_6A809A867A38__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000



class CCodeCnv  
{
private:

public:
	CCodeCnv();
	virtual ~CCodeCnv();
	void FileName(CString strFile);
	CString GetCodeName(CString strCodeNo, CString strCode, int intNo);
	int SetData();
	int SetLine(CString strLine, long lngCt);
};


#endif // !defined(AFX_CODECNV_H__4715F9C3_38C5_47B1_B93F_6A809A867A38__INCLUDED_)
