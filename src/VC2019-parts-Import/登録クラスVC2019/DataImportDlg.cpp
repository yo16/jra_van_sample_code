/*=======================================================================
  JRA-VAN Data Lab.�v���O���~���O�p�[�c�u�f�[�^�o�^�p�[�c�@�_�C�A���O�v

	   �쐬: JRA-VAN �\�t�g�E�F�A�H�[  2003�N6��26��

========================================================================
   (C) Copyright JRA SYSTEM SERVICE CO.,LTD. 2007 All rights reserved
========================================================================*/


// DataImportDlg.cpp : �����t�@�C��
//

#include "stdafx.h"
#include "DataImport.h"
#include "DataImportDlg.h"
#include "clsDBImport.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CDataImportDlg �_�C�A���O

	//�L�����Z���t���O
	bool DialogCancel;
	//JVOpen:���Ǎ��݃t�@�C����
    long ReadCount;                     
	//JVOpen:���_�E�����[�h�t�@�C����
    long DownloadCount;  

	//JVOpen:�^�C���X�^���v
	CString strLastFile;
	BSTR bstrLastFile;


CDataImportDlg::CDataImportDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CDataImportDlg::IDD, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
	m_iRadio = 0;
}

void CDataImportDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_JVLINK1, m_jvlink1);
	DDX_Control(pDX, IDC_EDIT3, m_txtCount);
	DDX_Control(pDX, IDC_EDIT2, m_txtFromDate);
	DDX_Control(pDX, IDC_EDIT1, m_txtDataSpec);

}

BEGIN_MESSAGE_MAP(CDataImportDlg, CDialog)
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	//}}AFX_MSG_MAP
	ON_BN_CLICKED(IDC_BUTTON1, OnBnClickedButton1)
	ON_BN_CLICKED(IDC_BUTTON2, OnBnClickedButton2)
	ON_BN_CLICKED(IDOK, OnBnClickedOk)
	ON_EN_CHANGE(IDC_EDIT1, OnEnChangeEdit1)
	ON_BN_CLICKED(IDC_RADIOTUJO, OnBnClickedRadiotujo)
	ON_BN_CLICKED(IDC_RADIOKONSHU, OnBnClickedRadiokonshu)
	ON_BN_CLICKED(IDC_RADIOSETUP, OnBnClickedRadiosetup)
	ON_BN_CLICKED(IDC_RADIOREAL, OnBnClickedRadioreal)
END_MESSAGE_MAP()

// CDataImportDlg ���b�Z�[�W �n���h��

BOOL CDataImportDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	// ���̃_�C�A���O�̃A�C�R����ݒ肵�܂��B�A�v���P�[�V�����̃��C�� �E�B���h�E���_�C�A���O�łȂ��ꍇ�A
	//  Framework �́A���̐ݒ�������I�ɍs���܂��B
	SetIcon(m_hIcon, TRUE);			// �傫���A�C�R���̐ݒ�
	SetIcon(m_hIcon, FALSE);		// �������A�C�R���̐ݒ�

	// TODO: �������������ɒǉ����܂��B
		long ReturnCode;                //JVLink�߂�l
        CString sid;
        sid = "UNKNOWN";               //���� JVInit:�\�t�g�E�F�AID

        //**********************
        //JVLink������
        //**********************
        //������ JVInit�� JVLink���\�b�h�g�p�O�i�A���AJVSetUIProPerties�������j�Ɍďo��
        ReturnCode = m_jvlink1.JVInit(sid);

		//������ɕϊ�
		CString strReturnCode;
		strReturnCode.Format("%d", ReturnCode);
	
	return TRUE;  // �t�H�[�J�X���R���g���[���ɐݒ肵���ꍇ�������ATRUE ��Ԃ��܂��B
}

// �_�C�A���O�ɍŏ����{�^����ǉ�����ꍇ�A�A�C�R����`�悷�邽�߂�
//  ���̃R�[�h���K�v�ł��B�h�L�������g/�r���[ ���f�����g�� MFC �A�v���P�[�V�����̏ꍇ�A
//  ����́AFramework �ɂ���Ď����I�ɐݒ肳��܂��B

void CDataImportDlg::OnPaint() 
{
	if (IsIconic())
	{
		CPaintDC dc(this); // �`��̃f�o�C�X �R���e�L�X�g

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// �N���C�A���g�̎l�p�`�̈���̒���
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// �A�C�R���̕`��
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialog::OnPaint();
	}
}

//���[�U�[���ŏ��������E�B���h�E���h���b�O���Ă���Ƃ��ɕ\������J�[�\�����擾���邽�߂ɁA
//  �V�X�e�������̊֐����Ăяo���܂��B
HCURSOR CDataImportDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

void CDataImportDlg::OnBnClickedButton1()
{
	// TODO : �����ɃR���g���[���ʒm�n���h�� �R�[�h��ǉ����܂��B

		long ReturnCode;					//JVLink�߂�l
		CString DataSpec;
		CString FromDate;
		int DataOption;
	
		//�����l�ݒ�
		DialogCancel=false;					//�L�����Z���t���O������	
		
		m_jvlink1.JVInit("UNKNOWN");


		UpdateData(true);
		m_txtDataSpec.GetWindowText(DataSpec);
		m_txtFromDate.GetWindowText(FromDate);

		clsDBImport mDBCon;
		
		if(MessageBox("�f�[�^���N���A���܂����H",NULL,MB_YESNO)==IDYES){
			mDBCon.ClearData("");
		}


		DataOption = m_iRadio;

		
		//**********************
		//JVLink�_�E�����[�h����
		//**********************
		if(DataOption==100){
			ReturnCode = m_jvlink1.JVRTOpen((LPCTSTR)DataSpec,(LPCTSTR)FromDate);
			DownloadCount=0;
		}else{
			ReturnCode = m_jvlink1.JVOpen((LPCTSTR)DataSpec,
													(LPCTSTR)FromDate,
													DataOption,
													&ReadCount,
													&DownloadCount,
													&bstrLastFile);
		}	
		//�G���[����
		if (ReturnCode != 0) {		   //�G���[
			//������ɕϊ�
			CString strReturnCode;
			strReturnCode.Format("%d", ReturnCode);
			//�I������
			JVClosing();

		}else{							//����

			//������ɕϊ�
			CString strReturnCode;
			CString strDownloadCount;
			CString strReadCount;
			strReturnCode.Format("%d", ReturnCode);
			strDownloadCount.Format("%d", DownloadCount);
			strReadCount.Format("%d", ReadCount);
			

			//�����ݒ�
			SetWindowText("�_�E�����[�h���E�E�E");
			while(m_jvlink1.JVStatus()!=DownloadCount || m_jvlink1.JVStatus()<0){
				Sleep(1000);
			}
			//�Ǎ��ݏ���
			JVReading();
			//�I������
			JVClosing();
			return;
		}
	
}

void CDataImportDlg::OnBnClickedButton2()
{
	// TODO : �����ɃR���g���[���ʒm�n���h�� �R�[�h��ǉ����܂��B
	long ReturnCode;

	//**********************
    //JVLink�ݒ��ʕ\��
    //**********************
	ReturnCode=m_jvlink1.JVSetUIProperties();

	CString strReturnCode;
	strReturnCode.Format("%d", ReturnCode);
}

void CDataImportDlg::OnBnClickedOk()
{
	// TODO : �����ɃR���g���[���ʒm�n���h�� �R�[�h��ǉ����܂��B
	OnOK();
}

void CDataImportDlg::OnEnChangeEdit1()
{
	// TODO :  ���ꂪ RICHEDIT �R���g���[���̏ꍇ�A�܂��ACDialog::OnInitDialog() �֐����I�[�o�[���C�h���āA
	// OR ��Ԃ� ENM_CORRECTTEXT �t���O���}�X�N�ɓ���āA
	// CRichEditCtrl().SetEventMask() ���Ăяo���Ȃ�����A
	// �R���g���[���́A���̒ʒm�𑗐M���܂���B

	// TODO :  �����ɃR���g���[���ʒm�n���h�� �R�[�h��ǉ����Ă��������B
}
//------------------------------------------------------------------------------
//		�I������
//------------------------------------------------------------------------------
void CDataImportDlg::JVClosing()
{
		long ReturnCode;		//JVLink�߂�l

		KillTimer(1);
		::SysFreeString(bstrLastFile);

		//***************
		//JVLink�I������
		//***************
		ReturnCode = m_jvlink1.JVClose();

		//������ɕϊ�
		CString strReturnCode;
		strReturnCode.Format("%d", ReturnCode);

}
//------------------------------------------------------------------------------
//		�Ǎ��ݏ���
//------------------------------------------------------------------------------
void CDataImportDlg::JVReading()
{
		long	ReturnCode; 					//JVLink�߂�l
		long	BuffSize;						//�o�b�t�@�T�C�Y
		int ReturnCodeDB;
		BuffSize = 110000;						//�o�b�t�@�T�C�Y�w��

		CString sBuff;							//�o�b�t�@
		BSTR bBuff;
		CString sBuffName;						//�o�b�t�@��
		BSTR bBuffName;
		CString sLineCount;

		//�o�b�t�@�̈�m��
		sBuff.GetBufferSetLength(BuffSize);
		bBuff=sBuff.AllocSysString();
		sBuffName.GetBuffer(32);
		bBuffName=sBuffName.AllocSysString();
		
		int 	JVReadingCount; 				//�Ǎ��݃t�@�C����
		int		LineCount;
		LineCount=0;

		//������ɕϊ�
		CString strReadCount;
		strReadCount.Format("%d", ReadCount);


		clsDBImport m_DBCon;
		m_DBCon.BeginTrans();
		m_txtCount.SetDlgItemText(IDC_EDIT3,"");
		//�����l
		ReturnCode=0;
		JVReadingCount=1;
		SetWindowText("�f�[�^�Ǎ��ݒ��D�D�D(0/" + strReadCount + ")");
		CString strReturnCode;
		CString strJVReadingCount;
		do {

			PumpMessages();

				//�L�����Z���������ꂽ�珈���𔲂���
				if (DialogCancel==true) return;

				//***************
				//JVLink�Ǎ��ݏ���
				//***************


				ReturnCode =  m_jvlink1.JVRead(&bBuff,&BuffSize,&bBuffName);


				//������ɕϊ�
				strReturnCode.Format("%d", ReturnCode);
				strJVReadingCount.Format("%d", JVReadingCount);
				
				//�G���[����
				if (ReturnCode > 0){		   //����I��

					SetWindowText("�f�[�^�Ǎ��ݒ��D�D�D(" + strJVReadingCount + "/" + strReadCount + ": " + bBuffName + ")");
					sBuff.GetBufferSetLength(ReturnCode);
					sBuff = bBuff;
					//�N���A
					ReturnCodeDB=m_DBCon.SetData(sBuff,sBuff.GetLength());
					if(ReturnCodeDB < 0) break;
					sBuff.Empty();
				}else if (ReturnCode == -1){   //�t�@�C���̐؂��
					//�t�@�C�����\��
					sBuffName.GetBufferSetLength(32);
					sBuffName = bBuffName;
					LineCount = 0;
					//�v���O���X�o�[�\��
					JVReadingCount++; //�J�E���g�A�b�v
					SetWindowText("�f�[�^�Ǎ��ݒ��D�D�D(" + strJVReadingCount + "/" + strReadCount + ")");
					//�N���A
					sBuff.Empty();					
				}else if (ReturnCode == 0){    //�S���R�[�h�Ǎ��ݏI��(EOF)
				strJVReadingCount.Format("%d", JVReadingCount-1);
					SetWindowText("�f�[�^�Ǎ��݊���(" + strJVReadingCount + "/" + strReadCount + ")");
					//Repeat�𔲂���
					break;
				}else if (ReturnCode < -3 ){	//�Ǎ��݃G���[
					//Repeat�𔲂���
				}else if (ReturnCode < -1 ){	//�Ǎ��݃G���[
					//Repeat�𔲂���
					break;
				}
				LineCount++;
				sLineCount.Format("%d",LineCount);
				SetDlgItemText(IDC_EDIT3,sLineCount);
		} while (1);
		if(ReturnCodeDB==0){
			m_DBCon.CommitTrans();
		}else if(ReturnCodeDB == -1){
			m_DBCon.RollbackTrans();
		}

		::SysFreeString(bBuff);
		::SysFreeString(bBuffName);
		sBuff.Empty();
		sBuffName.Empty();
}

//------------------------------------------------------------------------------
//		�o�b�N�O���E���h����
//------------------------------------------------------------------------------
void CDataImportDlg::PumpMessages()
{
		MSG msg;
		while (::PeekMessage(&msg, NULL, 0, 0, PM_REMOVE)) {
			   ::TranslateMessage(&msg);
			   ::DispatchMessage (&msg);
		}
}

void CDataImportDlg::OnBnClickedRadiotujo()
{
	// TODO : �����ɃR���g���[���ʒm�n���h�� �R�[�h��ǉ����܂��B
	m_iRadio = 1;
}

void CDataImportDlg::OnBnClickedRadiokonshu()
{
	// TODO : �����ɃR���g���[���ʒm�n���h�� �R�[�h��ǉ����܂��B
	m_iRadio = 2;
}

void CDataImportDlg::OnBnClickedRadiosetup()
{
	// TODO : �����ɃR���g���[���ʒm�n���h�� �R�[�h��ǉ����܂��B
	m_iRadio = 3;
}

void CDataImportDlg::OnBnClickedRadioreal()
{
	// TODO : �����ɃR���g���[���ʒm�n���h�� �R�[�h��ǉ����܂��B
	m_iRadio = 100;
}
