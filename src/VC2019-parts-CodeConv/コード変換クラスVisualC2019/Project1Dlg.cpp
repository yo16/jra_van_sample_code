// Project1Dlg.cpp : �C���v�������e�[�V���� �t�@�C��
//

#include "stdafx.h"
#include "Project1.h"
#include "Project1Dlg.h"
#include "DlgProxy.h"
#include "CodeCnv.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// �A�v���P�[�V�����̃o�[�W�������Ŏg���Ă��� CAboutDlg �_�C�A���O

class CAboutDlg : public CDialog
{
public:
	CAboutDlg();

// �_�C�A���O �f�[�^
	//{{AFX_DATA(CAboutDlg)
	enum { IDD = IDD_ABOUTBOX };
	//}}AFX_DATA

	// ClassWizard �͉��z�֐��̃I�[�o�[���C�h�𐶐����܂�
	//{{AFX_VIRTUAL(CAboutDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV �̃T�|�[�g
	//}}AFX_VIRTUAL

// �C���v�������e�[�V����
protected:
	//{{AFX_MSG(CAboutDlg)
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialog(CAboutDlg::IDD)
{
	//{{AFX_DATA_INIT(CAboutDlg)
	//}}AFX_DATA_INIT
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CAboutDlg)
	//}}AFX_DATA_MAP
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialog)
	//{{AFX_MSG_MAP(CAboutDlg)
		// ���b�Z�[�W �n���h��������܂���B
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CProject1Dlg �_�C�A���O

IMPLEMENT_DYNAMIC(CProject1Dlg, CDialog);

CProject1Dlg::CProject1Dlg(CWnd* pParent /*=NULL*/)
	: CDialog(CProject1Dlg::IDD, pParent)
{
	//{{AFX_DATA_INIT(CProject1Dlg)
	m_eCode = _T("");
	m_eValue = _T("");
	m_eNum = _T("");
	m_eOut = _T("");
	//}}AFX_DATA_INIT
	// ����: LoadIcon �� Win32 �� DestroyIcon �̃T�u�V�[�P���X��v�����܂���B
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
	m_pAutoProxy = NULL;
}

CProject1Dlg::~CProject1Dlg()
{
	// ���̃_�C�A���O�p�̃I�[�g���[�V���� �v���L�V������ꍇ�́A���̃_�C�A���O
	// �ւ̃|�C���^�� NULL �ɖ߂��܂��A����ɂ���ă_�C�A���O���폜���ꂽ����
	// ���킩��܂��B
	if (m_pAutoProxy != NULL)
		m_pAutoProxy->m_pDialog = NULL;
}

void CProject1Dlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CProject1Dlg)
	DDX_Text(pDX, IDC_EDIT1, m_eCode);
	DDX_Text(pDX, IDC_EDIT2, m_eValue);
	DDX_Text(pDX, IDC_EDIT3, m_eNum);
	DDX_Text(pDX, IDC_EDIT4, m_eOut);
	//}}AFX_DATA_MAP
}

BEGIN_MESSAGE_MAP(CProject1Dlg, CDialog)
	//{{AFX_MSG_MAP(CProject1Dlg)
	ON_WM_SYSCOMMAND()
	ON_WM_CLOSE()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CProject1Dlg ���b�Z�[�W �n���h��

BOOL CProject1Dlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	// "�o�[�W�������..." ���j���[���ڂ��V�X�e�� ���j���[�֒ǉ����܂��B

	// IDM_ABOUTBOX �̓R�}���h ���j���[�͈̔͂łȂ���΂Ȃ�܂���B
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		CString strAboutMenu;
		strAboutMenu.LoadString(IDS_ABOUTBOX);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// ���̃_�C�A���O�p�̃A�C�R����ݒ肵�܂��B�t���[�����[�N�̓A�v���P�[�V�����̃��C��
	// �E�B���h�E���_�C�A���O�łȂ����͎����I�ɐݒ肵�܂���B
	SetIcon(m_hIcon, TRUE);			// �傫���A�C�R����ݒ�
	SetIcon(m_hIcon, FALSE);		// �������A�C�R����ݒ�
	
	// TODO: ���ʂȏ��������s�����͂��̏ꏊ�ɒǉ����Ă��������B
	
	return TRUE;  // TRUE ��Ԃ��ƃR���g���[���ɐݒ肵���t�H�[�J�X�͎����܂���B
}

void CProject1Dlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialog::OnSysCommand(nID, lParam);
	}
}

// �����_�C�A���O�{�b�N�X�ɍŏ����{�^����ǉ�����Ȃ�΁A�A�C�R����`�悷��
// �R�[�h���ȉ��ɋL�q����K�v������܂��BMFC �A�v���P�[�V������ document/view
// ���f�����g���Ă���̂ŁA���̏����̓t���[�����[�N�ɂ�莩���I�ɏ�������܂��B

void CProject1Dlg::OnPaint() 
{
	if (IsIconic())
	{
		CPaintDC dc(this); // �`��p�̃f�o�C�X �R���e�L�X�g

		SendMessage(WM_ICONERASEBKGND, (WPARAM) dc.GetSafeHdc(), 0);

		// �N���C�A���g�̋�`�̈���̒���
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// �A�C�R����`�悵�܂��B
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialog::OnPaint();
	}
}

// �V�X�e���́A���[�U�[���ŏ����E�B���h�E���h���b�O���Ă���ԁA
// �J�[�\����\�����邽�߂ɂ������Ăяo���܂��B
HCURSOR CProject1Dlg::OnQueryDragIcon()
{
	return (HCURSOR) m_hIcon;
}

// �R���g���[�����I�u�W�F�N�g�� 1 �����܂��ێ����Ă���ꍇ�A
// �I�[�g���[�V�����T�[�o�[�̓��[�U�[�� UI �����ۂɏI����
// ���܂���B�����̃��b�Z�[�W �n���h���̓v���L�V���܂��g�p��
// ���ǂ������m�F���A���ꂩ�� UI ����\���ɂȂ�܂����_�C�A��
// �O�͂��ꂪ�����ꂽ�ꍇ���̏ꏊ�Ɏc��܂��B

void CProject1Dlg::OnClose() 
{
	if (CanExit())
		CDialog::OnClose();
}

void CProject1Dlg::OnOK() 
{
	char currentpath[MAX_PATH];
	CString strCSVPath;

	//�C���X�^���X����
	CCodeCnv *objCodeCnv = new CCodeCnv;
	
	//�R�[�h�\�̃p�X���擾
	GetCurrentDirectory(MAX_PATH,currentpath);
	strCSVPath = currentpath;
	strCSVPath += "\\CodeTable.csv";

	objCodeCnv->FileName(strCSVPath);

	UpdateData(TRUE);//�R���g���[��>>�����o�ϐ�
	m_eOut=objCodeCnv->GetCodeName(m_eCode,m_eValue,atoi(m_eNum));
	UpdateData(FALSE);//�����o�ϐ�>>�R���g���[��

	//�j��
	delete objCodeCnv;

	
}

void CProject1Dlg::OnCancel() 
{
	if (CanExit())
		CDialog::OnCancel();
}

BOOL CProject1Dlg::CanExit()
{
	// �v���L�V �I�u�W�F�N�g���܂��c���Ă���ꍇ�A�I�[�g���[�V����	
	// �R���g���[���͂��̃A�v���P�[�V�������܂��ێ����Ă��܂��B 
	// �_�C�A���O�̎��͎͂c���܂��� UI �͔�\���ɂȂ�܂��B
	if (m_pAutoProxy != NULL)
	{
		ShowWindow(SW_HIDE);
		return FALSE;
	}

	return TRUE;
}
