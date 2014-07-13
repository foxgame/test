
// excel_xmlDlg.cpp : ʵ���ļ�
//

#include "stdafx.h"
#include "excel_xml.h"
#include "excel_xmlDlg.h"
#include "afxdialogex.h"
#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// ����Ӧ�ó��򡰹��ڡ��˵���� CAboutDlg �Ի���

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// �Ի�������
	enum { IDD = IDD_ABOUTBOX };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV ֧��

// ʵ��
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialogEx(CAboutDlg::IDD)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialogEx)
END_MESSAGE_MAP()


// Cexcel_xmlDlg �Ի���



Cexcel_xmlDlg::Cexcel_xmlDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(Cexcel_xmlDlg::IDD, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void Cexcel_xmlDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(Cexcel_xmlDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDOK, &Cexcel_xmlDlg::OnBnClickedOk)
	ON_LBN_SELCHANGE(IDC_LIST1, &Cexcel_xmlDlg::OnLbnSelchangeList1)
END_MESSAGE_MAP()


// Cexcel_xmlDlg ��Ϣ�������

BOOL Cexcel_xmlDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// ��������...���˵�����ӵ�ϵͳ�˵��С�

	// IDM_ABOUTBOX ������ϵͳ���Χ�ڡ�
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		BOOL bNameValid;
		CString strAboutMenu;
		bNameValid = strAboutMenu.LoadString(IDS_ABOUTBOX);
		ASSERT(bNameValid);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// ���ô˶Ի����ͼ�ꡣ  ��Ӧ�ó��������ڲ��ǶԻ���ʱ����ܽ��Զ�
	//  ִ�д˲���
	SetIcon(m_hIcon, TRUE);			// ���ô�ͼ��
	SetIcon(m_hIcon, FALSE);		// ����Сͼ��

	// TODO:  �ڴ���Ӷ���ĳ�ʼ������

	return TRUE;  // ���ǽ��������õ��ؼ������򷵻� TRUE
}

void Cexcel_xmlDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialogEx::OnSysCommand(nID, lParam);
	}
}

// �����Ի��������С����ť������Ҫ����Ĵ���
//  �����Ƹ�ͼ�ꡣ  ����ʹ���ĵ�/��ͼģ�͵� MFC Ӧ�ó���
//  �⽫�ɿ���Զ���ɡ�

void Cexcel_xmlDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // ���ڻ��Ƶ��豸������

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// ʹͼ���ڹ����������о���
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// ����ͼ��
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

//���û��϶���С������ʱϵͳ���ô˺���ȡ�ù��
//��ʾ��
HCURSOR Cexcel_xmlDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}



void Cexcel_xmlDlg::OnBnClickedOk()
{
	// TODO:  �ڴ���ӿؼ�֪ͨ����������

	Workbooks books;
	_Workbook book;
	Worksheets sheets;
	_Worksheet sheet;
	Range range;
	_Application app;

	CString excelPath;
	CString notFind ;
	notFind = "û���ҵ�excel�ļ�!";
	CString haveFind;
	haveFind = "123.xlsm";
	excelPath = "C:\\Users\\Administrator\\Desktop\\123.xml";
	CFileFind filefind;
	COleVariant covOptional((long)DISP_E_PARAMNOTFOUND,VT_ERROR);
	if (CoInitialize(NULL) == E_INVALIDARG)
	{
		AfxMessageBox(_T("��ʼ��ʧ�ܣ�"));
		return;
	};
	if (!app.CreateDispatch(_T("Excel.Application")))
	{
		AfxMessageBox(_T("�޷�����excel!"));
		return;
	}





	if (!filefind.FindFile(excelPath))
	{
		AfxMessageBox(notFind);
		return;
	}
	else

	{
		//AfxMessageBox(haveFind);
	}

	LPDISPATCH lpDisp=NULL;
	books = app.GetWorkbooks();
	lpDisp = books.Open(excelPath, 
		covOptional, 
		covOptional, 
		covOptional, 
		covOptional,
		covOptional,
		covOptional,
		covOptional,
		covOptional, 
		covOptional, 
		covOptional,
		covOptional, 
		covOptional,
		covOptional,
		covOptional);
	book.AttachDispatch(lpDisp);
	if (!isExcel(book.GetName()))
	{
		AfxMessageBox(book.GetName()+_T("����excel�ļ���"));
		return;
	}
	else
	{
		AfxMessageBox(book.GetName());
	}
	sheets = book.GetSheets();
	sheet = sheets.GetItem(COleVariant((short)1));
	COleVariant tempValue;
	list1 = (CListBox*)GetDlgItem(IDC_LIST1);


	/////
	int temp1=0;
	for (int c = 1; c <= 20; ++c)
	{

		range = sheet.GetRange(COleVariant(indexToString(c, 1)), COleVariant(indexToString(c, 1)));
		tempValue = COleVariant(range.GetValue2());
		tempValue.ChangeType(VT_BSTR);
		if (temp1 >= 9)
		{
			list1->InsertString(c-1, CString(tempValue.bstrVal));
		}
		else
		{
			list1->AddString(CString(tempValue.bstrVal));
		}
		
		temp1++;
	};
	




	//AfxMessageBox(CString(tempValue.bstrVal));

	
	books.ReleaseDispatch();
	sheets.ReleaseDispatch();
	sheet.ReleaseDispatch();
	book.ReleaseDispatch();
	range.ReleaseDispatch();
	app.Quit();

	 //CDialogEx::OnOK();
}
bool Cexcel_xmlDlg::isExcel(CString name)
{
	int n;
	n=name.Find(_T("."));
	CString temp;
	temp = name.Mid(n+1, 10);
	if (temp ==_T("xls") || temp == _T("xlsm") || temp == _T("xlsx"))
		return true;
	else
		return false;
}
CString Cexcel_xmlDlg::indexToString(int row, int col)
{
	CString temp;
	if (col > 26)
	{
		temp.Format( _T("%c%c%d"), 'A' + (col - 1) / 26 - 1, 'A' + (col - 1) % 26, row);
	}
	else
	{
		temp.Format( _T("%c%d"), 'A' + (col - 1)%26, row);
	}
	return temp;
}

void Cexcel_xmlDlg::OnLbnSelchangeList1()
{


	// TODO:  �ڴ���ӿؼ�֪ͨ����������
}
