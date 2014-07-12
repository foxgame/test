
// excel_xmlDlg.cpp : 实现文件
//

#include "stdafx.h"
#include "excel_xml.h"
#include "excel_xmlDlg.h"
#include "afxdialogex.h"
#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// 用于应用程序“关于”菜单项的 CAboutDlg 对话框

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// 对话框数据
	enum { IDD = IDD_ABOUTBOX };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

// 实现
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


// Cexcel_xmlDlg 对话框



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


// Cexcel_xmlDlg 消息处理程序

BOOL Cexcel_xmlDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// 将“关于...”菜单项添加到系统菜单中。

	// IDM_ABOUTBOX 必须在系统命令范围内。
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

	// 设置此对话框的图标。  当应用程序主窗口不是对话框时，框架将自动
	//  执行此操作
	SetIcon(m_hIcon, TRUE);			// 设置大图标
	SetIcon(m_hIcon, FALSE);		// 设置小图标

	// TODO:  在此添加额外的初始化代码

	return TRUE;  // 除非将焦点设置到控件，否则返回 TRUE
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

// 如果向对话框添加最小化按钮，则需要下面的代码
//  来绘制该图标。  对于使用文档/视图模型的 MFC 应用程序，
//  这将由框架自动完成。

void Cexcel_xmlDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // 用于绘制的设备上下文

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// 使图标在工作区矩形中居中
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// 绘制图标
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

//当用户拖动最小化窗口时系统调用此函数取得光标
//显示。
HCURSOR Cexcel_xmlDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}



void Cexcel_xmlDlg::OnBnClickedOk()
{
	// TODO:  在此添加控件通知处理程序代码

	Workbooks books;
	_Workbook book;
	Worksheets sheets;
	_Worksheet sheet;
	Range range;
	_Application app;

	CString excelPath;
	CString notFind ;
	notFind = "没有找到excel文件!";
	CString haveFind;
	haveFind = "123.xlsm";
	excelPath = "C:\\Users\\Administrator\\Desktop\\123.xml";
	CFileFind filefind;
	COleVariant covOptional((long)DISP_E_PARAMNOTFOUND,VT_ERROR);
	if (CoInitialize(NULL) == E_INVALIDARG)
	{
		AfxMessageBox(_T("初始化失败！"));
		return;
	};
	if (!app.CreateDispatch(_T("Excel.Application")))
	{
		AfxMessageBox(_T("无法创建excel!"));
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
		AfxMessageBox(book.GetName()+_T("不是excel文件！"));
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


	// TODO:  在此添加控件通知处理程序代码
}
