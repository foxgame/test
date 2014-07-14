
// excel_xmlDlg.cpp : 实现文件
//

#include "stdafx.h"
#include "excel_xml.h"
#include "JDExcel.h"
#include "excel_xmlDlg.h"
#include "afxdialogex.h"
#include "JDfindFile.h"
#include <string>
#include <iostream>
#ifdef _DEBUG
#define new DEBUG_NEW
#endif
char* UnicodeToUtf8(CString unicode);
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
	ON_EN_CHANGE(IDC_EDIT3, &Cexcel_xmlDlg::OnEnChangeEdit1)
	ON_EN_CHANGE(IDC_EDIT2, &Cexcel_xmlDlg::OnEnChangeEdit2)
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

	string path = "C:\\Users\\Administrator\\Desktop\\新建文件夹 (4)";
	CString path3;
	path3 = "C:\\Users\\Administrator\\Desktop\\新建文件夹 (4)";
	list1 = (CListBox*)GetDlgItem(IDC_LIST1);
	CEdit* pathF = (CEdit*)GetDlgItem(IDC_EDIT2);
	CString temp1;
	pathF->GetWindowText(temp1);
	CString temp2;
	CEdit* pathB = (CEdit*)GetDlgItem(IDC_EDIT3);
	pathB->GetWindowText(temp2);
	JDExcel test;
	JDfindFile files(path+"\\*.*", list1);
	for (int i = 1; i <= files.cout; ++i)
	{
		path3 = path3 + "\\" + files.p[i];
		if (!test.openExcelBook(path3))
			return;

		temp1 += "\\";
		temp1 += files.getFileName(i);
		temp1 += ".xml";
		test.writeXml(temp1);
		//test.writeXml("C:\\Users\\Administrator\\Desktop\\1\\123.xml");
		test.saveExcel();
		list1->AddString(files.p[i]);
		list1->AddString(files.getFileName(i) + ".xml");
	}
	//CDialogEx::OnOK();
}

 char* UnicodeToUtf8(CString unicode)
{
	int len;
	len = WideCharToMultiByte(CP_UTF8, 0, (LPCWSTR)unicode, -1, NULL, 0, NULL, NULL);
	char *szUtf8 = new char[len + 1];
	memset(szUtf8, 0, len * 2 + 2);
	WideCharToMultiByte(CP_UTF8, 0, (LPCWSTR)unicode, -1, szUtf8, len, NULL, NULL);
	return szUtf8;
}




void Cexcel_xmlDlg::OnLbnSelchangeList1()
{


	// TODO:  在此添加控件通知处理程序代码
}


void Cexcel_xmlDlg::OnEnChangeEdit1()
{
	// TODO:  如果该控件是 RICHEDIT 控件，它将不
	// 发送此通知，除非重写 CDialogEx::OnInitDialog()
	// 函数并调用 CRichEditCtrl().SetEventMask()，
	// 同时将 ENM_CHANGE 标志“或”运算到掩码中。

	// TODO:  在此添加控件通知处理程序代码
}


void Cexcel_xmlDlg::OnEnChangeEdit2()
{
	// TODO:  如果该控件是 RICHEDIT 控件，它将不
	// 发送此通知，除非重写 CDialogEx::OnInitDialog()
	// 函数并调用 CRichEditCtrl().SetEventMask()，
	// 同时将 ENM_CHANGE 标志“或”运算到掩码中。

	// TODO:  在此添加控件通知处理程序代码
}
