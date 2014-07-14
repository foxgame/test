
// excel_xmlDlg.cpp : ʵ���ļ�
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
	ON_EN_CHANGE(IDC_EDIT3, &Cexcel_xmlDlg::OnEnChangeEdit1)
	ON_EN_CHANGE(IDC_EDIT2, &Cexcel_xmlDlg::OnEnChangeEdit2)
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

	string path = "C:\\Users\\Administrator\\Desktop\\�½��ļ��� (4)";
	CString path3;
	path3 = "C:\\Users\\Administrator\\Desktop\\�½��ļ��� (4)";
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


	// TODO:  �ڴ���ӿؼ�֪ͨ����������
}


void Cexcel_xmlDlg::OnEnChangeEdit1()
{
	// TODO:  ����ÿؼ��� RICHEDIT �ؼ���������
	// ���ʹ�֪ͨ��������д CDialogEx::OnInitDialog()
	// ���������� CRichEditCtrl().SetEventMask()��
	// ͬʱ�� ENM_CHANGE ��־�������㵽�����С�

	// TODO:  �ڴ���ӿؼ�֪ͨ����������
}


void Cexcel_xmlDlg::OnEnChangeEdit2()
{
	// TODO:  ����ÿؼ��� RICHEDIT �ؼ���������
	// ���ʹ�֪ͨ��������д CDialogEx::OnInitDialog()
	// ���������� CRichEditCtrl().SetEventMask()��
	// ͬʱ�� ENM_CHANGE ��־�������㵽�����С�

	// TODO:  �ڴ���ӿؼ�֪ͨ����������
}
