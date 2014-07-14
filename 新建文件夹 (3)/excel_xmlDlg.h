
// excel_xmlDlg.h : 头文件
//

#pragma once
// Cexcel_xmlDlg 对话框
class Cexcel_xmlDlg : public CDialogEx
{
// 构造
public:
	Cexcel_xmlDlg(CWnd* pParent = NULL);	// 标准构造函数

// 对话框数据
	enum { IDD = IDD_EXCEL_XML_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV 支持


// 实现
protected:
	HICON m_hIcon;

	// 生成的消息映射函数
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedOk();
	afx_msg void OnLbnSelchangeList1();
private:
	CListBox *list1;
	CString indexToString(int row, int col);
	bool isExcel(CString);
public:
	afx_msg void OnEnChangeEdit1();
	afx_msg void OnEnChangeEdit2();
};
