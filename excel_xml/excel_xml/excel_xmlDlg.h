
// excel_xmlDlg.h : ͷ�ļ�
//

#pragma once


// Cexcel_xmlDlg �Ի���
class Cexcel_xmlDlg : public CDialogEx
{
// ����
public:
	Cexcel_xmlDlg(CWnd* pParent = NULL);	// ��׼���캯��

// �Ի�������
	enum { IDD = IDD_EXCEL_XML_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV ֧��


// ʵ��
protected:
	HICON m_hIcon;

	// ���ɵ���Ϣӳ�亯��
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
};
