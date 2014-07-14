#include "stdafx.h"
#include "JDExcel.h"
#include <string>
#include <iostream>
#include "tinyxml.h"
_Application app;
Workbooks  books;
_Workbook book;
Worksheets sheets;
_Worksheet sheet;
Range range;
COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);

JDExcel::JDExcel()
{
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
}
JDExcel::~JDExcel()
{
	books.ReleaseDispatch();
	sheets.ReleaseDispatch();
	sheet.ReleaseDispatch();
	book.ReleaseDispatch();
	range.ReleaseDispatch();
	app.Quit();
	app.ReleaseDispatch();
	::CoUninitialize();
}
void JDExcel::createExcelBook()
{
	books = app.GetWorkbooks();
	book = books.Add(covOptional);
	sheets = book.GetSheets();
	sheet = sheets.GetItem(COleVariant((short)1));
}
bool JDExcel::openExcelBook(CString filename)
{
	CFileFind filefind;
	if (!filefind.FindFile(filename))
	{
		AfxMessageBox(_T("文件不存在！"));
		return false;
	};
	LPDISPATCH lpDisp = NULL;
	books = app.GetWorkbooks();
	lpDisp = books.Open(filename,
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
		AfxMessageBox(book.GetName() + _T("不是excel文件！"));
		return false;
	}
	sheets = book.GetSheets();
	sheet = sheets.GetItem(COleVariant((short)1));
	return true;
}
bool JDExcel::isExcel(CString name)
{
		int n;
		n = name.Find(_T("."));
		CString temp;
		temp = name.Mid(n + 1, 10);
		if (temp == _T("xls") || temp == _T("xlsm") || temp == _T("xlsx"))
			return true;
		else
			return false;
}
void JDExcel::saveExcel()
{
	book.SetSaved(true);
}
CString JDExcel::getCellValue(int row, int col)
{
	range = sheet.GetRange(COleVariant(indexToString(row, col)), COleVariant(indexToString(row, col)));
	COleVariant tempValue;
	tempValue = COleVariant(range.GetValue2());
	tempValue.ChangeType(VT_BSTR);
	return CString(tempValue.bstrVal);
}
CString JDExcel::indexToString(int row,int col)
{
	CString temp;
	if (col > 26)
	{
		temp.Format(_T("%c%c%d"), 'A' + (col - 1) / 26 - 1, 'A' + (col - 1) % 26, row);
	}
	else
	{
		temp.Format(_T("%c%d"), 'A' + (col - 1) % 26, row);
	}
	return temp;
}
int JDExcel::lastRow()
{
	int i;
	CString str;
	for (i = 4;; ++i)
	{
		str.Format(_T("%s"), this->getCellValue(i, 1).Trim());
		if (str.Compare(_T("")) == 0)
		{
			return i-1;
		}
	}
}
int JDExcel::laseCol()
{
	int i;
	CString str;
	for (i = 1;; i++)
	{
		str.Format(_T("%s"), this->getCellValue(4, i).Trim());
		if (str.Compare(_T("")) == 0)
		{
			return i-1;
		}
	}
}

void JDExcel::writeXml(CString path) {
	using namespace std;
	//const char * xmlFile = UnicodeToUtf8(path);
	TiXmlDocument doc;
	TiXmlDeclaration * decl = new TiXmlDeclaration("1.0", "", "");
	TiXmlElement * titleElement = new TiXmlElement("aaa");
	for (int i = 4; i <= this->lastRow(); ++i)
	{
		TiXmlElement * Element = new TiXmlElement(UnicodeToUtf8(this->getCellValue(3, 1)));
		for (int b = 1; b <= this->laseCol(); ++b)
		{
			Element->SetAttribute(UnicodeToUtf8(this->getCellValue(3, b)), UnicodeToUtf8(this->getCellValue(i, b)));
		}
		titleElement->LinkEndChild(Element);
	}
	doc.LinkEndChild(decl);
	doc.LinkEndChild(titleElement);
	char* www = UnicodeToUtf8(path);
	int n = sizeof(path);
	char*tt = new char[n+1];
	memcpy(tt, www, sizeof(path));
	doc.SaveFile(www);
}
char* JDExcel::UnicodeToUtf8(CString unicode)
{
	int len;
	len = WideCharToMultiByte(CP_UTF8, 0, (LPCWSTR)unicode, -1, NULL, 0, NULL, NULL);
	char *szUtf8 = new char[len + 1];
	memset(szUtf8, 0, len * 2 + 2);
	WideCharToMultiByte(CP_UTF8, 0, (LPCWSTR)unicode, -1, szUtf8, len, NULL, NULL);
	return szUtf8;
}