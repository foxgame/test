#pragma once
#include "stdafx.h"
#include "excel_xml.h"
#include "excel_xmlDlg.h"
#include "afxdialogex.h"
#include <string>
#include <iostream>
#include "tinyxml.h"
class JDExcel 
{
public:
	JDExcel();
	~JDExcel();
public:
	bool openExcelBook(CString filename);
	void createExcelBook();
	void openExcelApp();
	void saveExcel();
	void saveAsExcel(CString filename);
	void setCellValue(int row, int col, int value);
	CString getCellValue(int row, int col);
	CString indexToString(int row, int col);
	int lastRow();
	int laseCol();
	bool isExcel(CString name);
	//void writeXml(const char* path,const char* value);
	void writeXml(const char* path);

private:

};

