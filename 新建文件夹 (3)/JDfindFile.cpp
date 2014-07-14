#include "stdafx.h"
#include "JDfindFile.h"
#include <string>
#include <iostream>
char* UnicodeToUtf82(CString unicode);
JDfindFile::JDfindFile(string path ,CListBox*list1)
{
	long n;
	int i = 0;
	_finddata_t file;
	if ((n = _findfirst(path.c_str(), &file)) == -1)
		AfxMessageBox(_T("没有找到文件！"));
	else
	{
		//list1->AddString((CString)file.name);
		p[i] = file.name;
		i++;
		while (_findnext(n, &file) == 0)
		{
			//list1->AddString((CString)file.name);
			p[i] = file.name;
			i++;
			cout = --i;
		}
	}
	_findclose(n);
}


JDfindFile::~JDfindFile()
{
}

CString JDfindFile::getFileName(int id)
{
	int n;
	n = p[id].Find(_T("."));
	CString temp;
	temp = p[id].Mid(0, n);
	return temp;
}
 const char* JDfindFile::getfileName2(int id)
{
	int n;
	n = p[id].Find(_T("."));
	string temp;
	temp = UnicodeToUtf82(p[id].Mid(0, n));
	return temp.c_str();
}
 char* UnicodeToUtf82(CString unicode)
 {
	 int len;
	 len = WideCharToMultiByte(CP_UTF8, 0, (LPCWSTR)unicode, -1, NULL, 0, NULL, NULL);
	 char *szUtf8 = new char[len + 1];
	 memset(szUtf8, 0, len * 2 + 2);
	 WideCharToMultiByte(CP_UTF8, 0, (LPCWSTR)unicode, -1, szUtf8, len, NULL, NULL);
	 return szUtf8;
 }
