#include "stdafx.h"
#include "JDfindFile.h"


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