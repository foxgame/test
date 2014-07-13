// ConsoleApplication6.cpp : 定义控制台应用程序的入口点。
//

#include "stdafx.h"
#include <string>
#include <windows.h>
#include <iostream>
#include <fstream>
#include <io.h>
#include "stdlib.h"

#define MAXFILE 9999
//using std::string;
//using std::wstring;
//using std::cout;
//using std::endl;
using namespace std;
void xxxx(string&);
string GetExePath(void);
int _tmain(int argc, _TCHAR* argv[])
{
	string p[MAXFILE];
	string  temp = "\\*.xml";
	string path = GetExePath() +temp;
	long n;
	int i=0;
	_finddata_t file;
	if ((n = _findfirst(path.c_str(), &file)) == -1)
		cout << -1 << '\n';
	else
	{

		cout << file.name << '\n';
		p[i] = file.name;
		i++;
		while (_findnext(n, &file) == 0)
		{
			cout << file.name << '\n';
			p[i] = file.name;
			cout << p[i].empty() << '\n';
			i++;
		}
	}
	_findclose(n);

	for (i--; i >= 0; i--)
	{
		string temp1 = "\\";
		string temp = GetExePath() + temp1+ p[i];
		xxxx(temp);
		cout << "ok!"<<'\n';	
	}
	getchar();
	return 0;
}
string GetExePath(void)
{
	char szFilePath[MAX_PATH + 1] = { 0 };
	GetModuleFileNameA(NULL, szFilePath, MAX_PATH);
	(strrchr(szFilePath, '\\'))[0] = 0;
	string path = szFilePath;

	return path;
}
void xxxx(string& strFilePath)
{
	fstream in_stream;
	char inBOM[3];

	in_stream.open(strFilePath.c_str());
	in_stream.read(inBOM, 3);
	if (inBOM[0] + 256 == 0xef && inBOM[1] + 256 == 0xbb && inBOM[2] + 256 == 0xbf)  //UTF8 
	{
		in_stream.close();
		cout << "this one is UTF8！";
		return;
	}
	else
	{
		in_stream.seekg(0);
		char *ansi_text = NULL;
		WCHAR *unicdoe_text = NULL;
		char *utf8_text = NULL;
		int ansi_text_length = 9999999;
		ansi_text = new char[ansi_text_length];

		int utf8_text_length = 0;

		in_stream.seekg(0);
		in_stream.read(ansi_text, ansi_text_length);
		ansi_text_length = in_stream.gcount();
		in_stream.close();
		ansi_text[ansi_text_length] = '\0';

		int unicode_text_length = MultiByteToWideChar(CP_ACP, NULL, ansi_text, ansi_text_length, NULL, 0);
		unicdoe_text = new  WCHAR[unicode_text_length + 1];
		MultiByteToWideChar(CP_ACP, NULL, ansi_text, ansi_text_length, unicdoe_text, unicode_text_length);
		unicdoe_text[unicode_text_length] = WCHAR('\0');

		utf8_text_length = WideCharToMultiByte(CP_UTF8, NULL, unicdoe_text, unicode_text_length, NULL, 0, NULL, NULL);
		utf8_text = new char[utf8_text_length + 4];
		utf8_text[0] = 0xef;
		utf8_text[1] = 0xbb;
		utf8_text[2] = 0xbf;

		WideCharToMultiByte(CP_UTF8, NULL, unicdoe_text, unicode_text_length, &utf8_text[3], utf8_text_length, NULL, NULL);
		utf8_text_length += 3;
		utf8_text[utf8_text_length] = '\0';

		ofstream out_stream;
		out_stream.open(strFilePath.c_str());
		out_stream.write(utf8_text, utf8_text_length);
		out_stream.close();
	
		if (ansi_text)
		{
			delete[]ansi_text;
		}
		if (utf8_text)
		{
			delete[]utf8_text;
		}
		if (unicdoe_text)
		{
			delete[]unicdoe_text;
		}
	}
}
