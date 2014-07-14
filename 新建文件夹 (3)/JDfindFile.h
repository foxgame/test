#pragma once
#include "stdafx.h"
#include <string>
#include <windows.h>
#include <iostream>
#include <fstream>
#include <io.h>
#include "stdlib.h"
#define MAXFILE 999
using namespace std;
class JDfindFile
{
public:
	JDfindFile(string path, CListBox*);
	~JDfindFile();

public:
	CString getFileName(int id);
	const char* getfileName2(int id);
	CString p[MAXFILE];
	long cout;
private:


};

