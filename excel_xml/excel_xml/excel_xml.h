
// excel_xml.h : PROJECT_NAME Ӧ�ó������ͷ�ļ�
//

#pragma once

#ifndef __AFXWIN_H__
	#error "�ڰ������ļ�֮ǰ������stdafx.h�������� PCH �ļ�"
#endif

#include "resource.h"		// ������


// Cexcel_xmlApp: 
// �йش����ʵ�֣������ excel_xml.cpp
//

class Cexcel_xmlApp : public CWinApp
{
public:
	Cexcel_xmlApp();

// ��д
public:
	virtual BOOL InitInstance();

// ʵ��

	DECLARE_MESSAGE_MAP()
};

extern Cexcel_xmlApp theApp;