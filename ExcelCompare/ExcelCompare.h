
// ExcelCompare.h : PROJECT_NAME Ӧ�ó������ͷ�ļ�
//

#pragma once

#ifndef __AFXWIN_H__
	#error "�ڰ������ļ�֮ǰ������stdafx.h�������� PCH �ļ�"
#endif

#include "resource.h"		// ������


// CExcelCompareApp: 
// �йش����ʵ�֣������ ExcelCompare.cpp
//

class CExcelCompareApp : public CWinApp
{
public:
	CExcelCompareApp();

// ��д
public:
	virtual BOOL InitInstance();

    // ʵ��

	DECLARE_MESSAGE_MAP()
};

extern CExcelCompareApp theApp;