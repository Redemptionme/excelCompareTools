
// ExcelCompareDlg.h : 头文件
//


#pragma once

#include "CExcelCompareTools.h"

// CExcelCompareDlg 对话框
class CExcelCompareDlg : public CDialogEx,public IExcelCompareToolsDelegate
{
// 构造
public:
	CExcelCompareDlg(CWnd* pParent = NULL);	// 标准构造函数

// 对话框数据
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_EXCELCOMPARE_DIALOG };
#endif

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
    afx_msg void OnBnClickedOpenBtn1();
    afx_msg void OnBnClickedOpenBtn2();
    void initData();
    void doCompare();
    bool checkFileExist();
    void loadFile1();
    void loadFile2();
    void test();
    virtual void showTip(const char *str);
    virtual void showLog(const char *str);
private:
    CString m_fileName1;
    CString m_fileName2;
    CExcelCompareTools *m_pExcelComapreTools;
public:
    afx_msg void OnStnClickedStaticCont();
    afx_msg void OnBnClickedCancel();
    afx_msg void OnBnClickedCancel2();
    afx_msg void OnBnClickedCancel3();
};
