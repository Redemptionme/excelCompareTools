
// ExcelCompareDlg.cpp : 实现文件
//

#include "stdafx.h"
#include "ExcelCompare.h"
#include "ExcelCompareDlg.h"
#include "afxdialogex.h"
#include <xlnt/xlnt.hpp>
#include <shlwapi.h>
#pragma comment(lib,"Shlwapi.lib")
#include <iostream>
#include "../othercode/othercode.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

using namespace xlnt;
using namespace std;


// 用于应用程序“关于”菜单项的 CAboutDlg 对话框

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// 对话框数据
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_ABOUTBOX };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

// 实现
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialogEx(IDD_ABOUTBOX)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialogEx)
END_MESSAGE_MAP()


// CExcelCompareDlg 对话框



CExcelCompareDlg::CExcelCompareDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(IDD_EXCELCOMPARE_DIALOG, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CExcelCompareDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CExcelCompareDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
    ON_BN_CLICKED(IDOK, &CExcelCompareDlg::OnBnClickedOk)
    ON_BN_CLICKED(ID_OPEN_BTN_1, &CExcelCompareDlg::OnBnClickedOpenBtn1)
    ON_BN_CLICKED(ID_OPEN_BTN_2, &CExcelCompareDlg::OnBnClickedOpenBtn2)
END_MESSAGE_MAP()


// CExcelCompareDlg 消息处理程序

BOOL CExcelCompareDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// 将“关于...”菜单项添加到系统菜单中。

	// IDM_ABOUTBOX 必须在系统命令范围内。
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		BOOL bNameValid;
		CString strAboutMenu;
		bNameValid = strAboutMenu.LoadString(IDS_ABOUTBOX);
		ASSERT(bNameValid);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// 设置此对话框的图标。  当应用程序主窗口不是对话框时，框架将自动
	//  执行此操作
	SetIcon(m_hIcon, TRUE);			// 设置大图标
	SetIcon(m_hIcon, FALSE);		// 设置小图标

	// TODO: 在此添加额外的初始化代码
    initData();
	return TRUE;  // 除非将焦点设置到控件，否则返回 TRUE
}

void CExcelCompareDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialogEx::OnSysCommand(nID, lParam);
	}
}

// 如果向对话框添加最小化按钮，则需要下面的代码
//  来绘制该图标。  对于使用文档/视图模型的 MFC 应用程序，
//  这将由框架自动完成。

void CExcelCompareDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // 用于绘制的设备上下文

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// 使图标在工作区矩形中居中
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// 绘制图标
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

//当用户拖动最小化窗口时系统调用此函数取得光标
//显示。
HCURSOR CExcelCompareDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}



void CExcelCompareDlg::OnBnClickedOk()
{
    // TODO: 在此添加控件通知处理程序代码
    //CDialogEx::OnOK();
    doCompare();
}


void CExcelCompareDlg::OnBnClickedOpenBtn1()
{
    // TODO: 在此添加控件通知处理程序代码
    CFileDialog dlg(TRUE, _T("txt"), _T("b.txt"), OFN_FILEMUSTEXIST | OFN_HIDEREADONLY, _T("office文档|*.xlsx|office文档|*.xls|所有文件|*||"));
    if (dlg.DoModal() == IDOK) {
        CString pathname = dlg.GetPathName();
        CEdit* pMessage = (CEdit*)GetDlgItem(IDC_OPEN_FILE_TXT_1);
        if (pMessage) {
            pMessage->SetWindowTextW(pathname);
        }
    }
}


void CExcelCompareDlg::OnBnClickedOpenBtn2()
{
    // TODO: 在此添加控件通知处理程序代码
    CFileDialog dlg(TRUE, _T("txt"), _T("b.txt"), OFN_FILEMUSTEXIST | OFN_HIDEREADONLY, _T("office文档|*.xlsx|office文档|*.xls|所有文件|*||"));
    if (dlg.DoModal() == IDOK) {
        CString pathname = dlg.GetPathName();
        CEdit* pMessage = (CEdit*)GetDlgItem(IDC_OPEN_FILE_TXT_2);
        if (pMessage) {
            pMessage->SetWindowTextW(pathname);
        }
    }
}


void CExcelCompareDlg::initData() {
    CString pathName;
    CEdit* pMessage1 = (CEdit*) GetDlgItem(IDC_OPEN_FILE_TXT_1);  
    if (pMessage1) {
        pathName = "输入第一个文档的路径";
        pMessage1->SetWindowTextW(pathName);
    }
    CEdit* pMessage2 = (CEdit*) GetDlgItem(IDC_OPEN_FILE_TXT_2);  
    if (pMessage2) {
        pathName = "输入第二个文档的路径";
        pMessage2->SetWindowTextW(pathName);
    }
}

void CExcelCompareDlg::doCompare() {    
    if (checkFileExist()||true) 
    {
        
        CStringA charstr(m_fileName1);
        
        xlnt::workbook wb;
        wb.load((const char *)charstr);
        
        // The workbook class has begin and end methods so it can be iterated upon.
        // Each item is a sheet in the workbook.
        for (const auto sheet : wb)
        {
            // Print the title of the sheet on its own line.
            std::cout << sheet.get_title() << ": " << std::endl;

            // Iterating on a range, such as from worksheet::rows, yields cell_vectors.
            // Cell vectors don't actually contain cells to reduce overhead.
            // Instead they hold a reference to a worksheet and the current cell_reference.
            // Internally, calling worksheet::get_cell with the current cell_reference yields the next cell.
            // This allows easy and fast iteration over a row (sometimes a column) in the worksheet.
            for (auto row : sheet.rows())
            {
                for (auto cell : row)
                {
                    // cell::operator<< adds a string represenation of the cell's value to the stream.
                    std::cout << cell << ", ";
                    string txt = cell.to_string();
                    string newt = string_To_UTF8(txt);
                    string newt2 = UTF8_To_string(txt);
                    if (newt2 == "交易所编号") {
                        std::wstring stemp = std::wstring(newt2.begin(), newt2.end());
                        LPCWSTR sw = stemp.c_str();
                        OutputDebugString(sw);
                    }
                    
                }

                std::cout << std::endl;
            }
        }
        
    }
}

bool CExcelCompareDlg::checkFileExist() {
    CEdit* pMessage1 = (CEdit*)GetDlgItem(IDC_OPEN_FILE_TXT_1);
    if (!pMessage1) {
        return false;
    }
    pMessage1->GetWindowTextW(m_fileName1);
    if (!PathFileExists(m_fileName1))
    {
        CString tip = m_fileName1 + "不存在!!!";
        MessageBox(tip);
        return false;
    }
    
    CEdit* pMessage2 = (CEdit*)GetDlgItem(IDC_OPEN_FILE_TXT_2);
    if (!pMessage2) {
        return false;
    }
    pMessage2->GetWindowTextW(m_fileName2);
    if (!PathFileExists(m_fileName2))
    {
        CString tip = m_fileName2 + "不存在!!!";
        MessageBox(tip);
        return false;
    }
    
    return true;
}

