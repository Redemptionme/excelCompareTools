
// ExcelCompareDlg.cpp : ʵ���ļ�
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


// ����Ӧ�ó��򡰹��ڡ��˵���� CAboutDlg �Ի���

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// �Ի�������
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_ABOUTBOX };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV ֧��

// ʵ��
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


// CExcelCompareDlg �Ի���



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


// CExcelCompareDlg ��Ϣ�������

BOOL CExcelCompareDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// ��������...���˵�����ӵ�ϵͳ�˵��С�

	// IDM_ABOUTBOX ������ϵͳ���Χ�ڡ�
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

	// ���ô˶Ի����ͼ�ꡣ  ��Ӧ�ó��������ڲ��ǶԻ���ʱ����ܽ��Զ�
	//  ִ�д˲���
	SetIcon(m_hIcon, TRUE);			// ���ô�ͼ��
	SetIcon(m_hIcon, FALSE);		// ����Сͼ��

	// TODO: �ڴ���Ӷ���ĳ�ʼ������
    initData();
	return TRUE;  // ���ǽ��������õ��ؼ������򷵻� TRUE
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

// �����Ի��������С����ť������Ҫ����Ĵ���
//  �����Ƹ�ͼ�ꡣ  ����ʹ���ĵ�/��ͼģ�͵� MFC Ӧ�ó���
//  �⽫�ɿ���Զ���ɡ�

void CExcelCompareDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // ���ڻ��Ƶ��豸������

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// ʹͼ���ڹ����������о���
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// ����ͼ��
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

//���û��϶���С������ʱϵͳ���ô˺���ȡ�ù��
//��ʾ��
HCURSOR CExcelCompareDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}



void CExcelCompareDlg::OnBnClickedOk()
{
    // TODO: �ڴ���ӿؼ�֪ͨ����������
    //CDialogEx::OnOK();
    doCompare();
}


void CExcelCompareDlg::OnBnClickedOpenBtn1()
{
    // TODO: �ڴ���ӿؼ�֪ͨ����������
    CFileDialog dlg(TRUE, _T("txt"), _T("b.txt"), OFN_FILEMUSTEXIST | OFN_HIDEREADONLY, _T("office�ĵ�|*.xlsx|office�ĵ�|*.xls|�����ļ�|*||"));
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
    // TODO: �ڴ���ӿؼ�֪ͨ����������
    CFileDialog dlg(TRUE, _T("txt"), _T("b.txt"), OFN_FILEMUSTEXIST | OFN_HIDEREADONLY, _T("office�ĵ�|*.xlsx|office�ĵ�|*.xls|�����ļ�|*||"));
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
        pathName = "�����һ���ĵ���·��";
        pMessage1->SetWindowTextW(pathName);
    }
    CEdit* pMessage2 = (CEdit*) GetDlgItem(IDC_OPEN_FILE_TXT_2);  
    if (pMessage2) {
        pathName = "����ڶ����ĵ���·��";
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
                    if (newt2 == "���������") {
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
        CString tip = m_fileName1 + "������!!!";
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
        CString tip = m_fileName2 + "������!!!";
        MessageBox(tip);
        return false;
    }
    
    return true;
}

