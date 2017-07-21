
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
    ON_BN_CLICKED(IDCANCEL, &CExcelCompareDlg::OnBnClickedCancel)
    ON_BN_CLICKED(IDCANCEL2, &CExcelCompareDlg::OnBnClickedCancel2)
    ON_BN_CLICKED(IDCANCEL3, &CExcelCompareDlg::OnBnClickedCancel3)
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
	//test();
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
    m_pExcelComapreTools = new CExcelCompareTools();
    m_pExcelComapreTools->setDelegate(this);
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
    CEdit* pMessage3 = (CEdit*)GetDlgItem(IDC_INPUT_TXT1);
    if (pMessage3) {
        pathName = "��������";
        pMessage3->SetWindowTextW(pathName);
    }
    CEdit* pMessage4 = (CEdit*)GetDlgItem(IDC_INPUT_TXT2);
    if (pMessage4) {
        pathName = "����У���ֵ";
        pMessage4->SetWindowTextW(pathName);
    }
    CEdit* pMessage5 = (CEdit*)GetDlgItem(IDC_INPUT_TXT3);
    if (pMessage5) {
        pathName = "����sheet����";
        pMessage5->SetWindowTextW(pathName);
    }
}

void CExcelCompareDlg::doCompare() {    
    if (checkFileExist()) 
    {
        loadFile1();
        loadFile2();
        m_pExcelComapreTools->compare();
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

    /*TCHAR pBuf[MAX_PATH];
    GetCurrentDirectory(MAX_PATH, pBuf);
    CString path = pBuf;
    path += "\\testFile\\1";
    CString newfile1 = path;

    int test = m_fileName1.ReverseFind('.');
    CString tests;
    for (int i = test; i < m_fileName1.GetLength(); ++i) {
        tests += m_fileName1[i];
    }
    newfile1 += tests;

    bool x =CopyFile(m_fileName1, newfile1, false);*/
    
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

    CString key;
    string skey;

    GetDlgItem(IDC_INPUT_TXT1)->GetWindowTextW(key);
    skey = CT2A(key);
    m_pExcelComapreTools->setMainKey(skey);

    GetDlgItem(IDC_INPUT_TXT2)->GetWindowTextW(key);
    skey = CT2A(key);
    m_pExcelComapreTools->setMainValue(skey);

    GetDlgItem(IDC_INPUT_TXT3)->GetWindowTextW(key);
    skey = CT2A(key);
    m_pExcelComapreTools->setSheetName(skey);
    
    return true;
}

void CExcelCompareDlg::loadFile1() {
    CStringA charstr(m_fileName1);

    xlnt::workbook wb;
    wb.load((const char *)charstr);
    //worksheet a = wb.sheet_by_title("余额");
    CExcelCompareTools::SEXcel excel;
    // The workbook class has begin and end methods so it can be iterated upon.
    // Each item is a sheet in the workbook.
    for (const auto sheet : wb)
    {
        // Print the title of the sheet on its own line.
        //std::cout << sheet.get_title() << ": " << std::endl;
        string titletxt = UTF8_To_string(sheet.title());
        //OutputDebugStringA(titletxt.c_str());
        //OutputDebugStringA("\n========================\n");
        // Iterating on a range, such as from worksheet::rows, yields cell_vectors.
        // Cell vectors don't actually contain cells to reduce overhead.
        // Instead they hold a reference to a worksheet and the current cell_reference.
        // Internally, calling worksheet::get_cell with the current cell_reference yields the next cell.
        // This allows easy and fast iteration over a row (sometimes a column) in the worksheet.

        std::vector<std::string> keyVec;
        CExcelCompareTools::SExcelSheet excelSheet;
        for (auto row : sheet.rows())
        {
            int t = row.length();
            CExcelCompareTools::SExcelTableItemRow rowItem;
            for (auto cell : row)
            {
                // cell::operator<< adds a string represenation of the cell's value to the stream.
                auto rowline = cell.row();
                auto  rowColumn = cell.column().index;

                string txt = cell.to_string();
                string txtCont = UTF8_To_string(txt);
                //OutputDebugStringA(txtCont.c_str());
                if (rowline == 1) {
                    keyVec.push_back(txtCont);
                }
                else
                {
                    rowItem[keyVec[rowColumn - 1]] = txtCont;
                }
            }
            if (!rowItem.empty()) {
                excelSheet.push_back(rowItem);
            }

            //OutputDebugStringA("           \n");
            std::cout << std::endl;
        }
        excel.insert(make_pair(titletxt, excelSheet));
    }
    m_pExcelComapreTools->setFile1(excel);
}

void CExcelCompareDlg::loadFile2() {
    CStringA charstr(m_fileName2);

    xlnt::workbook wb;
    wb.load((const char *)charstr);
    //worksheet a = wb.sheet_by_title("余额");
    CExcelCompareTools::SEXcel excel;
    // The workbook class has begin and end methods so it can be iterated upon.
    // Each item is a sheet in the workbook.
    for (const auto sheet : wb)
    {
        // Print the title of the sheet on its own line.
        //std::cout << sheet.get_title() << ": " << std::endl;
        string titletxt = UTF8_To_string(sheet.title());
        //OutputDebugStringA(titletxt.c_str());
        //OutputDebugStringA("\n========================\n");
        // Iterating on a range, such as from worksheet::rows, yields cell_vectors.
        // Cell vectors don't actually contain cells to reduce overhead.
        // Instead they hold a reference to a worksheet and the current cell_reference.
        // Internally, calling worksheet::get_cell with the current cell_reference yields the next cell.
        // This allows easy and fast iteration over a row (sometimes a column) in the worksheet.

        std::vector<std::string> keyVec;
        CExcelCompareTools::SExcelSheet excelSheet;
        for (auto row : sheet.rows())
        {
            int t = row.length();
            CExcelCompareTools::SExcelTableItemRow rowItem;
            for (auto cell : row)
            {
                // cell::operator<< adds a string represenation of the cell's value to the stream.
                auto rowline = cell.row();
                auto  rowColumn = cell.column().index;

                string txt = cell.to_string();
                string txtCont = UTF8_To_string(txt);
                //OutputDebugStringA(txtCont.c_str());
                if (rowline == 1) {
                    keyVec.push_back(txtCont);
                }
                else
                {
                    rowItem[keyVec[rowColumn - 1]] = txtCont;
                }
            }
            if (!rowItem.empty()) {
                excelSheet.push_back(rowItem);
            }

            //OutputDebugStringA("           \n");
            std::cout << std::endl;
        }
        excel.insert(make_pair(titletxt, excelSheet));
    }
    m_pExcelComapreTools->setFile2(excel);
}



void CExcelCompareDlg::test() {
    TCHAR pBuf[MAX_PATH];
    GetCurrentDirectory(MAX_PATH, pBuf);
    CString path = pBuf;
    path += "//testFile//";
    m_fileName1 = path;
    m_fileName2 = path;
    m_fileName1 += "1.xlsx" ;
    m_fileName2 += "2.xlsx";
    CEdit* pMessage1 = (CEdit*)GetDlgItem(IDC_OPEN_FILE_TXT_1);
    if (!pMessage1) {
        return;
    }
    pMessage1->SetWindowTextW(m_fileName1);
    CEdit* pMessage2 = (CEdit*)GetDlgItem(IDC_OPEN_FILE_TXT_2);
    if (!pMessage2) {
        return;
    }
    pMessage2->SetWindowTextW(m_fileName2);

    CString key ("��������Ա���");
    GetDlgItem(IDC_INPUT_TXT1)->SetWindowTextW(key);
    
    CString value ("�˺ſ���");
    GetDlgItem(IDC_INPUT_TXT2)->SetWindowTextW(value);

    CString sheet("�ͻ�");
    GetDlgItem(IDC_INPUT_TXT3)->SetWindowTextW(sheet);

    doCompare();
}


void CExcelCompareDlg::showTip(const char *str) {
    MessageBox(CString(str), CString("Tip"));
}

void CExcelCompareDlg::showLog(const char *str) {   
    if (GetDlgItem(IDC_OUT_TXT)) {
        GetDlgItem(IDC_OUT_TXT)->SetWindowTextW(CString(str));
    }
}

void CExcelCompareDlg::OnBnClickedCancel()
{
    // TODO: �ڴ���ӿؼ�֪ͨ����������
    CDialogEx::OnCancel();
}


void CExcelCompareDlg::OnBnClickedCancel2()
{
    // TODO: �ڴ���ӿؼ�֪ͨ����������
    CString title("���ڰ���");
    CString content("1open �������ļ���·��\n2��������,�Ƚ�ֵ������sheet����\n3�����ʼ�Ƚϲ鿴���");
    MessageBox(content, title);
}


void CExcelCompareDlg::OnBnClickedCancel3()
{
    // TODO: �ڴ���ӿؼ�֪ͨ����������

    CString title("����");
    CString content("for С���� �ر��");
    MessageBox(content, title);
}
