#include "stdafx.h"
#include "CExcelCompareTools.h"


CExcelCompareTools::CExcelCompareTools():m_pDelegate(NULL)
{
}

CExcelCompareTools::~CExcelCompareTools()
{

}

void CExcelCompareTools::setFile1(SEXcel &file) {
    m_file1 = file;
}

void CExcelCompareTools::setFile2(SEXcel &file) {
    m_file2 = file;
}

void CExcelCompareTools::compare() {    
    if (m_SheetName.empty()||m_MainKey.empty()||m_MainValue.empty()) {
        showTip("sheet，key，value 出现空值");
        return;
    }
    
    std::unordered_map<std::string, std::string> map1;
    std::unordered_map<std::string, std::string> map2;
    pushLog("================校验文件1================");
    productCompareMap(m_MainKey, m_MainValue, m_SheetName,m_file1, map1);
    pushLog("================校验文件2================");
    productCompareMap(m_MainKey, m_MainValue, m_SheetName, m_file2, map2);
    pushLog("================对比结果================");
    std::string str;
    for (std::unordered_map<std::string, std::string>::iterator iter1 = map1.begin();
        iter1 !=map1.end();
        iter1++) {
        std::unordered_map<std::string, std::string>::iterator iter2 = map2.find(iter1->first);
        if (iter2 != map2.end()) {
            if (iter1->second != iter2->second) {
                str = iter1->first + "在表1为" + iter1->second + "表2为" + iter2->second;
                pushLog(str);
            }
        }
        else
        {
            str = "表一数据" + m_MainKey  + "为" + iter1->first + "在表二不存在";
            pushLog(str);
        }
    }

    //output windows
  //  outPutWindows(map1);
   // outPutWindows(map2);

    showLog();
}

void CExcelCompareTools::productCompareMap(const std::string &key, const std::string &value, const std::string sheetName, const SEXcel &file, OUT std::unordered_map<std::string, std::string> &outMap)  {
    std::string strTip;
    SEXcel::const_iterator iter1 = file.find(sheetName);
    if (iter1 == file.end()) {
        strTip = std::string("文件没有发现对应sheet") + m_SheetName;
        pushLog(strTip);
        return;
    }
    
    char temp[256];
    for (size_t i = 0; i < iter1->second.size(); i++){
        const SExcelTableItemRow &row = iter1->second[i];
        SExcelTableItemRow::const_iterator rowValue1 = row.find(key);
        SExcelTableItemRow::const_iterator rowValue2 = row.find(value);
        
        bool outKey = false;
        bool outValue = false;
        sprintf_s(temp, "%d", i+1);

        if (rowValue1 == row.end()) {                        
            strTip = std::string("第") + temp + std::string("行数据")  + std::string("未填写主键 ") + key;
            pushLog(strTip);
        } else {
            outKey = true;
        }

        if (rowValue2 == row.end()) {
            strTip = std::string("第") + temp + std::string("行数据") + std::string("未填写比较值 ") + value;
            pushLog(strTip);
        } else {
            outValue = true;
        }
        if (outKey&&outValue) {
            std::unordered_map<std::string, std::string>::iterator iter = outMap.find(rowValue1->second);
            if (iter != outMap.end()) {
                strTip = std::string("第") + temp + std::string("行数据") + std::string("key值有重复");
                pushLog(strTip);
                continue;
            }
            outMap[rowValue1->second] = rowValue2->second;
        }
    }
}

void CExcelCompareTools::showTip(const char *str) {
    if (m_pDelegate != NULL) {
        m_pDelegate->showTip(str);
    }
}

void CExcelCompareTools::outPutWindows(std::unordered_map<std::string, std::string> &map) {
    for (std::unordered_map<std::string, std::string>::iterator oiter = map.begin(); oiter != map.end(); oiter++) {
        OutputDebugStringA(oiter->first.c_str());
        OutputDebugStringA("    ");
        OutputDebugStringA(oiter->second.c_str());
        OutputDebugStringA("    \n");
    }
}

void CExcelCompareTools::pushLog(const std::string & str) {
    m_ResultLog += str;
    m_ResultLog += "\r\n";
}

void CExcelCompareTools::showLog() {
    if (m_pDelegate != NULL) {
        m_pDelegate->showLog(m_ResultLog.c_str());
        m_ResultLog.clear();
    }
}
