#include <iostream>
#include <string>
#include <vector>
#include <map>
#include <unordered_map>

class IExcelCompareToolsDelegate
{
public:
    virtual void showTip(const char *str) = 0;
    virtual void showLog(const char *str) = 0;
};

class CExcelCompareTools
{
public:
    CExcelCompareTools();
    ~CExcelCompareTools();

    typedef std::map<std::string, std::string> SExcelTableItemRow;
    typedef std::vector<SExcelTableItemRow> SExcelSheet;
    typedef std::map<std::string, SExcelSheet> SEXcel;

    void setFile1(SEXcel &file);
    void setFile2(SEXcel &file);

    void compare();

    void productCompareMap(const std::string &key, const std::string &value,const std::string sheetName,const  SEXcel &file, OUT std::unordered_map<std::string, std::string> &outMap);

    void showTip(const char *str);

    void outPutWindows(std::unordered_map<std::string, std::string> &map);

    std::string getMainKey() const { return m_MainKey; }
    void setMainKey(std::string val) { m_MainKey = val; }

    std::string getMainValue() const { return m_MainValue; }
    void setMainValue(std::string val) { m_MainValue = val; }

    std::string getSheetName() const { return m_SheetName; }
    void setSheetName(std::string val) { m_SheetName = val; }

    void setDelegate(IExcelCompareToolsDelegate * p) { m_pDelegate = p; };

    void pushLog(const std::string & str);
    void showLog();

private:
    SEXcel m_file1;
    SEXcel m_file2;
    std::string m_MainKey;
    std::string m_MainValue;
    std::string m_SheetName;
    IExcelCompareToolsDelegate *m_pDelegate;

    std::string m_ResultLog;
};