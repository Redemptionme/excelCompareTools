#include <iostream>
#include <string>
#include <vector>
#include <map>

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


    std::string getMainKey() const { return mainKey; }
    void setMainKey(std::string val) { mainKey = val; }

    std::string getMainValue() const { return mainValue; }
    void setMainValue(std::string val) { mainValue = val; }

private:
    SEXcel m_file1;
    SEXcel m_file2;
    std::string mainKey;
    std::string mainValue;
};