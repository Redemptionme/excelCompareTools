#include <string>
#include <xstring>

using namespace std;
static std::string string_To_UTF8(const std::string & str)
{
    int nwLen = ::MultiByteToWideChar(CP_ACP, 0, str.c_str(), -1, NULL, 0);

    wchar_t * pwBuf = new wchar_t[nwLen + 1];//һ��Ҫ��1����Ȼ�����β��  
    ZeroMemory(pwBuf, nwLen * 2 + 2);

    ::MultiByteToWideChar(CP_ACP, 0, str.c_str(), str.length(), pwBuf, nwLen);

    int nLen = ::WideCharToMultiByte(CP_UTF8, 0, pwBuf, -1, NULL, NULL, NULL, NULL);

    char * pBuf = new char[nLen + 1];
    ZeroMemory(pBuf, nLen + 1);

    ::WideCharToMultiByte(CP_UTF8, 0, pwBuf, nwLen, pBuf, nLen, NULL, NULL);

    std::string retStr(pBuf);

    delete[]pwBuf;
    delete[]pBuf;

    pwBuf = NULL;
    pBuf = NULL;

    return retStr;
}


static std::string UTF8_To_string(const std::string & str)
{
    int nwLen = MultiByteToWideChar(CP_UTF8, 0, str.c_str(), -1, NULL, 0);

    wchar_t * pwBuf = new wchar_t[nwLen + 1];//һ��Ҫ��1����Ȼ�����β��  
    memset(pwBuf, 0, nwLen * 2 + 2);

    MultiByteToWideChar(CP_UTF8, 0, str.c_str(), str.length(), pwBuf, nwLen);

    int nLen = WideCharToMultiByte(CP_ACP, 0, pwBuf, -1, NULL, NULL, NULL, NULL);

    char * pBuf = new char[nLen + 1];
    memset(pBuf, 0, nLen + 1);

    WideCharToMultiByte(CP_ACP, 0, pwBuf, nwLen, pBuf, nLen, NULL, NULL);

    std::string retStr = pBuf;

    delete[]pBuf;
    delete[]pwBuf;

    pBuf = NULL;
    pwBuf = NULL;

    return retStr;
}

void UTF_8ToUnicode(wchar_t* pOut, char *pText)
{
    char* uchar = (char *)pOut;
    uchar[1] = ((pText[0] & 0x0F) << 4) + ((pText[1] >> 2) & 0x0F);
    uchar[0] = ((pText[1] & 0x03) << 6) + (pText[2] & 0x3F);
}
void UnicodeToUTF_8(char* pOut, wchar_t* pText)
{
    // ע�� WCHAR�ߵ��ֵ�˳��,���ֽ���ǰ�����ֽ��ں�   
    char* pchar = (char *)pText;
    pOut[0] = (0xE0 | ((pchar[1] & 0xF0) >> 4));
    pOut[1] = (0x80 | ((pchar[1] & 0x0F) << 2)) + ((pchar[0] & 0xC0) >> 6);
    pOut[2] = (0x80 | (pchar[0] & 0x3F));
}
void UnicodeToGB2312(char* pOut, wchar_t uData)
{
    WideCharToMultiByte(CP_ACP, NULL, &uData, 1, pOut, sizeof(wchar_t), NULL, NULL);
}
void Gb2312ToUnicode(wchar_t* pOut, char *gbBuffer)
{
    ::MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, gbBuffer, 2, pOut, 1);
}
void GB2312ToUTF_8(string& pOut, char *pText, int pLen)
{
    char buf[4] = { 0 };
    int nLength = pLen * 3;
    char* rst = new char[nLength];
    memset(rst, 0, nLength);
    int i = 0, j = 0;
    while (i < pLen)
    {
        //�����Ӣ��ֱ�Ӹ��ƾͿ���   
        if (*(pText + i) >= 0)
        {
            rst[j++] = pText[i++];
        }
        else
        {
            wchar_t pbuffer;
            Gb2312ToUnicode(&pbuffer, pText + i);
            UnicodeToUTF_8(buf, &pbuffer);
            rst[j] = buf[0];
            rst[j + 1] = buf[1];
            rst[j + 2] = buf[2];
            j += 3;
            i += 2;
        }
    }

    rst[j] = '\n';   //���ؽ��    
    pOut = rst;
    delete[]rst;
    return;
}
void UTF_8ToGB2312(char*pOut, char *pText, int pLen)
{
    char Ctemp[4];
    memset(Ctemp, 0, 4);
    int i = 0, j = 0;
    while (i < pLen)
    {
        if (pText[i] >= 0)
        {
            pOut[j++] = pText[i++];
        }
        else
        {
            WCHAR Wtemp;
            UTF_8ToUnicode(&Wtemp, pText + i);
            UnicodeToGB2312(Ctemp, Wtemp);
            pOut[j] = Ctemp[0];
            pOut[j + 1] = Ctemp[1];
            i += 3;
            j += 2;
        }
    }
    pOut[j] = '\n';
    return;
}

bool DeleteFile(char * lpszPath)
{
    SHFILEOPSTRUCT FileOp = { 0 };
    FileOp.fFlags = FOF_ALLOWUNDO |   //����Żػ���վ
        FOF_NOCONFIRMATION; //������ȷ�϶Ի���
    FileOp.pFrom = CString(lpszPath);
    FileOp.pTo = NULL;      //һ��Ҫ��NULL
    FileOp.wFunc = FO_DELETE;    //ɾ������
    return SHFileOperation(&FileOp) == 0;
}
//�����ļ����ļ���
bool CopyFile(char *pTo, char *pFrom)
{
    SHFILEOPSTRUCT FileOp = { 0 };
    FileOp.fFlags = FOF_NOCONFIRMATION |   //������ȷ�϶Ի���
        FOF_NOCONFIRMMKDIR; //��Ҫʱֱ�Ӵ���һ���ļ���,�����û�ȷ��
    FileOp.pFrom = CString(pFrom);
    FileOp.pTo = CString(pTo);
    FileOp.wFunc = FO_COPY;
    return SHFileOperation(&FileOp) == 0;
}
//�ƶ��ļ����ļ���
bool MoveFile(char *pTo, char *pFrom)
{
    SHFILEOPSTRUCT FileOp = { 0 };
    FileOp.fFlags = FOF_NOCONFIRMATION |   //������ȷ�϶Ի���
        FOF_NOCONFIRMMKDIR; //��Ҫʱֱ�Ӵ���һ���ļ���,�����û�ȷ��
    FileOp.pFrom = CString(pFrom);
    FileOp.pTo = CString(pTo);
    FileOp.wFunc = FO_MOVE;
    return SHFileOperation(&FileOp) == 0;
}

//�������ļ����ļ���
bool ReNameFile(char *pTo, char *pFrom)
{
    SHFILEOPSTRUCT FileOp = { 0 };
    FileOp.fFlags = FOF_NOCONFIRMATION;   //������ȷ�϶Ի���
    FileOp.pFrom = CString(pFrom);
    FileOp.pTo = CString(pTo);
    FileOp.wFunc = FO_RENAME;
    return SHFileOperation(&FileOp) == 0;
}


