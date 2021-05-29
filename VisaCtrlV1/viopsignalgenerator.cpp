#include "viopsignalgenerator.h"
#include<qdebug.h>





VIOpSignalGenerator::VIOpSignalGenerator(QObject *parent,VIOPHead *head):VIOperation(parent,head)
{

}

/*
    BERT command
*/
//百分比、科学计数法
bool VIOpSignalGenerator::viop_displayMode(int type)
{
    if (type == PERCENT)
    {
        VISTATUS = viPrintf(VI, ":CALCulate:BERT:DISPlay:MODE PERCent\n");
    }
    else
    {
        VISTATUS = viPrintf(VI, ":CALCulate:BERT:DISPlay:MODE SCIentific\n");
    }
    return (viop_StatusCheck());
}
//设置阿尔法=1
bool VIOpSignalGenerator::viop_alpha()
{
    VISTATUS = viPrintf(VI, ":RADio:CUSTom:ALPHa 1.000\n");   //这个是对的 for real-time modulation
    //viPrintf(vi, ":RADio:DMODulation:ARB:FILTer:ALPHa 0.600\n");
    //viPrintf(vi, ":RADio:DMODulation:ARB:FILTer:ALPHa?\n");
    //viScanf(vi, "%li",&num);
    //printf("The Alpha is:%d\n", num);
    return viop_StatusCheck();
}

//85%
bool VIOpSignalGenerator::viop_alphaDepth()
{
    VISTATUS = viPrintf(VI, ":RADio:CUSTom:MODulation:ASK:DEPTh 85\n");
    return viop_StatusCheck();
}
//设置触发
bool VIOpSignalGenerator::viop_BERTtriggerSource(int source)
{
    switch (source)
    {
    case IMM:
        VISTATUS = viPrintf(VI, ":SENSe:BERT:TRIGger IMM\n");
        break;
    case EXT:
        VISTATUS = viPrintf(VI, ":SENSe:BERT:TRIGger EXT\n");
        break;
    case BUS:
        VISTATUS = viPrintf(VI, ":SENSe:BERT:TRIGger BUS\n");
        break;
    case TRI:
        VISTATUS = viPrintf(VI, ":SENSe:BERT:TRIGger TRI\n");
    default:
        break;
    }
    return viop_StatusCheck();
}

//symbol Rate
bool VIOpSignalGenerator::viop_symbolRate(int symRat)
{
    switch(symRat){
    case 512:VISTATUS = viPrintf(VI, ":RADio:CUSTom:BRATe 512000\n");break;    //默认sps
    case 256:VISTATUS = viPrintf(VI, ":RADio:CUSTom:BRATe 256000\n");break;
    case 28:VISTATUS = viPrintf(VI, ":RADio:CUSTom:BRATe 28000\n");break;
    default:break;
    }

    return viop_StatusCheck();
}
//ASK
bool VIOpSignalGenerator::viop_modulationType()
{
    VISTATUS = viPrintf(VI, ":RADio:CUSTom:MODulation ASK\n");
    return viop_StatusCheck();
}
bool VIOpSignalGenerator::viop_bertOn()
{
    VISTATUS = viPrintf(VI, ":SENSe:BERT:STATe ON\n");
    return viop_StatusCheck();
}
bool VIOpSignalGenerator::viop_bertOff()
{
    VISTATUS = viPrintf(VI, ":SENSe:BERT:STATe OFF\n");
    return viop_StatusCheck();
}
//bool VIOpSignalGenerator::viop_realtimeOn()
//{
//    VISTATUS = viPrintf(VI, ":RADio:CUSTom ON\n");
//    return viop_StatusCheck();
//}
bool VIOpSignalGenerator::viop_realtimeOn()
{
    VISTATUS = viPrintf(VI, ":RADio:CUSTom ON\n");
    return viop_StatusCheck();
}

bool VIOpSignalGenerator::viop_realtimeOff()
{
    VISTATUS = viPrintf(VI, ":RADio:CUSTom OFF\n");
    return viop_StatusCheck();
}

bool VIOpSignalGenerator::viop_loadData()
{
    VISTATUS = viPrintf(VI, ":RADio:CUSTom:DATA \"PN9-1FF-FM0@BIT\"\n");
    return viop_StatusCheck();
}
bool VIOpSignalGenerator::viop_loadDataWU()
{
    VISTATUS = viPrintf(VI, ":RADio:CUSTom:DATA \"14K_TEST_28@BIT\"\n");
    return viop_StatusCheck();
}
bool VIOpSignalGenerator::viop_totalBit(QString tem)
{

    QString tem1=":SENSe:BERT:TBITs "+tem+"\n";
    QByteArray ba=tem1.toLatin1();
    ViConstString com=ba.data();
    VISTATUS = viPrintf(VI, com);
    return viop_StatusCheck();
}

/*
    CW
*/
//设置频率
bool VIOpSignalGenerator::viop_cwFreq(QString frequency, int Unit)
{
//    QString str_1("FREQ ");

    QString str_1(":FREQuency ");
    QString str_khz("kHz\n");
    QString str_mhz("mHz\n");
    QString str_ghz("gHz\n");
    QString command;

    switch (Unit)
    {
    case KHz:
        command = str_1 + frequency + str_khz;
        break;
    case MHz:
        command = str_1 + frequency + str_mhz;
        break;
    case GHz:
        command = str_1 + frequency + str_ghz;
        break;
    }
    viop_sendCommand(&command);
    return viop_StatusCheck();
}

//设置ampl
bool VIOpSignalGenerator::viop_cwAmpl(QString amplitude)
{
    QString str_1("POW:AMPL ");
    QString str_dBm(" dBm\n");
    QString command;
    command = str_1 + amplitude + str_dBm;
    viop_sendCommand(&command);
    return viop_StatusCheck();
}

//RF
bool VIOpSignalGenerator::viop_RFOn()
{
    VISTATUS = viPrintf(VI, "OUTP:STAT ON\n");
    return viop_StatusCheck();
}
bool VIOpSignalGenerator::viop_RFOff()
{
    VISTATUS = viPrintf(VI,"OUTP:STAT OFF\n");
    return viop_StatusCheck();
}
//MOD
bool VIOpSignalGenerator::viop_MODOn()
{
    VISTATUS=viPrintf(VI,"MOD:STAT ON\n");
    return viop_StatusCheck();
}
bool VIOpSignalGenerator::viop_MODOff()
{
    VISTATUS=viPrintf(VI,"MOD:STAT OFF\n");
    return viop_StatusCheck();
}

bool VIOpSignalGenerator::viop_freqMode(int mode)
{
    switch (mode)
    {
    case LIST:
    {
        VISTATUS=viPrintf(VI, "FREQ:MODE LIST\n"); // Sets the sig gen freq mode to list
    }
    break;
    default:
        VISTATUS = false;
        break;
    }
    return viop_StatusCheck();
}

bool VIOpSignalGenerator::viop_listType(int type)
{
    switch (type)
    {
    case STEP:
    {
        VISTATUS = viPrintf(VI, "LIST:TYPE STEP\n"); // Sets sig gen LIST type to step
    }
    break;
    default:
        VISTATUS = 0;
        break;
    }
    return viop_StatusCheck();
}

bool VIOpSignalGenerator::viop_funcType(int type)  //选择信号波形 SQU--方波
{
    switch (type)
    {
    case SQU:
    {
        VISTATUS = viPrintf(VI, "FUNC SQU\n");
    }
    break;
    default:
        VISTATUS = 0;
    }
    return viop_StatusCheck();
}
/*
FM/AM command
*/
bool VIOpSignalGenerator::viop_AMSrc(int src)
{
    switch (src)
    {
    case AMSRC_EXT1:
        VISTATUS = viPrintf(VI, "AM:SOUR EXT1\n");   //am path 1
        break;
    case AMSRC_EXT2:
        VISTATUS = viPrintf(VI, "AM:SOUR EXT2\n");   //am path 1
        break;
    default:
        break;
    }
    return viop_StatusCheck();
}
bool VIOpSignalGenerator::viop_AMONOFF(bool status)
{
    if (status)
        VISTATUS = viPrintf(VI, "AM:STAT ON\n");
    else
        VISTATUS = viPrintf(VI, "AM:STAT OFF\n\n");
    return viop_StatusCheck();
}
bool VIOpSignalGenerator::viop_AMDepth(QString depth)
{
    QString str_1("SOURce:AM:DEPTh:LINear ");
    QString str_2("\n");
    QString command;
    command = str_1 + depth + str_2;
    viop_sendCommand(&command);
    return viop_StatusCheck();
}
bool VIOpSignalGenerator::viop_AMWBONOFF(bool status)
{
    if (status)
        VISTATUS = viPrintf(VI, "AM:WIDeband:STATe 1\n");
    else
        VISTATUS = viPrintf(VI, "AM:WIDeband:STATe 0\n");
    return viop_StatusCheck();
}
/*
    Quries command
*/
bool VIOpSignalGenerator::viop_Q_bertONOFF()
{
    int num;
    VISTATUS = viPrintf(VI, ":SENSe:BERT:STATe?\n");
    if (VISTATUS != 1)
    {
        viScanf(VI, "%li", &num);
        if (num > 0)
            return true;  // on
        else
            return false; //off
    }
    else
    {
        return false;
    }
}

bool VIOpSignalGenerator::viop_Q_realtimeONOFF()
{
    int num;
    VISTATUS = viPrintf(VI, ":RADio:CUSTom?\n");  // p515
    if (VISTATUS != 1)
    {
        viScanf(VI, "%li", &num);
        if (num > 0)
            return true;  // on
        else
            return false; //off
    }
    else
    {
        return false;
    }
}

bool VIOpSignalGenerator::viop_Q_rfONOFF()
{
    int num=0;
    VISTATUS = viPrintf(*vi, "OUTP?\n"); // Query the signal generator's RF state
    if (VISTATUS)
        return false;   //连接错误则返回未打开状态
    viScanf(VI, "%1i", &num); // Read the response
    // Confirm the on/off RF state
    if (num > 0) {
        return true;
    }
    else{
        return false;
    }
}
bool VIOpSignalGenerator::SetAmpl(QString ampl)
{
    QString s1="POW:AMPL "+ampl+"dBm\n";
    QByteArray ba=s1.toLatin1();
    ViConstString com=ba.data();
    VISTATUS = viPrintf(VI, com);
    return viop_StatusCheck();

}
char* VIOpSignalGenerator::QuestFre()
{
    char *str = nullptr;
    char *rdBuffer = (char*)malloc(256 * sizeof(char));
    int pos = 0;
    int count = 0;
    VISTATUS = viPrintf(VI, "POW:AMPL?\n");
    //:AMPLitude:VALue?
    //:FREQuency?
    if (VISTATUS != 1)
    {
        VISTATUS = viScanf(VI, "%t", rdBuffer);
        while (rdBuffer[pos] != '\0')
        {
            pos++;
        }
        count = pos + 1;
        str = (char*)malloc(count*sizeof(char));
        for (int i = 0; i < count; i++)  //copy rdBuffer to str including '\0'
        {
            *(str + i) = rdBuffer[i];
        }
    }
    return str;
}

bool VIOpSignalGenerator::SetFre(QString fre)
{
    QString command=":FREQuency "+fre+"gHz\n";
    QByteArray ba=command.toLatin1();
    ViConstString com=ba.data();
    VISTATUS = viPrintf(VI, com);
    return viop_StatusCheck();
}

char* VIOpSignalGenerator::viop_Q_BertResult1()
{
    char *str = nullptr;
    char *rdBuffer = (char*)malloc(256 * sizeof(char));
    int pos = 0;
    int count = 0;
    VISTATUS = viPrintf(VI, ":DATA:BERT? BER\n");
    if (VISTATUS != 1)
    {
        VISTATUS = viScanf(VI, "%t", rdBuffer);
        while (rdBuffer[pos] != '\0')
        {
            pos++;
        }
        count = pos + 1;
        str = (char*)malloc(count*sizeof(char));
        for (int i = 0; i < count; i++)  //copy rdBuffer to str including '\0'
        {
            *(str + i) = rdBuffer[i];
        }
    }
    return str;
}

// 返回BERT结果
char* VIOpSignalGenerator::viop_Q_BertResult()
{
    char *str = nullptr;
    char *rdBuffer = (char*)malloc(256 * sizeof(char));
    int pos = 0;
    int count = 0;
    VISTATUS = viPrintf(VI, ":DATA:BERT? ALL\n");
    if (VISTATUS != 1)
    {
        VISTATUS = viScanf(VI, "%t", rdBuffer);
        while (rdBuffer[pos] != '\0')
        {
            pos++;
        }
        count = pos + 1;
        str = (char*)malloc(count*sizeof(char));
        for (int i = 0; i < count; i++)  //copy rdBuffer to str including '\0'
        {
            *(str + i) = rdBuffer[i];
        }
    }
    return str;
}
char* VIOpSignalGenerator::viop_Q_rfFq(char *rdBuffer)
{
    viPrintf(VI, "FREQ:CW?\n"); // Querys the CW frequency
    VISTATUS = viScanf(VI, "%t", rdBuffer); // Reads response into rdBuffer
    if (*viStatus)
        return nullptr;
    else
        return rdBuffer;
}
bool VIOpSignalGenerator::viop_Q_amONOFF()
{
    int num;
    VISTATUS = viPrintf(VI, "AM:STAT?\n"); // Query the signal generator's RF state
    if (*viStatus)
        return false;
    viScanf(VI, "%1i", &num); // Read the response
    // Confirm the on/off RF state
    if (num > 0) {
        return true;
    }
    else{
        return false;
    }
}

char* VIOpSignalGenerator::viop_Q_rfAmpl(char *rdBuffer)
{
    viPrintf(VI, "POW:AMPL?\n"); // Querys the CW frequency
    VISTATUS = viScanf(VI, "%t", rdBuffer); // Reads response into rdBuffer
    if (*viStatus)
        return nullptr;
    else
        return rdBuffer;
}

char* VIOpSignalGenerator::viop_Q_freqMode(char *rdBuffer)
{
    viPrintf(VI, "FREQ:MODE?\n"); // Querys the CW frequency
    VISTATUS = viScanf(VI, "%t", rdBuffer); // Reads response into rdBuffer
    if (*viStatus)
        return nullptr;
    else
        return rdBuffer;
}


void Delay_MSec(unsigned int msec)
{
    QEventLoop loop;                              //定义一个新的事件循环
    QTimer::singleShot(msec, &loop, SLOT(quit()));//创建单次定时器，槽函数为事件循环的退出函数
    loop.exec();                                  //事件循环开始执行，程序会卡在这里，直到定时时间到，本循环被退出
}




