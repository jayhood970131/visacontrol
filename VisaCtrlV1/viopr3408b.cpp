#include "viopr3408b.h"
#include <QMessageBox>
#include <QDebug>
VIOpR3408B::VIOpR3408B(QObject *parent,VIOPHead *head):VIOpSignalGenerator(parent,head)
{

}

//设置开始频率
bool VIOpR3408B::viop_startfreq(QString startfreq, int Unit)
{
    QString str_1("FREQuency:STARt ");
    QString str_khz(" kHz\n");
    QString str_mhz(" mHz\n");
    QString str_ghz(" gHz\n");
    QString command = nullptr;
    switch (Unit)
    {
    case KHz:
        command = str_1 + startfreq + str_khz;
        break;
    case MHz:
        command = str_1 + startfreq + str_mhz;
        break;
    case GHz:
        command = str_1 + startfreq + str_ghz;
        break;
    default:
        command = "";
        break;
    }
    viop_sendCommand(&command);
    return viop_StatusCheck();
}

//设置停止频率
bool VIOpR3408B::viop_stopfreq(QString startfreq, int Unit)
{
    QString str_1("FREQuency:STOP ");
    QString str_khz(" kHz\n");
    QString str_mhz(" mHz\n");
    QString str_ghz(" gHz\n");
    QString command = nullptr;
    switch (Unit)
    {
    case KHz:
        command = str_1 + startfreq + str_khz;
        break;
    case MHz:
        command = str_1 + startfreq + str_mhz;
        break;
    case GHz:
        command = str_1 + startfreq + str_ghz;
        break;
    default:
        command = "";
        break;
    }
    viop_sendCommand(&command);
    return viop_StatusCheck();
}


//设置中心频率
bool VIOpR3408B::viop_CentralFre(QString fre,int index)
{
    int unit;
    // 获取组合框控件的列表框中选中项的索引
    switch (index)
    {
    case 0:
        unit = GHz;
        break;
    case 1:
        unit = MHz;
        break;
    case 2:
        unit = KHz;
        break;
    default:
        unit = GHz;
        break;
    }
    QString str_1("FREQuency:CENTEr ");
    QString str_khz(" kHz\n");
    QString str_mhz(" mHz\n");
    QString str_ghz(" gHz\n");
    QString command=nullptr;
    switch (unit)
    {
    case KHz:
        command = str_1 + fre + str_khz;
        break;
    case MHz:
        command = str_1 + fre + str_mhz;
        break;
    case GHz:
        command = str_1 + fre + str_ghz;
        break;
    default:
        command = "";
        break;
    }
    viop_sendCommand(&command);
    return viop_StatusCheck();

}


bool VIOpR3408B::viop_span(QString width, int nSel)  //设置频谱分析仪span参数
{

    int unit;
    // 获取组合框控件的列表框中选中项的索引
    switch (nSel)
    {
    case 0:
        unit = GHz;
        break;
    case 1:
        unit = MHz;
        break;
    case 2:
        unit = KHz;
        break;
    default:
        unit = GHz;
        break;
    }

    QString str_1("FREQuency:SPAN ");
    QString str_khz(" kHz\n");
    QString str_mhz(" mHz\n");
    QString str_ghz(" gHz\n");
    QString command=nullptr;
    switch (unit)
    {
    case KHz:
        command = str_1 + width + str_khz;
        break;
    case MHz:
        command = str_1 + width + str_mhz;
        break;
    case GHz:
        command = str_1 + width + str_ghz;
        break;
    default:
        command = "";
        break;
    }
    viop_sendCommand(&command);
    return viop_StatusCheck();
}

//打开MAKER1
bool VIOpR3408B::viop_markerX_ONOFF()  //显示/关闭MARKER
{
    QString str_1(":CALCulate1:");
    QString str_mark1("MARKer1:");
 //   QString str_mark2("MARKer2:");
    QString str_stateOn("STATe ON\n");
//    QString str_stateOff("STATe OFF\n");
    QString command = nullptr;
    command = str_1 + str_mark1 + str_stateOn;
    viop_sendCommand(&command);
    return viop_StatusCheck();
}

//将标记放在轨迹上的最大值点
bool VIOpR3408B::viop_markerMovePeak()
{
    QString str_1(":CALCulate1:");
    QString marker1("MARKer1:");
    QString Tomax("MAXimum\n");
    QString command = nullptr;
    command = str_1 + marker1 + Tomax;
    viop_sendCommand(&command);
    return viop_StatusCheck();
}

//将中心频率设置为标记位置的值。
bool VIOpR3408B::viop_markerMovePeak1()
{
    QString str_1(":CALCulate1:");
    QString marker1("MARKer1:");
    QString Tomax("[:SET]:CENTer\n");   //
    QString command = nullptr;
    command = str_1 + marker1 + Tomax;
    viop_sendCommand(&command);
    return viop_StatusCheck();
}

//返回数据
char* VIOpR3408B::viop_ReturnTestPower()
{
    char *str = nullptr;
    char *rdBuffer= (char*)malloc(30 * sizeof(char));
    char c='*';
    char*ch=&c;

    VISTATUS =viPrintf(VI, ":CALCulate1:MARKer1:Y?\n");

        if (VISTATUS != 1)
        {
            viScanf(VI, "%t", rdBuffer);
            str = (char*)malloc(13*sizeof(char));
            for (int i = 0; i < 13; i++)  //copy rdBuffer to str including '\0'
            {
                *(str + i) = rdBuffer[i];
            }

            *(str + 13) = '\0';
        }
        else
        {
            return ch;
        }

        return str;
}


QString charTostring(char* addr)
{
    int i=0;
    QString data;
    while(*(addr+i)!='\0')
    {
        data[i]=*(addr+i);
        i++;
    }
    return data;
}

double charTodouble(char* result)
{
    int i=0;
    double b;
    QString data;
    QString a;
    while(*(result+i)!='\0')
     {
         data[i]=*(result+i);
         i++;
     }
    a = data.mid(0,6);
    b=a.toDouble();
    b=b*10;
    return b;
}
























