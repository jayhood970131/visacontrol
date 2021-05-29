#include "vioperation.h"

VIOperation::VIOperation(QObject *parent,VIOPHead *head):QObject(parent)
{
    defaultRM = head->defaultRM;
    vi = head->vi;
    viStatus = head->viStatus;
}

/*
    Session init command
*/

bool VIOperation::viop_viOpenDefaultRM()
{
    VISTATUS = viOpenDefaultRM(defaultRM);
    return viop_StatusCheck();
}

bool VIOperation::viop_viOpen(char *address)
{
    VISTATUS = viOpen(*defaultRM, address, VI_NULL, VI_NULL, vi);
    return viop_StatusCheck();
}
bool VIOperation::viop_CLEAR()    // Clears the signal generator
{
    VISTATUS = viClear(VI);
    return viop_StatusCheck();
}
/*
    command operation
*/
bool VIOperation::viop_StatusCheck()
{
    if (VISTATUS!=0)
        return false;
    else
        return true;
}
bool VIOperation::viop_sendCommand(QString *command)  //发送一条命令 参数使用指针是为了对同一个CString进行操作，以免因形参和实参的复制出现问题
{
    char * cmd = command->toUtf8().data();
    VISTATUS = viPrintf(VI, cmd);
    if (VISTATUS==1)
        return false;
    else
        return true;
}

/*
    basic command
*/
bool VIOperation::viop_RST()
{
    //ViUInt32 retCnt;  //使用tekvisa测试
    VISTATUS = viPrintf(VI, "*RST\n");         // Initializes signal generator
    //VISTATUS = viWrite(VI, (ViBuf)"*RST", 4, &retCnt);  //tekvisa测试
    return viop_StatusCheck();
}
bool VIOperation::viop_WAI()
{
    VISTATUS = viPrintf(VI, "*WAI\n");  //收到*WAI命令后，在执行新的命令之前，函数发生器等待所有未执行的命令。
    return viop_StatusCheck();
}
bool VIOperation::viop_TRG()
{
    VISTATUS = viPrintf(VI, "*TRG\n");   //只有trigger source 设置为BUS时起作用。其余情况均忽略该命令
    return viop_StatusCheck();
}

bool VIOperation::viop_CLS()
{
    VISTATUS = viPrintf(VI, "*CLS\n"); // Clears the status byte register
    return viop_StatusCheck();
}
