#include "viop33250a.h"

VIOp33250A::VIOp33250A(QObject *parent,VIOPHead *head):VIOpSignalGenerator(parent,head)
{

}
bool VIOp33250A::viop33250A_BurstStateONOFF(int state)
{
    switch (state)
    {
    case OFF:
        VISTATUS = viPrintf(VI, "BURSt:STATe OFF\n");
        break;
    case ON:
        VISTATUS = viPrintf(VI, "BURSt:STATe ON\n");
        break;
    default:
        break;
    }
    return viop_StatusCheck();
}
bool VIOp33250A::viop33250A_BurstNcycles(QString cycles)
{
    QString str_1("BURSt:NCYCles ");
    QString strEnter("\n");
    QString command = str_1 + cycles + strEnter;
    viop_sendCommand(&command);
    return viop_StatusCheck();
}
bool VIOp33250A::viop33250A_TriggerSource(int source)
{
    switch (source)
    {
    case IMM:
        VISTATUS = viPrintf(VI, "TRIGger:SOURce IMM\n");
        break;
    case EXT:
        VISTATUS = viPrintf(VI, "TRIGger:SOURce EXT\n");
        break;
    case BUS:
        VISTATUS = viPrintf(VI, "TRIGger:SOURce BUS\n");
        break;
    default:
        break;
    }
    return viop_StatusCheck();
}
