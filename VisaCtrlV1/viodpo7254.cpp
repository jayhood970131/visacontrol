#include "viodpo7254.h"
#include <QMessageBox>
#include <QDebug>
VIOpD7254::VIOpD7254(QObject *parent,VIOPHead *head):VIOpSignalGenerator(parent,head)
{

}

char* VIOpD7254::viop_DPOQueValueMean()
{
    char *str = nullptr;
    char *rdBuffer= (char*)malloc( 200* sizeof(char));
    char c='*';
    char*ch=&c;
    int i = 0;

    VISTATUS =viPrintf(VI, "MEASUrement:ANNOTation:Y1?\n");
    if ( VISTATUS!=1)
    {
        viScanf(VI, "%s",rdBuffer);
        str = (char*)malloc(200*sizeof(char));
        for (;i<14;i++)                                     //copy rdBuffer to str including '\0'
        {
                *(str + i) = rdBuffer[i];
        }
        *(str + i) = '\0';
    }
    else
    {
        return ch;
    }
    free(rdBuffer);
     return str;
}
char* VIOpD7254::viop_Test()
{

    char *str = nullptr;
    char *rdBuffer= (char*)malloc(15 * sizeof(char));
    char c='*';
    char*ch=&c;
    int i = 0;
    //MEASUrement:ANNOTation:Y<x>?
    VISTATUS =viPrintf(VI, "MEASUrement:ANNOTation:Y1?\n");
    if ( VISTATUS!=1)
    {
        viScanf(VI, "%s",rdBuffer);
        str = (char*)malloc(15*sizeof(char));
        for (;i<14;i++)                                  //copy rdBuffer to str including '\0'
        {
                *(str + i) = rdBuffer[i];
        }
        *(str + i) = '\0';
    }
    else
    {
        return ch;
    }
    free(rdBuffer);
     return str;
}

//设置vertical scale
bool VIOpD7254::viop_VerticalScale()
{
    VISTATUS = viPrintf(VI, "CH1:SCALE  1.0000E+00\n");  //设置scale1.0V
    if (VISTATUS==1)
        return false;
    else
        return true;
}

//设置horize scale
//HORIZONTAL:MODE:SCALE2e-9 sets the horizontal scale to 2 ns per division.
bool VIOpD7254::viop_HorizeScale5()
{
    VISTATUS = viPrintf(VI, "HORIZONTAL:MODE:SCALE 500e-3\n");  //设置500ms
    if (VISTATUS==1)
        return false;
    else
        return true;
}

bool VIOpD7254::viop_HorizeScale2()
{   
    VISTATUS = viPrintf(VI, "HORIZONTAL:MODE:SCALE 2e-3\n");  //设置2ms
    if (VISTATUS==1)
        return false;
    else
        return true;
}


bool VIOpD7254::viop_set()
{
    VISTATUS = viPrintf(VI, "MEASUREMENT:IMMED:REFLEVEL:METHOD ABSOLUTE\n");
    return viop_StatusCheck();
}

//MEASUREMENT:IMMED:NOISEHIGH
bool VIOpD7254::viop_setHigh()
{
    VISTATUS = viPrintf(VI, "MEASUREMENT:IMMED:TYPE HIGH\n");
    return viop_StatusCheck();
}





















