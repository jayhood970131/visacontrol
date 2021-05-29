#ifndef VIODPO7254_H
#define VIODPO7254_H
#include "viopsignalgenerator.h"
class VIOpD7254:public VIOpSignalGenerator
{
public:
    VIOpD7254(QObject *parent=nullptr,VIOPHead *head=nullptr);
    char* viop_DPOQueValueMean();
    char* viop_Test();
    bool viop_VerticalScale();
    bool viop_HorizeScale5();
    bool viop_HorizeScale2();
    bool viop_set();
    bool viop_setHigh();
};
#endif // VIODPO7254_H
