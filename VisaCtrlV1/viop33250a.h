#ifndef VIOP33250A_H
#define VIOP33250A_H

#include "viopsignalgenerator.h"
class VIOp33250A:public VIOpSignalGenerator
{
public:
    VIOp33250A(QObject *parent=0,VIOPHead *head=0);
    bool viop33250A_BurstStateONOFF(int state);
    bool viop33250A_BurstNcycles(QString cycles);
    bool viop33250A_TriggerSource(int source);
};

#endif // VIOP33250A_H
