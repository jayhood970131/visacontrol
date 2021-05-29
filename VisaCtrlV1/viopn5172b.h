#ifndef VIOPN5172B_H
#define VIOPN5172B_H

#include "viopsignalgenerator.h"
class VIOpN5172B:public VIOpSignalGenerator
{
public:
    VIOpN5172B(QObject *parent=nullptr,VIOPHead *head=nullptr);
};

#endif // VIOPN5172B_H
