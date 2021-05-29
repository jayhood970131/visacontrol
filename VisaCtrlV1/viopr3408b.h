#ifndef VIOPR3408B_H
#define VIOPR3408B_H
#include "viopsignalgenerator.h"

class VIOpR3408B:public VIOpSignalGenerator
{
public:
  explicit VIOpR3408B(QObject *parent=0,VIOPHead *head=0);
    
     
 //void OnBnClickedTxtextEnterspanButton(QString fre,int nSel);

 bool viop_span(QString width, int Unit);
 bool viop_CentralFre(QString fre,int Unit);
 bool viop_markerX_ONOFF();  //显示/关闭MARKER
 bool viop_markerMovePeak();
 bool viop_markerMovePeak1();
 bool viop_markerMovePeakl();
 char* viop_ReturnTestPower();
 bool viop_startfreq(QString startfreq, int Unit);
 bool viop_stopfreq(QString startfreq, int Unit);
 QString charTostring(char* addr);

};
double charTodouble(char* result);
#endif // VIOPR3408B_H
