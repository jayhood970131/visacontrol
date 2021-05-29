#ifndef VIOPSIGNALGENERATOR_H
#define VIOPSIGNALGENERATOR_H
#include "vioperation.h"
#include <synchapi.h>
#include <QCoreApplication>
#include <QTimer>
#include <QDialog>
#include <QFileDialog>
#include <qaxobject.h>
//#include <visa.h>
#include <stdio.h>
#include <stdlib.h>
#include <assert.h>

class VIOpSignalGenerator:public VIOperation
{
public:
    explicit VIOpSignalGenerator(QObject *parent = nullptr,VIOPHead *head = nullptr);

//      bool exportToExcel();

    /*
        BERT command
        */
    bool viop_displayMode(int type);
    bool viop_alpha();
    bool viop_alphaDepth();
    bool viop_BERTtriggerSource(int source); //触发模式--IMM,EXT,BUS
    bool viop_symbolRate(int symRat);
    bool viop_modulationType();
    bool viop_bertOn();
    bool viop_bertOff();
    bool viop_realtimeOn();
    bool viop_realtimeOff();
    bool viop_loadData();
    bool viop_totalBit(QString tem);
    bool viop_loadDataWU();
    /*
        CW command
        */
    bool viop_cwFreq(QString frequency, int Unit);
    bool viop_cwAmpl(QString amplitude);
    bool viop_RFOn();
    bool viop_RFOff();
    bool viop_MODOn();
    bool viop_MODOff();
    bool viop_freqMode(int mode);
    bool viop_listType(int type);
    bool viop_funcType(int type);
    /*
        FM/AM command
        */
    bool viop_AMSrc(int src);
    bool viop_AMONOFF(bool);
    bool viop_AMDepth(QString depth);
    bool viop_AMWBONOFF(bool);
      /*
        Quries command
        */
    char* viop_Q_BertResult();
    bool viop_Q_bertONOFF();
    bool viop_Q_realtimeONOFF();
    bool viop_Q_rfONOFF();
    char* viop_Q_rfFq(char*);
    bool viop_Q_amONOFF();
    char* viop_Q_rfAmpl(char*);
    char* viop_Q_freqMode(char*);
    char* viop_Q_BertResult1();
    char* QuestFre();
    bool SetFre(QString fre);
    bool SetAmpl(QString ampl);
};


void Delay_MSec(unsigned int msec);
#endif // VIOPSIGNALGENERATOR_H
