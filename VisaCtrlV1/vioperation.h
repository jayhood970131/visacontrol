#ifndef VIOPERATION_H
#define VIOPERATION_H

#define DEFAULTRM (*defalutRM)
//#define Visession VI

#define VI (*vi)
#define VISTATUS (*viStatus)

#include <QObject>
#include "vipara.h"
#include "viparaunit.h"
#include "viophead.h"
#include "include/visa.h"
#include <QObject>

class VIOperation : public QObject
{
    Q_OBJECT
public:
    explicit VIOperation(QObject *parent = 0 ,VIOPHead *head = 0);

    ViSession *defaultRM;
    ViSession *vi;
    ViStatus *viStatus;

        /*
            Session command
        */
    bool viop_viOpenDefaultRM();
    bool viop_viOpen(char *address);
    bool viop_CLEAR();
        /*
            command operation
        */
    bool viop_StatusCheck();                  //true--操作正确 false--操作错误
    bool viop_sendCommand(QString *command);  //ture--发送成功 false--发送失败
    bool viop_sendCommand_viwrite(QString command); //使用viwrite（）函数
        /*
            basic command 以*开头的各种命令
        */
    bool viop_RST();
    bool viop_WAI();
    bool viop_TRG();
    bool viop_CLS();
signals:

public slots:
};

#endif // VIOPERATION_H
