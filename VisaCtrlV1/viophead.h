#ifndef VIOPHEAD_H
#define VIOPHEAD_H

#include <QObject>
#include <include/visa.h>

class VIOPHead : public QObject
{
    Q_OBJECT
public:
    explicit VIOPHead(QObject *parent = nullptr);

    ViSession *defaultRM;
    ViSession *vi;
    ViStatus *viStatus;
signals:
public slots:
};


#endif // VIOPHEAD_H
