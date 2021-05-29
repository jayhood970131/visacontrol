#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include <QMessageBox>
#include <QPushButton>
#include <process.h>
#include "QProcess"
#include "viop33250a.h"
#include "viopn5172b.h"
#include "viopr3408b.h"
#include "viodpo7254.h"

namespace Ui {
class MainWindow;
}

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    explicit MainWindow(QWidget *parent = 0);
    ~MainWindow();
public:
    static VIOpN5172B * ptr_N5172B;     //使用这个指针来完成对仪器的各种操作
    static VIOp33250A * ptr_33250A;
    static VIOpR3408B * ptr_R3408B;
    static VIOpD7254 *  ptr_D7254;
private slots:
    //void on_EditFreUp_cursorPositionChanged(int arg1, int arg2);

private:
    Ui::MainWindow *ui;

    //各类标志位
    bool isN5172BConnected = false;
    bool is33250AConnected = false;
    bool isR3408BConnected = false;
    bool isD7245Connected = false;
    bool isRFOn=false;
    bool isAMOn=true;
    bool bertKey=true;

//
    const QString addr_N5172B = "GPIB0::19::INSTR";
    const QString addr_33250A = "GPIB0::10::INSTR";
    const QString addr_R3408B = "TCPIP::192.168.1.6::INSTR";
    const QString addr_DPO7254= "TCPIP::192.168.1.15::INSTR";
 //
//    const QString  addr_DPO7254="Variable:Value \"SetRemoteApplication\",\"DPOPWR\"";
};
#endif // MAINWINDOW_H


























