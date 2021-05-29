/*
 * 传感器核心硬件关键参数嵌入式自动测量软件
 * 2020-11-30
 * Chasny Chang
 *
*/
#include "mainwindow.h"
#include "ui_mainwindow.h"
#include "vioneceoperate.h"
#include "viopr3408b.h"
#include <stdlib.h>
#include <QDebug>
#include "libxl.h"
//#define _Q(s) (QString::fromLocal8Bit(s))
static char *result =nullptr;
static int chipNum=1;
static QString str;
static QString str_totalBits;
static QString str_errorBits;
static QString str_BER;
static char *p=nullptr;
//static char *d = ",";
static int AmplLower,AmplUpper,AmplStep=1;
static double FreDown,FreUp;
//,FreStep;
static double BerStand;
static QString ToalBits;
static int row=2;                            //接收灵敏度测试数据开始填写行号
static QString windowText="已连接";
static int Txrow=1;
static bool TestFirst=true;
static bool RxMultTextFirst=true;
static bool RxMulPotsTestFirst=true;
static bool RxFreRange=true;
static bool RxFreRangeMul=true;
static bool RxMaxFirst=true;
static bool TestFirstWu=true;
static bool TestFirstWuFre=true;
static bool TestFirstWuMax=true;
static QString BoardNum;
static bool EnSpiDsrc_HI = false;
static QString filepath;
static QAxObject *workbook;
static QAxObject *excel = new QAxObject();
static QAxObject *worksheet ;
static QAxObject *worksheets;
static QAxObject *workbooks;

VIOpN5172B* MainWindow::ptr_N5172B = nullptr;    //静态成员初始化
VIOp33250A* MainWindow::ptr_33250A = nullptr;
VIOpR3408B* MainWindow::ptr_R3408B=nullptr;
VIOpD7254*  MainWindow::ptr_D7254=nullptr;

static QSerialPort *serial;
static int PortOneFrameSize=100;

using namespace libxl;
//static Book* book;
//static Sheet* sheet;

MainWindow::MainWindow(QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::MainWindow)
{
    ui->setupUi(this);
    this->setWindowTitle("传感器核心硬件关键参数嵌入式自动测量");
    QSize size = this->size();
    this->setFixedSize(size);

    QStringList listFreq;
    listFreq<<"5.83 GHz"<<"5.84 GHz";

    //状态栏 最多有一个
    QStatusBar * stBar = statusBar();
    //设置到窗口中
    setStatusBar(stBar);
    //放标签控件
    QLabel * labelLinkStatus = new QLabel("连接状态：N5712B-未建立连接 DPO7254-未建立连接 RSA3408B-未建立连接",this);
    stBar->addWidget(labelLinkStatus);

    VIOPHead *head = new VIOPHead();   //N5172B的会话
    head->defaultRM = new ViSession();
    head->vi = new ViSession();
    head->viStatus = new ViStatus();
    MainWindow::ptr_N5172B = new VIOpN5172B(parent,head);
    MainWindow::ptr_N5172B->viop_viOpenDefaultRM();
    delete head;

    head = new VIOPHead();         //33250A的会话
    head->defaultRM = new ViSession();
    head->vi = new ViSession();
    head->viStatus = new ViStatus();
    MainWindow::ptr_33250A = new VIOp33250A(parent,head);
    MainWindow::ptr_33250A->viop_viOpenDefaultRM();

    head = new VIOPHead();          //3408B的会话
    head->defaultRM = new ViSession();
    head->vi = new ViSession();
    head->viStatus = new ViStatus();
    MainWindow::ptr_R3408B = new VIOpR3408B(parent,head);
    MainWindow::ptr_R3408B->viop_viOpenDefaultRM();

    head = new VIOPHead();              //7254的会话
    head->defaultRM = new ViSession();
    head->vi = new ViSession();
    head->viStatus = new ViStatus();
    MainWindow::ptr_D7254 = new VIOpD7254(parent,head);
    MainWindow::ptr_D7254->viop_viOpenDefaultRM();


    isN5172BConnected=MainWindow::ptr_N5172B->viop_viOpen(addr_N5172B.toUtf8().data());
    is33250AConnected=MainWindow::ptr_33250A->viop_viOpen(addr_33250A.toUtf8().data());
    isR3408BConnected=MainWindow::ptr_R3408B->viop_viOpen(addr_R3408B.toUtf8().data());
    isD7245Connected =MainWindow::ptr_D7254->viop_viOpen(addr_DPO7254.toUtf8().data());

     if(isN5172BConnected)
     {
         windowText="N5172B信号源 "+windowText;
         labelLinkStatus->setText(windowText);
     }
     if(isR3408BConnected)
     {
         windowText="R3408B频谱仪 "+windowText;
         labelLinkStatus->setText(windowText);
     }

     if(isD7245Connected)
     {
         windowText="D7254示波器 "+windowText;
         labelLinkStatus->setText(windowText);
     }

     ui->EditBoard->setValidator(new QRegExpValidator(QRegExp("[0-9]+$")));     //限制只能输入数字
     ui->EditamplUP->setValidator(new QIntValidator(-100,100,this));
     ui->EditamplLW->setValidator(new QIntValidator(-100,100,this));

    //设置table widget表格自适应
    ui->tblWidget->horizontalHeader()->setSectionResizeMode(QHeaderView::Stretch);
    ui->tblWidget->verticalHeader()->setSectionResizeMode(QHeaderView::Stretch);
    ui->tblWidget->setVerticalHeaderLabels(QStringList()<<"频率"<<"功率");          //设置表头
    QStringList frepots;
    frepots<<"5.8258"<<"5.8283"<<"5.8315"<<"5.8316"<<"5.8317"<<"5.8383"<<"5.8385"<<"5.8415"<<"5.8417";
    for(int i=0;i<9;i++)
    {
       ui->tblWidget->setItem(0,i,new QTableWidgetItem(frepots[i]));
    }

     //Tx输出功率范围设置
      ui->cbxTxFreUnits->addItem("GHz");
      ui->cbxTxFreUnits->addItem("MHz");
      ui->cbxTxFreUnits->addItem("KHz");
      ui->cbxTxSpanUnits->addItem("GHz");
      ui->cbxTxSpanUnits->addItem("MHz");
      ui->cbxTxSpanUnits->addItem("KHz");
      ui->cbxTxStarFreUnits->addItem("GHz");
      ui->cbxTxStarFreUnits->addItem("MHz");
      ui->cbxTxStarFreUnits->addItem("KHz");
      ui->cbxTxStopFreUnits->addItem("GHz");
      ui->cbxTxStopFreUnits->addItem("MHz");
      ui->cbxTxStopFreUnits->addItem("KHz");

      //选择文件
      connect(ui->btnTxChoseExcel,&QPushButton::clicked,[&](){

          filepath =QFileDialog::getOpenFileName(this,("选择文件"),("C:\\Users\\ztp\\Desktop"),("*.xls *.xlsx *.csv *.xlsm")); //获取保存路径
          if(filepath.isEmpty())
          {
              QMessageBox::critical(this, "错误信息", "没有找到EXCEL");
                  return;
          }
           excel->setControl("ket.Application");                   //连接WPS
           excel->dynamicCall("SetVisible(bool Visible)","false"); //不显示窗体
           excel->setProperty("DisplayAlerts", false);             //不显示任何警告信息。如果为true那么在关闭是会出现类似“文件已修改，是否保存”的提示
           workbooks = excel->querySubObject("WorkBooks");         //获取工作簿集合
           workbooks->dynamicCall("Open(const QString&)",QDir::toNativeSeparators(filepath));  //打开打开已存在的工作簿
           workbook = excel->querySubObject("ActiveWorkBook");     //获取当前工作簿
           worksheets = workbook->querySubObject("WorkSheets");     //获取工作表集合
           worksheet = worksheets->querySubObject("Item(int)",1);   //获取工作表集合的工作表1，即sheet1
           QMessageBox::information(this,"提示","选择成功");
      });

      //写入数据
      connect(ui->btnTxWrite,&QPushButton::clicked,[&](){
           QString BoardInfo=ui->EditTxBoardNum->text()+"号板";

           ptr_R3408B->VIOpR3408B::viop_markerMovePeak();                               //marker移动到peak
           result=ptr_R3408B->VIOpR3408B::viop_ReturnTestPower();
           double c=charTodouble(result);
           QAxObject *cellA = worksheet->querySubObject("Cells(int,int)",Txrow,1);
           cellA->dynamicCall("SetValue(const QVariant&)",QVariant(BoardInfo));
           QAxObject *cellB = worksheet->querySubObject("Cells(int,int)",Txrow,2);
           cellB->dynamicCall("SetValue(const QVariant&)",QVariant(c));
           workbook->dynamicCall("SaveAs(const QString&)",QDir::toNativeSeparators(filepath));
           Txrow++;
           QMessageBox::information(this,"提示","数据已写入");
      });

      //关闭文件
      connect(ui->btnTxEnd,&QPushButton::clicked,[&]()
      {
          int numchip=ui->EditTxBoardNum->text().toInt();
          numchip++;
          ui->EditTxBoardNum->setText(QString::number(numchip));
          //关闭WPS
          workbook->dynamicCall("SaveAs(const QString&)",QDir::toNativeSeparators(filepath));
          workbook->dynamicCall("Close(boolean)",false);                              //关闭工作簿
          excel->dynamicCall("Quit(void)");                                  //关闭excel
          delete excel;
          excel=nullptr;
          QMessageBox::information(this,"提示","Tx测试结束");
      });

      connect(ui->btnWuSet,&QPushButton::clicked,[&]()
      {
          //CH1:SCALE  1.0000E+00 1V
          ptr_D7254->VIOpD7254::viop_VerticalScale();
         //设置horize Scale 500ms  2.0ms
//          ptr_D7254->VIOpD7254::viop_HorizeScale5();
          ptr_D7254->VIOpD7254::viop_HorizeScale2();
          //set high
          ptr_D7254->VIOpD7254::viop_set();
      });


      connect(ui->btnWuN5172B,&QPushButton::clicked,[&](){
          //设置信号源配置
          QString amplWu=ui->EditAmplWuUp->text();
          QString freWu=ui->EditWuFre->text();
          ptr_N5172B->VIOpSignalGenerator::viop_realtimeOn();                  //realtime
          ptr_N5172B->VIOpSignalGenerator::viop_modulationType();              //设置ASK
          ptr_N5172B->VIOpSignalGenerator::viop_alpha();                       //设置阿尔法为1
          ptr_N5172B->VIOpSignalGenerator::viop_alphaDepth();                  //85%
          ptr_N5172B->VIOpSignalGenerator::viop_symbolRate(28);               //symbol rate 默认单位Ksps
          ptr_N5172B->VIOpSignalGenerator::viop_loadDataWU();                  //load文件14K_TEST_28@BIT
          ptr_N5172B->VIOpSignalGenerator::viop_displayMode(SCIENTIFIC);       //科学计数法显示误码率
          ptr_N5172B->VIOpSignalGenerator::viop_BERTtriggerSource(IMM);        //trigger-immediate
          ptr_N5172B->VIOpSignalGenerator::SetAmpl(amplWu);                  //设置ampl-50
          ptr_N5172B->VIOpSignalGenerator::SetFre(freWu);
          ptr_N5172B->VIOpSignalGenerator::viop_MODOn();
          ptr_N5172B->VIOpSignalGenerator::viop_RFOn();                        //打开RF
//          ptr_N5172B->VIOpSignalGenerator::viop_bertOn();                      //打开
      });


      //唤醒灵敏度测试
      connect(ui->btnWuTest,&QPushButton::clicked,[&](){
          //测试步骤
          QString filepathWu;
          filepathWu =QFileDialog::getOpenFileName(nullptr,("选择文件"),("C:/Users/ztp/Desktop"),("*.xls *.xlsx *.csv")); //获取保存路径
          if(filepathWu.isEmpty())
          {
              QMessageBox::critical(nullptr, "错误信息", "没有找到EXCEL");
                  return;
          }
          QAxObject *excelwu = new QAxObject();
          excelwu->setControl("ket.Application");                                     //连接WPS
          excelwu->dynamicCall("SetVisible(bool Visible)","false");                   //不显示窗体
          excelwu->setProperty("DisplayAlerts", false);                               //不显示任何警告信息。如果为true那么在关闭是会出现类似“文件已修改，是否保存”的提示
          QAxObject* workbookswu = excelwu->querySubObject("WorkBooks");                //获取工作簿集合
          workbookswu->dynamicCall("Open(const QString&)",filepathWu);                  //打开打开已存在的工作簿
          QAxObject *workbookwu = excelwu->querySubObject("ActiveWorkBook");            //获取当前工作簿
          QAxObject *worksheetswu = workbookwu->querySubObject("WorkSheets");           //获取工作表集合
          QAxObject *worksheetwu = worksheetswu->querySubObject("Item(int)",1);         //获取工作表集合的工作表1，即sheet1
          QAxObject *cellA;
          if(TestFirstWu)
          {
               cellA = worksheetwu->querySubObject("Cells(int,int)",1,1);
               cellA->dynamicCall("SetValue(const QVariant&)","No.");         //写入数据
               cellA->setProperty("HorizontalAlignment", -4108);                     //居中
               cellA->setProperty("VerticalAlignment", -4108);                       //居中
               cellA = worksheetwu->querySubObject("Cells(int,int)",1,2);
               cellA->dynamicCall("SetValue(const QVariant&)","5.83");         //写入数据
               cellA->setProperty("HorizontalAlignment", -4108);                     //居中
               cellA->setProperty("VerticalAlignment", -4108);                       //居中
               cellA = worksheetwu->querySubObject("Cells(int,int)",1,3);
               cellA->dynamicCall("SetValue(const QVariant&)","5.84");         //写入数据
               cellA->setProperty("HorizontalAlignment", -4108);                     //居中
               cellA->setProperty("VerticalAlignment", -4108);                       //居中
               TestFirstWu=false;
          }

          cellA = worksheetwu->querySubObject("Cells(int,int)",row,1);
          cellA->dynamicCall("SetValue(const QVariant&)",ui->EditBoardWu->text());   //芯片号
          cellA->setProperty("HorizontalAlignment", -4108);                     //居中
          cellA->setProperty("VerticalAlignment", -4108);                       //居中
           int stringminus=ui->Wuline->text().toInt();
           ptr_N5172B->VIOpSignalGenerator::SetAmpl(ui->EditAmplWuUp->text());
           Delay_MSec(2000);
           result=ptr_D7254->VIOpD7254::viop_DPOQueValueMean();      //求出第一个值
           double standResult=QString(result).toDouble();
           int up=ui->EditAmplWuUp->text().toInt();
           int down=ui->EditAmplWuDown->text().toInt();
           int stp=ui->EditWuAmplStep->text().toInt();

           ptr_N5172B->VIOpSignalGenerator::SetFre("5.83");                //设置频率
          for(int i=up;i>=down;i-=stp)
          {
              ptr_N5172B->VIOpSignalGenerator::SetAmpl(QString::number(i));   //设置功率
              Delay_MSec(2500);
              p=ptr_D7254->VIOpD7254::viop_DPOQueValueMean();
              double c=QString(p).toDouble();
              if(abs(c-standResult)>=1)
              {
                  cellA = worksheetwu->querySubObject("Cells(int,int)",row,2);
                  cellA->dynamicCall("SetValue(const QVariant&)",QString::number(i+1+stringminus));
                  i=down-1;
              }
          }
         //设置频率5.84
         ptr_N5172B->VIOpSignalGenerator::SetFre("5.84");                //设置频率
         ptr_N5172B->VIOpSignalGenerator::SetAmpl(ui->EditAmplWuUp->text());   //设置功率
         Delay_MSec(3000);
         result=ptr_D7254->VIOpD7254::viop_DPOQueValueMean();             //求出第一个值
         standResult=QString(result).toDouble();
         for(int i=up;i>=down;i-=stp)
         {
             ptr_N5172B->VIOpSignalGenerator::SetAmpl(QString::number(i));   //设置功率
             Delay_MSec(2500);
             p=ptr_D7254->VIOpD7254::viop_DPOQueValueMean();
             double c=QString(p).toDouble();
             if(abs(c-standResult)>=1)
             {
                 cellA = worksheetwu->querySubObject("Cells(int,int)",row,3);
                 cellA->dynamicCall("SetValue(const QVariant&)",QString::number(i+1+stringminus));
                 i=down-1;
             }
         }
         chipNum=ui->EditBoardWu->text().toInt();
         chipNum++;
         ui->EditBoardWu->setText(QString::number(chipNum));    //更新芯片号
         chipNum=1;
         row++;
          ui->WulineRow->setText(QString::number(row));
          free(result);
          free(p);
          result=nullptr;
          p=nullptr;
          //关闭文件
          workbookwu->dynamicCall("SaveAs(const QString&)",QDir::toNativeSeparators(filepathWu));
          workbookwu->dynamicCall("Close()");                               //关闭工作簿
          excelwu->dynamicCall("Quit()");                                   //关闭excel
          delete excelwu;
          excelwu=nullptr;
          QMessageBox::information(this,tr("完成"),tr("测试完成"));
      });









      //span开始频率
      connect(ui->btnTxStarFre,&QPushButton::clicked,[&](){
         ptr_R3408B->VIOpR3408B::viop_startfreq(ui->EditTxStarFre->text(), ui->cbxTxStarFreUnits->currentIndex());

       });
      //span结束频率
      connect(ui->btnTxStopFre,&QPushButton::clicked,[&](){
          ptr_R3408B->VIOpR3408B::viop_stopfreq(ui->EditTxStopFre->text(), ui->cbxTxStopFreUnits->currentIndex());
       });

      //中心频率
      connect(ui->btnTxCenterFre,&QPushButton::clicked,[&](){
           ptr_R3408B->VIOpR3408B::viop_CentralFre(ui->EditTxFre->text(),ui->cbxTxFreUnits->currentIndex());
      });

      //span设置
      connect(ui->btnTxSpan,&QPushButton::clicked,[&](){
       ptr_R3408B->VIOpR3408B::VIOpR3408B::viop_span(ui->EditTxSpan->text() ,ui->cbxTxSpanUnits->currentIndex());
      });

       //Wu定频点测试
      connect(ui->WuFreTest,&QPushButton::clicked,[&](){
        //测试步骤
        QString filepathWu;
        filepathWu =QFileDialog::getOpenFileName(nullptr,("选择文件"),("C:/Users/ztp/Desktop"),("*.xls *.xlsx *.csv")); //获取保存路径
        if(filepathWu.isEmpty())
        {
            QMessageBox::critical(nullptr, "错误信息", "没有找到EXCEL");
                return;
        }
        QAxObject *excelwu = new QAxObject();
        excelwu->setControl("ket.Application");                                     //连接WPS
        excelwu->dynamicCall("SetVisible(bool Visible)","false");                   //不显示窗体
        excelwu->setProperty("DisplayAlerts", false);                               //不显示任何警告信息。如果为true那么在关闭是会出现类似“文件已修改，是否保存”的提示
        QAxObject* workbookswu = excelwu->querySubObject("WorkBooks");                //获取工作簿集合
        workbookswu->dynamicCall("Open(const QString&)",filepathWu);                  //打开打开已存在的工作簿
        QAxObject *workbookwu = excelwu->querySubObject("ActiveWorkBook");            //获取当前工作簿
        QAxObject *worksheetswu = workbookwu->querySubObject("WorkSheets");           //获取工作表集合
        QAxObject *worksheetwu = worksheetswu->querySubObject("Item(int)",1);         //获取工作表集合的工作表1，即sheet1
        QAxObject *cellA;

        QString wufre=ui->EditWuFre->text();
        int stringminus=ui->Wuline->text().toInt();

        if(TestFirstWuFre)
        {
             cellA = worksheetwu->querySubObject("Cells(int,int)",1,1);
             cellA->dynamicCall("SetValue(const QVariant&)","No.");         //写入数据
             cellA->setProperty("HorizontalAlignment", -4108);                     //居中
             cellA->setProperty("VerticalAlignment", -4108);                       //居中
             cellA = worksheetwu->querySubObject("Cells(int,int)",1,2);
             cellA->dynamicCall("SetValue(const QVariant&)",wufre);         //写入数据
             cellA->setProperty("HorizontalAlignment", -4108);                     //居中
             cellA->setProperty("VerticalAlignment", -4108);                       //居中
             TestFirstWuFre=false;
        }

        cellA = worksheetwu->querySubObject("Cells(int,int)",row,1);
        cellA->dynamicCall("SetValue(const QVariant&)",ui->EditBoardWu->text());   //芯片号
        cellA->setProperty("HorizontalAlignment", -4108);                     //居中
        cellA->setProperty("VerticalAlignment", -4108);                       //居中


         ptr_N5172B->VIOpSignalGenerator::SetAmpl(ui->EditAmplWuUp->text());
         Delay_MSec(2000);
         result=ptr_D7254->VIOpD7254::viop_DPOQueValueMean();               //求出第一个值

         double standResult=QString(result).toDouble();

         int up=ui->EditAmplWuUp->text().toInt();
         int down=ui->EditAmplWuDown->text().toInt();
         int stp=ui->EditWuAmplStep->text().toInt();
         ptr_N5172B->VIOpSignalGenerator::SetFre(wufre);                     //设置频率

        for(int i=up;i>=down;i-=stp)
        {
            ptr_N5172B->VIOpSignalGenerator::SetAmpl(QString::number(i));   //设置功率
            Delay_MSec(2500);
            p=ptr_D7254->VIOpD7254::viop_DPOQueValueMean();
            double c=QString(p).toDouble();
            if(abs(c-standResult)>=1)
            {
                cellA = worksheetwu->querySubObject("Cells(int,int)",row,2);
                cellA->dynamicCall("SetValue(const QVariant&)",QString::number(i+1+stringminus));
                i=down-1;
            }
        }
       chipNum=ui->EditBoardWu->text().toInt();
       chipNum++;
       ui->EditBoardWu->setText(QString::number(chipNum));
//      chipNum=1;
        row++;
        ui->WulineRow->setText(QString::number(row));
        free(result);
        free(p);
        result=nullptr;
        p=nullptr;
        //关闭文件
        workbookwu->dynamicCall("SaveAs(const QString&)",QDir::toNativeSeparators(filepathWu));
        workbookwu->dynamicCall("Close()");                               //关闭工作簿
        excelwu->dynamicCall("Quit()");                                   //关闭excel
        delete excelwu;
        excelwu=nullptr;
        QMessageBox::information(this,tr("完成"),tr("测试完成"));
    });

      //Wu大功率测试
     connect(ui->WuMaxAmplTest,&QPushButton::clicked,[&]()
     {
         QString filepathWu;
         filepathWu =QFileDialog::getOpenFileName(nullptr,("选择文件"),("C:/Users/ztp/Desktop"),("*.xls *.xlsx *.csv")); //获取保存路径
         if(filepathWu.isEmpty())
         {
             QMessageBox::critical(nullptr, "错误信息", "没有找到EXCEL");
                 return;
         }
         QAxObject *excelwu = new QAxObject();
         excelwu->setControl("ket.Application");                                     //连接WPS
         excelwu->dynamicCall("SetVisible(bool Visible)","false");                   //不显示窗体
         excelwu->setProperty("DisplayAlerts", false);                               //不显示任何警告信息。如果为true那么在关闭是会出现类似“文件已修改，是否保存”的提示
         QAxObject* workbookswu = excelwu->querySubObject("WorkBooks");                //获取工作簿集合
         workbookswu->dynamicCall("Open(const QString&)",filepathWu);                  //打开打开已存在的工作簿
         QAxObject *workbookwu = excelwu->querySubObject("ActiveWorkBook");            //获取当前工作簿
         QAxObject *worksheetswu = workbookwu->querySubObject("WorkSheets");           //获取工作表集合
         QAxObject *worksheetwu = worksheetswu->querySubObject("Item(int)",1);         //获取工作表集合的工作表1，即sheet1
         QAxObject *cellA;

         if(TestFirstWuMax)
         {
              cellA = worksheetwu->querySubObject("Cells(int,int)",1,1);
              cellA->dynamicCall("SetValue(const QVariant&)","Fre");         //写入数据
              cellA->setProperty("HorizontalAlignment", -4108);                     //居中
              cellA->setProperty("VerticalAlignment", -4108);                       //居中
              cellA = worksheetwu->querySubObject("Cells(int,int)",1,2);
              cellA->dynamicCall("SetValue(const QVariant&)","Amp");         //写入数据
              cellA->setProperty("HorizontalAlignment", -4108);                     //居中
              cellA->setProperty("VerticalAlignment", -4108);                       //居中
              TestFirstWuMax=false;
         }

         cellA = worksheetwu->querySubObject("Cells(int,int)",row++,1);
         cellA->dynamicCall("SetValue(const QVariant&)",ui->EditBoardWu->text()+"号芯片");   //芯片号
         cellA->setProperty("HorizontalAlignment", -4108);                     //居中
         cellA->setProperty("VerticalAlignment", -4108);                       //居中

         double freup=ui->WUFreUp->text().toDouble();
         double fredown=ui->WUFreDOWN->text().toDouble();
         double frestep=ui->WUFreStep->text().toDouble()/1000;

         int ampup1=ui->WUAmpUp->text().toInt();
         int ampdown=ui->WuAmpDown->text().toInt();
         int ampstep=ui->WUAmpStep->text().toInt();

         for(;fredown<=freup;fredown+=frestep)
         {
             ampup1=ui->WUAmpUp->text().toInt();
             ptr_N5172B->VIOpSignalGenerator::SetFre(QString::number(fredown)); //设置频率
             ptr_N5172B->VIOpSignalGenerator::SetAmpl(QString::number(ampup1+1)); //设置功率
             result=ptr_D7254->VIOpD7254::viop_DPOQueValueMean();               //求出第一个值
             double standResult=QString(result).toDouble();
             Delay_MSec(2000);
             for(;ampup1>=ampdown;ampup1=ampup1-ampstep)
             {
               ptr_N5172B->VIOpSignalGenerator::SetAmpl(QString::number(ampup1)); //设置功率
               Delay_MSec(2500);
               p=ptr_D7254->VIOpD7254::viop_DPOQueValueMean();
               double c=QString(p).toDouble();
               if(abs(c-standResult)>=1)
               {
                   cellA = worksheetwu->querySubObject("Cells(int,int)",row,1);
                   cellA->dynamicCall("SetValue(const QVariant&)",QString::number(fredown));
                   cellA = worksheetwu->querySubObject("Cells(int,int)",row++,2);
                   cellA->dynamicCall("SetValue(const QVariant&)",QString::number(ampup1+ampstep));
                   continue;
               }
             }
         }

         chipNum=ui->EditBoardWu->text().toInt();
         chipNum++;
         ui->EditBoardWu->setText(QString::number(chipNum));
         row++;
         free(result);
         free(p);
         result=nullptr;
         p=nullptr;
         //关闭文件
         workbookwu->dynamicCall("SaveAs(const QString&)",QDir::toNativeSeparators(filepathWu));
         workbookwu->dynamicCall("Close()");                               //关闭工作簿
         excelwu->dynamicCall("Quit()");                                   //关闭excel
         delete excelwu;
         excelwu=nullptr;
         QMessageBox::information(this,tr("完成"),tr("测试完成"));
     });





    //打开远程连接
    connect(ui->btnRMTCtrl_Open,&QPushButton::clicked,[&]()
    {
        isN5172BConnected=MainWindow::ptr_N5172B->viop_viOpen(addr_N5172B.toUtf8().data());
        is33250AConnected=MainWindow::ptr_33250A->viop_viOpen(addr_33250A.toUtf8().data());
        isR3408BConnected=MainWindow::ptr_R3408B->viop_viOpen(addr_R3408B.toUtf8().data());
        if(isN5172BConnected ||is33250AConnected||isR3408BConnected)
        {
            //设置按钮颜色
            QPalette pal = ui->btnRMTCtrl_Open->palette();
            pal.setColor(QPalette::Button,Qt::green);
            ui->btnRMTCtrl_Open->setPalette(pal);
            ui->btnRMTCtrl_Open->setAutoFillBackground(true);
            ui->btnRMTCtrl_Open->setFlat(true);
            //labelLinkStatus->setText("连接状态：N5712B-已连接 33250A-已连接 R3408B-已连接");
            labelLinkStatus->setText("当前仪器连接成功");
        }
        else
        {
            QMessageBox::warning(this,"错误","建立连接失败，请确认连接正确",QMessageBox::Ok,QMessageBox::Cancel);
        }
    });


    //关闭远程连接
    connect(ui->btnRMTCtrl_Close,&QPushButton::clicked,[&](){
        viClose(*ptr_N5172B->vi);                    // Closes session
        viClose(*ptr_N5172B->defaultRM);             // Closes default session
        viClose(*ptr_33250A->vi);
        viClose(*ptr_33250A->defaultRM);
        viClose(*ptr_R3408B->vi);
        viClose(*ptr_R3408B->defaultRM);
    });

    //RF  ON/OFF
    connect(ui->btnRFOnOff,&QPushButton::clicked,[&](){
        if(isRFOn)
        {
            ptr_N5172B->VIOpSignalGenerator::viop_RFOn();
            isRFOn = true;
        }
        else
        {
            ptr_N5172B->VIOpSignalGenerator::viop_RFOff();
            isRFOn = false;
        }
    });

    //AM   ON/OF
    connect(ui->btnAMOnOff,&QPushButton::clicked,[&](){
      if(isAMOn)
      {
         ptr_N5172B->viop_AMONOFF(isAMOn);
         isAMOn=false;
      }
      else
      {
          ptr_N5172B->viop_AMONOFF(isAMOn);
          isAMOn=true;
      }
    });

    //打开串口并配置文件
    //检测串口
      connect(ui->btnDect,&QPushButton::clicked,[&]()
      {
          ui->cbxComPort->clear();
          foreach (const QSerialPortInfo &info,QSerialPortInfo::availablePorts())
             {
                 QSerialPort serial;
                 serial.setPort(info);
                 if(serial.open(QIODevice::ReadWrite))
                 {
                      ui->cbxComPort->addItem(serial.portName());
                      serial.close();
                 }
             }
       });

      //打开串口
      connect(ui->btnOpen,&QPushButton::clicked,[&]()
     {
         QString open ="打开串口";
         QString close="关闭串口";

         /* 先来判断对象是不是为空 */
            if(serial == nullptr)
            {
                /* 新建串口对象 */
               serial  = new QSerialPort();
            }
            /* 判断是要打开串口，还是关闭串口 */
            if(serial ->isOpen())
            {
                /* 串口已经打开，现在来关闭串口 */
               ui->btnOpen->setText(open);
               serial ->close();
               QMessageBox::information(this,"提示","串口已关闭");
            }
            else
            {
                /* 判断是否有可用串口 */
                if(ui->cbxComPort->count() != 0)
                {
                    /* 串口已经关闭，现在来打开串口 */
                    /* 设置串口名称 */
                   serial->setPortName(ui->cbxComPort->currentText());
                    /* 设置波特率 */
                   serial ->setBaudRate(115200);
                    /* 设置数据位数 */
                   serial ->setDataBits(QSerialPort::Data8);
                   /* 设置停止位 */
                   serial ->setStopBits(QSerialPort::OneStop);
                    /* 设置奇偶校验 */
                   serial ->setParity(QSerialPort::NoParity);
                    /* 设置流控制 */
                   serial ->setFlowControl(QSerialPort::NoFlowControl);
                    /* 打开串口 */
                   serial ->open(QIODevice::ReadWrite);
                    /* 设置串口缓冲区大小，这里必须设置为这么大 */
                   serial ->setReadBufferSize(PortOneFrameSize);
                   ui->btnOpen->setText(close);
                   QMessageBox::information(this,"提示","串口已打开");
                }
                else
                {
                       // 警告对话框
                   QMessageBox::warning(this,tr("警告"),tr("没有可用串口，请重新尝试扫描串口！"),QMessageBox::Ok);}
                }
      });

      //  enable chip
      connect(ui->btnEnChip,&QPushButton::clicked,[&](){
          if(ui->cbxComPort->count() != 0)
          {
           int tempchip=0xff;
           QString ss;
           ss=QString::number(tempchip);
           QByteArray tempChip=ss.toLatin1();
           serial->write(tempChip);
          }
          else
          {
          QMessageBox::warning(this,tr("警告"),tr("没有打开串口，请打开串口！"),QMessageBox::Ok);
          }
      });

      //   enable spi
      connect(ui->btnEnSPI,&QPushButton::clicked,[&]()
      {
          if(ui->cbxComPort->count() != 0){
          if(EnSpiDsrc_HI==false)
          {
          int tempchip=0xbf;
          QString ss;
          ss=QString::number(tempchip);
          QByteArray tempChip=ss.toLatin1();
          serial->write(tempChip);
          EnSpiDsrc_HI=true;
          }
          else
           {
              int tempchip=0xaf;
              QString ss;
              ss=QString::number(tempchip);
              QByteArray tempChip=ss.toLatin1();
              serial->write(tempChip);
              EnSpiDsrc_HI=false;
           }
          }
          else
          {
           QMessageBox::warning(this,tr("警告"),tr("没有打开串口，请打开串口！"),QMessageBox::Ok);
          }
      });

      //读取文件并发送
      connect(ui->btnSend,&QPushButton::clicked,[&]()
      {
          QString filepath;
          filepath =QFileDialog::getOpenFileName(this,("选择文件"),("C:/Users/ztp/Desktop"),("*.txt")); //获取保存路径
          QFile file(filepath);
          if(filepath.isEmpty())
          {
              QMessageBox::critical(this, "错误信息", "没有选择文件");
                  return;
          }

          //打开txt
          if(!file.open(QIODevice::ReadOnly | QIODevice::Text))                                                         //只读 文本文档
          {
             QMessageBox::information(this,"警告","无法打开文件");
          }
          while(!file.atEnd())                                                                                          //没有读到最后一行
          {
              QByteArray line = file.readLine();
              QString str(line);
              QStringList s=str.split(" ");
              QString  a1= s[0];
              QString  b=binTohex(s[1]);                     //二进制转十六进制
              QString  c=binTohex1(s[2]);
              QString a2;
              QString firOpera="00";
               for(int i=2;i<4;i++)
              {
                  a2.append(a1[i]);
              }
              b.append(c);
              a2.append(b);
              firOpera.append(a2);
              QByteArray qByteArray=QByteArray::fromHex(firOpera.toLatin1().data());
              serial->write(qByteArray);
              a1="";
              b="";
              c="";
              a2="";
              Delay_MSec(30);
          }
             file.close();
          QMessageBox::information(this,"信息","配置完成");
      });

      //Rx带宽测试
     connect(ui->RxFreRange,&QPushButton::clicked,[&]()
      {
         //设置信号源
         ToalBits=ui->EditTolBits->text();
         ptr_N5172B->VIOpSignalGenerator::viop_realtimeOn();                  //realtime
         ptr_N5172B->VIOpSignalGenerator::viop_modulationType();              //设置ASK
         ptr_N5172B->VIOpSignalGenerator::viop_alpha();                       //设置阿尔法为1
         ptr_N5172B->VIOpSignalGenerator::viop_alphaDepth();                  //85%
         ptr_N5172B->VIOpSignalGenerator::viop_symbolRate(512);               //symbol rate 默认单位Ksps
         ptr_N5172B->VIOpSignalGenerator::viop_loadData();                    //load文件PN9-1FF-FM0
         ptr_N5172B->VIOpSignalGenerator::viop_displayMode(SCIENTIFIC);       //科学计数法显示误码率
         ptr_N5172B->VIOpSignalGenerator::viop_BERTtriggerSource(IMM);        //trigger-immediate
         ptr_N5172B->VIOpSignalGenerator::viop_totalBit(ToalBits);            //设置totalBit为1千万
         ptr_N5172B->VIOpSignalGenerator::viop_MODOn();
         ptr_N5172B->VIOpSignalGenerator::viop_RFOn();                         //打开RF
         ptr_N5172B->VIOpSignalGenerator::viop_bertOn();                       //打开

         //  根据实际赋值
         double freUp=ui->RxRangeFreUp->text().toDouble();
         double freDown=ui->RxRangeFreDown->text().toDouble();
         double freStep=ui->RxRangStep->text().toDouble()/1000000;
//         qDebug()<<freUp;
//         qDebug()<<freDown;
//         qDebug()<<freStep;


         BerStand=ui->EditBer->text().toDouble();   //误码率

         //选择文件
         QString filepath;
         filepath =QFileDialog::getOpenFileName(nullptr,("选择文件"),("C:/Users/ztp/Desktop"),("*.xls *.xlsx *.csv")); //获取保存路径
         if(filepath.isEmpty())
         {
             QMessageBox::critical(nullptr, "错误信息", "没有找到EXCEL");
                 return;
         }

         QAxObject *excel = new QAxObject();
         excel->setControl("ket.Application");                                     //连接WPS
         excel->dynamicCall("SetVisible(bool Visible)","false");                   //不显示窗体
         excel->setProperty("DisplayAlerts", false);                               //不显示任何警告信息。如果为true那么在关闭是会出现类似“文件已修改，是否保存”的提示
         QAxObject* workbooks = excel->querySubObject("WorkBooks");                //获取工作簿集合
         workbooks->dynamicCall("Open(const QString&)",filepath);                  //打开打开已存在的工作簿
         QAxObject *workbook = excel->querySubObject("ActiveWorkBook");            //获取当前工作簿
         QAxObject *worksheets = workbook->querySubObject("WorkSheets");           //获取工作表集合
         QAxObject *worksheet = worksheets->querySubObject("Item(int)",1);         //获取工作表集合的工作表1，即sheet1
         QAxObject *cellA1;

         if(RxFreRange)
         {
              cellA1 = worksheet->querySubObject("Cells(int,int)",1,1);
              cellA1->dynamicCall("SetValue(const QVariant&)","Rx带宽测试");         //写入数据
              cellA1->setProperty("HorizontalAlignment", -4108);                     //居中
              cellA1->setProperty("VerticalAlignment", -4108);                       //居中
             RxFreRange=false;
         }

         //写入芯片号
         int tempRxAmpl=ui->RxRangePoint->text().toInt()+6;
         QString textLab=QString::number(tempRxAmpl)+"dBm "+ui->RxRangStep->text()+"KHz";
         textLab=ui->EditBoard->text()+"号 "+textLab;
         cellA1= worksheet->querySubObject("Cells(int,int)",row++,1);
         cellA1->dynamicCall("SetValue(const QVariant&)",textLab);                   //写入板号
         cellA1->setProperty("HorizontalAlignment", -4108);                          //居中
         cellA1->setProperty("VerticalAlignment", -4108);                            //居中
         int colum=2;

         // 第二行数据进行操作
         cellA1 = worksheet->querySubObject("Cells(int,int)",row++,1);
         cellA1->dynamicCall("SetValue(const QVariant&)","频率(GHz)");         //写入功率
         cellA1->setProperty("HorizontalAlignment", -4108);                   //居中
         cellA1->setProperty("VerticalAlignment", -4108);                     //居中
         cellA1 = worksheet->querySubObject("Cells(int,int)",row++,1);
         cellA1->dynamicCall("SetValue(const QVariant&)","误码率");             //写入误码率
         cellA1->setProperty("HorizontalAlignment", -4108);                    //居中
         cellA1->setProperty("VerticalAlignment", -4108);                      //居中
         cellA1 = worksheet->querySubObject("Cells(int,int)",row,1);
         cellA1->dynamicCall("SetValue(const QVariant&)","结果");               //写入结果
         cellA1->setProperty("HorizontalAlignment", -4108);                    //居中
         cellA1->setProperty("VerticalAlignment", -4108);                      //居中

         //
         ptr_N5172B->VIOpSignalGenerator::SetAmpl(QString::number(tempRxAmpl));     //输入ampl
             //功率循环
          for(;freDown<=freUp;freDown+=freStep)
          {
            ptr_N5172B->VIOpSignalGenerator::SetFre(QString::number(freDown));//设置频率
            Delay_MSec(100);
            ptr_N5172B->VIOpSignalGenerator::viop_bertOff();                         //关闭bert
            ptr_N5172B->VIOpSignalGenerator::viop_bertOn();                          //打开bert
            Delay_MSec(2000);                                                        //延时约4秒等待BER稳定
            //返回BERT结果
            result = ptr_N5172B ->viop_Q_BertResult1();                               //调用viop_Q_BertResult()函数，得到bert测试各项结果的值
            str_BER = QString(result);	               //得到bert测试下 ber的值 根据测试要求这个值小于10的负五次方的话为合格
            double strBER=str_BER.toDouble();

                 //写入excel
            cellA1 = worksheet->querySubObject("Cells(int,int)",row-2,colum);
            cellA1->dynamicCall("SetValue(const QVariant&)",QString::number(freDown));   //写入频率
            cellA1->setProperty("HorizontalAlignment", -4108);           //居中
            cellA1->setProperty("VerticalAlignment", -4108);             //居中

            cellA1 = worksheet->querySubObject("Cells(int,int)",row-1,colum);
            cellA1->dynamicCall("SetValue(const QVariant&)",str_BER);    //写入误码率
            cellA1->setProperty("HorizontalAlignment", -4108);           //居中
            cellA1->setProperty("VerticalAlignment", -4108);             //居中

            if(strBER<=BerStand)
            {
                    cellA1 = worksheet->querySubObject("Cells(int,int)",row,colum);
                    cellA1->dynamicCall("SetValue(const QVariant&)","T");   //写入结果
                    cellA1->setProperty("HorizontalAlignment", -4108);         //居中
                    cellA1->setProperty("VerticalAlignment", -4108);           //居中
             }
            if(strBER>BerStand)
            {
             cellA1 = worksheet->querySubObject("Cells(int,int)",row,colum);
             cellA1->dynamicCall("SetValue(const QVariant&)","F");    //写入结果
//           QAxObject *font = cellA1->querySubObject("Font");            //获取单元格字体
//           font->setProperty("Bold", "true");                        //设置单元格字体颜色（红色）
//           font->setProperty("Color", QColor(255, 0, 0));               //设置单元格字体颜色（红色）
             cellA1->setProperty("HorizontalAlignment", -4108);           //居中
            cellA1->setProperty("VerticalAlignment", -4108);             //居中
             }
                 colum++;
           }
         colum=2;
         row++;
         ui->lineEdit_Row->setText(QString::number(row));
         //芯片号增加
         chipNum=ui->EditBoard->text().toInt();
         chipNum++;
         ui->EditBoard->setText(QString::number(chipNum));   //显示下一个芯片号
         free(result);
         //关闭文件
         workbook->dynamicCall("SaveAs(const QString&)",QDir::toNativeSeparators(filepath));
         workbook->dynamicCall("Close()");                               //关闭工作簿
         excel->dynamicCall("Quit()");                                   //关闭excel
         delete excel;
         excel=nullptr;
        QMessageBox::information(this,"RX带宽测试","完成");
      });


     //Rx带宽测试-多灵敏度
    connect(ui->RxFreRangeMul,&QPushButton::clicked,[&]()
     {
        //设置信号源
        ToalBits=ui->EditTolBits->text();
        ptr_N5172B->VIOpSignalGenerator::viop_realtimeOn();                  //realtime
        ptr_N5172B->VIOpSignalGenerator::viop_modulationType();              //设置ASK
        ptr_N5172B->VIOpSignalGenerator::viop_alpha();                       //设置阿尔法为1
        ptr_N5172B->VIOpSignalGenerator::viop_alphaDepth();                  //85%
        ptr_N5172B->VIOpSignalGenerator::viop_symbolRate(512);               //symbol rate 默认单位Ksps
        ptr_N5172B->VIOpSignalGenerator::viop_loadData();                    //load文件PN9-1FF-FM0
        ptr_N5172B->VIOpSignalGenerator::viop_displayMode(SCIENTIFIC);       //科学计数法显示误码率
        ptr_N5172B->VIOpSignalGenerator::viop_BERTtriggerSource(IMM);        //trigger-immediate
        ptr_N5172B->VIOpSignalGenerator::viop_totalBit(ToalBits);            //设置totalBit为1千万
        ptr_N5172B->VIOpSignalGenerator::viop_MODOn();
        ptr_N5172B->VIOpSignalGenerator::viop_RFOn();                         //打开RF
        ptr_N5172B->VIOpSignalGenerator::viop_bertOn();                       //打开

        //  根据实际赋值
        double freUp=ui->RxRangeFreUp->text().toDouble();         //频率上限
        double freDown=ui->RxRangeFreDown->text().toDouble();     //频率下限
        double freStep=ui->RxRangStep->text().toDouble()/1000000; //频率步长
        int amplUp=ui->RxRangePointUp->text().toInt();            //功率上限
        int amplDown=ui->RxRangePoint->text().toInt()+6;            //功率下限,自动+6
        int amplStep=ui->RxRangeAmplStep->text().toInt();         //功率步长

//         qDebug()<<freUp;
//         qDebug()<<freDown;
//         qDebug()<<freStep;


        BerStand=ui->EditBer->text().toDouble();   //误码率

        //选择文件
        QString filepath;
        filepath =QFileDialog::getOpenFileName(nullptr,("选择文件"),("C:/Users/ztp/Desktop"),("*.xls *.xlsx *.csv")); //获取保存路径
        if(filepath.isEmpty())
        {
            QMessageBox::critical(nullptr, "错误信息", "没有找到EXCEL");
                return;
        }

        QAxObject *excel = new QAxObject();
        excel->setControl("ket.Application");                                     //连接WPS
        excel->dynamicCall("SetVisible(bool Visible)","false");                   //不显示窗体
        excel->setProperty("DisplayAlerts", false);                               //不显示任何警告信息。如果为true那么在关闭是会出现类似“文件已修改，是否保存”的提示
        QAxObject* workbooks = excel->querySubObject("WorkBooks");                //获取工作簿集合
        workbooks->dynamicCall("Open(const QString&)",filepath);                  //打开打开已存在的工作簿
        QAxObject *workbook = excel->querySubObject("ActiveWorkBook");            //获取当前工作簿
        QAxObject *worksheets = workbook->querySubObject("WorkSheets");           //获取工作表集合
        QAxObject *worksheet = worksheets->querySubObject("Item(int)",1);         //获取工作表集合的工作表1，即sheet1
        QAxObject *cellA1;

        if(RxFreRangeMul)
        {
             cellA1 = worksheet->querySubObject("Cells(int,int)",1,1);
             cellA1->dynamicCall("SetValue(const QVariant&)","Rx带宽测试-灵敏度上下限");         //写入数据
             cellA1->setProperty("HorizontalAlignment", -4108);                     //居中
             cellA1->setProperty("VerticalAlignment", -4108);                       //居中
            RxFreRangeMul=false;
        }


      for(;amplDown<=amplUp;amplDown+=amplStep)
      {
          //写入芯片号
          QString textLab=QString::number(amplDown)+"dBm";
          textLab=ui->EditBoard->text()+"号 "+textLab;
          cellA1= worksheet->querySubObject("Cells(int,int)",row++,1);
          cellA1->dynamicCall("SetValue(const QVariant&)",textLab);                   //写入板号
          cellA1->setProperty("HorizontalAlignment", -4108);                          //居中
          cellA1->setProperty("VerticalAlignment", -4108);                            //居中
          int colum=2;

          // 第二行数据进行操作
          cellA1 = worksheet->querySubObject("Cells(int,int)",row++,1);
          cellA1->dynamicCall("SetValue(const QVariant&)","频率(GHz)");         //写入功率
          cellA1->setProperty("HorizontalAlignment", -4108);                   //居中
          cellA1->setProperty("VerticalAlignment", -4108);                     //居中
          cellA1 = worksheet->querySubObject("Cells(int,int)",row++,1);
          cellA1->dynamicCall("SetValue(const QVariant&)","误码率");             //写入误码率
          cellA1->setProperty("HorizontalAlignment", -4108);                    //居中
          cellA1->setProperty("VerticalAlignment", -4108);                      //居中
          cellA1 = worksheet->querySubObject("Cells(int,int)",row,1);
          cellA1->dynamicCall("SetValue(const QVariant&)","结果");               //写入结果
          cellA1->setProperty("HorizontalAlignment", -4108);                    //居中
          cellA1->setProperty("VerticalAlignment", -4108);                      //居中
          //
          ptr_N5172B->VIOpSignalGenerator::SetAmpl(QString::number(amplDown));     //输入ampl
              //功率循环
           for(;freDown<=freUp;freDown+=freStep)
           {
             ptr_N5172B->VIOpSignalGenerator::SetFre(QString::number(freDown));//设置频率
             Delay_MSec(100);
             ptr_N5172B->VIOpSignalGenerator::viop_bertOff();                         //关闭bert
             ptr_N5172B->VIOpSignalGenerator::viop_bertOn();                          //打开bert
             Delay_MSec(2000);                                                        //延时约4秒等待BER稳定
             //返回BERT结果
             result = ptr_N5172B ->viop_Q_BertResult1();                               //调用viop_Q_BertResult()函数，得到bert测试各项结果的值
             str_BER = QString(result);	               //得到bert测试下 ber的值 根据测试要求这个值小于10的负五次方的话为合格
             double strBER=str_BER.toDouble();

                  //写入excel
             cellA1 = worksheet->querySubObject("Cells(int,int)",row-2,colum);
             cellA1->dynamicCall("SetValue(const QVariant&)",QString::number(freDown));   //写入频率
             cellA1->setProperty("HorizontalAlignment", -4108);           //居中
             cellA1->setProperty("VerticalAlignment", -4108);             //居中

             cellA1 = worksheet->querySubObject("Cells(int,int)",row-1,colum);
             cellA1->dynamicCall("SetValue(const QVariant&)",str_BER);    //写入误码率
             cellA1->setProperty("HorizontalAlignment", -4108);           //居中
             cellA1->setProperty("VerticalAlignment", -4108);             //居中

             if(strBER<=BerStand)
             {
                     cellA1 = worksheet->querySubObject("Cells(int,int)",row,colum);
                     cellA1->dynamicCall("SetValue(const QVariant&)","T");   //写入结果
                     cellA1->setProperty("HorizontalAlignment", -4108);         //居中
                     cellA1->setProperty("VerticalAlignment", -4108);           //居中
              }
             if(strBER>BerStand)
             {
              cellA1 = worksheet->querySubObject("Cells(int,int)",row,colum);
              cellA1->dynamicCall("SetValue(const QVariant&)","F");    //写入结果
              cellA1->setProperty("HorizontalAlignment", -4108);           //居中
             cellA1->setProperty("VerticalAlignment", -4108);             //居中
              }
                  colum++;
            }
          freDown=ui->RxRangeFreDown->text().toDouble();
          colum=2;
          row++;
      }

        ui->lineEdit_Row->setText(QString::number(row));
        //芯片号增加
        chipNum=ui->EditBoard->text().toInt();
        chipNum++;
        ui->EditBoard->setText(QString::number(chipNum));   //显示下一个芯片号
        free(result);
        //关闭文件
        workbook->dynamicCall("SaveAs(const QString&)",QDir::toNativeSeparators(filepath));
        workbook->dynamicCall("Close()");                               //关闭工作簿
        excel->dynamicCall("Quit()");                                   //关闭excel
        delete excel;
        excel=nullptr;
       QMessageBox::information(this,"RX带宽测试","完成");
     });


      //Rx接收灵敏度默认选5.83GHz
       ui->btnRx5830->setChecked(true);
      //Rx灵敏度测试
          connect(ui->btnRX,&QPushButton::clicked,[&]()
          {
               QString textLab;
               ToalBits=ui->EditTolBits->text();
               ptr_N5172B->VIOpSignalGenerator::viop_realtimeOn();                  //realtime
               ptr_N5172B->VIOpSignalGenerator::viop_modulationType();              //设置ASK
               ptr_N5172B->VIOpSignalGenerator::viop_alpha();                       //设置阿尔法为1
               ptr_N5172B->VIOpSignalGenerator::viop_alphaDepth();                  //85%
               ptr_N5172B->VIOpSignalGenerator::viop_symbolRate(512);               //symbol rate 默认单位Ksps
               ptr_N5172B->VIOpSignalGenerator::viop_loadData();                    //load文件PN9-1FF-FM0
               ptr_N5172B->VIOpSignalGenerator::viop_displayMode(SCIENTIFIC);       //科学计数法显示误码率
               ptr_N5172B->VIOpSignalGenerator::viop_BERTtriggerSource(IMM);        //trigger-immediate
               ptr_N5172B->VIOpSignalGenerator::viop_totalBit(ToalBits);            //设置totalBit为1千万
               ptr_N5172B->VIOpSignalGenerator::viop_MODOn();
               ptr_N5172B->VIOpSignalGenerator::viop_RFOn();                         //打开RF
               ptr_N5172B->VIOpSignalGenerator::viop_bertOn();                       //打开

               if(ui->btnRx5830->isChecked())
               {
                    ptr_N5172B->VIOpSignalGenerator::SetFre("5.83");
                    textLab="号5.83GHz";
               }
               if(ui->btnRx5840->isChecked())
               {
                  ptr_N5172B->VIOpSignalGenerator::SetFre("5.84");
                  textLab="号5.84GHz";
               }

               //  根据实际赋值
               AmplUpper=ui->EditamplUP->text().toInt();                               //ampl上限
               AmplLower=ui->EditamplLW->text().toInt();                               //ampl下限
               AmplStep=ui->Editamplstep->text().toInt();                              //ampl步长
               BerStand=ui->EditBer->text().toDouble();                                   //误码率

               //检查输入情况
               if(AmplUpper<AmplLower||FreUp<FreDown)
                  {
                    QMessageBox::information(this,"错误","数值输入错误");
                      return;
                  }
               //选择文件
               QString filepath;
               filepath =QFileDialog::getOpenFileName(nullptr,("选择文件"),("C:/Users/ztp/Desktop"),("*.xls *.xlsx *.csv")); //获取保存路径
               if(filepath.isEmpty())
               {
                   QMessageBox::critical(nullptr, "错误信息", "没有找到EXCEL");
                       return;
               }

               QAxObject *excel = new QAxObject();
               excel->setControl("ket.Application");                                     //连接WPS
               excel->dynamicCall("SetVisible(bool Visible)","false");                   //不显示窗体
               excel->setProperty("DisplayAlerts", false);                               //不显示任何警告信息。如果为true那么在关闭是会出现类似“文件已修改，是否保存”的提示
               QAxObject* workbooks = excel->querySubObject("WorkBooks");                //获取工作簿集合
               workbooks->dynamicCall("Open(const QString&)",filepath);                  //打开打开已存在的工作簿
               QAxObject *workbook = excel->querySubObject("ActiveWorkBook");            //获取当前工作簿
               QAxObject *worksheets = workbook->querySubObject("WorkSheets");           //获取工作表集合
               QAxObject *worksheet = worksheets->querySubObject("Item(int)",1);         //获取工作表集合的工作表1，即sheet1
               QAxObject *cellA1;

               if(TestFirst)
               {
                    cellA1 = worksheet->querySubObject("Cells(int,int)",1,1);
                    cellA1->dynamicCall("SetValue(const QVariant&)","Rx灵敏度测试");         //写入数据
                    cellA1->setProperty("HorizontalAlignment", -4108);                     //居中
                    cellA1->setProperty("VerticalAlignment", -4108);                       //居中
                    TestFirst=false;
               }

               textLab=ui->EditBoard->text()+ textLab;
               cellA1= worksheet->querySubObject("Cells(int,int)",row++,1);
               cellA1->dynamicCall("SetValue(const QVariant&)",textLab);                   //写入板号
               cellA1->setProperty("HorizontalAlignment", -4108);                          //居中
               cellA1->setProperty("VerticalAlignment", -4108);                            //居中
               int colum=2;

               // 第二行数据进行操作
               cellA1 = worksheet->querySubObject("Cells(int,int)",row++,1);
               cellA1->dynamicCall("SetValue(const QVariant&)","功率(dBm)");         //写入功率
               cellA1->setProperty("HorizontalAlignment", -4108);                   //居中
               cellA1->setProperty("VerticalAlignment", -4108);                     //居中

               cellA1 = worksheet->querySubObject("Cells(int,int)",row++,1);
               cellA1->dynamicCall("SetValue(const QVariant&)","误码率");             //写入误码率
               cellA1->setProperty("HorizontalAlignment", -4108);                    //居中
               cellA1->setProperty("VerticalAlignment", -4108);                      //居中
               cellA1 = worksheet->querySubObject("Cells(int,int)",row,1);
               cellA1->dynamicCall("SetValue(const QVariant&)","结果");               //写入结果
               cellA1->setProperty("HorizontalAlignment", -4108);                    //居中
               cellA1->setProperty("VerticalAlignment", -4108);                      //居中

                   //功率循环
                   for(;AmplUpper>=AmplLower;AmplUpper=AmplUpper-AmplStep)
                    {
                       QString TemAmplS=QString::number(AmplUpper);
                       ptr_N5172B->VIOpSignalGenerator::SetAmpl(TemAmplS);                  //输入ampl
                       Delay_MSec(100);
                       ptr_N5172B->VIOpSignalGenerator::viop_bertOff();                         //关闭bert
                       ptr_N5172B->VIOpSignalGenerator::viop_bertOn();                          //打开bert
                       Delay_MSec(2000);                                                        //延时约4秒等待BER稳定
                       //返回BERT结果
                       result = ptr_N5172B ->viop_Q_BertResult1();                               //调用viop_Q_BertResult()函数，得到bert测试各项结果的值
//                       p = strtok(result,d);
//                       str_totalBits=QString(p);           //得到bert测试下 totalBits的值
//                       p = strtok(nullptr, d);
//                       str_errorBits = QString(p);        //得到bert测试下 errorBits
//                       p = strtok(nullptr, d);
                       str_BER = QString(result);	               //得到bert测试下 ber的值 根据测试要求这个值小于10的负五次方的话为合格
                       double strBER=str_BER.toDouble();

                       //写入excel
                       cellA1 = worksheet->querySubObject("Cells(int,int)",row-2,colum);
                       cellA1->dynamicCall("SetValue(const QVariant&)",TemAmplS);   //写入功率
                       cellA1->setProperty("HorizontalAlignment", -4108);           //居中
                       cellA1->setProperty("VerticalAlignment", -4108);             //居中

                       cellA1 = worksheet->querySubObject("Cells(int,int)",row-1,colum);
                       cellA1->dynamicCall("SetValue(const QVariant&)",str_BER);    //写入误码率
                       cellA1->setProperty("HorizontalAlignment", -4108);           //居中
                       cellA1->setProperty("VerticalAlignment", -4108);             //居中

                       if(strBER<=BerStand)
                       {
                          cellA1 = worksheet->querySubObject("Cells(int,int)",row,colum);
                          cellA1->dynamicCall("SetValue(const QVariant&)","T");   //写入结果
                          cellA1->setProperty("HorizontalAlignment", -4108);         //居中
                          cellA1->setProperty("VerticalAlignment", -4108);           //居中
                       }
                       if(strBER>BerStand)
                       {
                          cellA1 = worksheet->querySubObject("Cells(int,int)",row,colum);
                          cellA1->dynamicCall("SetValue(const QVariant&)","F");    //写入结果
//                          QAxObject *font = cellA1->querySubObject("Font");            //获取单元格字体
//                          font->setProperty("Bold", "true");                        //设置单元格字体颜色（红色）
//                          font->setProperty("Color", QColor(255, 0, 0));               //设置单元格字体颜色（红色）
                          cellA1->setProperty("HorizontalAlignment", -4108);           //居中
                          cellA1->setProperty("VerticalAlignment", -4108);             //居中
                       }
                       colum++;
                   }
               colum=2;
               row++;
               ui->lineEdit_Row->setText(QString::number(row));
               //芯片号增加
               chipNum=ui->EditBoard->text().toInt();
               chipNum++;
               ui->EditBoard->setText(QString::number(chipNum));//显示下一个芯片号
               free(result);
               //关闭文件
               workbook->dynamicCall("SaveAs(const QString&)",QDir::toNativeSeparators(filepath));
               workbook->dynamicCall("Close()");                               //关闭工作簿
               excel->dynamicCall("Quit()");                                   //关闭excel
               delete excel;
               excel=nullptr;
              QMessageBox::information(this,"RX接收灵敏度测试","完成");
           });


          //重置N5172B
          connect(ui->btnRxRest,&QPushButton::clicked,[&](){
               ptr_N5172B->viop_RST();
               QMessageBox::information(this,"提示","仪器已经初始化");
               return;
          });

          //配置N5172B为Rx测试状态
          connect(ui->btnRxSet,&QPushButton::clicked,[&](){
              ToalBits=ui->EditTolBits->text();
              ptr_N5172B->VIOpSignalGenerator::viop_realtimeOn();                  //realtime
              ptr_N5172B->VIOpSignalGenerator::viop_modulationType();              //设置ASK
              ptr_N5172B->VIOpSignalGenerator::viop_alpha();                       //设置阿尔法为1
              ptr_N5172B->VIOpSignalGenerator::viop_alphaDepth();                  //85%
              ptr_N5172B->VIOpSignalGenerator::viop_symbolRate(512);               //symbol rate 默认单位Ksps
              ptr_N5172B->VIOpSignalGenerator::viop_loadData();                    //load文件PN9-1FF-FM0
              ptr_N5172B->VIOpSignalGenerator::viop_displayMode(SCIENTIFIC);       //科学计数法显示误码率
              ptr_N5172B->VIOpSignalGenerator::viop_BERTtriggerSource(IMM);        //trigger-immediate
              ptr_N5172B->VIOpSignalGenerator::viop_totalBit(ToalBits);            //设置totalBit为1千万
              ptr_N5172B->VIOpSignalGenerator::SetAmpl("-30");                 //设置ampl-30
              ptr_N5172B->VIOpSignalGenerator::viop_MODOn();
              ptr_N5172B->VIOpSignalGenerator::viop_RFOn();                         //打开RF
              ptr_N5172B->VIOpSignalGenerator::viop_bertOn();                       //打开
              QMessageBox::information(this,"提示","仪器已经配置");
          });

          connect(ui->btnRxLineRow,&QPushButton::clicked,[&](){

              if(ui->lineEdit_Row->text().isEmpty())
              {
                   QMessageBox::information(this,"警告","请输入合法数值");
                   return;
              }
              else
              {
                   row=ui->lineEdit_Row->text().toInt();
                   QMessageBox::information(this,"提示","excel行数已设置");
              }

          });

        //Wu测试中更改Excel行数
        connect(ui->WuexcelRow,&QPushButton::clicked,[&]()
          {
           if(ui->WulineRow->text().isEmpty()||ui->WulineRow->text().toInt()==0)
            {
               QMessageBox::information(this,"警告","请输入合法数值:1-999999");
               return;
            }
           else
           {
             row=ui->WulineRow->text().toInt();
             QMessageBox::information(this,"提示","excel行数已设置");
          }
          });

          //Rx多频点测量
          connect(ui->btnMulRx,&QPushButton::clicked,[&]()
          {
              //检查
              if(ui->tblWidget->item(0,0)->text().isEmpty())
              {
                  QMessageBox::information(this,"提示","请输入测试频点");
                  return;
              }

              AmplUpper=ui->EditamplUP->text().toInt();                               //ampl上限
              AmplLower=ui->EditamplLW->text().toInt();                               //ampl下限
              if(AmplUpper<AmplLower)
               {
                   QMessageBox::information(this,"错误","功率下限不能超过功率上限");
                     return;
               }

              //读取测试频点
              int columnCount=ui->tblWidget->columnCount();                       //表格列数
               QStringList frepotss;
              for(int i=0;i<columnCount;i++)
              {
                  if(ui->tblWidget->item(0,i)!=nullptr)
                  {
                      if(!(ui->tblWidget->item(0,i)->text().isEmpty()))
                      {
                           frepotss.append(ui->tblWidget->item(0,i)->text());
                      }
                  }
                  else
                  {
                      continue;
                  }
              }
               ToalBits=ui->EditTolBits->text();
               ptr_N5172B->VIOpSignalGenerator::viop_realtimeOn();                  //realtime
               ptr_N5172B->VIOpSignalGenerator::viop_modulationType();              //设置ASK
               ptr_N5172B->VIOpSignalGenerator::viop_alpha();                       //设置阿尔法为1
               ptr_N5172B->VIOpSignalGenerator::viop_alphaDepth();                  //85%
               ptr_N5172B->VIOpSignalGenerator::viop_symbolRate(512);               //symbol rate 默认单位Ksps
               ptr_N5172B->VIOpSignalGenerator::viop_loadData();                    //load文件PN9-1FF-FM0
               ptr_N5172B->VIOpSignalGenerator::viop_displayMode(SCIENTIFIC);       //科学计数法显示误码率
               ptr_N5172B->VIOpSignalGenerator::viop_BERTtriggerSource(IMM);        //trigger-immediate
               ptr_N5172B->VIOpSignalGenerator::viop_totalBit(ToalBits);            //设置totalBit为1千万
               ptr_N5172B->VIOpSignalGenerator::SetAmpl("-30");                 //设置ampl-30
               ptr_N5172B->VIOpSignalGenerator::viop_MODOn();
               ptr_N5172B->VIOpSignalGenerator::viop_RFOn();                         //打开RF
               ptr_N5172B->VIOpSignalGenerator::viop_bertOn();                       //打开
               AmplStep=ui->Editamplstep->text().toInt();                              //ampl步长
               BerStand=ui->EditBer->text().toInt();                                   //误码率
               int p583=0,p584=0;
//               bool TestFlag=false;
               for(int i=0;i<2;i++)
               {
                   //设置频率
                   if(i==0)
                   {
                     ptr_N5172B->VIOpSignalGenerator::SetFre("5.83");
                   }
                   else
                   {
                     ptr_N5172B->VIOpSignalGenerator::SetFre("5.84");
                   }
                   AmplUpper=ui->EditamplUP->text().toInt();                               //ampl上限
                   AmplLower=ui->EditamplLW->text().toInt();                               //ampl下限
                   AmplStep=ui->Editamplstep->text().toInt();                              //ampl步长
                   //5.83频点处的灵敏度为0，则程序中断
                   if(i==1&&p583==0)
                   {
                       QMessageBox::information(this,"提示","5.83GHz灵敏度为0，程序结束");
                       return;
                   }
                  //设置功率
                   for(;AmplUpper>=AmplLower;AmplUpper=AmplUpper-AmplStep)
                    {
                       QString TemAmplS=QString::number(AmplUpper);
                       ptr_N5172B->VIOpSignalGenerator::SetAmpl(TemAmplS);                  //输入ampl
                       //QMessageBox::information(this,"此时功率","此时功率"+TemAmplS);            //可以显示此时功率
                       Delay_MSec(200);
                       ptr_N5172B->VIOpSignalGenerator::viop_bertOff();                         //关闭bert
                       ptr_N5172B->VIOpSignalGenerator::viop_bertOn();                          //打开bert
                       Delay_MSec(2000);                                                        //延时约3秒等待BER稳定
                      //返回BERT结果
                       result = ptr_N5172B ->viop_Q_BertResult1();                                //调用viop_Q_BertResult()函数，得到bert测试各项结果的值
                       str_BER = QString(result);	                                                 //得到bert测试下 ber的值 根据测试要求这个值小于10的负五次方的话为合格
                       double strBER=str_BER.toDouble();
                       //写入excel
                       if(strBER>BerStand)
                       {
                           p583=AmplUpper+1;
                           AmplLower=AmplUpper+1;
//                           TestFlag=true;
                       }
                       if(strBER>BerStand&&(i!=0))
                       {
                           p584=AmplUpper+1;
                           AmplLower=AmplUpper+1;
                       }
                   }
               }

               QMessageBox::information(this,"此时功率","5.83GHz接收灵敏度"+QString::number(p583,10)+" 5.84GHz接收灵敏度"+QString::number(p584,10));   //可以显示此时功率

               //5.83和5.84两个频点灵敏度差值过大
               if(qAbs(p583-p584)>=3)
               {
                   QMessageBox::information(this,"提示","灵敏度相差过大，程序结束");
                   return;
               }

               if(!(ui->lineEditMeansRx->text().isEmpty()))
               {
                    int MeansAmpl=ui->lineEditMeansRx->text().toInt();
                    //当灵敏度大过平均灵敏度4个点
                    if(qAbs(p583-MeansAmpl)==4)
                    {
                        QMessageBox::information(this,"提示","5.83GHz的功率为"+QString::number(p583,10));
                        return;
                    }
                    //当灵敏度大过平均灵敏度5个点
                    if(qAbs(p583-MeansAmpl)==5)
                    {
                         QMessageBox::information(this,"提示","5.83GHz的功率为"+QString::number(p583,10));
                         return;
                    }
               }

               //选择文件
               QString filepath;
               filepath =QFileDialog::getOpenFileName(nullptr,("选择文件"),("C:/Users/ztp/Desktop"),("*.xls *.xlsx *.csv")); //获取保存路径
               if(filepath.isEmpty())
               {
                   QMessageBox::critical(nullptr, "错误信息", "没有找到EXCEL");
                       return;
               }

               QAxObject *excel = new QAxObject();
               excel->setControl("ket.Application");                                     //连接WPS
               excel->dynamicCall("SetVisible(bool Visible)","false");                   //不显示窗体
               excel->setProperty("DisplayAlerts", false);                               //不显示任何警告信息。如果为true那么在关闭是会出现类似“文件已修改，是否保存”的提示
               QAxObject* workbooks = excel->querySubObject("WorkBooks");                //获取工作簿集合
               workbooks->dynamicCall("Open(const QString&)",filepath);                  //打开打开已存在的工作簿
               QAxObject *workbook = excel->querySubObject("ActiveWorkBook");            //获取当前工作簿
               QAxObject *worksheets = workbook->querySubObject("WorkSheets");           //获取工作表集合
               QAxObject *worksheet = worksheets->querySubObject("Item(int)",1);         //获取工作表集合的工作表1，即sheet1
               QAxObject *cellA;

               if(RxMultTextFirst)
               {
                    cellA= worksheet->querySubObject("Cells(int,int)",1,1);
                    cellA->dynamicCall("SetValue(const QVariant&)","No.");           //写入数据
                    cellA->setProperty("HorizontalAlignment", -4108);                     //居中
                    cellA->setProperty("VerticalAlignment", -4108);                       //居中
                    cellA = worksheet->querySubObject("Cells(int,int)",1,2);
                    cellA->dynamicCall("SetValue(const QVariant&)","5.83");         //写入数据
                    cellA->setProperty("HorizontalAlignment", -4108);                     //居中
                    cellA->setProperty("VerticalAlignment", -4108);                       //居中]
                    cellA = worksheet->querySubObject("Cells(int,int)",1,3);
                    cellA->dynamicCall("SetValue(const QVariant&)","5.84");         //写入数据
                    cellA->setProperty("HorizontalAlignment", -4108);                     //居中
                    cellA->setProperty("VerticalAlignment", -4108);                       //居中
                    for(int i=0;i<frepotss.length();i++)
                    {
                        cellA= worksheet->querySubObject("Cells(int,int)",1,4+i);
                        cellA->dynamicCall("SetValue(const QVariant&)",frepotss[i]);         //写入数据
                        cellA->setProperty("HorizontalAlignment", -4108);                     //居中
                        cellA->setProperty("VerticalAlignment", -4108);                       //居中
                    }
                    RxMultTextFirst=false;
               }

               cellA = worksheet->querySubObject("Cells(int,int)",row,1);
               cellA->dynamicCall("SetValue(const QVariant&)",ui->EditBoard->text());                //
               cellA->setProperty("HorizontalAlignment", -4108);
               cellA->setProperty("VerticalAlignment", -4108);

               cellA = worksheet->querySubObject("Cells(int,int)",row,2);
               cellA->dynamicCall("SetValue(const QVariant&)",p583);                                 //
               cellA->setProperty("HorizontalAlignment", -4108);
               cellA->setProperty("VerticalAlignment", -4108);

               cellA = worksheet->querySubObject("Cells(int,int)",row,3);
               cellA->dynamicCall("SetValue(const QVariant&)",p584);                                       //
               cellA->setProperty("HorizontalAlignment", -4108);
               cellA->setProperty("VerticalAlignment", -4108);

             //判断5.83  频率点
              //功率确定
              AmplUpper=p583+6;
              ptr_N5172B->VIOpSignalGenerator::SetAmpl(QString::number(AmplUpper));                  //输入ampl
              for(int i=0;i<5;i++)
              {
                    ptr_N5172B->VIOpSignalGenerator::SetFre(frepotss[i]);
                    ptr_N5172B->VIOpSignalGenerator::SetAmpl(QString::number(AmplUpper));                  //输入ampl
                    Delay_MSec(200);
                    ptr_N5172B->VIOpSignalGenerator::viop_bertOff();                         //关闭bert
                    ptr_N5172B->VIOpSignalGenerator::viop_bertOn();                          //打开bert
                    Delay_MSec(2000);                                                        //延时约3秒等待BER稳定
                   //返回BERT结果
                    result = ptr_N5172B ->viop_Q_BertResult1();                                //调用viop_Q_BertResult()函数，得到bert测试各项结果的值
                    str_BER = QString(result);	                                                 //得到bert测试下 ber的值 根据测试要求这个值小于10的负五次方的话为合格
                    double strBER=str_BER.toDouble();
                    if(strBER<=BerStand)                   //判断数据
                    {
                        QAxObject *cellA0 = worksheet->querySubObject("Cells(int,int)",row,4+i);
                        cellA0->dynamicCall("SetValue(const QVariant&)","T");
                        cellA0->setProperty("HorizontalAlignment", -4108);                          //居中
                        cellA0->setProperty("VerticalAlignment", -4108);                            //居中
                    }
                    else
                    {
                        QAxObject *cellA0 = worksheet->querySubObject("Cells(int,int)",row,4+i);
                        cellA0->dynamicCall("SetValue(const QVariant&)","F");                   //写入板号
                        QAxObject *font = cellA0->querySubObject("Font");            //获取单元格字体
                        font->setProperty("Color", QColor(255, 0, 0));               //设置单元格字体颜色（红色）
                        font->setProperty("Bold", true);
                        cellA0->setProperty("HorizontalAlignment", -4108);                          //居中
                        cellA0->setProperty("VerticalAlignment", -4108);                            //居中
                    }
              }
              //判断5.84  频率点
              //功率确定
              AmplUpper=p584+6;
              for(int i=5;i<(frepotss.length());i++)
              {
                    ptr_N5172B->VIOpSignalGenerator::SetFre(frepotss[i]);
                    ptr_N5172B->VIOpSignalGenerator::SetAmpl(QString::number(AmplUpper));                  //输入ampl
                    Delay_MSec(200);
                    ptr_N5172B->VIOpSignalGenerator::viop_bertOff();                         //关闭bert
                    ptr_N5172B->VIOpSignalGenerator::viop_bertOn();                          //打开bert
                    Delay_MSec(2000);                                                        //延时约3秒等待BER稳定
                   //返回BERT结果
                    result = ptr_N5172B ->viop_Q_BertResult();                                //调用viop_Q_BertResult()函数，得到bert测试各项结果的值
                    str_BER = QString(result);	                                                 //得到bert测试下 ber的值 根据测试要求这个值小于10的负五次方的话为合格
                    double strBER=str_BER.toDouble();
                    if(strBER<=BerStand)                   //判断数据
                    {
                        cellA= worksheet->querySubObject("Cells(int,int)",row,4+i);
                        cellA->dynamicCall("SetValue(const QVariant&)","T");
                        cellA->setProperty("HorizontalAlignment", -4108);                          //居中
                        cellA->setProperty("VerticalAlignment", -4108);                            //居中
                    }
                    else
                    {
                        cellA = worksheet->querySubObject("Cells(int,int)",row,4+i);
                        cellA->dynamicCall("SetValue(const QVariant&)","F");
                        QAxObject *font = cellA->querySubObject("Font");            //获取单元格字体
                        font->setProperty("Color", QColor(255, 0, 0));               //设置单元格字体颜色（红色）
                        font->setProperty("Bold", true);                             //设置单元格字体加粗
                        cellA->setProperty("HorizontalAlignment", -4108);                          //居中
                        cellA->setProperty("VerticalAlignment", -4108);                            //居中
                    }
              }
               //excel行数增加
               row++;    //全局变量，行数加一
               ui->lineEdit_Row->setText(QString::number(row));

               //芯片号增加
               int chiNumMulRx;
               chiNumMulRx=ui->EditBoard->text().toInt();                               //记录芯片号
               chiNumMulRx++;
               ui->EditBoard->setText(QString::number(chiNumMulRx));//显示下一个芯片号
               frepotss.clear();
               free(result);
//               delete [] frepotss;
               //关闭文件
               workbook->dynamicCall("SaveAs(const QString&)",QDir::toNativeSeparators(filepath));
               workbook->dynamicCall("Close()");                              //关闭工作簿
               excel->dynamicCall("Quit()");                                  //关闭excel
               delete excel;
               excel=nullptr;
               QMessageBox::information(this,"RX接收灵敏度测试","完成");
           });

          //Rx最大功率测量
      connect(ui->btnRxMax,&QPushButton::clicked,[&]()
     {
          QString textLab;
          ToalBits=ui->EditTolBits->text();
          ptr_N5172B->VIOpSignalGenerator::viop_realtimeOn();                  //realtime
          ptr_N5172B->VIOpSignalGenerator::viop_modulationType();              //设置ASK
          ptr_N5172B->VIOpSignalGenerator::viop_alpha();                       //设置阿尔法为1
          ptr_N5172B->VIOpSignalGenerator::viop_alphaDepth();                  //85%
          ptr_N5172B->VIOpSignalGenerator::viop_symbolRate(512);               //symbol rate 默认单位Ksps
          ptr_N5172B->VIOpSignalGenerator::viop_loadData();                    //load文件PN9-1FF-FM0
          ptr_N5172B->VIOpSignalGenerator::viop_displayMode(SCIENTIFIC);       //科学计数法显示误码率
          ptr_N5172B->VIOpSignalGenerator::viop_BERTtriggerSource(IMM);        //trigger-immediate
          ptr_N5172B->VIOpSignalGenerator::viop_totalBit(ToalBits);            //设置totalBit为1千万
          ptr_N5172B->VIOpSignalGenerator::viop_MODOn();
          ptr_N5172B->VIOpSignalGenerator::viop_RFOn();                         //打开RF
          ptr_N5172B->VIOpSignalGenerator::viop_bertOn();                       //打开

         ptr_N5172B->VIOpSignalGenerator::SetFre(ui->RxMaxFre->text());                       //设置频率
         textLab=ui->EditBoard->text()+"号 "+ui->RxMaxFre->text();


          //  根据实际赋值
          AmplUpper=ui->RxMaxRan->text().toInt();                               //ampl上限
          AmplLower=ui->RxMaxPoint->text().toInt()+6;                           //ampl下限
          AmplStep=ui->RxMaxStep->text().toInt();                            //ampl步长

          BerStand=ui->EditBer->text().toDouble();                            //误码率

          //检查输入情况
          if(AmplUpper<AmplLower)
             {
               QMessageBox::information(this,"错误","数值输入错误");
                 return;
             }
          //选择文件
          QString filepath;
          filepath =QFileDialog::getOpenFileName(nullptr,("选择文件"),("C:/Users/ztp/Desktop"),("*.xls *.xlsx *.csv")); //获取保存路径
          if(filepath.isEmpty())
          {
              QMessageBox::critical(nullptr, "错误信息", "没有找到EXCEL");
                  return;
          }
          QAxObject *excel = new QAxObject();
          excel->setControl("ket.Application");                                     //连接WPS
          excel->dynamicCall("SetVisible(bool Visible)","false");                   //不显示窗体
          excel->setProperty("DisplayAlerts", false);                               //不显示任何警告信息。如果为true那么在关闭是会出现类似“文件已修改，是否保存”的提示
          QAxObject* workbooks = excel->querySubObject("WorkBooks");                //获取工作簿集合
          workbooks->dynamicCall("Open(const QString&)",filepath);                  //打开打开已存在的工作簿
          QAxObject *workbook = excel->querySubObject("ActiveWorkBook");            //获取当前工作簿
          QAxObject *worksheets = workbook->querySubObject("WorkSheets");           //获取工作表集合
          QAxObject *worksheet = worksheets->querySubObject("Item(int)",1);         //获取工作表集合的工作表1，即sheet1
          QAxObject *cellA1;

          if(RxMaxFirst)
          {
               cellA1 = worksheet->querySubObject("Cells(int,int)",1,1);
               cellA1->dynamicCall("SetValue(const QVariant&)","Rx最大功率测试");         //写入数据
               cellA1->setProperty("HorizontalAlignment", -4108);                     //居中
               cellA1->setProperty("VerticalAlignment", -4108);                       //居中
               RxMaxFirst=false;
          }

          textLab=ui->EditBoard->text()+ textLab;
          cellA1= worksheet->querySubObject("Cells(int,int)",row++,1);
          cellA1->dynamicCall("SetValue(const QVariant&)",textLab);                   //写入板号
          cellA1->setProperty("HorizontalAlignment", -4108);                          //居中
          cellA1->setProperty("VerticalAlignment", -4108);                            //居中
          int colum=2;

          // 第二行数据进行操作
          cellA1 = worksheet->querySubObject("Cells(int,int)",row++,1);
          cellA1->dynamicCall("SetValue(const QVariant&)","功率(dBm)");         //写入功率
          cellA1->setProperty("HorizontalAlignment", -4108);                   //居中
          cellA1->setProperty("VerticalAlignment", -4108);                     //居中

          cellA1 = worksheet->querySubObject("Cells(int,int)",row++,1);
          cellA1->dynamicCall("SetValue(const QVariant&)","误码率");             //写入误码率
          cellA1->setProperty("HorizontalAlignment", -4108);                    //居中
          cellA1->setProperty("VerticalAlignment", -4108);                      //居中
          cellA1 = worksheet->querySubObject("Cells(int,int)",row,1);
          cellA1->dynamicCall("SetValue(const QVariant&)","结果");               //写入结果
          cellA1->setProperty("HorizontalAlignment", -4108);                    //居中
          cellA1->setProperty("VerticalAlignment", -4108);                      //居中

              //功率循环
              for(;AmplLower<AmplUpper;AmplLower=AmplLower+AmplStep)
               {
                  QString TemAmplS=QString::number(AmplLower);
                  ptr_N5172B->VIOpSignalGenerator::SetAmpl(TemAmplS);                  //输入ampl
                  Delay_MSec(100);
                  ptr_N5172B->VIOpSignalGenerator::viop_bertOff();                         //关闭bert
                  ptr_N5172B->VIOpSignalGenerator::viop_bertOn();                          //打开bert
                  Delay_MSec(2000);                                                        //延时约4秒等待BER稳定
                  //返回BERT结果
                  result = ptr_N5172B ->viop_Q_BertResult1();                               //调用viop_Q_BertResult()函数，得到bert测试各项结果的值
//                       p = strtok(result,d);
//                       str_totalBits=QString(p);           //得到bert测试下 totalBits的值
//                       p = strtok(nullptr, d);
//                       str_errorBits = QString(p);        //得到bert测试下 errorBits
//                       p = strtok(nullptr, d);
                  str_BER = QString(result);	               //得到bert测试下 ber的值 根据测试要求这个值小于10的负五次方的话为合格
                  double strBER=str_BER.toDouble();

                  //写入excel
                  cellA1 = worksheet->querySubObject("Cells(int,int)",row-2,colum);
                  cellA1->dynamicCall("SetValue(const QVariant&)",TemAmplS);   //写入功率
                  cellA1->setProperty("HorizontalAlignment", -4108);           //居中
                  cellA1->setProperty("VerticalAlignment", -4108);             //居中

                  cellA1 = worksheet->querySubObject("Cells(int,int)",row-1,colum);
                  cellA1->dynamicCall("SetValue(const QVariant&)",str_BER);    //写入误码率
                  cellA1->setProperty("HorizontalAlignment", -4108);           //居中
                  cellA1->setProperty("VerticalAlignment", -4108);             //居中

                  if(strBER<=BerStand)
                  {
                     cellA1 = worksheet->querySubObject("Cells(int,int)",row,colum);
                     cellA1->dynamicCall("SetValue(const QVariant&)","T");   //写入结果
                     cellA1->setProperty("HorizontalAlignment", -4108);         //居中
                     cellA1->setProperty("VerticalAlignment", -4108);           //居中
                  }
                  if(strBER>BerStand)
                  {
                     cellA1 = worksheet->querySubObject("Cells(int,int)",row,colum);
                     cellA1->dynamicCall("SetValue(const QVariant&)","F");    //写入结果
//                          QAxObject *font = cellA1->querySubObject("Font");            //获取单元格字体
//                          font->setProperty("Bold", "true");                        //设置单元格字体颜色（红色）
//                          font->setProperty("Color", QColor(255, 0, 0));               //设置单元格字体颜色（红色）
                     cellA1->setProperty("HorizontalAlignment", -4108);           //居中
                     cellA1->setProperty("VerticalAlignment", -4108);             //居中
                  }
                  colum++;
              }
          colum=2;
          row++;
          ui->lineEdit_Row->setText(QString::number(row));
          //芯片号增加
          chipNum=ui->EditBoard->text().toInt();
          chipNum++;
          ui->EditBoard->setText(QString::number(chipNum));//显示下一个芯片号
          free(result);
          //关闭文件
          workbook->dynamicCall("SaveAs(const QString&)",QDir::toNativeSeparators(filepath));
          workbook->dynamicCall("Close()");                               //关闭工作簿
          excel->dynamicCall("Quit()");                                   //关闭excel
          delete excel;
          excel=nullptr;
         QMessageBox::information(this,"RX接收灵敏度测试","完成");

     });

    //设置功率
      connect(ui->btn_power,&QPushButton::clicked,[&]()
      {
         ptr_N5172B->VIOpSignalGenerator::viop_cwAmpl(ui->lEdit_N5172B_Power->text());
      });

//       connect(ui->btn_TRx,&QPushButton::clicked,[&]()
//       {
//           ptr_N5172B->VIOpSignalGenerator::viop_bertOn();                       //打开
//            ptr_N5172B->VIOpSignalGenerator::SetFre("5.83");
//           for(int i=-70;i>-80;i--)
//           {
//                ptr_N5172B->VIOpSignalGenerator::SetAmpl(QString::number(i));
//                Delay_MSec(200);
//                ptr_N5172B->VIOpSignalGenerator::viop_bertOff();                       //打开
//                ptr_N5172B->VIOpSignalGenerator::viop_bertOn();                       //打开
//                Delay_MSec(2000);
//                result= ptr_N5172B ->viop_Q_BertResult1();
//               QString s=result;
//               QMessageBox::information(this,"数据读取完成",s);
//           }
//          QMessageBox::information(this,"数据读取完成","完成");

//      });


     //设置频率
      connect(ui->btn_fre,&QPushButton::clicked,[&]()
      {
           int Unit;
           if(ui->btnRadioKHz->isChecked())
           {
               Unit=KHz;
           }
           else if(ui->btnRadioMHz->isChecked())
           {
               Unit=MHz;
           }
           else{
               Unit=GHz;
           }
         ptr_N5172B-> VIOpSignalGenerator::viop_cwFreq(ui->lEdit_N5172B_Freq->text(), Unit);
       });

    //打开/关闭bert
    connect(ui->btnBERTOnOff,&QPushButton::clicked,[&]()
    {
        if(bertKey)
        {
          ptr_N5172B->VIOpSignalGenerator::viop_bertOn();               //打开
          bertKey=false;
        }else
        {
           ptr_N5172B->VIOpSignalGenerator::viop_bertOff();             //关闭
           bertKey=true;
        }
    });

    //重置仪器
    connect(ui->btnRST,&QPushButton::clicked,[&]()
    {
        ptr_N5172B->viop_RST();
    });

    //添加列
    connect(ui->btnTablAddColum,&QPushButton::clicked,[&]()
    {
        int columIndex = ui->tblWidget->columnCount();  //列总数
        int columCurrent = ui->tblWidget-> currentColumn(); //当前列
        if(columCurrent!=-1)
        {
            ui->tblWidget->insertColumn(columCurrent+1);    //当前列前面增加1
        }
        else
        {
            ui->tblWidget->setColumnCount(columIndex + 1);  //总列数增加1
        }
         ui->tblWidget->setVerticalHeaderLabels(QStringList()<<"频率"<<"功率");          //设置表头
         ui->tblWidget->horizontalHeader()->setSectionResizeMode(QHeaderView::Stretch);
         ui->tblWidget->verticalHeader()->setSectionResizeMode(QHeaderView::Stretch);
    });

    //添加行
    connect(ui->btnTablAddRow,&QPushButton::clicked,[&]()
    {
         int rowIndex = ui->tblWidget->rowCount();         //行总数
         int rowCurrent = ui->tblWidget->currentRow();     //当前行
         if(rowCurrent!=-1)
         {
             ui->tblWidget->insertRow(rowCurrent+1);    //当前列前面增加1
         }
         else
         {
             ui->tblWidget->setRowCount(rowIndex + 1);  //总列数增加1
         }
         ui->tblWidget->setVerticalHeaderLabels(QStringList()<<"频率"<<"功率");          //设置表头
         ui->tblWidget->horizontalHeader()->setSectionResizeMode(QHeaderView::Stretch);
         ui->tblWidget->verticalHeader()->setSectionResizeMode(QHeaderView::Stretch);
    });

    //删除选中列
    connect(ui->btnTableDelColum,&QPushButton::clicked,[&]()
    {
        //resizeColumnsToContents();                      根据内容调整列宽
        QMessageBox *msgBox;
        msgBox=new QMessageBox("删除列","是否删除选中列",QMessageBox::Information,	QMessageBox::Yes | QMessageBox::Default,QMessageBox::No | QMessageBox::Escape,0);
        if(msgBox->exec() == QMessageBox::No)
        {
             return;
        }
        int columIndex = ui->tblWidget-> currentColumn();
        if (columIndex != -1)
            ui->tblWidget->removeColumn(columIndex);
        ui->tblWidget->setVerticalHeaderLabels(QStringList()<<"频率"<<"功率");          //设置表头
        ui->tblWidget->horizontalHeader()->setSectionResizeMode(QHeaderView::Stretch);
        ui->tblWidget->verticalHeader()->setSectionResizeMode(QHeaderView::Stretch);
    });

    //删除选中行
    connect(ui->btnTableDelRow,&QPushButton::clicked,[&]()
    {
        QMessageBox *msgBox;
        msgBox=new QMessageBox("删除行","是否删除选中行",QMessageBox::Information,	QMessageBox::Yes | QMessageBox::Default,QMessageBox::No | QMessageBox::Escape,0);
        if(msgBox->exec() == QMessageBox::No)
        {
             return;
        }
        int rowIndex = ui->tblWidget->currentRow();
        if (rowIndex != -1)
            ui->tblWidget->removeRow(rowIndex);
        ui->tblWidget->setVerticalHeaderLabels(QStringList()<<"频率"<<"功率");          //设置表头
        ui->tblWidget->horizontalHeader()->setSectionResizeMode(QHeaderView::Stretch);
        ui->tblWidget->verticalHeader()->setSectionResizeMode(QHeaderView::Stretch);
    });

     connect(ui->btnRxLoadData,&QPushButton::clicked,[&]()
      {
         QAxObject *cell;
         QString filepath;
         QVariant cell_value;
         QAxObject *range;
         QString strVal;

         int rowIndex = ui->tblWidget->rowCount();         //ui界面 行总数
         int columIndex = ui->tblWidget->columnCount();    //ui界面 列总数
         filepath =QFileDialog::getOpenFileName(nullptr,("选择文件"),("C:/Users/ztp/Desktop"),("*.xls *.xlsx *.csv")); //获取保存路径
         if(filepath.isEmpty())
         {
             QMessageBox::critical(nullptr, "错误信息", "没有找到EXCEL");
                 return;
         }
         QAxObject *excel = new QAxObject();
         excel->setControl("ket.Application");                                     //连接WPS
         excel->dynamicCall("SetVisible(bool Visible)","false");                   //不显示窗体
         excel->setProperty("DisplayAlerts", false);                               //不显示任何警告信息。如果为true那么在关闭是会出现类似“文件已修改，是否保存”的提示
         QAxObject* workbooks = excel->querySubObject("WorkBooks");                //获取工作簿集合
         workbooks->dynamicCall("Open(const QString&)",filepath);                  //打开打开已存在的工作簿
         QAxObject *workbook = excel->querySubObject("ActiveWorkBook");            //获取当前工作簿
         QAxObject *worksheets = workbook->querySubObject("WorkSheets");           //获取工作表集合
         QAxObject *worksheet = worksheets->querySubObject("Item(int)",1);         //获取工作表集合的工作表1，即sheet1
         QAxObject *usedRange = worksheet->querySubObject("UsedRange");            //获取表格中的数据范围
         QAxObject *rows =usedRange->querySubObject("Rows");
         QAxObject *columns = usedRange->querySubObject("Columns");
         int row_start =usedRange->property("Row").toInt();              //获取起始行
         int column_start = usedRange->property("Column").toInt();       //获取起始列
         int row_count = rows->property("Count").toInt();                //获取行数
         int column_count = columns->property("Count").toInt();         //获取列数
         //防止数据量大于当前单元格数量

         if(row_count>rowIndex)
         {
             ui->tblWidget->setRowCount(row_count);
         }
         if(column_count>columIndex)
         {
             ui->tblWidget->setRowCount(column_count);
         }

         //写数据
         for(int i=row_start;i<=row_count;i++)
         {
             for(int j=column_start;j<=column_count;j++)
             {
                  cell= worksheet->querySubObject("Cells(int,int)", i, j);
                  cell_value= cell->property("Value");                //获取单元格内容
                  range= worksheet->querySubObject("Cells(int,int)",i,j); //获取cell的值
                  strVal= range->dynamicCall("Value2()").toString();
                  ui->tblWidget->setItem(i-1,j-1,new QTableWidgetItem(strVal));
             }
         }

         //清除其余单元格数据
         for(int i=0;i<rowIndex;i++)
         {
             for(int j=column_count;j<columIndex;j++)
             {
                    ui->tblWidget->setItem(i,j,new QTableWidgetItem(""));
             }
         }
       //关闭文件
         workbook->dynamicCall("SaveAs(const QString&)",QDir::toNativeSeparators(filepath));
         workbook->dynamicCall("Close()");                              //关闭工作簿
         excel->dynamicCall("Quit()");                                  //关闭excel
         delete excel;
         excel=nullptr;
         QMessageBox::information(this,"数据读取完成","完成");
       });

     //Rx跳点测试
     connect(ui->btnMulPots,&QPushButton::clicked,[&]()
     {
       if((ui->tblWidget->item(0,0)->text().isEmpty())||(ui->tblWidget->item(0,1)->text().isEmpty()))
         {
           QMessageBox::information(this,"请检查输入数据","确认");
            return;
         }
        QStringList frepotss,amplpots;
//        int rowIndex = ui->tblWidget->rowCount();         //ui界面 行总数
        int columIndex = ui->tblWidget->columnCount();    //ui界面 列总数

        for(int i=0;i<2;i++)
        {
            for(int j=0;j<columIndex;j++)
            {
                if((ui->tblWidget->item(i,j)!=nullptr) &&(!ui->tblWidget->item(i,j)->text().isEmpty()))
                {
                    if(i==0)
                    {
                      frepotss.append(ui->tblWidget->item(i,j)->text());
                    }
                    else
                    {
                        amplpots.append(ui->tblWidget->item(i,j)->text());
                    }
                }

            }
        }
//         QMessageBox::information(this,"",QString::number(frepotss.length()));
//          QMessageBox::information(this,"",QString::number(amplpots.length()));
//          return;

        //选择文件
        QString filepaths;
        filepaths =QFileDialog::getOpenFileName(nullptr,("选择文件"),("C:/Users/ztp/Desktop"),("*.xls *.xlsx *.csv")); //获取保存路径
        if(filepaths.isEmpty())
        {
            QMessageBox::critical(nullptr, "错误信息", "没有找到EXCEL");
                return;
        }

        QAxObject *excel = new QAxObject();
        excel->setControl("ket.Application");                                     //连接WPS
        excel->dynamicCall("SetVisible(bool Visible)","false");                   //不显示窗体
        excel->setProperty("DisplayAlerts", false);                               //不显示任何警告信息。如果为true那么在关闭是会出现类似“文件已修改，是否保存”的提示
        QAxObject* workbooks = excel->querySubObject("WorkBooks");                //获取工作簿集合
        workbooks->dynamicCall("Open(const QString&)",filepaths);                  //打开打开已存在的工作簿
        QAxObject *workbook = excel->querySubObject("ActiveWorkBook");            //获取当前工作簿
        QAxObject *worksheets = workbook->querySubObject("WorkSheets");           //获取工作表集合
        QAxObject *worksheet = worksheets->querySubObject("Item(int)",1);         //获取工作表集合的工作表1，即sheet1
        QAxObject *cellA0;
        QAxObject *font;
        int excelColum=2;

        if(RxMulPotsTestFirst)
        {
            cellA0= worksheet->querySubObject("Cells(int,int)",1,1);
            cellA0->dynamicCall("SetValue(const QVariant&)","No.");      //序列
            font= cellA0->querySubObject("Font");            //获取单元格字体
//            font->setProperty("Bold", true);                             //设置单元格字体加粗
            cellA0->setProperty("HorizontalAlignment", -4108);                          //居中
            cellA0->setProperty("VerticalAlignment", -4108);                            //居中

            for(int i=0;i<frepotss.length();i++)
            {
                for(int j=0;j<amplpots.length();j++)
                {
                      cellA0= worksheet->querySubObject("Cells(int,int)",1,excelColum++);
                      cellA0->dynamicCall("SetValue(const QVariant&)",frepotss[i]+"GHz"+amplpots[j]+"dBm");
                      font=cellA0->querySubObject("Font");                                        //获取单元格字体
//                      font->setProperty("Bold", true);                                            //设置单元格字体加粗
                      cellA0->setProperty("HorizontalAlignment", -4108);                          //居中
                      cellA0->setProperty("VerticalAlignment", -4108);                            //居中
                }
            }
            RxMulPotsTestFirst=false;
        }
        excelColum=2;
        cellA0= worksheet->querySubObject("Cells(int,int)",row,1);
        cellA0->dynamicCall("SetValue(const QVariant&)",ui->EditBoard->text());      //序列号

        //配置文件
        chipNum=ui->EditBoard->text().toInt();                               //记录芯片号
        ToalBits=ui->EditTolBits->text();
        BerStand=ui->EditBer->text().toInt();                                //误码率
        ptr_N5172B->VIOpSignalGenerator::viop_realtimeOn();                  //realtime
        ptr_N5172B->VIOpSignalGenerator::viop_modulationType();              //设置ASK
        ptr_N5172B->VIOpSignalGenerator::viop_alpha();                       //设置阿尔法为1
        ptr_N5172B->VIOpSignalGenerator::viop_alphaDepth();                  //85%
        ptr_N5172B->VIOpSignalGenerator::viop_symbolRate(512);               //symbol rate 默认单位Ksps
        ptr_N5172B->VIOpSignalGenerator::viop_loadData();                    //load文件PN9-1FF-FM0
        ptr_N5172B->VIOpSignalGenerator::viop_displayMode(SCIENTIFIC);       //科学计数法显示误码率
        ptr_N5172B->VIOpSignalGenerator::viop_BERTtriggerSource(IMM);        //trigger-immediate
        ptr_N5172B->VIOpSignalGenerator::viop_totalBit(ToalBits);            //设置totalBit为实际数值
        ptr_N5172B->VIOpSignalGenerator::SetAmpl("-30");                     //设置ampl-30
        ptr_N5172B->VIOpSignalGenerator::viop_MODOn();
        ptr_N5172B->VIOpSignalGenerator::viop_RFOn();                         //打开RF
        ptr_N5172B->VIOpSignalGenerator::viop_bertOn();                       //打开

        for(int i=0;i<=frepotss.length()-1;i++)
        {
            ptr_N5172B->VIOpSignalGenerator::SetFre(frepotss[i]);             //配置频率
            Delay_MSec(500);
            for(int j=0;j<=amplpots.length()-1;j++)
            {
               ptr_N5172B->VIOpSignalGenerator::SetAmpl(amplpots[j]);               //输入ampl
               Delay_MSec(200);
               ptr_N5172B->VIOpSignalGenerator::viop_bertOff();                         //关闭bert
               ptr_N5172B->VIOpSignalGenerator::viop_bertOn();                          //打开bert
               Delay_MSec(2000);                                                        //延时约3秒等待BER稳定
              //返回BERT结果
               result = ptr_N5172B ->viop_Q_BertResult1();                                //调用viop_Q_BertResult()函数，得到bert测试各项结果的值
               str_BER = QString(result);	                                                 //得到bert测试下 ber的值 根据测试要求这个值小于10的负五次方的话为合格
               double strBER=str_BER.toDouble();
               if(strBER<=BerStand)                                                      //判断数据
               {
                   cellA0 = worksheet->querySubObject("Cells(int,int)",row,excelColum++);
                   cellA0->dynamicCall("SetValue(const QVariant&)","T");
                   cellA0->setProperty("HorizontalAlignment", -4108);                          //居中
                   cellA0->setProperty("VerticalAlignment", -4108);                            //居中
               }
               else
               {
                   cellA0 = worksheet->querySubObject("Cells(int,int)",row,excelColum++);
                   cellA0->dynamicCall("SetValue(const QVariant&)","F");                   //写入板号
                   cellA0->setProperty("HorizontalAlignment", -4108);                      //居中
                   cellA0->setProperty("VerticalAlignment", -4108);                         //居中
               }
            }
        }

        QMessageBox::information(this,"RX接收灵敏度测试","信号源测试完成");
        row++;
        ui->lineEdit_Row->setText(QString::number(row));
        //芯片号增加
        chipNum++;
        ui->EditBoard->setText(QString::number(chipNum));//显示下一个芯片号
        frepotss.empty();
        amplpots.empty();
        free(result);
        frepotss.clear();
        amplpots.clear();
        //关闭文件
        workbook->dynamicCall("SaveAs(const QString&)",QDir::toNativeSeparators(filepaths));
        workbook->dynamicCall("Close()");                              //关闭工作簿
        excel->dynamicCall("Quit()");                                  //关闭excel
        delete excel;
        excel=nullptr;
        QMessageBox::information(this,"RX接收灵敏度测试","完成");
     });

     connect(ui->btnTxEx,&QPushButton::clicked,[&]()
     {
         QString str="Tx输出功率范围测试\
                 测试目的：该测试利用输出芯片输出信号在频谱仪上观察波形峰值处的功率值，程序自动记录    \
                 连接设置：pc连接芯片，芯片输出端连接频谱仪。PC和频谱仪需要用网线连接。    \
                 芯片配置文件：  \
                 CPC5801803_TX_test_5790_MAX_20200831.txt   \
                 具体流程：       \
                 1）确保所有的数据线都已经连接好，给芯片配置好文件   \
                 2）根据芯片配置文件设置频谱仪频率和span参数    \
                 3）选择写入的excel文件。注意只有在第一颗芯片测量过程中，点击一次该按钮。\
                 4）写入数据。整个测量过程主要点击写入按钮。    \
                 这一步程序会自动检测频谱仪peak处的功率值并记录在excel中\
                 5）结束程序，释放excel。注意最后一颗芯片测完才能点击\
                 由于这项测试针对芯片所进行的操作比较少，所以为了减少程序对计算机的资源消耗，整个excel操作的线程不释放，直到点击结束按钮。";

        ui->plainTextEdit->setPlainText(str);
     });
     connect(ui->btnRxEx,&QPushButton::clicked,[&]()
     {
         QString str="1、Rx灵敏度测试\
                 测试目的：判断芯片在某一频率下、能够满足误码率要求（通常是，其他涉及误码率的测试，均以此值为标准）的灵敏度下限。\
                 连接设置：pc连接芯片，芯片连接N5172B信号发生器，注意连接s1和s2线。      \
                 芯片配置文件：\
                 CPC5801803_RX_test_2020_cg_2_filter_5_lindao_20200902.txt\
                 信号发生器配置：symbol rate 512ksps 、alpha 1.0、文件配置PN9-1FF-FM0、Total Bits 两百万或者一千万、realtiem设置为on、打开RF。    \
                 具体流程：    \
                 1）确保所有的数据线都已经连接好，给芯片配置好文件。\
                 2）设置信号源配置\
                 功率上界和下界的值要根据实际情况设置，为了保证测量的准确性可以设的宽一些。Total Bits一般两百万即可，设置过大时可能导致耗时较长或者测试错误。\
                 3）选择频率\
                 4）点击按钮开始测试，然后在弹窗处选择提前准备好的excel文件。        \
                 点击该按钮时，程序会自动给信号源按照设置的参数进行配置。\
                 频率选择5.83GHz时，测试5.83GHZ通道，从功率上界开始，以步长为单位扫描到功率下届。测试完成后，可以选择5.84GHz，再次点击开始按钮，进行5.84GHz通道的接收灵敏度测试。\
                 5）程序测试完成后，会弹出测试结束的窗口，同时芯片号文本框自动加一。程序中的所有文本框都是可以操作的，即手动更改数据。\
                 2、Rx频点测试\
                 该测试会按照普通接收灵敏度测试模式先在5.83GHz和5.84GHz两个通道下进行接收灵敏度测试并且进行比较。当找不到某一个通道的灵敏度或者两个通道的灵敏度值相差过大（绝对值大于等于3）时，程序终止。       \
                 当5.83GHz和5.84GHz两个通道的灵敏度绝对值小于等于2时，程序会判断两个通道的接收灵敏度与平均灵敏度的差值。平均灵敏度的值需要提前设置，一般是多次测量接收灵敏度后得到的均值。如果某个通道的灵敏度与平均灵敏度的值大于等于4，则程序终止。\
                 以上步骤完成后，程序将两个通道的灵敏度分别加六然后在提前设置好的频点下进行灵敏度测量。即固定功率，变更频率，看是否满足误码率要求。  \
                 目前经常测量的频点为：\
                 5.8283/5.8285/5.8315/5.8316/5.8317/5.8383/5.8385/5.8415/5.8417GHz，而且程序会检测5.83GHz通道下的接收灵敏度加六后在\
                 5.8283/5.8285/5.8315/5.8316/5.8317频点下是否满足误码率要求，\
                 以及5.84GHz通道下的接收灵敏度加六后在5.8383/5.8385/5.8415/\
                 5.8417GHz频点下的误码率情况。\
                 测试目的：当芯片的5.83和5.84GHz两个通道的接收灵敏度误差不大的情况下，检测灵敏度+6在不同频点下的接收误码率情况。\
                 连接设置：pc连接芯片，芯片连接N5172B信号发生器，注意连接s1和s2线。      \
                 芯片配置文件：\
                 CPC5801803_RX_test_2020_cg_2_filter_5_lindao_20200902.txt\
                 信号发生器配置：symbol rate 512ksps 、alpha 1.0、文件配置PN9-1FF-FM0、Total Bits 两百万或者一千万、realtiem设置为on、打开RF。    \
                 具体流程：  \
                 1）确保所有的数据线都已经连接好，给芯片配置好文件\
                 2）设置信号源配置\
                 3）设置好测试频点\
                 频点设置可以手动设置，直接在软件表格中更改，单位均为GHz。或者选择已经准备好的excel文件，将频点读入软件表格。点击读取测试数据按钮可以读取数据。数据在excel中只需填写频点即可，且只能写在第一行如下图所示。\
                 当手动输入频点时，可以根据实际需要使用添加和删除按钮操作软件表格的行与列。但是频点是能填写在软件表格的第一行，第二行是为填写功率做准备的，在Rx跳点测试中使用。   \
                 4）点击按钮开始测试，然后在弹窗处选择提前准备好的excel文件\
                 Rx频点测试在测试过程中会根据实际测试情况进行不同的判定与结果输出，所以一定要注意程序在测试过程中的提示。\
                 3、Rx跳点测试  \
                 Rx跳点测试是根据指定的频率与功率组合，测量误码率的测试项目。在连接好数据线等之后，需要先把频率和功率值读进软件表格。这一步可以手动填写，也可以读取excel文件实现。  \
                 测试目的：测试频率点在不同功率点下的误码率情况。频率和功率值都是不连续的。\
                 连接设置：pc连接芯片，芯片连接N5172B信号发生器，注意连接s1和s2线。      \
                 芯片配置文件：\
                 CPC5801803_RX_test_2020_cg_2_filter_5_lindao_20200902.txt\
                 信号发生器配置：symbol rate 512ksps 、alpha 1.0、文件配置PN9-1FF-FM0、Total Bits 两百万或者一千万、realtiem设置为on、打开RF。    \
                 具体流程：    \
                 1）确保所有的数据线都已经连接好，给芯片配置好文件\
                 2）设置信号源配置\
                 3）设置好测试频点和功率点\
                 频点和功率设置可以手动设置，直接在软件表格中更改。或者选择已经准备好的excel文件，将频点读入软件表格。点击读取测试数据按钮可以读取数据。\
                 当手动输入频点时，可以根据实际需要使用添加和删除按钮操作软件表格的行与列。   \
                 4）点击按钮开始测试，然后在弹窗处选择提前准备好的excel文件\
                 Rx跳点测试会读取依次读取频点，在信号源中设置。然后测试每一个频点在各个功率值条件下的误码率。即5.8258GHz在-43、-52、-67dBm下的误码率，以此类推。    \
                 4、Rx带宽测试-单点      \
                 测试目的：该测试项目主要是固定功率值，测量设定频率范围下的误码率。       \
                 连接设置：pc连接芯片，芯片连接N5172B信号发生器，注意连接s1和s2线。    \
                 芯片配置文件：\
                 CPC5801803_RX_test_2020_cg_2_filter_5_lindao_20200902.txt   \
                 信号发生器配置：symbol rate 512ksps 、alpha 1.0、文件配置PN9-1FF-FM0、Total Bits 两百万或者一千万、realtiem设置为on、打开RF。   \
                 具体流程：      \
                 1）确保所有的数据线都已经连接好，给芯片配置好文件\
                 2）设置信号源配置\
                 此处主要设置Total Bits、误码率标准以及芯片号\
                 3）设置好功率值和带宽范围\
                 此处设置灵敏度，程序会在运行过程中自动加六，然后对信号源进行设置。还需要设置频率上限和频率下限，以及步长。这两个频率的差值就是所需要的测量的带宽范围。在频率和带宽设置过程中，需要注意单位，以确保填写数据的准确。\
                 4）点击按钮开始测试，然后在弹窗处选择提前准备好的excel文件     \
                 5、Rx带宽测试-多点\
                 测试目的：该测试的目的是给定带宽范围和功率范围，测量每一个频点在给定功率范围下的误码率。\
                 连接设置：pc连接芯片，芯片连接N5172B信号发生器，注意连接s1和s2线。      \
                 芯片配置文件：   \
                 CPC5801803_RX_test_2020_cg_2_filter_5_lindao_20200902.txt\
                 信号发生器配置：symbol rate 512ksps 、alpha 1.0、文件配置PN9-1FF-FM0、Total Bits 两百万或者一千万、realtiem设置为on、打开RF。     \
                 具体流程：\
                 1）确保所有的数据线都已经连接好，给芯片配置好文件\
                 2）设置信号源配置\
                 此处主要设置Total Bits、误码率标准以及芯片号\
                 3）设置好带宽范围，功率范围以及各自相应的步长\
                 这里填写的灵敏度会自动加六然后设置在信号发生器中。程序从频率下限开始依据每一步的步长扫描到频率表上限，在每一个频点都从灵敏度+6的范围依据步长扫描到灵敏度上限，记录每一个频点、功率值下的误码率。\
                 4）、点击按钮开始测试，然后在弹窗处选择提前准备好的excel文件         \
                 6、Rx最大功率测试\
                 测试目的：在接收灵敏度测量中，求的是下限。这里的程序设置可以求得接收灵敏度的上限。\
                 连接设置：pc连接芯片，芯片连接N5172B信号发生器，注意连接s1和s2线。       \
                 芯片配置文件：\
                 CPC5801803_RX_test_2020_cg_2_filter_5_lindao_20200902.txt\
                 信号发生器配置：symbol rate 512ksps 、alpha 1.0、文件配置PN9-1FF-FM0、Total Bits 两百万或者一千万、realtiem设置为on、打开RF。   \
                 具体流程：\
                 1）确保所有的数据线都已经连接好，给芯片配置好文件\
                 2）设置信号源配置\
                 此处主要设置Total Bits、误码率标准以及芯片号\
                 3）在这里设置功率下限，即接收灵敏度。设置功率上限，以防止功率设的太大。这里只能设置一个频率点。     \
                 在设置信号发生器参数时，程序会将这里的灵敏度加六。\
                 4）点击按钮开始测试，然后在弹窗处选择提前准备好的excel文件";
        ui->plainTextEdit->setPlainText(str);
     });
     connect(ui->btnWuEx,&QPushButton::clicked,[&]()
     {
         QString str="1、固定测试  \
                 测试目的：该测试利用观察14K方波的方式测量芯片的唤醒灵敏度，这个测试项目固定测试5.83GHz和5.84GHz两个通道的唤醒灵敏度。\
                 连接设置：pc连接芯片，芯片连接N5172B信号发生器，这里不需要使用s1和s2线。DPO7254示波器需要打开TeKVISA Line Server Control，然后用探针或者线连接芯片板的WU-OUT孔。示波器和PC通过网线连接。\
                 芯片配置文件：\
                 CPC5801803_WU_Normal_work_WKRX_TGAP0s3_BST_WU4ms-20200902_55dbm.txt      \
                 信号发生器配置：symbol rate 28ksps 、alpha 1.0、文件配置14K_TEST_28@BIT、realtiem设置为on、打开RF。\
                 具体流程： \
                 1）确保所有的数据线都已经连接好，给芯片配置好文件\
                 2）配置信号源\
                 3）配置示波器\
                 进行完这三步的设置后，在示波器上可以看到很明显的方波\
                 4）设置好程序要扫描的功率范围\
                 注:这里无需设置频率\
                 5）点击按钮开始测试，然后在弹窗处选择提前准备好的excel文件\
                 程序会自动从功率上界开始，依据步长扫描到功率下界，当示波器方波消失的时候，程序自动记录当前功率值。程序会自动扫描5.83和5.84GHz两个频点下功率范围，分别记录灵敏度。\
                 2、固定测试  \
                 测试目的：该测试利用观察14K方波的方式测量芯片的唤醒灵敏度。\
                 这个测试项目可以测试任意频率通道的唤醒灵敏度，单次只能测量一个通道。    \
                 连接设置：pc连接芯片，芯片连接N5172B信号发生器，这里不需要使用s1和s2线。DPO7254示波器需要打开TeKVISA Line Server Control，然后用探针或者线连接芯片板的WU-OUT孔。示波器和PC通过网线连接。\
                 芯片配置文件：\
                 CPC5801803_WU_Normal_work_WKRX_TGAP0s3_BST_WU4ms-20200902_55dbm.txt    \
                 信号发生器配置：symbol rate 28ksps 、alpha 1.0、文件配置14K_TEST_28@BIT、realtiem设置为on、打开RF。\
                 具体流程：  \
                 1）确保所有的数据线都已经连接好，给芯片配置好文件\
                 2）配置信号源\
                 3）配置示波器\
                 进行完这三步的设置后，在示波器上可以看到很明显的方波\
                 4）设置好程序要扫描的功率范围\
                 5）点击按钮开始测试，然后在弹窗处选择提前准备好的excel文件\
                 程序运行过程同固定测试，但是只能测量指定通道。\
                 3、唤醒最大功率测试\
                 测试目的：该测试利用观察14K方波的方式测量芯片的唤醒灵敏度。\
                 测量的是指定通道，唤醒的最大功率值。       \
                 连接设置：pc连接芯片，芯片连接N5172B信号发生器，这里不需要使用s1和s2线。DPO7254示波器需要打开TeKVISA Line Server Control，然后用探针或者线连接芯片板的WU-OUT孔。示波器和PC通过网线连接。\
                 芯片配置文件：\
                 CPC5801803_WU_Normal_work_WKRX_TGAP0s3_BST_WU4ms-20200902_55dbm.txt      \
                 信号发生器配置：symbol rate 28ksps 、alpha 1.0、文件配置14K_TEST_28@BIT、realtiem设置为on、打开RF\
                 具体流程：  \
                 1）确保所有的数据线都已经连接好，给芯片配置好文件\
                 2）配置信号源\
                 3）配置示波器\
                 进行完这三步的设置后，在示波器上可以看到很明显的方波\
                 4）设置好程序要扫描的功率范围\
                 5）点击按钮开始测试，然后在弹窗处选择提前准备好的excel文件\
                 程序会自动从功率下界开始依据步长扫描到功率上届，当无法检测到方波时，记录当前功率值。单次只能检测一个通道。\
                 4、其他  \
                 Excel行数：在所有唤醒灵敏度测试模式中，可以通过这里设置调整excel行数    \
                 芯片号：在所有唤醒灵敏度测试模式中，此处可以自动更新芯片号，也可以手动变更";

        ui->plainTextEdit->setPlainText(str);
     });

}

MainWindow::~MainWindow()
{
    delete ui;
}


