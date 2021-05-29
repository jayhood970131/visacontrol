#ifndef VIOVECEOPERATE_H
#define VIOVECEOPERATE_H
#include <QAxObject.h>
#include "viopsignalgenerator.h"
#include <synchapi.h>
#include <QCoreApplication>
#include <QTime>
#include <QDialog>
#include <QFileDialog>                     //读取文件
#include <QtSerialPort/QtSerialPort>                    //串口通信
#include <QtSerialPort/QSerialPortInfo>
#include <QSerialPort>
#include <mainwindow.h>
#include <iostream>
#include <fstream>
#include <QFile>
#define byte unsigned char

//void exportToExcel();
QString decTobin (QString strDec);
int hex2(unsigned char ch);
QString hexToDec(QString strHex);
void readTxtData();
QString hexTobin(QString hex);
QString binTohex(QString bin);
int binToIntHex(QString s1,QString s2);
int hex2int(char c);
QString binTohex1(QString bin);
void WriteExcel( QAxObject *cellSample,QString data );
QAxObject* Findpath();



class MultiTypeNumber
   {
       /*
         构造函数-1
         通过一个byte类型参数，并将其转换成对应二进制与十六进制形式并存储。
        */
private:
        byte byteNumber = 0;
        QString byteNumber_Bin = "";
        QString byteNumber_Hex = "";
public:
    MultiTypeNumber(byte num)
     {
           this->byteNumber = num;
           //byte转成16进制显示
           QString str_H = "";
           QString str_L = "";

           int temp = 0;

           temp = num & 0xf0;  //高四位
           switch(temp)
           {
               case 0x00:
                   str_H = "0";
                   break;
               case 0x10:
                   str_H = "1";
                   break;
               case 0x20:
                   str_H = "2";
                   break;
               case 0x30:
                   str_H = "3";
                   break;
               case 0x40:
                   str_H = "4";
                   break;
               case 0x50:
                   str_H = "5";
                   break;
               case 0x60:
                   str_H = "6";
                   break;
               case 0x70:
                   str_H = "7";
                   break;
               case 0x80:
                   str_H = "8";
                   break;
               case 0x90:
                   str_H = "9";
                   break;
               case 0xa0:
                   str_H = "a";
                   break;
               case 0xb0:
                   str_H = "b";
                   break;
               case 0xc0:
                   str_H = "c";
                   break;
               case 0xd0:
                   str_H = "d";
                   break;
               case 0xe0:
                   str_H = "e";
                   break;
               case 0xf0:
                   str_H = "f";
                   break;
               default:
                   str_H = "err";
                   break;
          }

           temp = num & 0x0f;  //低四位
           switch (temp)
           {
               case 0x00:
                   str_L = "0";
                   break;
               case 0x01:
                   str_L = "1";
                   break;
               case 0x02:
                   str_L = "2";
                   break;
               case 0x03:
                   str_L = "3";
                   break;
               case 0x04:
                   str_L = "4";
                   break;
               case 0x05:
                   str_L = "5";
                   break;
               case 0x06:
                   str_L = "6";
                   break;
               case 0x07:
                   str_L = "7";
                   break;
               case 0x08:
                   str_L = "8";
                   break;
               case 0x09:
                   str_L = "9";
                   break;
               case 0x0a:
                   str_L = "a";
                   break;
               case 0x0b:
                   str_L = "b";
                   break;
               case 0x0c:
                   str_L = "c";
                   break;
               case 0x0d:
                   str_L = "d";
                   break;
               case 0x0e:
                   str_L = "e";
                   break;
               case 0x0f:
                   str_L = "f";
                   break;
               default:
                   str_L = "or";
                   break;
           }

           this->byteNumber_Hex = "0x" + str_H + str_L;
       }

       /*
        * 设置数据
        */
    void setByteNumber(byte num)
       {
           this->byteNumber = num;

           //byte转成16进制显示
           QString str_H = "";
           QString str_L = "";
           int temp = 0;
           temp = num & 0xf0;  //高四位
           switch (temp)
           {
               case 0x00:
                   str_H = "0";
                   break;
               case 0x10:
                   str_H = "1";
                   break;
               case 0x20:
                   str_H = "2";
                   break;
               case 0x30:
                   str_H = "3";
                   break;
               case 0x40:
                   str_H = "4";
                   break;
               case 0x50:
                   str_H = "5";
                   break;
               case 0x60:
                   str_H = "6";
                   break;
               case 0x70:
                   str_H = "7";
                   break;
               case 0x80:
                   str_H = "8";
                   break;
               case 0x90:
                   str_H = "9";
                   break;
               case 0xa0:
                   str_H = "a";
                   break;
               case 0xb0:
                   str_H = "b";
                   break;
               case 0xc0:
                   str_H = "c";
                   break;
               case 0xd0:
                   str_H = "d";
                   break;
               case 0xe0:
                   str_H = "e";
                   break;
               case 0xf0:
                   str_H = "f";
                   break;
               default:
                   str_H = "err";
                   break;
           }

           temp = num & 0x0f;  //低四位
           switch (temp)
           {
               case 0x00:
                   str_L = "0";
                   break;
               case 0x01:
                   str_L = "1";
                   break;
               case 0x02:
                   str_L = "2";
                   break;
               case 0x03:
                   str_L = "3";
                   break;
               case 0x04:
                   str_L = "4";
                   break;
               case 0x05:
                   str_L = "5";
                   break;
               case 0x06:
                   str_L = "6";
                   break;
               case 0x07:
                   str_L = "7";
                   break;
               case 0x08:
                   str_L = "8";
                   break;
               case 0x09:
                   str_L = "9";
                   break;
               case 0x0a:
                   str_L = "a";
                   break;
               case 0x0b:
                   str_L = "b";
                   break;
               case 0x0c:
                   str_L = "c";
                   break;
               case 0x0d:
                   str_L = "d";
                   break;
               case 0x0e:
                   str_L = "e";
                   break;
               case 0x0f:
                   str_L = "f";
                   break;
               default:
                   str_L = "or";
                   break;
           }

           this->byteNumber_Hex = "0x" + str_H + str_L;
       }

       /*
        * 返回十六进制形式显示的数据
        */
        QString getByteNumber_Hex()
       {
           return this->byteNumber_Hex;
       }
};


#endif // VIOVECEOPERATE_H
