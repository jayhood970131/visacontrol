#include "vioneceoperate.h"

QString decTobin (QString strDec)
{  //十进制转二进制
    int nDec = strDec.toInt();
    int nYushu, nShang;
    QString strBin, strTemp;
    //TCHAR buf[2];
    bool bContinue = true;
    while ( bContinue )
    {
        nYushu = nDec % 2;
        nShang = nDec / 2;
        strBin=QString::number(nYushu)+strBin;       //qDebug()<<strBin;
        strTemp = strBin;
        //strBin.Format("%s%s", buf, strTemp);
        nDec = nShang;
        if ( nShang == 0 )
            bContinue = false;
    }
    int nTemp = strBin.length() % 4;
    switch(nTemp)
    {
    case 1:
        //strTemp.Format(_T("000%s"), strBin);
        strTemp = "000"+strBin;
        strBin = strTemp;
        break;
    case 2:
        //strTemp.Format(_T("00%s"), strBin);
        strTemp = "00"+strBin;
        strBin = strTemp;
        break;
    case 3:
        //strTemp.Format(_T("0%s"), strBin);
        strTemp = "0"+strBin;
        strBin = strTemp;
        break;
    default:
        break;
    }
    return strBin;
}


int hex2(unsigned char ch){           //十六进制转换工具
    if((ch >= '0') && (ch <= '9'))
        return ch-0x30;
    else if((ch >= 'A') && (ch <= 'F'))
        return ch-'A'+10;
    else if((ch >= 'a') && (ch <= 'f'))
        return ch-'a'+10;
    return 0;
}


QString hexToDec(QString strHex)         //十六进制转十进制
{
    int i;
    int v = 0;
    for(i=0;i<strHex.length();i++)
    {
        v*=16;
        v+=hex2(strHex[i].toLatin1());
    }
    return QString::number(v);
}

//读取文本文档
void readTxtData( QSerialPort serial)
{
    QString filepath;
  // QString displayString;
//    QString SbinData;
//    char *binData;
    filepath =QFileDialog::getOpenFileName(nullptr,("选择文件"),("C:/Users/SZUIoT-2/Desktop"),("*.txt")); //获取保存路径
  //  filepath =QFileDialog::getOpenFileName(NULL,("选择文件"),(filepath),("*.txt")); //获取保存路径

    QFile file(filepath);                                                                                         //打开txt

    if(!file.open(QIODevice::ReadOnly | QIODevice::Text))                                                         //只读 文本文档
    {
       // qDebug()<<"Can't open the file!"<<endl;
       QMessageBox::information(nullptr,"警告","无法打开文件");
    }

    while(!file.atEnd())                                                                                          //没有读到最后一行
    {

        QByteArray line = file.readLine();
        QString str(line);
        QByteArray addr;
        QByteArray data;

        while(str!=nullptr)
        {
            QStringList s=str.split(" ");
            QString a=s[0];
            QString b=s[1];
            QString c=s[2];
            QString binary = decTobin(hexToDec(a));       //十六进制转成二进制
            b.append(c);
            binary.append(b);
            addr=binary.toLatin1();
            data=b.toLatin1();
            addr.append(data);
            serial.write(addr);
        }  
    }
}

QString hexTobin(QString hex)
{
    int len =hex.length();
    QString bin;
    for(int i=2;i<len;i++)
    {
       if(hex[i]=='0')bin.append("0000");
       else if(hex[i]=='1')bin.append("0001");
       else if(hex[i]=='2')bin.append("0010");
       else if(hex[i]=='3')bin.append("0011");
       else if(hex[i]=='4')bin.append("0100");
       else if(hex[i]=='5')bin.append("0101");
       else if(hex[i]=='6')bin.append("0110");
       else if(hex[i]=='7')bin.append("0111");
       else if(hex[i]=='8')bin.append("1000");
       else if(hex[i]=='9')bin.append("1001");
       else if(hex[i]=='a')bin.append("1010");
       else if(hex[i]=='b')bin.append("1011");
       else if(hex[i]=='c')bin.append("1100");
       else if(hex[i]=='d')bin.append("1101");
       else if(hex[i]=='e')bin.append("1110");
       else bin.append("1111");
    }
    return bin;
}
//二进制转成十六进制
QString binTohex(QString bin)
{
    QString hex;
//    if(bin=="0000")hex="0";
//    else if(bin=="0001")hex="1";
//    else if(bin=="0010")hex="2";
//    else if(bin=="0011")hex="3";
//    else if(bin=="0100")hex="4";
//    else if(bin=="0101")hex="5";
//    else if(bin=="0110")hex="6";
//    else if(bin=="0111")hex="7";
//    else if(bin=="1000")hex="8";
//    else if(bin=="1001")hex="9";
//    else if(bin=="1010")hex="a";
//    else if(bin=="1011")hex="b";
//    else if(bin=="1100")hex="c";
//    else if(bin=="1101")hex="d";
//    else if(bin=="1110")hex="e";
//    else hex="f";
//    return hex;

        if(bin=="1111")
        {
            hex="F";
        }
        if(bin=="1110")
        {
            hex="E";
        }
        if(bin=="1101")
        {
           hex="D";
        }
        if(bin=="1100")
        {
            hex="C";
        }
        if(bin=="1011")
        {
           hex="B";
        }
        if(bin=="1010")
        {
            hex="A";
        }
        if(bin=="1001")
        {
            hex="9";
        }
        if(bin=="1000")
        {
            hex="8";
        }
        if(bin=="0111")
        {
            hex="7";
        }
        if(bin=="0100")
        {
           hex="6";
        }
        if(bin=="0101")
        {
           hex="5";
        }
        if(bin=="0100")
        {
           hex="4";
        }
        if(bin=="0011")
        {
           hex="3";
        }
        if(bin=="0010")
        {
            hex="2";
        }
        if(bin=="0001")
        {
            hex="1";
        }
        if(bin=="0000")
        {
            hex="0";
        }
        return hex;
}


QString binTohex1(QString bin)
{
     QString hex;
    if(bin=="1111\n")
    {
        hex="F";
    }
    if(bin=="1110\n")
    {
        hex="E";
    }
    if(bin=="1101\n")
    {
       hex="D";
    }
    if(bin=="1100\n")
    {
        hex="C";
    }
    if(bin=="1011\n")
    {
       hex="B";
    }
    if(bin=="1010\n")
    {
        hex="A";
    }
    if(bin=="1001\n")
    {
        hex="9";
    }
    if(bin=="1000\n")
    {
        hex="8";
    }
    if(bin=="0111\n")
    {
        hex="7";
    }
    if(bin=="0100\n")
    {
       hex="6";
    }
    if(bin=="0101\n")
    {
       hex="5";
    }
    if(bin=="0100")
    {
       hex="4";
    }
    if(bin=="0011\n")
    {
       hex="3";
    }
    if(bin=="0010\n")
    {
        hex="2";
    }
    if(bin=="0000\n")
    {
        hex="1";
    }
    if(bin=="0000\n")
    {
        hex="0";
    }
    return hex;
}


int hex2int(char c)
{
 if ((c >= 'A') && (c <= 'Z'))
 {
 return c - 'A' + 16;
 }
 else if ((c >= 'a') && (c <= 'z'))
 {
 return c - 'a' + 16;
 }
 else if ((c >= '0') && (c <= '9'))
 {
 return c - '0';
 }
}

//bin to int(hex)
//int binToIntHex(QString s1,QString s2)
/*{
    int s3,s4;
    int s;
    int bintohex;
    int m;
    if(s1=="0000")s3='0';
    else if(s1=="0001")s3=1;
    else if(s1=="0010")s3=2;
    else if(s1=="0011")s3=3;
    else if(s1=="0100")s3=4;
    else if(s1=="0101")s3=5;
    else if(s1=="0110")s3=6;
    else if(s1=="0111")s3=7;
    else if(s1=="1000")s3=8;
    else if(s1=="1001")s3=9;
    else if(s1=="1010")s3=10;
    else if(s1=="1011")s3=11;
    else if(s1=="1100")s3=12;
    else if(s1=="1101")s3=13;
    else if(s1=="1110")s3=14;
    else s3=15;
    if(s2=="0000")s4=0;
    else if(s2=="0001")s4=1;
    else if(s2=="0010")s4=2;
    else if(s2=="0011")s4=3;
    else if(s2=="0100")s4=4;
    else if(s2=="0101")s4=5;
    else if(s2=="0110")s4=6;
    else if(s2=="0111")s4=7;
    else if(s2=="1000")s4=8;
    else if(s2=="1001")s4=9;
    else if(s2=="1010")s4=10;
    else if(s2=="1011")s4=11;
    else if(s2=="1100")s4=12;
    else if(s2=="1101")s4=13;
    else if(s2=="1110")s4=14;
    else s4=15;
    s=s3*10+s4;
    bintohex=s&0x00ff;
    int ms=0bintohex);
    return 0;
//}
*/
QAxObject* Findpath()
{
    QString filepath;
    QAxObject *excel = new QAxObject();
    excel->setControl("ket.Application");                                     //连接WPS
    excel->dynamicCall("SetVisible(bool Visible)","false");                   //不显示窗体
    excel->setProperty("DisplayAlerts", false);                               //不显示任何警告信息。如果为true那么在关闭是会出现类似“文件已修改，是否保存”的提示
    QAxObject* workbooks = excel->querySubObject("WorkBooks");                //获取工作簿集合
    workbooks->dynamicCall("Open(const QString&)",filepath);                  //打开打开已存在的工作簿
    QAxObject *workbook = excel->querySubObject("ActiveWorkBook");            //获取当前工作簿
    QAxObject *worksheets = workbook->querySubObject("WorkSheets");           //获取工作表集合
    QAxObject *worksheet = worksheets->querySubObject("Item(int)",1);         //获取工作表集合的工作表1，即sheet
    QAxObject *cellA = worksheet->querySubObject("Cells(int,int)",1,1);
    return cellA;
 //   return filepath;
}

void WriteExcel( QAxObject *cellSample,QString data )
{
    cellSample->dynamicCall("SetValue(const QVariant&)",data);
    cellSample->setProperty("HorizontalAlignment", -4108);             //居中
    cellSample->setProperty("VerticalAlignment", -4108);               //居中

}

























































