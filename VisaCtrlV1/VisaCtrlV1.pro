#-------------------------------------------------
#
# Project created by QtCreator 2019-10-22T10:53:59
#
#-------------------------------------------------

QT       += core gui

greaterThan(QT_MAJOR_VERSION, 4): QT += widgets
TARGET = VisaCtrlV1
TEMPLATE = app

SOURCES += main.cpp\
        mainwindow.cpp \
    viodpo7254.cpp \
    vioperation.cpp \
    viophead.cpp \
    viopsignalgenerator.cpp \
    viop33250a.cpp \
    viopn5172b.cpp \
    vioneceoperate.cpp \
    viopr3408b.cpp

HEADERS  += mainwindow.h \
    viodpo7254.h \
    vioperation.h \
    vipara.h \
    viparaunit.h \
    viophead.h \
    viopsignalgenerator.h \
    viop33250a.h \
    viopn5172b.h \
    vioneceoperate.h \
    viopr3408b.h \
    IBookT.h \
    libxl.h

FORMS    += mainwindow.ui
QT       += axcontainer
QT       += serialport

#LIBS += -LC:\Users\zhang\Desktop\VISA20-9-23\code\VisaCtrlV1\lib -lktvisa32
LIBS += -LD:\Codes\QT\VisaCtrlV1\lib -ltKVisa64       #laptop
LIBS += -LD:\Codes\QT\VisaCtrlV1\lib -llibxl
