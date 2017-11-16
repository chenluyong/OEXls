#-------------------------------------------------
#
# Project created by QtCreator 2017-11-02T16:23:07
#
#-------------------------------------------------

QT       -= gui

DEFINES += QT_IDE

DESTDIR = $$PWD/../state/lib/
TARGET = auditxls
Release:TEMPLATE = lib
Debug:TEMPLATE = app

INCLUDEPATH += $$PWD/../state/include/
INCLUDEPATH += $$PWD/../state/include/libxl/
INCLUDEPATH += $$PWD/OE/
INCLUDEPATH += $$PWD/xlslib/

LIBS += -L$$PWD/../state/lib/libxl -llibxl

DEFINES += AUDITXLS_LIBRARY

include(xlslib/xlslib.pri)

SOURCES += \
    main.cpp \
    oeexcelwriter.cpp \
    oeexcelreader.cpp

HEADERS +=\
        auditxls_global.h \
    oeexcelwriter.h \
    oeexcelreader.h

unix {
    target.path = /usr/lib
    INSTALLS += target
}

DISTFILES += \
    xlslib/xlslib/formula.txt

# 版本信息
VERSION = 1.0.0.0

## 图标
#RC_ICONS = Images/MyApp.ico

# 公司名称
QMAKE_TARGET_COMPANY = "eshanren for WenZhou"

# 产品名称
QMAKE_TARGET_PRODUCT = "excel base"

# 文件说明
QMAKE_TARGET_DESCRIPTION = "excel base for eshanren(cross platform)"

# 版权信息
QMAKE_TARGET_COPYRIGHT = "Copyright 2015-2018 The eshanren Company Ltd. All rights reserved."

# 中文（简体）
RC_LANG = 0x0004
