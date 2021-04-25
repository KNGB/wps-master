QT += core gui network
greaterThan(QT_MAJOR_VERSION, 4): QT += widgets
QMAKE_CXXFLAGS += -std=c++0x -Wno-attributes
TARGET = wpsDemo
TEMPLATE = app

exists(/opt/kingsoft/wps-office/office6/libstdc++.so.6){
        system(ln -s /opt/kingsoft/wps-office/office6/libstdc++.so.6  libstdc++.so.6)
	LIBS += libstdc++.so.6
}

QMAKE_LFLAGS += -Wl,--rpath=\'\$\$ORIGIN\':$$[QT_INSTALL_LIBS]:/opt/kingsoft/wps-office/office6
QMAKE_LIBDIR =  ./ $$[QT_INSTALL_LIBS]  /opt/kingsoft/wps-office/office6

greaterThan(QT_MAJOR_VERSION, 4){
        LIBS += -lrpcwpsapi_sysqt5 -lrpcetapi_sysqt5 -lrpcwppapi_sysqt5
	exists(/opt/kingsoft/wps-office/office6/libc++abi.so.1){
		system(ln -sf /opt/kingsoft/wps-office/office6/libc++abi.so.1  libc++abi.so.1)
		LIBS += libc++abi.so.1
	}
}
else{
        LIBS += -lrpcwpsapi -lrpcetapi -lrpcwppapi
}


INCLUDEPATH = . \
                ./et \
                ./wps \
                ./wpp \
                ../include/common \
                ../include/wps \
                ../include/wpp \
                ../include/et
SOURCES += main.cpp\
		buttonlistwnd.cpp \
                et/etmainwindow.cpp \
                wpp/wppmainwindow.cpp \
                wps/wpsmainwindow.cpp \
                mainwindow.cpp

HEADERS  += wps/wpsmainwindow.h \
                et/etmainwindow.h \
                wpp/wppmainwindow.h \
                buttonlistwnd.h \
                mainwindow.h

FORMS += \
    mainwindow.ui


DISTFILES +=

RESOURCES += \
    translation.qrc
