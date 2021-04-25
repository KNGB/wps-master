/*
** Copyright @ 2012-2019, Kingsoft office,All rights reserved.
** Redistribution and use in source and binary forms, with or without
** modification, are permitted provided that the following conditions are met:
**
** 1. Redistributions of source code must retain the above copyright notice,
**    this list of conditions and the following disclaimer.
** 2. Redistributions in binary form must reproduce the above copyright notice,
**    this list of conditions and the following disclaimer in the documentation
**    and/or other materials provided with the distribution.
** 3. Neither the name of the copyright holder nor the names of its
**    contributors may be used to endorse or promote products derived from this
**    software without specific prior written permission.
**
** THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
** AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
** IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE
** ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE
** LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR
** CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF
** SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS
** INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN
** CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE)
** ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE
** POSSIBILITY OF SUCH DAMAGE.
*/

#ifndef WPPMAINWINDOW_H
#define WPPMAINWINDOW_H
#include <QVBoxLayout>
#include <QPushButton>
#include <QHBoxLayout>
#include <QLineEdit>
#include <QtDebug>
#include "wppapi.h"
#include "kfc/comsptr.h"
#include "wpsapiex.h"
#if (QT_VERSION <= QT_VERSION_CHECK(5,0,0))
#include <QX11EmbedContainer>
#endif

#if (QT_VERSION < QT_VERSION_CHECK(5,0,0))
class WPPWindow : public QX11EmbedContainer
#else
class WPPWindow : public QWidget
#endif
{
	Q_OBJECT
public:
	WPPWindow(QWidget *parent = NULL);
	~WPPWindow();

	IKRpcClient * initWpsApplication();
	void destroyWpsApplication();
public slots:
	void onEmbeded();
	void addContainerWin(QWidget* containerWin);
private:
	IKRpcClient* m_rpcClient;
	QHBoxLayout *m_hlayout;
	QWidget* m_containerWin;
};

class WPPMainWindow : public QWidget
{
	Q_OBJECT
public:
	WPPMainWindow(QWidget *parent = NULL);
	~WPPMainWindow();

public:
	void initApp();
	void newDoc();
	void openDoc();
	void saveAs();
	void forceBackUpEnabled();
	void regiestPrintAfterEvent();
	void printOut();

	void slideSize();
	void addTable();
	void newSection();
	void addAudio();
	void hyperlink();
	void insertTextbox();
	//隐藏菜单按钮
	void hideToolBar();
	void hideToolBtn();
	//注册事件
	void registerCloseEvent();
	void registerSaveEvent();
	//其他
	void closeApp();
	void insertPicture();
	void insertSlide();
	void startPlay();
public slots:
	void slotButtonClick(const QString& name);
private:
	IKRpcClient * m_rpcClient;
	WPPWindow *m_mainArea;
	kfc::ks_stdptr<wppapi::_Application> m_spApplication;
	kfc::ks_stdptr<wppapi::Presentations> m_spDocs;
	QHBoxLayout *m_hlayout;
	kfc::ks_stdptr<wpsapiex::_ApplicationEx> m_spApplicationEx;
};
#endif // WPPMAINWINDOW_H
