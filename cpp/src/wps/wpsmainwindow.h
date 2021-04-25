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

#ifndef WPSMAINWINDOW_H
#define WPSMAINWINDOW_H
#include <QWidget>
#include "wpsapi.h"
#include "kfc/comsptr.h"
#include "wpsapiex.h"
#include <QPushButton>
#include <QHBoxLayout>
#include <QEvent>
#include <QMap>
#include <QFileDialog>
#if (QT_VERSION <= QT_VERSION_CHECK(5,0,0))
#include <QX11EmbedContainer>
#endif

class ButtonListWnd;
#if (QT_VERSION < QT_VERSION_CHECK(5,0,0))
class WPSWindow : public QX11EmbedContainer
#else
class WPSWindow : public QWidget
#endif
{
	Q_OBJECT
public:
	WPSWindow(QWidget *parent = NULL);
	~WPSWindow();
	IKRpcClient * initWpsApplication();
	void destroyWpsApplication();
	void addHotKey(const QString& hotKey, bool isEnable = false);
	void removeHotKey(const QString& hotKey);
protected:
	bool eventFilter(QObject *Object, QEvent *Event);
public slots:
	void onEmbeded();
	void addContainerWin(QWidget* containerWin);
private:
	IKRpcClient* m_rpcClient;
	QMap<QString, bool> m_enableHotKeysMap;
	QHBoxLayout* m_hlayout;
	QWidget* m_containerWin;
};

class WPSMainWindow : public QWidget
{
	Q_OBJECT
public:
	WPSMainWindow(QWidget *parent = NULL);
	~WPSMainWindow();
	void closeApp();
	void initApp();
	void newDoc();
	void openDoc();
	void saveAs();
	void closeDoc();
	void printOutDoc();
	void forceBackUpEnabled();
	void setHeaderFooter();
	void getSaved();
	void getPagesNum();
	void disableHotKey();
	void enableProtect();
	void disableProtect();

	void showAllTool();
	void hideToolBar();
	void hideToolBtn();

	void listTemplates();
	void hyperlinks();
	void highlight();
	void inserttable();
	void setFontSize();
	void insertEllipse();

	//注册事件
	void registerCloseEvent();
	void registerSaveEvent();

	void slotButtonClick(const QString& name);
private:
	IKRpcClient* m_rpcClient;
	WPSWindow* m_mainArea;
	kfc::ks_stdptr<wpsapi::_Application> m_spApplication;
	kfc::ks_stdptr<wpsapiex::_ApplicationEx> m_spApplicationEx;
	kfc::ks_stdptr<wpsapi::Documents> m_spDocs;
	QHBoxLayout* m_hlayout;
};
#endif // WPSMAINWINDOW_H
