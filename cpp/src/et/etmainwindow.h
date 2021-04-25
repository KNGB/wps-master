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

#ifndef ETMAINWINDOW_H
#define ETMAINWINDOW_H
#include <QPushButton>
#include <QHBoxLayout>
#include <QLineEdit>
#include "etapi.h"
#include "wpsapiex.h"
#include "kfc/comsptr.h"
#if (QT_VERSION <= QT_VERSION_CHECK(5,0,0))
#include <QX11EmbedContainer>
#endif

#define ET_LCID ((long)2052)

class ButtonListWnd;
#if (QT_VERSION < QT_VERSION_CHECK(5,0,0))
class ETWindow : public QX11EmbedContainer
#else
class ETWindow : public QWidget
#endif
{
	Q_OBJECT
public:
	ETWindow(QWidget *parent = NULL);
	~ETWindow();
	IKRpcClient * initWpsApplication();
	void destroyWpsApplication();
	void addHotKey(const QString& hotKey, bool isEnable = false);
	void removeHotKey(const QString& hotKey);
public slots:
	void onEmbeded();
	void addContainerWin(QWidget* containerWin);

protected:
	bool eventFilter(QObject *Object, QEvent *Event);
private:
	IKRpcClient* m_rpcClient;
	QMap<QString, bool> m_enableHotKeysMap;
	QHBoxLayout* m_hlayout;
	QWidget* m_containerWin;
};

class EtMainWindow : public QWidget
{
	Q_OBJECT
public:
	EtMainWindow(QWidget *parent = NULL);
	~EtMainWindow();
public:
	void closeApp();
	void initApp();
	void newDoc();
	void openDoc();
	void saveAs();
	void saveDoc();
	void forceBackUpEnabled();
	void hideToolBar();
	void setHeaderFooter();
	void getSaved();
	void getPagesNum();
	void disableSaveAsHotKey();
	void setOrientation();
	void exportPdf();
	void protect();
	void unProtect();
	void hideToolBtn();

	void FontSize();
	void AddSheet();
	void AddComment();
	void PrintArea();
	void InsertPic();
	void InsertHistogram();

	//单元格设值
	void setRangeValue();
	void rangeMerge();
	void setHeigthWidth();
	//注册事件
	void registerCloseEvent();
	void registerSaveEvent();
public slots:
	void slotButtonClick(const QString& name);
private:
	IKRpcClient* m_rpcClient;
	ETWindow* m_mainArea;
	kfc::ks_stdptr<etapi::_Application> m_spApplication;
	kfc::ks_stdptr<wpsapiex::_ApplicationEx> m_spApplicationEx;
	kfc::ks_stdptr<etapi::Workbooks> m_spWorkbooks;
	QHBoxLayout* m_hlayout;
};
#endif // ETMAINWINDOW_H
