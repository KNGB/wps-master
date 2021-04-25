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
#include "mainwindow.h"
#include "ui_mainwindow.h"
#include "wpsmainwindow.h"
#include "etmainwindow.h"
#include "wppmainwindow.h"
#include <QPaintEvent>
#include <QStylePainter>

MainWindow::MainWindow(QWidget *parent)
	: QWidget(parent)
	, m_curIndex((int)appWps)
	, m_ui(new Ui::MainWindow)
{
	m_ui->setupUi(this);
	m_wps = new WPSMainWindow(this);
	m_et = new EtMainWindow(this);
	m_wpp = new WPPMainWindow(this);
	m_ui->tabWidget->addTab(m_wps, "WPS");
	m_ui->tabWidget->addTab(m_et, "ET");
	m_ui->tabWidget->addTab(m_wpp, "WPP");
	m_ui->treeWidget->initButtonTree(m_curIndex);
	setWindowTitle("WPS C++ DEMO");
	connect(m_ui->treeWidget, SIGNAL(itemClicked(QTreeWidgetItem*, int)), this, SLOT(slotItemClicked(QTreeWidgetItem*, int)));
	connect(m_ui->tabWidget, SIGNAL(currentChanged(int)), this, SLOT(slotTabCurrentChanged(int)));

}

MainWindow::~MainWindow()
{
	disconnect(m_ui->treeWidget, SIGNAL(itemClicked(QTreeWidgetItem*, int)), this, SLOT(slotItemClicked(QTreeWidgetItem*, int)));
	disconnect(m_ui->tabWidget, SIGNAL(currentChanged(int)), this, SLOT(slotTabCurrentChanged(int)));

	m_wps->deleteLater();
	m_et->deleteLater();
	m_wpp->deleteLater();
	delete m_ui;
}

void MainWindow::slotTabCurrentChanged(int index)
{
	if (m_ui->treeWidget)
	{
		m_ui->treeWidget->initButtonTree(index);
		m_curIndex = index;
	}
}

void MainWindow::slotItemClicked(QTreeWidgetItem *item, int column)
{
	QString name = item->text(column);
	if (m_curIndex == appWps)
		m_wps->slotButtonClick(name);
	else if (m_curIndex == appEt)
		m_et->slotButtonClick(name);
	else if (m_curIndex == appWpp)
		m_wpp->slotButtonClick(name);
}
