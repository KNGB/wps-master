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
#include "wppmainwindow.h"
#include "mainwindow.h"
#include "variant.h"
#include <QDebug>
#include <QTextCodec>
#include <QMessageBox>
#if (QT_VERSION >= QT_VERSION_CHECK(5,0,0))
#include <QWindow>
#endif

using namespace wppapi;
using namespace kfc;
using namespace wpsapiex;

WPPWindow::WPPWindow(QWidget *parent)
#if (QT_VERSION < QT_VERSION_CHECK(5,0,0))
	: QX11EmbedContainer(parent)
	, m_containerWin(NULL)
	, m_hlayout(NULL)
{
	connect(this, SIGNAL(clientIsEmbedded()), this, SLOT(onEmbeded()));
}
#else
	: QWidget(parent),
	m_containerWin(NULL)
{
	m_hlayout  = new QHBoxLayout;
	m_hlayout->setContentsMargins(1, 1, 1, 1);
	setLayout(m_hlayout);
}
#endif

WPPWindow::~WPPWindow()
{
	if (m_containerWin)
	{
		delete m_containerWin;
		m_containerWin = NULL;
	}
	if (m_hlayout)
	{
		delete m_hlayout;
		m_hlayout = NULL;
	}
}

IKRpcClient * WPPWindow::initWpsApplication()
{
	IKRpcClient *pWpsRpcClient = NULL;
	HRESULT hr = createWppRpcInstance(&pWpsRpcClient);
	if (hr != S_OK)
	{
		qDebug() <<"get WpsRpcClient erro";
		return NULL;
	}
	BSTR StrWpsAddress = SysAllocString(__X("/opt/kingsoft/wps-office/office6/wpp"));
	pWpsRpcClient->setProcessPath(StrWpsAddress);
	SysFreeString(StrWpsAddress);

	QVector<BSTR> vArgs;
	vArgs.resize(6);
	vArgs[0] =  SysAllocString(__X("-shield"));
	vArgs[1] =  SysAllocString(__X("-multiply"));
	vArgs[2] =  SysAllocString(__X("-x11embed"));
	vArgs[3] =  SysAllocString(QString::number(winId()).utf16());
	vArgs[4] =  SysAllocString(QString::number(width()).utf16());
	vArgs[5] =  SysAllocString(QString::number(height()).utf16());
	pWpsRpcClient->setProcessArgs(vArgs.size(), vArgs.data());
	for (int i = 0; i < vArgs.size(); i++)
	{
		SysFreeString(vArgs.at(i));
	}
	m_rpcClient = pWpsRpcClient;
	return pWpsRpcClient;
}

void WPPWindow::destroyWpsApplication()
{
	m_rpcClient = NULL;
}

void WPPWindow::onEmbeded()
{
#if (QT_VERSION < QT_VERSION_CHECK(5,0,0))
	if (m_rpcClient)
		m_rpcClient->setWpsWid(this->clientWinId());
#endif
}
void WPPWindow::addContainerWin(QWidget *containerWin)
{
	if (m_containerWin)
	{
		delete m_containerWin;
		m_containerWin = NULL;
	}
	m_containerWin = containerWin;
	m_hlayout->addWidget(m_containerWin);
}

WPPMainWindow::WPPMainWindow(QWidget *parent)
	: QWidget(parent)
	, m_mainArea(NULL)
{
	m_mainArea = new WPPWindow(this);
	m_hlayout  = new QHBoxLayout(this);
	m_hlayout->setContentsMargins(1, 1, 1, 1);
	m_hlayout->addWidget(m_mainArea);
	m_hlayout->setStretchFactor(m_mainArea,140);
	setLayout(m_hlayout);
	adjustSize();
}

WPPMainWindow::~WPPMainWindow()
{
	if (m_spApplication)
	{
		m_spApplication->Quit();
		m_spApplication.clear();
	}
	m_mainArea->destroyWpsApplication();
	m_rpcClient = NULL;
	m_mainArea->deleteLater();
	m_hlayout->deleteLater();
}

static HRESULT wppPrintOutPageSetCallback(_Presentation* pPres, PrintoutPageEx* pageEx)
{
	ks_bstr sPages;
	int iRange = 0;
	HRESULT hr = pageEx->get_PrintOutPages(&sPages);
	if (FAILED(hr))
		return S_OK;
	hr = pageEx->get_PrintOutRange(&iRange);
	if (FAILED(hr))
		return S_OK;

	if (iRange == 1)
		qDebug() << QString::fromUtf8("全选 ") << QString::fromUtf16((BSTR)sPages);
	else if (iRange == 2)
		qDebug() << QString::fromUtf8("选中 ") << QString::fromUtf16((BSTR)sPages);
	else if (iRange == 3)
		qDebug() <<  QString::fromUtf8("当前 ") << QString::fromUtf16((BSTR)sPages);
	else
		qDebug() <<  QString::fromUtf8("选定范围 ") << QString::fromUtf16((BSTR)sPages);

	return S_OK;
}

void WPPMainWindow::closeApp()
{
	if (m_spApplication != NULL)
	{
		m_spApplication->Quit();
		m_spDocs.clear();
		m_spApplicationEx.clear();
		m_spApplication.clear();
		m_mainArea->destroyWpsApplication();
	}
}

void WPPMainWindow::initApp()
{
	if (!m_spApplication)
	{
		IKRpcClient *pRpcClient = m_mainArea->initWpsApplication();
		m_rpcClient = pRpcClient;
		if (pRpcClient && pRpcClient->getWppApplication((IUnknown**)&m_spApplication) == S_OK)
		{
			m_spApplication->get_Presentations(&m_spDocs);
			m_spApplication->QueryInterface(IID__WppApplicationEx, (void**)&m_spApplicationEx);
#if (QT_VERSION >= QT_VERSION_CHECK(5,0,0))
			if (m_spApplicationEx)
			{
				unsigned long wid = 0;
				m_spApplicationEx->get_EmbedWid(&wid);
				QWidget* container = QWidget::createWindowContainer(QWindow::fromWinId(wid), this);
				m_mainArea->addContainerWin(container);
			}
#endif
		}
	}
}

void WPPMainWindow::newDoc()
{
	if (m_spDocs)
	{
		ks_stdptr<_Presentation> spPresentation;
		wppapi::MsoTriState withWindow = msoCTrue;
		HRESULT hr = m_spDocs->Add(withWindow, (Presentation**)&spPresentation);
		if (SUCCEEDED(hr) && spPresentation != NULL)
		{
			ks_stdptr<Slides> spSlides;
			hr = spPresentation->get_Slides(&spSlides);
			if (SUCCEEDED(hr) && spSlides != NULL)
			{
				ks_stdptr<_Slide> spSlide;
				PpSlideLayout Layout = ppLayoutTitle;
				hr = spSlides->Add(1, Layout, (Slide**)&spSlide);
				if (SUCCEEDED(hr))
					qDebug() << "New ok";
			}
		}
		else
			qDebug() << "New fail";
	}
	m_mainArea->setFocus();
}

void WPPMainWindow::openDoc()
{
	if (!m_spDocs)
	{
		qDebug() << "open fail";
		return;
	}
	QString openFile = QFileDialog::getOpenFileName((QWidget*)parent(), QString::fromUtf8("打开文档"));
	ks_bstr filename(openFile.utf16());
	wppapi::MsoTriState readOnly = msoFalse;
	wppapi::MsoTriState untitled = msoFalse;
	wppapi::MsoTriState withWindow = msoTrue;
	ks_stdptr<_Presentation> spPresentation;
	HRESULT hr = m_spDocs->Open(filename, readOnly, untitled, withWindow, (Presentation**)&spPresentation);
	if (SUCCEEDED(hr))
		qDebug() << "open ok";
	else
		qDebug() << "open fail";

	m_mainArea->setFocus();
}

void WPPMainWindow::saveAs()
{
	if (!m_spApplication)
	{
		qDebug() << "open fail";
		return;
	}
	ks_stdptr<_Presentation> spDoc;
	m_spApplication->get_ActivePresentation((Presentation**)&spDoc);
	if (!spDoc)
	{
		qDebug() << "Save erro";
		return;
	}
	QString saveFile = QFileDialog::getSaveFileName((QWidget*)parent(), QString::fromUtf8("保存文档"));
	if (!saveFile.isEmpty())
	{
		ks_bstr filename(saveFile.utf16());
		PpSaveAsFileType FileFormat = ppSaveAsPresentation;
		wppapi::MsoTriState EmbedTrueTypeFonts = msoTriStateMixed;
		HRESULT hr = spDoc->SaveAs(filename, FileFormat, EmbedTrueTypeFonts);
		if (SUCCEEDED(hr))
			qDebug() << "Save  ok";
		else
			qDebug() << "Save  fail";
	}

	m_mainArea->setFocus();
}

void WPPMainWindow::forceBackUpEnabled()
{
	qDebug() << "ForceBackUpEnabledSlot";
	static bool enable = true;
	enable = !enable;
	if (m_spApplicationEx)
		m_spApplicationEx->put_ForceBackupEnable(enable ? VARIANT_TRUE : VARIANT_FALSE);
	m_mainArea->setFocus();
}

void WPPMainWindow::regiestPrintAfterEvent()
{
	if (m_spApplicationEx)
	{
		HRESULT hr = m_rpcClient->registerEvent(m_spApplicationEx, DIID_ApplicationEventsEx, __X("DocumentAfterPrint"), (PVOID)wppPrintOutPageSetCallback);
		if (hr == S_OK)
		{
			qDebug() << "registerEvent succes" ;
		}
	}
	m_mainArea->setFocus();
}

void WPPMainWindow::printOut()
{
	if (m_spApplication)
	{
		ks_stdptr<_Presentation> spPresentation;
		m_spApplication->get_ActivePresentation((Presentation**)&spPresentation);
		if (!spPresentation)
			return;

		ks_stdptr<PrintOptions> spPrintOptions;
		spPresentation->get_PrintOptions(&spPrintOptions);
		if (spPrintOptions)
		{
			HRESULT hr = spPrintOptions->put_OutputType(ppPrintOutputSlides);
			if (SUCCEEDED(hr))
			{
				QString filePath = QFileDialog::getSaveFileName((QWidget*)parent(), QString::fromUtf8("保存路径"));
				ks_bstr fileName(filePath.utf16());
				HRESULT hr = spPresentation->PrintOut(-1, -1, fileName);
				if (SUCCEEDED(hr))
					qDebug() << "PrintOut succes" ;
			}
		}
	}
	m_mainArea->setFocus();
}

void WPPMainWindow::hideToolBar()
{
	/*
	"Standard"							常用
	"Formatting"						格式
	"Drawing"							绘图
	"Symbol"							符号栏
	"3-D Settings"						三维设置
	"Shadow Settings"					阴影设置
	"Reviewing"							审阅
	*/
	static bool enable = false;
	ks_stdptr<ksoapi::_CommandBars> spCommandBars;
	if (m_spApplication)
	{
		m_spApplication->get_CommandBars((ksoapi::CommandBars**)&spCommandBars);
		if (spCommandBars)
		{
			ks_stdptr<ksoapi::CommandBar> spCommandBar;
			KComVariant varIndex(__X("Standard"));
			spCommandBars->get_Item(varIndex, &spCommandBar);
			if (spCommandBar)
			{
				spCommandBar->put_Visible(enable ? VARIANT_TRUE : VARIANT_FALSE);
				enable = !enable;
			}
		}
	}
	m_mainArea->setFocus();
}

void WPPMainWindow::hideToolBtn()
{
	if (!m_spApplication)
		return;

	static bool enable = false;
	HRESULT ret = E_FAIL;
	ks_stdptr<ksoapi::_CommandBars> spCommandBars;
	ret = m_spApplication->get_CommandBars((ksoapi::CommandBars**)&spCommandBars);
	if (SUCCEEDED(ret) && spCommandBars)
	{
		KComVariant vCommandBarIndex(__X("Standard"));
		ks_stdptr<ksoapi::CommandBar> spCommandBar;
		ret = spCommandBars->get_Item(vCommandBarIndex, &spCommandBar);
		if (SUCCEEDED(ret) && spCommandBar)
		{
			ks_stdptr<ksoapi::CommandBarControls> spControls;
			spCommandBar->get_Controls(&spControls);
			int count = 0;
			ret = spControls->get_Count(&count);
			if (FAILED(ret))
				return;

			for (int i = 1; i <= count; i++)
			{
				KComVariant vControlIndex(i);
				ks_stdptr<ksoapi::CommandBarControl> spControl;
				ret = spControls->get_Item(vControlIndex, &spControl);
				if (SUCCEEDED(ret) && spControl)
				{
					spControl->put_Visible(enable ? VARIANT_TRUE : VARIANT_FALSE);
					//button对应的index和name，用户可根据index隐藏自己所需按钮
					//ks_bstr ksName;
					//spControl->get_Caption(&ksName);
					//qDebug() << "index, name: " << i << QString::fromUtf16(ksName);
				}
			}
			enable = !enable;
		}
	}
	m_mainArea->setFocus();
}

static HRESULT wppCloseCallback(_Presentation* pdoc)
{
	QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("关闭回调"));
	message.exec();
	return S_OK;
}

void WPPMainWindow::registerCloseEvent()
{
	if (m_spApplication)
	{
		HRESULT ret = m_rpcClient->registerEvent(m_spApplication, IID_EApplication, __X("PresentationClose"), (void*)wppCloseCallback);
		QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), "");
		if (ret == S_OK)
			message.setText(QString::fromUtf8("注册成功"));
		else
			message.setText(QString::fromUtf8("注册失败"));
		message.exec();
	}
	m_mainArea->setFocus();
}

static HRESULT wppSaveCallback(_Presentation* pdoc)
{
	QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("保存回调"));
	message.exec();
	return S_OK;
}

void WPPMainWindow::registerSaveEvent()
{
	if (m_spApplication)
	{
		HRESULT ret = m_rpcClient->registerEvent(m_spApplication, IID_EApplication, __X("PresentationSave"), (void*)wppSaveCallback);
		QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), "");
		if (ret == S_OK)
			message.setText(QString::fromUtf8("注册成功"));
		else
			message.setText(QString::fromUtf8("注册失败"));
		message.exec();
	}
	m_mainArea->setFocus();
}

void WPPMainWindow::insertPicture()
{
	if (m_spApplication)
	{
		//向当前幻灯片的第一页 插入图片
		ks_stdptr<_Presentation> spPresentation;
		HRESULT hr = m_spApplication->get_ActivePresentation((Presentation**) &spPresentation);
		if (hr != S_OK)
		{
			qDebug() << "can not get ActivePresentation";
			return;
		}
		ks_stdptr<Slides> spSlides;
		hr = spPresentation->get_Slides(&spSlides);
		if (hr != S_OK || spSlides == NULL)
		{
			qDebug() << "can not get Slides";
			return;
		}
		KComVariant indexVar(1);
		ks_stdptr<_Slide> spSlide;
		hr = spSlides->Item(indexVar,(Slide**)&spSlide);
		if (hr != S_OK || spSlide == NULL)
		{
			qDebug() << "can not get Slide 1";
			return;
		}
		ks_stdptr<Shapes> spShapes;
		hr = spSlide->get_Shapes((Shapes**)&spShapes);
		if (hr != S_OK || spShapes == NULL)
		{
			qDebug() << "can not get shapes";
			return;
		}
		ks_stdptr<Shape> spShapePic;
		QString filePath = QFileDialog::getOpenFileName((QWidget*)parent(), QString::fromUtf8("插入图片"));
		ks_bstr fileName(filePath.utf16());
		hr = spShapes->AddPicture(fileName, msoTrue, msoTrue, 20, 20, 500, 500, &spShapePic);
		if (hr != S_OK || spShapePic == NULL)
		{
			qDebug() << "can not insert picture";
			return;
		}
		qDebug() << "InsertPicture:" << (hr == S_OK ? true : false);
	}
	m_mainArea->setFocus();
}

void WPPMainWindow::insertSlide()
{
	if (m_spApplication)
	{
		ks_stdptr<_Presentation> spDoc;
		m_spApplication->get_ActivePresentation((Presentation**)&spDoc);
		if (spDoc)
		{
			ks_stdptr<Slides> spSlides;
			spDoc->get_Slides(&spSlides);
			if (spSlides)
			{
				ks_stdptr<_Slide> spSlide;
				spSlides->Add(1, ppLayoutBlank, (Slide**)&spSlide);
			}
		}
	}
	m_mainArea->setFocus();
}

void WPPMainWindow::startPlay()
{
	if (m_spApplication)
	{
		ks_stdptr<_Presentation> spDoc;
		m_spApplication->get_ActivePresentation((Presentation**)&spDoc);
		if (spDoc)
		{
			ks_stdptr<SlideShowSettings> spSlideShowSettings;
			spDoc->get_SlideShowSettings(&spSlideShowSettings);
			if (spSlideShowSettings)
			{
				ks_stdptr<SlideShowWindow> spSlideShowWindow;
				spSlideShowSettings->Run(&spSlideShowWindow);
			}
		}
	}
	m_mainArea->setFocus();
}

void WPPMainWindow::slideSize()
{
	if (m_spApplication)
	{
		ks_stdptr<wppapi::_Presentation> spPresentation;
		HRESULT hr = m_spApplication->get_ActivePresentation((wppapi::Presentation**)&spPresentation);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("获取Presentation对象失败"));
			message.exec();
			return;
		}

		ks_stdptr<wppapi::PageSetup> spPageSetup;
		hr = spPresentation->get_PageSetup(&spPageSetup);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("获取PageSetup对象失败"));
			message.exec();
			return;
		}

		hr = spPageSetup->put_SlideSize(wppapi::ppSlideSize35MM);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("设置幻灯片大小失败"));
			message.exec();
			return;
		}
	}
	m_mainArea->setFocus();
}
void WPPMainWindow::addTable()
{
	if (m_spApplication)
	{
		ks_stdptr<wppapi::_Presentation> spPresentation;
		HRESULT hr = m_spApplication->get_ActivePresentation((wppapi::Presentation**)&spPresentation);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("获取Presentation对象失败"));
			message.exec();
			return;
		}

		ks_stdptr<wppapi::Slides> spSlides;
		hr = spPresentation->get_Slides(&spSlides);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("获取Slides对象失败"));
			message.exec();
			return;
		}

		KComVariant index(1);
		ks_stdptr<wppapi::_Slide> spSlide;
		hr = spSlides->Item(index, (wppapi::Slide**)&spSlide);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("获取Slide对象失败"));
			message.exec();
			return;
		}

		ks_stdptr<wppapi::Shapes> spShapes;
		hr = spSlide->get_Shapes(&spShapes);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("获取Shapes对象失败"));
			message.exec();
			return;
		}

		int NumRows = 4;
		int NumColumns = 5;
		single Left = 10;
		single Top = 10;
		single Width = 500;
		single Height = 200;
		ks_stdptr<wppapi::Shape> spShape;
		hr = spShapes->AddTable(NumRows, NumColumns, Left, Top, Width, Height, &spShape);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("添加表格对象失败"));
			message.exec();
			return;
		}
	}
	m_mainArea->setFocus();
}

void WPPMainWindow::newSection()
{
	if (m_spApplication)
	{
		ks_stdptr<wppapi::_Presentation> spPresentation;
		HRESULT hr = m_spApplication->get_ActivePresentation((wppapi::Presentation**)&spPresentation);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("获取Rpesentation对象失败"));
			message.exec();
			return;
		}

		ks_stdptr<wppapi::SectionProperties> spSectionProperties;
		hr = spPresentation->get_SectionProperties(&spSectionProperties);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("回球SectionProperties对象失败"));
			message.exec();
			return;
		}

		long count = 0;
		hr = spSectionProperties->get_Count(&count);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("获取节数量失败"));
			message.exec();
			return;
		}

		count++;
		QString qstrName(QString::fromUtf8("新增节 %1").arg(count));
		KComVariant name(qstrName.utf16());

		int section= 0;
		hr = spSectionProperties->AddSection(count, name, &section);
	}
	m_mainArea->setFocus();
}
void WPPMainWindow::addAudio()
{
	if (m_spApplication)
	{
		ks_stdptr<wppapi::DocumentWindow> spDocumentWindow;
		HRESULT hr = m_spApplication->get_ActiveWindow(&spDocumentWindow);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("获取DocumentWind对象失败"));
			message.exec();
			return;
		}

		ks_stdptr<wppapi::Selection> spSelection;
		hr = spDocumentWindow->get_Selection(&spSelection);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("获取Selection对象失败"));
			message.exec();
			return;
		}

		ks_stdptr<wppapi::SlideRange> spSlideRange;
		hr = spSelection->get_SlideRange(&spSlideRange);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("获取SlideRange对象失败"));
			message.exec();
			return;
		}

		ks_stdptr<wppapi::Shapes> spShapes;
		hr = spSlideRange->get_Shapes(&spShapes);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("获取Shapes对象失败"));
			message.exec();
			return;
		}

		QString name = QFileDialog::getOpenFileName(this, QString::fromUtf8("选择一个音频文件"));
		ks_bstr FileName(name.utf16());
		single Left = 300;
		single Top = 300;
		single Width = 50;
		single Height = 50;
		ks_stdptr<wppapi::Shape> spShape;
		hr = spShapes->AddMediaObject(FileName, Left, Top, Width, Height, &spShape);

		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("添加音频文件失败"));
			message.exec();
			return;
		}

		hr = spShape->Select();
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("选中对象失败"));
			message.exec();
			return;
		}
	}
	m_mainArea->setFocus();
}

void WPPMainWindow::hyperlink()
{
	if (m_spApplication)
	{
		ks_stdptr<wppapi::DocumentWindow> spDocumentWindow;
		HRESULT hr = m_spApplication->get_ActiveWindow(&spDocumentWindow);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("获取DocumentWindow对象失败"));
			message.exec();
			return;
		}

		ks_stdptr<wppapi::Selection> spSelection;
		hr = spDocumentWindow->get_Selection(&spSelection);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("获取Selection对象失败"));
			message.exec();
			return;
		}

		ks_stdptr<wppapi::TextRange> spTextRange;
		hr = spSelection->get_TextRange(&spTextRange);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("获取TextRange对象失败"));
			message.exec();
			return;
		}

		ks_stdptr<wppapi::ActionSettings> spActionSettings;
		hr = spTextRange->get_ActionSettings(&spActionSettings);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("获取ActionSettings对象失败"));
			message.exec();
			return;
		}

		ks_stdptr<wppapi::ActionSetting> spActionSetting;
		hr = spActionSettings->Item(wppapi::ppMouseClick, &spActionSetting);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("获取ActionSetting对象失败"));
			message.exec();
			return;
		}

		ks_stdptr<wppapi::Hyperlink> spHyperlink;
		hr = spActionSetting->get_Hyperlink(&spHyperlink);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("获取Hyperlink对象失败"));
			message.exec();
			return;
		}

		QString name = QFileDialog::getOpenFileName(this, QString::fromUtf8("选择一个文件"));
		ks_bstr Address(name.utf16());
		hr = spHyperlink->put_Address(Address);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("设置链接路径失败"));
			message.exec();
			return;
		}

		hr = spHyperlink->put_TextToDisplay(Address);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("设置显示文本失败"));
			message.exec();
			return;
		}

		ks_stdptr<wppapi::TextRange> spCharacters;
		hr = spTextRange->Characters(1, 0, &spCharacters);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("获取Characters对象失败"));
			message.exec();
			return;
		}

		hr = spTextRange->put_Text(Address);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("输入文字失败"));
			message.exec();
			return;
		}
	}
	m_mainArea->setFocus();
}

void WPPMainWindow::insertTextbox()
{
	if (m_spApplication)
	{
		ks_stdptr<wppapi::DocumentWindow> spDocumentWindow;
		HRESULT hr = m_spApplication->get_ActiveWindow(&spDocumentWindow);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("获取DocumentWindow对象失败"));
			message.exec();
			return;
		}

		ks_stdptr<wppapi::Selection> spSelection;
		hr = spDocumentWindow->get_Selection(&spSelection);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("获取Selection对象失败"));
			message.exec();
			return;
		}

		ks_stdptr<wppapi::SlideRange> spSlideRange;
		hr = spSelection->get_SlideRange(&spSlideRange);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("获取SlideRange对象失败"));
			message.exec();
			return;
		}

		ks_stdptr<wppapi::Shapes> spShapes;
		hr = spSlideRange->get_Shapes(&spShapes);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("获取Shapes对象失败"));
			message.exec();
			return;
		}

		double Left = 100;
		double Top = 100;
		double Width = 480;
		double Height = 320;
		ks_stdptr<wppapi::Shape> spShape;
		hr = spShapes->AddTextbox(msoTextOrientationHorizontal, Left, Top, Width, Height, &spShape);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("添加文本框对象失败"));
			message.exec();
			return;
		}

		ks_stdptr<wppapi::TextFrame> spTextFrame;
		hr = spShape->get_TextFrame(&spTextFrame);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("获取TextFrame对象失败"));
			message.exec();
			return;
		}

		ks_stdptr<wppapi::TextRange> spTextRange;
		hr = spTextFrame->get_TextRange(&spTextRange);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("获取TextRange对象失败"));
			message.exec();
			return;
		}

		ks_bstr text(__X("插入文本框"));
		hr = spTextRange->put_Text(text);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("输入文字失败"));
			message.exec();
			return;
		}
	}
	m_mainArea->setFocus();
}

void WPPMainWindow::slotButtonClick(const QString& name)
{
	typedef void (WPPMainWindow::*WppOperationFun)(void);
	static QMap<QString, WppOperationFun> s_operationFunMap;
	if (s_operationFunMap.isEmpty())
	{
		s_operationFunMap.insert(QString::fromUtf8("初始化"), &WPPMainWindow::initApp);
		s_operationFunMap.insert(QString::fromUtf8("新建文档"), &WPPMainWindow::newDoc);
		s_operationFunMap.insert(QString::fromUtf8("打开文档"), &WPPMainWindow::openDoc);
		s_operationFunMap.insert(QString::fromUtf8("另存文档"), &WPPMainWindow::saveAs);
		s_operationFunMap.insert(QString::fromUtf8("打印到pdf"), &WPPMainWindow::printOut);
		s_operationFunMap.insert(QString::fromUtf8("隐藏工具栏"), &WPPMainWindow::hideToolBar);
		s_operationFunMap.insert(QString::fromUtf8("隐藏按钮"), &WPPMainWindow::hideToolBtn);
		s_operationFunMap.insert(QString::fromUtf8("注册关闭事件"), &WPPMainWindow::registerCloseEvent);
		s_operationFunMap.insert(QString::fromUtf8("注册保存事件"), &WPPMainWindow::registerSaveEvent);
		s_operationFunMap.insert(QString::fromUtf8("关闭WPP"), &WPPMainWindow::closeApp);
		s_operationFunMap.insert(QString::fromUtf8("插入图片"), &WPPMainWindow::insertPicture);
		s_operationFunMap.insert(QString::fromUtf8("添加幻灯片"), &WPPMainWindow::insertSlide);
		s_operationFunMap.insert(QString::fromUtf8("开始播放"), &WPPMainWindow::startPlay);
		s_operationFunMap.insert(QString::fromUtf8("新增节"), &WPPMainWindow::newSection);
		s_operationFunMap.insert(QString::fromUtf8("插入超链接"), &WPPMainWindow::hyperlink);
		s_operationFunMap.insert(QString::fromUtf8("插入音频"), &WPPMainWindow::addAudio);
		s_operationFunMap.insert(QString::fromUtf8("设置幻灯片大小"), &WPPMainWindow::slideSize);
		s_operationFunMap.insert(QString::fromUtf8("插入文本框"), &WPPMainWindow::insertTextbox);
		s_operationFunMap.insert(QString::fromUtf8("插入表格"), &WPPMainWindow::addTable);
	}
	if (s_operationFunMap.contains(name))
		(this->*s_operationFunMap.value(name))();
}
