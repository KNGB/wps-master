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
#include "wpsmainwindow.h"
#include "mainwindow.h"
#include "variant.h"
#include "QDebug"
#include <QMessageBox>
#include <QKeyEvent>
#include <QMouseEvent>
#include <QApplication>
#if (QT_VERSION >= QT_VERSION_CHECK(5,0,0))
#include <QWindow>
#endif

using namespace wpsapi;
using namespace kfc;
using namespace wpsapiex;

WPSWindow::WPSWindow(QWidget *parent)
#if (QT_VERSION < QT_VERSION_CHECK(5,0,0))
	: QX11EmbedContainer(parent)
	, m_containerWin(NULL)
	, m_hlayout(NULL)
{
	connect(this, SIGNAL(clientIsEmbedded()), this, SLOT(onEmbeded()));
}
#else
	: QWidget(parent)
	, m_containerWin(NULL)
{
	m_hlayout  = new QHBoxLayout;
	m_hlayout->setContentsMargins(1, 1, 1, 1);
	setLayout(m_hlayout);
}
#endif

WPSWindow::~WPSWindow()
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

IKRpcClient * WPSWindow::initWpsApplication()
{
	IKRpcClient *pWpsRpcClient = NULL;
	HRESULT hr = createWpsRpcInstance(&pWpsRpcClient);
	if (hr != S_OK)
	{
		qDebug() <<"get WpsRpcClient erro";
		return NULL;
	}
	BSTR StrWpsAddress = SysAllocString(__X("/opt/kingsoft/wps-office/office6/wps"));
	pWpsRpcClient->setProcessPath(StrWpsAddress);
	SysFreeString(StrWpsAddress);

	QVector<BSTR> vArgs;
	vArgs.resize(7);
	vArgs[0] =  SysAllocString(__X("-shield"));
	vArgs[1] =  SysAllocString(__X("-multiply"));
	vArgs[2] =  SysAllocString(__X("-x11embed"));
	vArgs[3] =  SysAllocString(QString::number(winId()).utf16());
	vArgs[4] =  SysAllocString(QString::number(width()).utf16());
	vArgs[5] =  SysAllocString(QString::number(height()).utf16());
	//-hidentp默认隐藏任务窗格
	//vArgs[6] =  SysAllocString(__X("-hidentp"));
	pWpsRpcClient->setProcessArgs(vArgs.size(), vArgs.data());

	for (int i = 0; i < vArgs.size(); i++)
	{
		SysFreeString(vArgs.at(i));
	}
	m_rpcClient = pWpsRpcClient;
	return pWpsRpcClient;
}

void WPSWindow::destroyWpsApplication()
{
	m_rpcClient = NULL;
}

void WPSWindow::addHotKey(const QString& hotKey, bool isEnable)
{
	m_enableHotKeysMap.insert(hotKey, isEnable);
}

void WPSWindow::removeHotKey(const QString& hotKey)
{
	m_enableHotKeysMap.remove(hotKey);
}

bool WPSWindow::eventFilter(QObject *Object, QEvent *Event)
{
	if (Event->type() == QEvent::KeyPress)
	{
		QKeyEvent *keyEvent = static_cast<QKeyEvent*>(Event);
		int keyInt = keyEvent->key();
		Qt::KeyboardModifiers modifiers = keyEvent->modifiers();
		if (modifiers & Qt::ShiftModifier)
			keyInt += Qt::SHIFT;
		if (modifiers & Qt::ControlModifier)
			keyInt += Qt::CTRL;
		if (modifiers & Qt::AltModifier)
			keyInt += Qt::ALT;
		if (modifiers & Qt::MetaModifier)
			keyInt += Qt::META;

		QKeySequence hotkey(keyInt);
		QMap<QString, bool>::iterator pos = m_enableHotKeysMap.begin();
		for (; pos != m_enableHotKeysMap.end(); pos++)
		{
			if (QKeySequence::ExactMatch == hotkey.matches(pos.key()))
			{
				if (!pos.value())
				{
					return true;
				}
			}
		}
	}
#if (QT_VERSION < QT_VERSION_CHECK(5,0,0))
	return QX11EmbedContainer::eventFilter(Object, Event);
#else
	return QWidget::eventFilter(Object, Event);
#endif
}

void WPSWindow::onEmbeded()
{
#if (QT_VERSION < QT_VERSION_CHECK(5,0,0))
	if (m_rpcClient)
		m_rpcClient->setWpsWid(this->clientWinId());
#endif
}

void WPSWindow::addContainerWin(QWidget *containerWin)
{
	if (m_containerWin)
	{
		m_hlayout->removeWidget(m_containerWin);
		delete m_containerWin;
		m_containerWin = NULL;
	}
	m_containerWin = containerWin;
	m_hlayout->addWidget(m_containerWin);
	m_hlayout->setStretchFactor(m_containerWin, 140);
}

WPSMainWindow::WPSMainWindow(QWidget *parent)
	: QWidget(parent)
	, m_mainArea(NULL)
{
	m_mainArea = new WPSWindow(this);
	m_hlayout = new QHBoxLayout;
	m_hlayout->setContentsMargins(1, 1, 1, 1);
	m_hlayout->addWidget(m_mainArea);
	m_hlayout->setStretchFactor(m_mainArea, 140);
	setLayout(m_hlayout);
	adjustSize();
}

WPSMainWindow::~WPSMainWindow()
{
	if (m_spApplication)
	{
		KComVariant vars[3];
		m_spApplication->Quit(vars, vars+1, vars+2);
		m_spApplication.clear();
	}
	m_mainArea->destroyWpsApplication();
	m_rpcClient = NULL;
	delete m_mainArea;
	delete m_hlayout;
}

void WPSMainWindow::closeApp()
{
	KComVariant vars[3];
	if (m_spApplication != NULL)
	{
		m_spApplication->Quit(vars, vars+1, vars+2);
		m_mainArea->destroyWpsApplication();
		m_spDocs = NULL;
		m_spApplication = NULL;
		m_spApplicationEx = NULL;
	}
}

void WPSMainWindow::initApp()
{
	if (!m_spApplication)
	{
		IKRpcClient *pRpcClient = m_mainArea->initWpsApplication();
		m_rpcClient = pRpcClient;
		if (pRpcClient && pRpcClient->getWpsApplication((IUnknown**)&m_spApplication) == S_OK)
		{
			m_spApplication->get_Documents(&m_spDocs);
			m_spApplication->QueryInterface(IID__WpsApplicationEx, (void**)&m_spApplicationEx);
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

void WPSMainWindow::newDoc()
{
	if (m_spApplication)
	{
		bool ret = m_spApplication->createDocument("wps");
		if (!ret)
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("新建失败"));
			message.exec();
		}
	}
	m_mainArea->setFocus();
}

void WPSMainWindow::openDoc()
{
	if (m_spApplication)
	{
		QString openFile = QFileDialog::getOpenFileName((QWidget*)parent(), QString::fromUtf8("打开文档"));
		if (!m_spApplication->openDocument(openFile, false))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("打开失败"));
			message.exec();
		}
	}
	m_mainArea->setFocus();
}

void WPSMainWindow::saveAs()
{
	if (m_spApplication)
	{
		QString saveFile = QFileDialog::getSaveFileName((QWidget*)parent(), QString::fromUtf8("保存文档"));
		if (!saveFile.isEmpty())
			m_spApplication->saveAs(saveFile);
	}
	m_mainArea->setFocus();
}

void WPSMainWindow::closeDoc()
{
	if (m_spApplication)
	{
		ks_stdptr<_Document> spDoc;
		m_spApplication->get_ActiveDocument((Document**)&spDoc);
		if (spDoc)
		{
			KComVariant varSaveOptions(wdPromptToSaveChanges);
			KComVariant varOriginalFormat(wdWordDocument);
			KComVariant varRouteDocument;
			spDoc->Close(&varSaveOptions, &varOriginalFormat, &varRouteDocument);
		}
	}
	m_mainArea->setFocus();
}

void WPSMainWindow::printOutDoc()
{
	if (m_spApplication)
	{
		ks_stdptr<_Document> spDoc;
		m_spApplication->get_ActiveDocument((Document**)&spDoc);
		if (spDoc)
		{
			KComVariant varEmpty;
			V_VT(&varEmpty) = VT_ERROR;
			V_ERROR(&varEmpty) = DISP_E_PARAMNOTFOUND;
			HRESULT ret = spDoc->PrintOut(&varEmpty, &varEmpty, &varEmpty,&varEmpty, &varEmpty, &varEmpty, &varEmpty, &varEmpty, &varEmpty, &varEmpty, &varEmpty,
							&varEmpty, &varEmpty, &varEmpty, &varEmpty, &varEmpty, &varEmpty, &varEmpty);
		}
	}
	m_mainArea->setFocus();
}

void WPSMainWindow::showAllTool()
{
	static bool enable = false;
	if (m_spApplication)
	{
		m_spApplication->setToolbarAllVisible(enable);
		enable = !enable;
	}
	m_mainArea->setFocus();
}

void WPSMainWindow::forceBackUpEnabled()
{
	static bool enable = false;
	if (m_spApplicationEx)
	{
		m_spApplicationEx->put_ForceBackupEnable(enable ? VARIANT_TRUE : VARIANT_FALSE);
		enable = !enable;
	}
	m_mainArea->setFocus();
}

void WPSMainWindow::hideToolBar()
{
	/*
	"Standard"							常用
	"Formatting"						格式
	"Tables and Borders"				表格和边框
	"Align"								对象对齐
	"Reviewing"							审阅
	"Extended Formatting"				其他格式
	"Drawing"							绘图
	"Symbol"							符号栏
	"3-D Settings"						三维设置
	"Shadow Settings"					阴影设置
	"Outlining"							大纲
	"Forms"								窗体
	*/
	static bool enable = false;
	ks_stdptr<ksoapi::_CommandBars> spCommandBars;
	if (m_spApplication)
	{
		m_spApplication->get_CommandBars((ksoapi::CommandBars**)&spCommandBars);
		if (spCommandBars)
		{
			ks_stdptr<ksoapi::CommandBar> spCommandBar;
			KComVariant vIndex(__X("Standard"));
			spCommandBars->get_Item(vIndex, &spCommandBar);
			if (spCommandBar)
			{
				spCommandBar->put_Visible(enable ? VARIANT_TRUE : VARIANT_FALSE);
				enable = !enable;
			}
		}
	}
	m_mainArea->setFocus();
}

void WPSMainWindow::setHeaderFooter()
{
	if (!m_spApplication)
		return;
	ks_stdptr<_Document> spDoc;
	ks_stdptr<Section> spSection;
	m_spApplication->get_ActiveDocument((Document**)&spDoc);
	if (spDoc)
	{
		ks_stdptr<Sections> spSections;
		spDoc->get_Sections(&spSections);
		if (spSections)
		{
			spSections->Item(1, &spSection);
			if (!spSection)
				return;
		}
	}

	ks_stdptr<HeadersFooters> spHeaders;
	spSection->get_Headers(&spHeaders);
	if (spHeaders)
	{
		ks_stdptr<HeaderFooter> spHeader;
		spHeaders->Item(wdHeaderFooterPrimary, &spHeader);
		if (spHeader)
		{
			ks_stdptr<Range> spRange;
			spHeader->get_Range(&spRange);
			if (spRange)
			{
				ks_bstr text( __X("Header") );
				spRange->put_Text(text);
			}
		}
	}
	ks_stdptr<HeadersFooters> spFooters;
	spSection->get_Footers(&spFooters);
	if (spFooters)
	{
		ks_stdptr<HeaderFooter> spFooter;
		spFooters->Item(wdHeaderFooterPrimary, &spFooter);
		if (spFooter)
		{
			ks_stdptr<Range> spRange;
			spFooter->get_Range(&spRange);
			if (spRange)
			{
				ks_bstr text(__X("Footer"));
				spRange->put_Text(text);
			}
		}
	}


	m_mainArea->setFocus();
}

void WPSMainWindow::getSaved()
{
	if (m_spApplication)
	{
		ks_stdptr<_Document> spDoc;
		m_spApplication->get_ActiveDocument((Document**)&spDoc);
		if (spDoc)
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), "");
			VARIANT_BOOL isSaved = VARIANT_FALSE;
			spDoc->get_Saved(&isSaved);
			if (isSaved == VARIANT_FALSE)
				message.setText(QString::fromUtf8("未保存"));
			else
				message.setText(QString::fromUtf8("已保存"));
			message.exec();
		}
	}
	m_mainArea->setFocus();
}

void WPSMainWindow::getPagesNum()
{
	if (m_spApplication)
	{
		ks_stdptr<_Document> spDoc;
		m_spApplication->get_ActiveDocument((Document**)&spDoc);
		if (spDoc)
		{
			long nPages = 2;
			KComVariant vIncludeFootnotesAndEndnotes;
			spDoc->ComputeStatistics(wdStatisticPages, &vIncludeFootnotesAndEndnotes, &nPages);
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("当前页数: %1").arg(nPages));
			message.exec();
		}
	}
	m_mainArea->setFocus();
}

void WPSMainWindow::disableHotKey()
{
	m_mainArea->addHotKey("ctrl+s", false);
}

void WPSMainWindow::enableProtect()
{
	if (m_spApplication)
	{
		ks_stdptr<_Document> spDoc;
		m_spApplication->get_ActiveDocument((Document**)&spDoc);
		if (spDoc)
		{
			KComVariant varEmpty;
			V_VT(&varEmpty) = VT_ERROR;
			V_ERROR(&varEmpty) = DISP_E_PARAMNOTFOUND;
			KComVariant vNoReset;
			vNoReset.AssignBOOL(TRUE);
			KComVariant vPassword(__X("123"));
			spDoc->Protect(wdAllowOnlyReading, &vNoReset, &vPassword, &varEmpty, &varEmpty);
		}
	}
	m_mainArea->setFocus();
}

void WPSMainWindow::disableProtect()
{
	if (m_spApplication)
	{
		ks_stdptr<_Document> spDoc;
		m_spApplication->get_ActiveDocument((Document**)&spDoc);
		if (spDoc)
		{
			KComVariant vPassword(__X("123"));
			spDoc->Unprotect(&vPassword);
		}
	}
	m_mainArea->setFocus();
}

void WPSMainWindow::hideToolBtn()
{
	if (m_spApplication)
	{
		static bool enable = false;
		ks_stdptr<ksoapi::_CommandBars> spCommandBars;
		m_spApplication->get_CommandBars((ksoapi::CommandBars**)&spCommandBars);
		if (spCommandBars)
		{
			ks_stdptr<ksoapi::CommandBar> spCommandBar;
			KComVariant vCommandBarIndex(__X("Formatting"));
			spCommandBars->get_Item(vCommandBarIndex, &spCommandBar);
			if (spCommandBar)
			{
				ks_stdptr<ksoapi::CommandBarControls> spControls;
				spCommandBar->get_Controls(&spControls);
				int iCount = 0;
				HRESULT ret = spControls->get_Count(&iCount);
				if (FAILED(ret))
					return;

				for (int i = 1; i <= iCount; i++)
				{
					KComVariant vControlIndex(i);
					ks_stdptr<ksoapi::CommandBarControl> spControl;
					spControls->get_Item(vControlIndex, &spControl);
					if (spControl)
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
	}
	m_mainArea->setFocus();
}

void WPSMainWindow::listTemplates()
{
	if (m_spApplication)
	{
		ks_stdptr<wpsapi::Selection> spSelection;
		HRESULT hr = m_spApplication->get_Selection(&spSelection);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("获取Selection对象失败"));
			message.exec();
			return;
		}

		ks_bstr context(__X("编号一"));
		hr = spSelection->TypeText(context);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("输入文字失败"));
			message.exec();
			return;
		}

		ks_stdptr<wpsapi::Range> spRange;
		hr = spSelection->get_Range(&spRange);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("获取Range对象失败"));
			message.exec();
			return;
		}

		ks_stdptr<wpsapi::ListFormat> spListFormat;
		hr = spRange->get_ListFormat(&spListFormat);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("回去ListFormat对象失败"));
			message.exec();
			return;
		}

		ks_stdptr<wpsapi::ListGalleries> spListGalleries;
		hr = m_spApplication->get_ListGalleries(&spListGalleries);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("获取ListGalleries对象失败"));
			message.exec();
			return;
		}

		ks_stdptr<wpsapi::ListGallery> spListGallery;
		hr = spListGalleries->Item(wpsapi::wdNumberGallery, &spListGallery);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("获取ListGallerie对象失败"));
			message.exec();
			return;
		}

		ks_stdptr<wpsapi::ListTemplates> spListTemplates;
		hr = spListGallery->get_ListTemplates(&spListTemplates);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("获取ListTemplates对象失败"));
			message.exec();
			return;
		}

		KComVariant Index(1);
		ks_stdptr<wpsapi::ListTemplate> spListTemplate;
		spListTemplates->Item(&Index, &spListTemplate);

		KComVariant ApplyTo(wdListApplyToSelection);
		KComVariant ContinuePreviousList(wdContinueList);
		KComVariant DefaultListBehavior(262144);
		hr = spListFormat->ApplyListTemplate(spListTemplate, &ContinuePreviousList, &ApplyTo, &DefaultListBehavior);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("添加列表失败"));
			message.exec();
			return;
		}
	}
	m_mainArea->setFocus();
}

void WPSMainWindow::hyperlinks()
{
	if (m_spApplication)
	{
		ks_stdptr<wpsapi::_Document> spActiveDoc;
		HRESULT hr = m_spApplication->get_ActiveDocument((wpsapi::Document**)&spActiveDoc);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("获取document对象失败"));
			message.exec();
			return;
		}

		ks_stdptr<wpsapi::Selection> spSelection;
		hr = m_spApplication->get_Selection(&spSelection);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("获取selection对象失败"));
			message.exec();
			return;
		}

		ks_stdptr<wpsapi::Range> spRange;
		hr = spSelection->get_Range(&spRange);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("获取range对象失败"));
			message.exec();
			return;
		}

		ks_stdptr<wpsapi::Hyperlinks> spHyperlinks;
		hr = spActiveDoc->get_Hyperlinks(&spHyperlinks);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("获取Hyperlinks对象失败"));
			message.exec();
			return;
		}

		ks_stdptr<wpsapi::Hyperlink> spHyperlink;
		QString strAddress = QFileDialog::getOpenFileName(this, QString::fromUtf8("选择一个文件"));
		LPCOLESTR uStr = strAddress.utf16();
		KComVariant Address(uStr);
		KComVariant varEmpty;
		V_VT(&varEmpty) = VT_ERROR;
		V_ERROR(&varEmpty) = DISP_E_PARAMNOTFOUND;
		hr = spHyperlinks->Add((IDispatch*)spRange, &Address, &varEmpty, &varEmpty, &varEmpty, &varEmpty, &spHyperlink);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("添加超链接失败"));
			message.exec();
			return;
		}
	}
	m_mainArea->setFocus();
}

void WPSMainWindow::highlight()
{
	if (m_spApplication)
	{
		ks_stdptr<wpsapi::Selection> spSelection;
		HRESULT hr = m_spApplication->get_Selection(&spSelection);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("获取selection对象失败"));
			message.exec();
			return;
		}

		ks_bstr context(__X("突出显示"));;
		spSelection->TypeText(context);

		long end = 0;
		spSelection->get_End(&end);
		spSelection->SetRange(end - 4, end);

		ks_stdptr<wpsapi::Range> spRange;
		hr = spSelection->get_Range(&spRange);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("获取range对象失败"));
			message.exec();
			return;
		}

		hr = spRange->put_HighlightColorIndex(wpsapi::wdYellow);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("设置突出显示失败"));
			message.exec();
			return;
		}
	}
	m_mainArea->setFocus();
}

void WPSMainWindow::inserttable()
{
	if (m_spApplication)
	{
		ks_stdptr<wpsapi::_Document> spActiveDoc;
		HRESULT hr = m_spApplication->get_ActiveDocument((wpsapi::Document**)&spActiveDoc);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("获取Document对象失败"));
			message.exec();
			return;
		}

		ks_stdptr<wpsapi::Range> spRange;
		hr = spActiveDoc->get_Content(&spRange);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("获取Range对象失败"));
			message.exec();
			return;
		}

		ks_stdptr<wpsapi::Tables> spTables;
		hr = spActiveDoc->get_Tables(&spTables);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("获取Tables对象失败"));
			message.exec();
			return;
		}

		long NumRows = 3;
		long NumColumns = 5;
		KComVariant DefaultTableBehavior;
		KComVariant AutoFitBehavior;
		ks_stdptr<wpsapi::Table> spTable;
		hr = spTables->Add(spRange, NumRows, NumColumns, &DefaultTableBehavior, &AutoFitBehavior, &spTable);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("添加表格失败"));
			message.exec();
			return;
		}
	}
	m_mainArea->setFocus();
}

void WPSMainWindow::setFontSize()
{
	if (m_spApplication)
	{
		ks_stdptr<wpsapi::Selection> spSelection;
		HRESULT hr = m_spApplication->get_Selection(&spSelection);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("获取Selection对象失败"));
			message.exec();
			return;
		}

		ks_stdptr<wpsapi::_Font> spFont;
		hr = spSelection->get_Font((wpsapi::Font**)&spFont);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("获取Font对象失败"));
			message.exec();
			return;
		}

		hr = spFont->put_Size(36);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("设置字体大小失败"));
			message.exec();
			return;
		}

		ks_bstr context(__X("字体设为36"));;
		spSelection->TypeText(context);
	}
	m_mainArea->setFocus();
}

void WPSMainWindow::insertEllipse()
{
	if (m_spApplication)
	{
		ks_stdptr<wpsapi::_Document> spDocumment;
		HRESULT hr = m_spApplication->get_ActiveDocument((wpsapi::Document**)&spDocumment);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("获取Document对象失败"));
			message.exec();
			return;
		}

		ks_stdptr<wpsapi::Shapes> spShapes;
		hr = spDocumment->get_Shapes(&spShapes);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("获取shapes对象失败"));
			message.exec();
			return;
		}

		long Type = (long)msoShapeOval;
		single Left = 150;
		single Top = 122;
		single Width = 118;
		single Height = 83;
		KComVariant Anchor;
		ks_stdptr<wpsapi::Shape> spShape;
		hr = spShapes->AddShape(Type, Left, Top, Width, Height, &Anchor, &spShape);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("添加图形失败"));
			message.exec();
			return;
		}
	}
	m_mainArea->setFocus();
}

static HRESULT wpsDocumentBeforeCloseCallback(_Document* pdoc, VARIANT_BOOL* Cancel)
{
	*Cancel = VARIANT_FALSE;
	QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("关闭回调"));
	message.exec();
}

void WPSMainWindow::registerCloseEvent()
{
	if (m_spApplication)
	{
		HRESULT ret = m_rpcClient->registerEvent(m_spApplication, DIID_ApplicationEvents4, __X("DocumentBeforeClose"), (void*)wpsDocumentBeforeCloseCallback);
		QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), "");
		if (ret == S_OK)
			message.setText(QString::fromUtf8("注册成功"));
		else
			message.setText(QString::fromUtf8("注册失败"));
		message.exec();
	}
	m_mainArea->setFocus();
}

static HRESULT wpsDocumentBeforeSaveCallback(_Document* pdoc, VARIANT_BOOL* SaveAsUI, VARIANT_BOOL* Cancel)
{
	*SaveAsUI = VARIANT_TRUE;
	QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("保存回调"));
	message.exec();
}

void WPSMainWindow::registerSaveEvent()
{
	if (m_spApplication)
	{
		HRESULT ret = m_rpcClient->registerEvent(m_spApplication, DIID_ApplicationEvents4, __X("DocumentBeforeSave"), (void*)wpsDocumentBeforeSaveCallback);
		QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), "");
		if (ret == S_OK)
			message.setText(QString::fromUtf8("注册成功"));
		else
			message.setText(QString::fromUtf8("注册失败"));
		message.exec();
	}
	m_mainArea->setFocus();
}


void WPSMainWindow::slotButtonClick(const QString& name)
{
	typedef void (WPSMainWindow::*WpsOperationFun)(void);
	static QMap<QString, WpsOperationFun> s_operationFunMap;
	if (s_operationFunMap.isEmpty())
	{
		s_operationFunMap.insert(QString::fromUtf8("初始化"), &WPSMainWindow::initApp);
		s_operationFunMap.insert(QString::fromUtf8("新建文档"), &WPSMainWindow::newDoc);
		s_operationFunMap.insert(QString::fromUtf8("打开文档"), &WPSMainWindow::openDoc);
		s_operationFunMap.insert(QString::fromUtf8("另存文档"), &WPSMainWindow::saveAs);
		s_operationFunMap.insert(QString::fromUtf8("关闭文档"), &WPSMainWindow::closeDoc);
		s_operationFunMap.insert(QString::fromUtf8("打印文档"), &WPSMainWindow::printOutDoc);
		s_operationFunMap.insert(QString::fromUtf8("隐藏菜单"), &WPSMainWindow::showAllTool);
		s_operationFunMap.insert(QString::fromUtf8("隐藏工具栏"), &WPSMainWindow::hideToolBar);
		s_operationFunMap.insert(QString::fromUtf8("隐藏按钮"), &WPSMainWindow::hideToolBtn);
		s_operationFunMap.insert(QString::fromUtf8("保护"), &WPSMainWindow::enableProtect);
		s_operationFunMap.insert(QString::fromUtf8("取消保护"), &WPSMainWindow::disableProtect);
		s_operationFunMap.insert(QString::fromUtf8("是否保存"), &WPSMainWindow::getSaved);
		s_operationFunMap.insert(QString::fromUtf8("注册关闭事件"), &WPSMainWindow::registerCloseEvent);
		s_operationFunMap.insert(QString::fromUtf8("注册保存事件"), &WPSMainWindow::registerSaveEvent);
		s_operationFunMap.insert(QString::fromUtf8("关闭WPS"), &WPSMainWindow::closeApp);
		s_operationFunMap.insert(QString::fromUtf8("设置页眉页脚"), &WPSMainWindow::setHeaderFooter);
		s_operationFunMap.insert(QString::fromUtf8("获取页数"), &WPSMainWindow::getPagesNum);
		s_operationFunMap.insert(QString::fromUtf8("设置自动备份"), &WPSMainWindow::forceBackUpEnabled);
		s_operationFunMap.insert(QString::fromUtf8("添加编号"), &WPSMainWindow::listTemplates);
		s_operationFunMap.insert(QString::fromUtf8("插入超链接"), &WPSMainWindow::hyperlinks);
		s_operationFunMap.insert(QString::fromUtf8("突出显示"), &WPSMainWindow::highlight);
		s_operationFunMap.insert(QString::fromUtf8("插入表格"), &WPSMainWindow::inserttable);
		s_operationFunMap.insert(QString::fromUtf8("设置字体大小"), &WPSMainWindow::setFontSize);
		s_operationFunMap.insert(QString::fromUtf8("插入椭圆"), &WPSMainWindow::insertEllipse);

	}
	if (s_operationFunMap.contains(name))
		(this->*s_operationFunMap.value(name))();

}
