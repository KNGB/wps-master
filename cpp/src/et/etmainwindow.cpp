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
#include "etmainwindow.h"
#include "mainwindow.h"
#include "variant.h"
#include <QDebug>
#include <QTextCodec>
#include <QMessageBox>
#include <QKeyEvent>
#if (QT_VERSION >= QT_VERSION_CHECK(5,0,0))
#include <QWindow>
#endif

using namespace etapi;
using namespace kfc;
using namespace wpsapiex;

ETWindow::ETWindow(QWidget *parent)
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

ETWindow::~ETWindow()
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

IKRpcClient * ETWindow::initWpsApplication()
{
	IKRpcClient *pWpsRpcClient = NULL;
	HRESULT hr = createEtRpcInstance(&pWpsRpcClient);
	if (hr != S_OK)
	{
		qDebug() <<"get WpsRpcClient erro";
		return NULL;
	}
	BSTR StrWpsAddress = SysAllocString(__X("/opt/kingsoft/wps-office/office6/et"));
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

void ETWindow::destroyWpsApplication()
{
	m_rpcClient = NULL;
}

void ETWindow::addHotKey(const QString& hotKey, bool isEnable)
{
	m_enableHotKeysMap.insert(hotKey, isEnable);
}

void ETWindow::removeHotKey(const QString& hotKey)
{
	m_enableHotKeysMap.remove(hotKey);
}

bool ETWindow::eventFilter(QObject *Object, QEvent *Event)
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

void ETWindow::onEmbeded()
{
#if (QT_VERSION < QT_VERSION_CHECK(5,0,0))
	if (m_rpcClient)
		m_rpcClient->setWpsWid(this->clientWinId());
#endif
}

void ETWindow::addContainerWin(QWidget *containerWin)
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

EtMainWindow::EtMainWindow(QWidget *parent)
	: QWidget(parent)
	, m_mainArea(NULL)
{
	m_mainArea = new ETWindow(this);
	m_hlayout  = new QHBoxLayout(this);
	m_hlayout->setContentsMargins(1, 1, 1, 1);
	m_hlayout->addWidget(m_mainArea);
	m_hlayout->setStretchFactor(m_mainArea,140);
	setLayout(m_hlayout);
	adjustSize();
}

EtMainWindow::~EtMainWindow()
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

void EtMainWindow::closeApp()
{
	if (m_spApplication != NULL)
	{
		m_spWorkbooks.clear();
		m_spApplicationEx.clear();
		m_spApplication->Quit();
		m_spApplication.clear();
		m_mainArea->destroyWpsApplication();
	}
}

void EtMainWindow::initApp()
{
	if (!m_spApplication)
	{
		IKRpcClient *pRpcClient = m_mainArea->initWpsApplication();
		m_rpcClient = pRpcClient;
		if (pRpcClient && pRpcClient->getEtApplication((IUnknown**)&m_spApplication) == S_OK)
		{
			m_spApplication->get_Workbooks(&m_spWorkbooks);
			m_spApplication->QueryInterface(IID__EtApplicationEx, (void**)&m_spApplicationEx);
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

void EtMainWindow::newDoc()
{
	if (m_spWorkbooks)
	{
		ks_stdptr<_Workbook> spWorkbook;
		KComVariant vars;
		HRESULT hr = m_spWorkbooks->Add(vars, ET_LCID, (Workbook**)&spWorkbook);
		if (hr == S_OK)
		{
			spWorkbook->Activate(ET_LCID);
		}
		else
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("新建失败"));
			message.exec();
		}

	}
	m_mainArea->setFocus();
}

void EtMainWindow::openDoc()
{
	if (m_spWorkbooks)
	{
		QString filePath = QFileDialog::getOpenFileName((QWidget*)parent(), QString::fromUtf8("打开文档"));
		ks_bstr kstrFilename(filePath.utf16());
		ks_stdptr<_Workbook> spWorkbook;
		KComVariant vars[12];
		HRESULT hr = m_spWorkbooks->_Open(kstrFilename, vars[0], vars[1], vars[2], vars[3], vars[4], vars[5], vars[6], vars[7]
				, vars[8], vars[9], vars[10], vars[11], ET_LCID, (Workbook**)&spWorkbook);
		if (SUCCEEDED(hr))
		{
			spWorkbook->Activate(ET_LCID);
		}
		else
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("打开失败"));
			message.exec();
		}
	}
	m_mainArea->setFocus();
}


void EtMainWindow::saveDoc()
{
	if (m_spApplication)
	{
		ks_stdptr<_Workbook> spWorkbook;
		m_spApplication->get_ActiveWorkbook((Workbook**)&spWorkbook);
		if (spWorkbook)
		{
			HRESULT hr = spWorkbook->Save(ET_LCID);
			if (FAILED(hr))
			{
				QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("保存失败"));
				message.exec();
			}
		}
	}
	m_mainArea->setFocus();
}

void EtMainWindow::saveAs()
{
	if (m_spApplication)
	{
		ks_stdptr<_Workbook> spWorkbook;
		m_spApplication->get_ActiveWorkbook((Workbook**)&spWorkbook);
		if (spWorkbook)
		{
			QString filePath = QFileDialog::getSaveFileName((QWidget*)parent(), QString::fromUtf8("保存文档"));
			KComVariant fnv(filePath.utf16());
			XlSaveAsAccessMode mode = xlNoChange;
			KComVariant vars[9];
			HRESULT hr = spWorkbook->_SaveAs(fnv, vars[0], vars[1], vars[2], vars[3], vars[4], mode, vars[5], vars[6], vars[7],
					vars[8], ET_LCID);
			if (hr != S_OK)
			{
				QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("另存失败"));
				message.exec();
			}
		}
	}
	m_mainArea->setFocus();
}

void EtMainWindow::forceBackUpEnabled()
{
	static bool enable = true;
	enable = !enable;
	if (m_spApplicationEx)
		m_spApplicationEx->put_ForceBackupEnable(enable ? VARIANT_TRUE : VARIANT_FALSE);
	m_mainArea->setFocus();
}

void EtMainWindow::hideToolBar()
{
	/*
	"Standard"							常用工具栏
	"Formatting"						格式工具栏
	"Drawing"							绘图
	"3-D Settings"						三维设置
	"Shadow Settings"					阴影设置
	"Protection"						保护
	"Forms"								窗体
	"Borders"							边框
	*/
	static bool enable = false;
	if (m_spApplication)
	{
		ks_stdptr<ksoapi::_CommandBars> spCommandBars;
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

void EtMainWindow::setHeaderFooter()
{
	if (m_spApplication)
	{
		ks_stdptr<etapi::_Worksheet> spWorksheet;
		m_spApplication->get_ActiveSheet((IDispatch**)&spWorksheet);
		if (spWorksheet)
		{
			ks_stdptr<IPageSetup> spPageSetup;
			spWorksheet->get_PageSetup((PageSetup**)&spPageSetup);
			if (spPageSetup)
			{
				ks_bstr leftHeader(__X("leftHeader"));
				spPageSetup->put_LeftHeader(leftHeader);				//页眉左
				ks_bstr rightHeader(__X("rightHeader"));
				spPageSetup->put_RightHeader(rightHeader);				//页眉右
				ks_bstr centerHeader(__X("centerHeader"));
				spPageSetup->put_CenterHeader(centerHeader);			//页眉中
				ks_bstr leftFooter(__X("leftFooter"));
				spPageSetup->put_LeftFooter(leftFooter);				//页脚左
				ks_bstr rightFooter(__X("rightFooter"));
				spPageSetup->put_RightFooter(rightFooter);				//页脚右
				ks_bstr centerFooter(__X("centerFooter"));
				spPageSetup->put_CenterFooter(centerFooter);			//页脚中
			}
		}
	}
	m_mainArea->setFocus();
}

void EtMainWindow::getSaved()
{
	if (m_spApplication)
	{
		ks_stdptr<_Workbook> spWorkbook;
		m_spApplication->get_ActiveWorkbook((Workbook**)&spWorkbook);
		if (spWorkbook)
		{
			VARIANT_BOOL isSaved = VARIANT_FALSE;
			spWorkbook->get_Saved(0, &isSaved);
			if (isSaved == VARIANT_FALSE)
			{
				QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("未保存"));
				message.exec();
			}
			else
			{
				QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("已保存"));
				message.exec();
			}
		}
	}
	m_mainArea->setFocus();
}

void EtMainWindow::getPagesNum()
{
	if (m_spApplication)
	{
		ks_stdptr<etapi::_Worksheet> spWorksheet;
		m_spApplication->get_ActiveSheet((IDispatch**)&spWorksheet);
		if (spWorksheet)
		{
			ks_stdptr<etapi::IHPageBreaks> spHPageBreaks;
			spWorksheet->get_HPageBreaks((etapi::HPageBreaks**)&spHPageBreaks);
			ks_stdptr<etapi::IVPageBreaks> spVPageBreaks;
			spWorksheet->get_VPageBreaks((etapi::VPageBreaks**)&spVPageBreaks);
			if (spHPageBreaks && spVPageBreaks)
			{
				long hPages = 0, vPages = 0;
				spHPageBreaks->get_Count(&hPages);
				spVPageBreaks->get_Count(&vPages);
				long pagesNum = (hPages + 1) * (vPages + 1);
				QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("当前页数: %1").arg(pagesNum));
				message.exec();
			}
		}
	}
	m_mainArea->setFocus();
}

void EtMainWindow::disableSaveAsHotKey()
{
	m_mainArea->addHotKey("ctrl+s", false);
}

void EtMainWindow::setOrientation()
{
	static bool isLandscape = false;
	if (m_spApplication)
	{
		ks_stdptr<etapi::_Worksheet> spWorksheet;
		m_spApplication->get_ActiveSheet((IDispatch**)&spWorksheet);
		if (spWorksheet)
		{
			ks_stdptr<IPageSetup> spPageSetup;
			spWorksheet->get_PageSetup((PageSetup**)&spPageSetup);
			if (spPageSetup)
			{
				//xlLandscape: 横版  xlPortrait: 纵版
				if (isLandscape)
					spPageSetup->put_Orientation(xlPortrait);
				else
					spPageSetup->put_Orientation(xlLandscape);
				isLandscape = !isLandscape;
			}
		}
	}
	m_mainArea->setFocus();
}

void EtMainWindow::exportPdf()
{
	if (m_spApplication)
	{
		ks_stdptr<_Workbook> spWorkbook;
		m_spApplication->get_ActiveWorkbook((Workbook**)&spWorkbook);
		if (spWorkbook)
		{
			QString filePath = QFileDialog::getSaveFileName((QWidget*)parent(), QString::fromUtf8("保存文档"));
			KComVariant vPdfPath(filePath.utf16());
			KComVariant vEmpty;
			V_VT(&vEmpty) = VT_ERROR;
			V_ERROR(&vEmpty) = DISP_E_PARAMNOTFOUND;
			spWorkbook->ExportAsFixedFormat(etapi::xlTypePDF, vPdfPath, vEmpty, vEmpty, vEmpty, vEmpty, vEmpty, vEmpty, vEmpty);
		}
	}
	m_mainArea->setFocus();
}

void EtMainWindow::protect()
{
	if (m_spApplication)
	{
		ks_stdptr<etapi::_Worksheet> spWorkSheet;
		m_spApplication->get_ActiveSheet((IDispatch **)&spWorkSheet);
		if (!spWorkSheet)
			return;

		KComVariant vPassWord(__X("123"));
		KComVariant vEmpty;
		V_VT(&vEmpty) = VT_ERROR;
		V_ERROR(&vEmpty) = DISP_E_PARAMNOTFOUND;
		spWorkSheet->Protect(vPassWord, vEmpty, vEmpty, vEmpty, vEmpty, vEmpty, vEmpty, vEmpty,
							 vEmpty, vEmpty, vEmpty, vEmpty, vEmpty, vEmpty, vEmpty, vEmpty);
	}
	m_mainArea->setFocus();
}

void EtMainWindow::unProtect()
{
	if (m_spApplication)
	{
		ks_stdptr<etapi::_Worksheet> spWorkSheet;
		m_spApplication->get_ActiveSheet((IDispatch **)&spWorkSheet);
		if (!spWorkSheet)
			return;

		KComVariant vPassWord(__X("123"));
		spWorkSheet->Unprotect(vPassWord, ET_LCID);
	}
	m_mainArea->setFocus();
}

void EtMainWindow::hideToolBtn()
{
	if (m_spApplication)
	{
		static bool enable = false;
		ks_stdptr<ksoapi::_CommandBars> spCommandBars;
		m_spApplication->get_CommandBars((ksoapi::CommandBars**)&spCommandBars);
		if (spCommandBars)
		{
			ks_stdptr<ksoapi::CommandBar> spCommandBar;
			KComVariant varIndex(__X("Standard"));
			spCommandBars->get_Item(varIndex, &spCommandBar);
			if (!spCommandBar)
				return;
			ks_stdptr<ksoapi::CommandBarControls> spControls;
			spCommandBar->get_Controls(&spControls);
			if (spControls)
			{
				int nCount = 0;
				HRESULT hr = spControls->get_Count(&nCount);
				if (FAILED(hr))
					return;
				for (int i = 0; i < nCount; i++)
				{
					ks_stdptr<ksoapi::CommandBarControl> spControl;
					KComVariant varIndexControl(i);
					spControls->get_Item(varIndexControl, &spControl);
					if (spControl)
					{
						spControl->put_Visible(enable ? VARIANT_TRUE : VARIANT_FALSE);
						//ks_bstr ksName;
						//spControl->get_Caption(&ksName);
						//qDebug() << "index, name: " << i << QString::fromUtf16(ksName);		//button对应的index和name，用户可根据index隐藏自己所需按钮
					}
				}
				enable = !enable;
			}
		}
	}
	m_mainArea->setFocus();
}

void EtMainWindow::FontSize()
{
	if (m_spApplication)
	{
		ks_stdptr<etapi::_Workbook> spWorkbook;
		HRESULT hr = m_spApplication->get_ActiveWorkbook((etapi::Workbook**)&spWorkbook);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("获取workbook对象失败"));
			message.exec();
			return;
		}

		ks_stdptr<etapi::_Worksheet> spWorksheet;
		hr = spWorkbook->get_ActiveSheet((IDispatch**)&spWorksheet);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("获取Worksheet对象失败"));
			message.exec();
			return;
		}

		KComVariant cell(__X("A1"));
		ks_stdptr<etapi::IRange> spRange;
		hr = spWorksheet->get_Range(cell, cell, (etapi::Range**)&spRange);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("获取Range对象失败"));
			message.exec();
			return;
		}

		ks_stdptr<etapi::IFont> spFont;
		hr = spRange->get_Font((etapi::Font**)&spFont);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("获取Font对象失败"));
			message.exec();
			return;
		}

		KComVariant size(36);
		hr = spFont->put_Size(size);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("设置字体大小失败"));
			message.exec();
			return;
		}
	}
	m_mainArea->setFocus();
}

void EtMainWindow::AddSheet()
{
	if (m_spApplication)
	{
		ks_stdptr<etapi::Sheets> spSheets;
		HRESULT hr = m_spApplication->get_Sheets(&spSheets);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("获取Sheets对象失败"));
			message.exec();
			return;
		}

		long count = 0;
		hr = spSheets->get_Count(&count);

		KComVariant varCount(count);
		ks_stdptr<etapi::_Worksheet> spWorksheet;
		hr = spSheets->get_Item(varCount, (IDispatch**)&spWorksheet);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("获取Worksheet对象失败"));
			message.exec();
			return;
		}

		KComVariant Before;
		KComVariant After(spWorksheet);
		KComVariant Type((long)etapi::xlWorksheet);
		KComVariant newSheetCount(1);
		ks_stdptr<IDispatch> spDsp;
		hr = spSheets->Add(Before, After, newSheetCount, Type, 0, (IDispatch**) &spDsp);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("新建工作表失败"));
			message.exec();
			return;
		}
	}
	m_mainArea->setFocus();
}

void EtMainWindow::AddComment()
{
	if (m_spApplication)
	{
		ks_stdptr<etapi::IRange> spRange;
		HRESULT hr = m_spApplication->get_ActiveCell((etapi::Range**)&spRange);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("获取Range对象失败"));
			message.exec();
			return;
		}

		KComVariant text(__X("新建批注"));
		ks_stdptr<etapi::Comment> spFont;
		hr = spRange->AddComment(text, &spFont);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("新建批注失败"));
			message.exec();
			return;
		}
	}
	m_mainArea->setFocus();
}

void EtMainWindow::PrintArea()
{
	if (m_spApplication)
	{
		ks_stdptr<etapi::_Workbook> spWorkbook;
		HRESULT hr = m_spApplication->get_ActiveWorkbook((etapi::Workbook**)&spWorkbook);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("获取Workbook对象失败"));
			message.exec();
			return;
		}
		ks_stdptr<etapi::_Worksheet> spWorksheet;
		hr = spWorkbook->get_ActiveSheet((IDispatch**)&spWorksheet);

		KComVariant cell1(__X("A1"));
		KComVariant cell2(__X("G12"));
		ks_stdptr<etapi::IRange> spRange;
		hr = spWorksheet->get_Range(cell1, cell2, (etapi::Range**)&spRange);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("获取Range对象失败"));
			message.exec();
			return;
		}

		KComVariant rhs;
		hr = spRange->Select(&rhs);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("选中对象失败"));
			message.exec();
			return;
		}

		ks_stdptr<etapi::IPageSetup> spPageSetup;
		hr = spWorksheet->get_PageSetup((etapi::PageSetup**)&spPageSetup);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("获取PageSetup对象失败"));
			message.exec();
			return;
		}

		ks_bstr area(__X("$A$1:$G$12"));
		hr = spPageSetup->put_PrintArea(area);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("设置打印区域失败"));
			message.exec();
			return;
		}
	}
	m_mainArea->setFocus();
}

void EtMainWindow::InsertPic()
{
	if (m_spApplication)
	{
		ks_stdptr<etapi::_Workbook> spWorkbook;
		HRESULT hr = m_spApplication->get_ActiveWorkbook((etapi::Workbook**)&spWorkbook);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("获取Workbook对象失败"));
			message.exec();
			return;
		}

		ks_stdptr<etapi::_Worksheet> spWorksheet;
		hr = spWorkbook->get_ActiveSheet((IDispatch**)&spWorksheet);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("获取Worksheet对象失败"));
			message.exec();
			return;
		}

		ks_stdptr<etapi::IShapes> spShapes;
		hr = spWorksheet->get_Shapes((etapi::Shapes**)&spShapes);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("获取Shapes对象失败"));
			message.exec();
			return;
		}

		QString path = QFileDialog::getOpenFileName(this, QString::fromUtf8("选择一个文件"));
		ks_bstr Filename(path.utf16());
		ks_stdptr<etapi::Shape> spShape;
		MsoTriState LinkToFile = etapi::msoCTrue;
		MsoTriState SaveWithDocument = etapi::msoCTrue;
		single Left = 0;
		single Top = 0;
		single Width = 100;
		single Height = 100;
		hr = spShapes->AddPicture(Filename, LinkToFile, SaveWithDocument, Left, Top, Width, Height, &spShape);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("插入图片失败"));
			message.exec();
			return;
		}
	}
	m_mainArea->setFocus();
}

void EtMainWindow::InsertHistogram()
{
	if (m_spApplication)
	{
		ks_stdptr<etapi::_Worksheet> spWorksheet;
		HRESULT hr = m_spApplication->get_ActiveSheet((IDispatch**)&spWorksheet);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("获取Worksheet对象失败"));
			message.exec();
			return;
		}

		KComVariant defCell;
		KComVariant varO3(__X("O3"));
		ks_stdptr<etapi::IRange>spRangeO3;
		hr = spWorksheet->get_Range(varO3, defCell, (etapi::Range**)&spRangeO3);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("获取Range对象失败"));
			message.exec();
			return;
		}

		KComVariant o3Val(__X("Item1"));
		KComVariant RangeValueDataType(etapi::xlRangeValueDefault);
		hr = spRangeO3->put_Value(RangeValueDataType, 0, o3Val);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("O3单元格设置值失败"));
			message.exec();
			return;
		}


		KComVariant varN4(__X("N4"));
		ks_stdptr<etapi::IRange>spRangeN4;
		hr = spWorksheet->get_Range(varN4, defCell, (etapi::Range**)&spRangeN4);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("获取Range失败"));
			message.exec();
			return;
		}
		KComVariant n4Val(3);
		hr = spRangeN4->put_Value(RangeValueDataType, 0, n4Val);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("N4单元格设置值失败"));
			message.exec();
			return;
		}

		KComVariant varN3(__X("N3"));
		ks_stdptr<etapi::IRange>spRangeN3;
		hr = spWorksheet->get_Range(varN3, defCell, (etapi::Range**)&spRangeN3);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("获取Range对象失败"));
			message.exec();
			return;
		}

		KComVariant n3Val(__X("Item2"));
		hr = spRangeN3->put_Value(RangeValueDataType, 0, n3Val);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("N3单元格设置值失败"));
			message.exec();
			return;
		}

		KComVariant varO4(__X("O4"));
		ks_stdptr<etapi::IRange>spRangeO4;
		hr = spWorksheet->get_Range(varO4, defCell, (etapi::Range**)&spRangeO4);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("获取Range对象失败"));
			message.exec();
			return;
		}

		KComVariant O4Val(4);
		hr = spRangeO4->put_Value(RangeValueDataType, 0, O4Val);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("O4单元格设置值失败"));
			message.exec();
			return;
		}
		KComVariant varN3O4(__X("N3:O4"));
		ks_stdptr<etapi::IRange>spRangeN3O4;
		hr = spWorksheet->get_Range(varN3O4, defCell, (etapi::Range**)&spRangeN3O4);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("获取Range对象失败"));
			message.exec();
			return;
		}

		KComVariant selectRhs;
		hr = spRangeN3O4->Select(&selectRhs);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("选择对象失败"));
			message.exec();
			return;
		}
		ks_stdptr<etapi::IShapes> spShapes;
		spWorksheet->get_Shapes((etapi::Shapes**)&spShapes);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("获取Shapes对象失败"));
			message.exec();
			return;
		}
		KComVariant Top(0);
		KComVariant Left(0);
		KComVariant Width(480);
		KComVariant Height(320);
		KComVariant XlChartType(51);
		ks_stdptr<etapi::IShape> spShape;
		hr = spShapes->AddChart(XlChartType, Left, Top, Width, Height, (etapi::Shape**)&spShape);
		if (FAILED(hr))
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("获取Spape对象失败"));
			message.exec();
			return;
		}
	}
	m_mainArea->setFocus();
}

void EtMainWindow::setRangeValue()
{
	if (m_spApplication)
	{
		ks_stdptr<_Worksheet> spWorkSheet;
		m_spApplication->get_ActiveSheet((IDispatch **)&spWorkSheet);
		if (!spWorkSheet)
			return;

		ks_stdptr<IRange> spCellsRange;
		KComVariant cell1(__X("A1"));
		KComVariant cell2(__X("D10"));
		//获得单元格A1:D10的区域
		HRESULT hr = spWorkSheet->get_Range(cell1, cell2,(Range**) &spCellsRange);
		if (hr != S_OK || !spCellsRange)
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("获取A1:D10失败"));
			message.exec();
			return;
		}
		KComVariant varRangeValueDataType(xlRangeValueDefault);
		KComVariant varValue(__X("kingsoft"));
		spCellsRange->put_Value(varRangeValueDataType, ET_LCID, varValue);
	}
	m_mainArea->setFocus();
}

void EtMainWindow::rangeMerge()
{
	if (m_spApplication)
	{
		static bool isMerge = true;
		ks_stdptr<_Worksheet> spWorkSheet;
		m_spApplication->get_ActiveSheet((IDispatch **)&spWorkSheet);
		if (!spWorkSheet)
			return;

		ks_stdptr<IRange> spCellsRange;
		KComVariant cell1(__X("A1"));
		KComVariant cell2(__X("D10"));
		//获得单元格A1:D10的区域
		HRESULT hr = spWorkSheet->get_Range(cell1, cell2,(Range**) &spCellsRange);
		if (hr != S_OK || !spCellsRange)
		{
			QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("获取A1:D10失败"));
			message.exec();
			return;
		}
		if (isMerge)
		{
			KComVariant varAcross;
			varAcross.AssignBOOL(TRUE);
			spCellsRange->Merge(varAcross);
		}
		else
		{
			spCellsRange->UnMerge();
		}
		isMerge = !isMerge;
	}
	m_mainArea->setFocus();
}

void EtMainWindow::setHeigthWidth()
{
	if (m_spApplication)
	{
		ks_stdptr<IRange> spSelection;
		m_spApplication->get_Selection(ET_LCID, (IDispatch **)&spSelection);
		if (spSelection)
		{
			KComVariant varNum(20);
			spSelection->put_RowHeight(varNum);
			spSelection->put_ColumnWidth(varNum);
		}
	}
	m_mainArea->setFocus();
}

static HRESULT wpsWorkbookBeforeCloseCallback(_Workbook* pWorkbook, VARIANT_BOOL* Cancel)
{
	*Cancel = VARIANT_FALSE;
	QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("关闭回调"));
	message.exec();
	return S_OK;
}

void EtMainWindow::registerCloseEvent()
{
	if (m_spApplication)
	{
		HRESULT ret = m_rpcClient->registerEvent(m_spApplication, DIID_AppEvents, __X("WorkbookBeforeClose"), (void*)wpsWorkbookBeforeCloseCallback);
		QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), "");
		if (ret == S_OK)
			message.setText(QString::fromUtf8("注册成功"));
		else
			message.setText(QString::fromUtf8("注册失败"));
		message.exec();
	}
	m_mainArea->setFocus();
}

static HRESULT wpsWorkbookBeforeSaveCallback(_Workbook* pdoc, VARIANT_BOOL* SaveAsUI, VARIANT_BOOL* Cancel)
{
	QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), QString::fromUtf8("保存回调"));
	message.exec();
	return S_OK;
}

void EtMainWindow::registerSaveEvent()
{
	if (m_spApplication)
	{
		HRESULT ret = m_rpcClient->registerEvent(m_spApplication, DIID_AppEvents, __X("WorkbookBeforeSave"), (void*)wpsWorkbookBeforeSaveCallback);
		QMessageBox message(QMessageBox::NoIcon, QString::fromUtf8("提示"), "");
		if (ret == S_OK)
			message.setText(QString::fromUtf8("注册成功"));
		else
			message.setText(QString::fromUtf8("注册失败"));
		message.exec();
	}
	m_mainArea->setFocus();
}

void EtMainWindow::slotButtonClick(const QString&  name)
{
	typedef void (EtMainWindow::*EtOperationFun)(void);
	static QMap<QString, EtOperationFun> s_operationFunMap;
	if (s_operationFunMap.isEmpty())
	{
		s_operationFunMap.insert(QString::fromUtf8("初始化"), &EtMainWindow::initApp);
		s_operationFunMap.insert(QString::fromUtf8("新建文档"), &EtMainWindow::newDoc);
		s_operationFunMap.insert(QString::fromUtf8("打开文档"), &EtMainWindow::openDoc);
		s_operationFunMap.insert(QString::fromUtf8("另存文档"), &EtMainWindow::saveAs);
		s_operationFunMap.insert(QString::fromUtf8("隐藏工具栏"), &EtMainWindow::hideToolBar);
		s_operationFunMap.insert(QString::fromUtf8("隐藏按钮"), &EtMainWindow::hideToolBtn);
		s_operationFunMap.insert(QString::fromUtf8("是否已保存"), &EtMainWindow::getSaved);
		s_operationFunMap.insert(QString::fromUtf8("文档保护"), &EtMainWindow::protect);
		s_operationFunMap.insert(QString::fromUtf8("取消保护"), &EtMainWindow::unProtect);
		s_operationFunMap.insert(QString::fromUtf8("注册关闭事件"), &EtMainWindow::registerCloseEvent);
		s_operationFunMap.insert(QString::fromUtf8("注册保存事件"), &EtMainWindow::registerSaveEvent);
		s_operationFunMap.insert(QString::fromUtf8("关闭ET"), &EtMainWindow::closeApp);
		s_operationFunMap.insert(QString::fromUtf8("单元格设值"), &EtMainWindow::setRangeValue);
		s_operationFunMap.insert(QString::fromUtf8("合并取消单元格"), &EtMainWindow::rangeMerge);
		s_operationFunMap.insert(QString::fromUtf8("设置行高列宽"), &EtMainWindow::setHeigthWidth);
		s_operationFunMap.insert(QString::fromUtf8("设置字号"), &EtMainWindow::FontSize);
		s_operationFunMap.insert(QString::fromUtf8("添加工作表"), &EtMainWindow::AddSheet);
		s_operationFunMap.insert(QString::fromUtf8("新建批注"), &EtMainWindow::AddComment);
		s_operationFunMap.insert(QString::fromUtf8("设置打印区域"), &EtMainWindow::PrintArea);
		s_operationFunMap.insert(QString::fromUtf8("插入图片"), &EtMainWindow::InsertPic);
		s_operationFunMap.insert(QString::fromUtf8("插入图表"), &EtMainWindow::InsertHistogram);
	}
	if (s_operationFunMap.contains(name))
		(this->*s_operationFunMap.value(name))();
}

