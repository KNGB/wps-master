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
#include "buttonlistwnd.h"
#include <QStringList>

ButtonListWnd::ButtonListWnd(QWidget *parent)
	: QTreeWidget(parent)
{
}

ButtonListWnd::~ButtonListWnd()
{
}

void ButtonListWnd::initButtonTree(int type)
{
	clear();
	switch (type)
	{
	case appWps:
			initWpsButtonTree();
			break;
	case appEt:
			initEtButtonTree();
			break;
	case appWpp:
			initWppButtonTree();
			break;
	default:
			break;
	}
	//默认展开第一行
	itemAt(0, 0)->setExpanded(true);
}

void ButtonListWnd::initWpsButtonTree()
{
	QStringList items;
	items << QString::fromUtf8("常用操作") << QString::fromUtf8("初始化")\
				<< QString::fromUtf8("新建文档") << QString::fromUtf8("打开文档")\
				<< QString::fromUtf8("另存文档") << QString::fromUtf8("关闭文档")\
				<< QString::fromUtf8("打印文档") << QString::fromUtf8("添加编号")\
				<< QString::fromUtf8("插入超链接") << QString::fromUtf8("插入表格")\
				<< QString::fromUtf8("设置字体大小") << QString::fromUtf8("插入椭圆")\
				<< QString::fromUtf8("突出显示");
	addRootItem(items);
	items.clear();
	items << QString::fromUtf8("隐藏菜单按钮") << QString::fromUtf8("隐藏菜单")\
				<< QString::fromUtf8("隐藏工具栏") << QString::fromUtf8("隐藏按钮");
	addRootItem(items);
	items.clear();
	items << QString::fromUtf8("文档保护") << QString::fromUtf8("保护")\
				<< QString::fromUtf8("取消保护") << QString::fromUtf8("是否保存");
	addRootItem(items);
	items.clear();
	items << QString::fromUtf8("事件注册") << QString::fromUtf8("注册关闭事件")\
				<< QString::fromUtf8("注册保存事件");
	addRootItem(items);
	items.clear();
	items << QString::fromUtf8("其他") << QString::fromUtf8("关闭WPS")\
				<< QString::fromUtf8("设置页眉页脚") << QString::fromUtf8("获取页数")\
				<< QString::fromUtf8("设置自动备份");
	addRootItem(items);
}

void ButtonListWnd::initEtButtonTree()
{
	QStringList items;
	items << QString::fromUtf8("常用操作") << QString::fromUtf8("初始化")\
				<< QString::fromUtf8("新建文档") << QString::fromUtf8("打开文档")\
				<< QString::fromUtf8("另存文档") <<QString::fromUtf8("设置字号")\
				<< QString::fromUtf8("添加工作表");
	addRootItem(items);
	items.clear();
	items << QString::fromUtf8("隐藏菜单按钮") << QString::fromUtf8("隐藏工具栏")\
				<< QString::fromUtf8("隐藏按钮");
	addRootItem(items);
	items.clear();
	items << QString::fromUtf8("文档保护") << QString::fromUtf8("文档保护") << QString::fromUtf8("取消保护") << QString::fromUtf8("是否已保存");
	addRootItem(items);
	items.clear();
	items << QString::fromUtf8("单元格设置") << QString::fromUtf8("单元格设值")\
				<< QString::fromUtf8("合并取消单元格") << QString::fromUtf8("设置行高列宽")\
				<< QString::fromUtf8("新建批注") << QString::fromUtf8("设置打印区域");
	addRootItem(items);
	items.clear();
	items << QString::fromUtf8("图形操作") << QString::fromUtf8("插入图片")\
				<< QString::fromUtf8("插入图表");
	addRootItem(items);
	items.clear();
	items << QString::fromUtf8("事件注册") << QString::fromUtf8("注册关闭事件") \
				<< QString::fromUtf8("注册保存事件");
	addRootItem(items);
	items.clear();
	items << QString::fromUtf8("其他") << QString::fromUtf8("关闭ET");
	addRootItem(items);
}

void ButtonListWnd::initWppButtonTree()
{
	QStringList items;
	items << QString::fromUtf8("常用操作") << QString::fromUtf8("初始化")\
				<< QString::fromUtf8("新建文档") << QString::fromUtf8("打开文档")\
				<< QString::fromUtf8("另存文档") << QString::fromUtf8("打印到pdf");
	addRootItem(items);
	items.clear();
	items << QString::fromUtf8("隐藏菜单按钮") << QString::fromUtf8("隐藏工具栏")\
				<< QString::fromUtf8("隐藏按钮");
	addRootItem(items);
	items.clear();
	items << QString::fromUtf8("事件注册") << QString::fromUtf8("注册关闭事件")\
				<< QString::fromUtf8("注册保存事件");
	addRootItem(items);
	items.clear();
	items << QString::fromUtf8("其他") << QString::fromUtf8("关闭WPP")\
				<< QString::fromUtf8("插入图片") << QString::fromUtf8("添加幻灯片")\
				<< QString::fromUtf8("设置幻灯片大小") << QString::fromUtf8("插入文本框")\
				<< QString::fromUtf8("插入表格") << QString::fromUtf8("插入音频")\
				<< QString::fromUtf8("插入超链接") << QString::fromUtf8("新增节")\
				<< QString::fromUtf8("开始播放");
	addRootItem(items);
}

void ButtonListWnd::addRootItem(const QStringList& childItems)
{
	if (childItems.isEmpty())
		return;
	QTreeWidgetItem *root = new QTreeWidgetItem(this);
	root->setText(0, childItems.at(0));
	for (int i = 1; i < childItems.count(); i++)
	{
		QTreeWidgetItem *child = new QTreeWidgetItem(root);
		child->setText(0, childItems.at(i));
		child->addChild(child);
	}
	insertTopLevelItem(0, root);
}
