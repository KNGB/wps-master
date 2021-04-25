/** 
 * @file scenes04.js
 * @desc 只读打开服务器文件，设置保护等属性
 * @author wps-developer
 */

/**
 * @desc 定义插件菜单枚举（仅仅在windows平台才会使用）
 */
var menuItems = {
	FILE: 1 << 0,
	EDIT: 1 << 1,
	VIEW: 1 << 2,
	INSERT: 1 << 3,
	FORMAT: 1 << 4,
	TOOL: 1 << 5,
	CHART: 1 << 6,
	HELP: 1 << 7
};
var FileSubmenuItems = {
	NEW: 1 << 0,
	OPEN: 1 << 1,
	CLOSE: 1 << 2,
	SAVE: 1 << 3,
	SAVEAS: 1 << 4,
	PAGESETUP: 1 << 5,
	PRINT: 1 << 6,
	PROPERTY: 1 << 7
};

/**
 * @desc 隐藏一些常用的按钮
 * @param { bool } visible true显示按钮，false隐藏按钮
 */
function setVisibleStandardButton(visible) {
	if (isWindowsPlatform()) {
		setVisibleFileSubMenuItem(FileSubmenuItems.NEW, visible);
		setVisibleFileSubMenuItem(FileSubmenuItems.OPEN, visible);
		setVisibleFileSubMenuItem(FileSubmenuItems.CLOSE, visible);
		setVisibleFileSubMenuItem(FileSubmenuItems.SAVE, visible);
		setVisibleFileSubMenuItem(FileSubmenuItems.SAVEAS, visible);
		setVisibleFileSubMenuItem(FileSubmenuItems.PAGESETUP, visible);
		setVisibleFileSubMenuItem(FileSubmenuItems.PRINT, visible);
		setVisibleFileSubMenuItem(FileSubmenuItems.PROPERTY, visible);
	} else {
		setVisibleToolButton(visible);
	}
}

/**
 * @desc 隐藏工具按钮（windows系统调用）
 * @param { Enum } item 菜单子项 
 * @param { bool } visible true显示，false隐藏
 */
function setVisibleFileSubMenuItem(item, visible) {
	try {
		if (visible) {
			wpsFrame.pluginObj.FileSubmenuItems |= item;
		} else {
			wpsFrame.pluginObj.FileSubmenuItems &= ~item;
		}
	} catch (error) {
		console.log(error.message);
	}
}

/**
 * @desc 隐藏工具按钮（linux系统调用）
 * @param { bool } visible 
 */
function setVisibleToolButton(visible) {
	try {
		var app, commandBars;

		app = wpsFrame.Application;
		commandBars = app.get_CommandBars();
		// 隐藏选项卡，其他选项卡都可参照该写法
		commandBars.get_Item('Menu Bar').Controls.get_Item('文件(&F)').Visible = visible;

		// 隐藏选项卡下的某个选项，名称都是对应中文的英文例：文件（File）、编辑（Edit）、视图（View）等等
		commandBars.get_Item('File').Controls.get_Item('新建(&N)...').Visible = visible;
		commandBars.get_Item('Edit').Controls.get_Item('查找(&F)...').Visible = visible;

		// 隐藏其他工具栏按钮
		commandBars.get_Item('Standard').Controls.get_Item('新建(&N)...').Visible = visible;
		commandBars.get_Item('Standard').Controls.get_Item('打开(&O)...').Visible = visible;
	} catch (error) {
		console.log(error.message);
	}
}