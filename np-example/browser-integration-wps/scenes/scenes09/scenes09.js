/** 
 * @file scenes05.js
 * @desc 插入图片相关处理
 * @author wps-developer
 */

 
 /**
 * @desc 打开本地或者服务器上的文档
 */
function openDocumentFromSever() {
	var fileURL = getPathByFileName('default.wps', 'file');
	if (checkLocalMode()) {
		wpsFrame.openDocumentByLocal(fileURL, false);
	} else {
		wpsFrame.openDocumentBySever(fileURL, false);
	}
}

/**
 * 初始化代码
 */
function init() {
	wpsFrame.initPlugin('wps', openDocumentFromSever);
}

/**
 * @desc 注册WPS内置事件回调
 */
function registerEventListener() {
	if (isWindowsPlatform()) {
		alert('windows平台注册方式不同，请咨询技术支持！');
		return;
	}
	var pluginObj = wpsFrame.pluginObj;

	pluginObj.ApplicationEx.registerEvent("DIID_WpsApplicationEventsEx", "DocumentBeforeNew", "DocumentBeforeNewCallBack");
	pluginObj.registerEvent("DIID_ApplicationEvents4", "NewDocument", "NewDocumentCallBack");
	pluginObj.ApplicationEx.registerEvent("DIID_WpsApplicationEventsEx", "DocumentBeforeOpen", "DocumentBeforeOpenCallBack");
	pluginObj.registerEvent("DIID_ApplicationEvents4", "DocumentOpen", "DocumentOpenCallBack");
	pluginObj.registerEvent("DIID_ApplicationEvents4", "DocumentBeforeSave", "DocumentSaveAndSaveAsCallBack");
	pluginObj.registerEvent("DIID_ApplicationEvents4", "DocumentBeforePrint", "DocumentBeforePrintCallBack");
	pluginObj.ApplicationEx.registerEvent("DIID_WpsApplicationEventsEx", "DocumentViewFocusIn", "DocumentViewFocusInCallBack");
	pluginObj.registerEvent("DIID_ApplicationEvents4", "DocumentBeforeClose", "DocumentBeforeCloseCallBack");
	pluginObj.ApplicationEx.registerEvent("DIID_WpsApplicationEventsEx", "DocumentAfterSave", "DocumentSaveAsOrSaveDoneCallBack");
	pluginObj.registerEvent("DIID_ApplicationEvents4", "ProtectedViewWindowActivate", "DocumentAfterCloseCallBack");
	pluginObj.ApplicationEx.registerEvent("DIID_WpsApplicationEventsEx", "DocumentViewFocusOut", "DocumentViewFocusOutCallBack");
	pluginObj.ApplicationEx.registerEvent("DIID_WpsApplicationEventsEx", "ContentChange", "DocumentContentChangeCallBack");
}

/**
 * @desc 取消注册WPS内置事件回调
 */
function unRegisterEventListener() {
	if (isWindowsPlatform()) {
		alert('windows平台注册方式不同，请咨询技术支持！');
		return;
	}
	var pluginObj = wpsFrame.pluginObj;

	pluginObj.ApplicationEx.unRegisterEvent("DIID_WpsApplicationEventsEx", "DocumentBeforeNew", "DocumentBeforeNewCallBack");
	pluginObj.unRegisterEvent("DIID_ApplicationEvents4", "NewDocument", "NewDocumentCallBack");
	pluginObj.ApplicationEx.unRegisterEvent("DIID_WpsApplicationEventsEx", "DocumentBeforeOpen", "DocumentBeforeOpenCallBack");
	pluginObj.unRegisterEvent("DIID_ApplicationEvents4", "DocumentOpen", "DocumentOpenCallBack");
	pluginObj.unRegisterEvent("DIID_ApplicationEvents4", "DocumentBeforeSave", "DocumentSaveAndSaveAsCallBack");
	pluginObj.unRegisterEvent("DIID_ApplicationEvents4", "DocumentBeforePrint", "DocumentBeforePrintCallBack");
	pluginObj.ApplicationEx.unRegisterEvent("DIID_WpsApplicationEventsEx", "DocumentViewFocusIn", "DocumentViewFocusInCallBack");
	pluginObj.unRegisterEvent("DIID_ApplicationEvents4", "DocumentBeforeClose", "DocumentBeforeCloseCallBack");
	pluginObj.ApplicationEx.unRegisterEvent("DIID_WpsApplicationEventsEx", "DocumentAfterSave", "DocumentSaveAsOrSaveDoneCallBack");
	pluginObj.unRegisterEvent("DIID_ApplicationEvents4", "ProtectedViewWindowActivate", "DocumentAfterCloseCallBack");
	pluginObj.ApplicationEx.unRegisterEvent("DIID_WpsApplicationEventsEx", "DocumentViewFocusOut", "DocumentViewFocusOutCallBack");
	pluginObj.ApplicationEx.unRegisterEvent("DIID_WpsApplicationEventsEx", "ContentChange", "DocumentContentChangeCallBack");
}

/**
 * @desc 新建文档之前事件回调
 * @param { Object } cancel	该对象是WPS插件对象，可以控制是否取消新建操作，设置方式为
 * 						cancle.SetValue(true)取消新建，cancle.SetValue(false)执行正常操作
 */
function DocumentBeforeNewCallBack(cancel) {
	console.log('新文档开始创建');
}

/**
 * @desc 新建文档完成之后事件回调
 * @param { Object } doc 该对象是WPS的Document对象，可以操控文档的一些设置
 */
function NewDocumentCallBack(doc) {
	console.log('新文档已被创建');
}

/**
 * @desc 执行打印操作回调事件（打印之前的事件）
 * @param { Object } doc 该对象是WPS的Document对象，可以操控文档的一些设置
 * @param { Object } cancel 该对象是WPS插件对象，可以控制是否取消打印操作，设置方式为
 * 				cancle.SetValue(true)取消打印，cancle.SetValue(false)执行正常打印
 */
function DocumentBeforePrintCallBack(doc, cancel) {
	console.log('即将打印文档');
}

/**
 * @desc 窗口激活事件回调（文档显示发生变化触发(新建、打开等等)，或者切换文档显示的时候触发，比如有隐藏文档和当前显示的文档进行切换的时候）
 * @param { Object } doc 该对象是WPS的Document对象，可以操控文档的一些设置
 * @param { Object } window 该对象是WPS的Window对象，可以操控文档窗口的一些设置
 */
function WindowActivateCallBack(doc, window) {
	console.log('文档已被激活');
}

/**
 * @desc 文档关闭之前事件回调
 * @param { Object } doc 该对象是WPS的Document对象，可以操控文档的一些设置
 * @param { Object } cancel 该对象是WPS插件对象，可以控制是否取消关闭文档操作，设置方式为
 *						cancle.SetValue(true)取消关闭操作，cancle.SetValue(false)执行正常关闭文档
 */
function DocumentBeforeCloseCallBack(doc, cancel) {
	console.log('文档将要被关闭');
}

/**
 * @desc 文档关闭之后事件回调
 * @param { Object } doc 该对象是WPS的Document对象，可以操控文档的一些设置
 */
function DocumentAfterCloseCallBack(doc) {
	console.log('文档已被关闭');
}

/**
 * @desc 文档保存或者另存为之前事件回调，通过参数区分是保存还是另存为
 * @param { Object } doc 该对象是WPS的Document对象，可以操控文档的一些设置
 * @param { Object } saveAsUI 该对象是WPS插件对象，可以判断当前是另存为操作还是保存操作
 * 						saveAsUI.GetValue()返回一个bool值，如果为true则证明是另存为操作，false为保存操作
 * @param { Object } Cancel 该对象是WPS插件对象，可以控制是否取消关闭文档操作，设置方式为
 * 					 	cancle.SetValue(true)取消关闭操作，cancle.SetValue(false)执行正常关闭文档
 */
function DocumentSaveAndSaveAsCallBack(doc, saveAsUI, Cancel) {
	var ui = saveAsUI.GetValue();
	if (ui) {
		isSaveAs = true;
		console.log('将要另存为文档');
	} else {
		isSaveAs = false;
		console.log('将要保存文档');
	}
}

/**
 * @desc 文档保存或者另存为之后事件回调
 * @param { Object } doc 该对象是WPS的Document对象，可以操控文档的一些设置
 * @param { bool } result true表示另存为成功，false表示另存为失败
 * 
 * 注意：此函数没发判断是另存为还是保存，所以我们要注册保存之前事件，里面用一个变量标记是保存还是另存为(isSaveAs)
 */
function DocumentSaveAsOrSaveDoneCallBack(doc, result) {
	if (isSaveAs && result) {
		console.log('已用新文件名保存文档');
	} else {
		console.log('文档已保存');
	}
}

/**
 * @desc 打开文档之前事件回调
 * @param { Object } cancel 该对象是WPS插件对象，可以控制是否取消打开文档操作
 * 						cancle.SetValue(true)取消打开文档操作，cancle.SetValue(false)执行正常打开文档操作
 */
function DocumentBeforeOpenCallBack(cancel) {
	console.log('新文档开始装入');
}

/**
 * @desc 打开文档完成事件回调
 * @param { Object } doc 该对象是WPS的Document对象，可以操控文档的一些设置
 */
function DocumentOpenCallBack(doc) {
	console.log('新文档已被装入');
}

/**
 * @desc 文档编辑区域获得了焦点事件回调
 * @param { Object } doc 该对象是WPS的Document对象，可以操控文档的一些设置
 */
function DocumentViewFocusInCallBack(doc) {
	console.log(doc.Name + "得到焦点");
}

/**
 * @desc 文档编辑区域失去焦点事件回调
 * @param { Object } doc 该对象是WPS的Document对象，可以操控文档的一些设置
 */
function DocumentViewFocusOutCallBack(doc) {
	console.log(doc.Name + "失去焦点");
}

/**
 * @desc 文档内容发生改变事件回调
 * @param { Object } doc 该对象是WPS的Document对象，可以操控文档的一些设置
 * @param { Object } range 该对象是WPS的Range对象，可以改变区域内容，该区域指定的是发生改变的区域
 * @param { Number } type	0 插入文本
							1 删除文本
							2 style改变
							4 段属性改变
							5 插入批注
							6 插入图片
							7 插入水印
							8 插入文本框
							9 插入文件
							10 插入表格
 */
function DocumentContentChangeCallBack(doc, range, type) {
	console.log(doc.Name + "发生改变，改变内容类型：" + type + "，改变内容：" + range.Text);
}