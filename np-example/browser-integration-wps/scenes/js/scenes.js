/** 
 * @file scenes.js
 * @desc 场景通用化js
 * @author wps-developer
 */

/**
 * @desc 判断环境是在服务器上还是在本地上
 * @returns { bool } true代表本地，false代表服务器
 */
function checkLocalMode() {
	return window.location.host.length < 1;
}

/**
 * @desc 返回文件全路径
 * @param { String } fileName 文件名
 * @param { Number } type 文件类型
 * @returns { String } 文件全路径
 */
function getPathByFileName(fileName, type) {
    var hostName;

    if (checkLocalMode()) {
        var index = window.location.pathname.lastIndexOf('browser-integration-wps');
        var start = 0;
        if (isWindowsPlatform()) {
            start = 1;
        }
        hostName = window.location.pathname.substr(start, index - 1);
    } else {
        hostName = window.location.href;
        var index = window.location.href.lastIndexOf('browser-integration-wps');
        hostName = hostName.substr(0, index - 1);
    }
    var realPath = hostName + '/' + 'browser-integration-wps/' + type + "/";
    if (undefined != fileName) {
        realPath = realPath + fileName
    }

    realPath = decodeURI(realPath);

    return realPath;
}

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
    
    if (isWindowsPlatform()) {
        wpsFrame.Application = wpsFrame.pluginObj.Application
    }
}

/**
 * @desc 新建一个文档
 */
function newDocument() {
    wpsFrame.pluginObj.createDocument('wps');
    if (isWindowsPlatform()) {
        wpsFrame.Application = wpsFrame.pluginObj.Application
    }
}

/**
 * 初始化代码
 */
function init(isNewDocument) {
    if (undefined == isNewDocument || false == isNewDocument) {
        wpsFrame.initPlugin('wps', openDocumentFromSever);
    } else {
        wpsFrame.initPlugin('wps', newDocument);
    }
}


