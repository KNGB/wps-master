/** 
 * @file scenes02.js
 * @desc 打开服务器文件，设置修订模式相关的信息
 * @author wps-developer
 */

/**
 * @desc 打开服务器文档，并且进入修订模式
 */
function openSeverDocumentAndSetRevision() {
	var fileURL = getPathByFileName('default.wps', 'file');
    if (checkLocalMode()) {
        wpsFrame.openDocumentByLocal(fileURL, false);
    } else {
        wpsFrame.openDocumentBySever(fileURL, false);
    }
	wpsFrame.pluginObj.enableRevision(true);
}

/**
 * @desc 保存当前文档到服务器
 */
function saveFileForSever() {
    var uploadURL = prompt("请输入要上传的服务器地址");
    if (0 == uploadURL.length || null == uploadURL) {
        alert('服务器地址不能为空');
    } else {
        wpsFrame.saveFileToSever(uploadURL);
    }
}

/**
 * @desc 获取当前修订人姓名
 */
function getUserName() {
	document.getElementById('auth').value = wpsFrame.pluginObj.getUserName();
}

/**
 * @desc 设置修订人姓名
 */
function setUserName() {
	wpsFrame.pluginObj.setUserName(document.getElementById('auth').value);
}

/**
 * @desc 以带标记的原始修订，在文档中显示
 */
function showRevision0() {
	wpsFrame.pluginObj.showRevision(0);
}

/**
 * @desc 以原始修订，在文档中显示
 */
function showRevision1() {
	wpsFrame.pluginObj.showRevision(1);
}

/**
 * @desc 以最终修订，在文档中显示
 */
function showRevision2() {
	wpsFrame.pluginObj.showRevision(2);
}