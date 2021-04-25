/** 
 * @file scenes02.js
 * @desc 只读打开服务器文件，设置保护等属性
 * @author wps-developer
 */


 /**
  * @desc 只读打开服务器文档
  */
function readOnlyOpenDocumentFromServer() {
	var fileURL = getPathByFileName('default.wps', 'file');
    if (checkLocalMode()) {
        wpsFrame.openDocumentByLocal(fileURL, true);
    } else {
        wpsFrame.openDocumentBySever(fileURL, true);
    }
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
 * @desc 设置或取消保护文档
 * @param { bool } isProtect 
 */
function enableProtect(isProtect) {
	try {
		var ret = wpsFrame.pluginObj.enableProtect(isProtect);
		if (ret) {
			if (isProtect) {
				alert('设置文档保护成功');
			} else {
				alert('取消文档保护成功');
			}
		}
	} catch (error) {
		console.log(error.message);
	}
}

/**
 * @desc 禁用或启动拷贝功能
 * @param { bool } bit 
 */
function enableCopy(isEnable) {
	try {
		var ret = wpsFrame.pluginObj.enableCopy(isEnable);
		if (ret) {
			if (isEnable) {
				alert('启动拷贝功能成功');
			} else {
				alert('禁用拷贝成功');
			}
		}	
	} catch (error) {
		console.log(error.message);
	}
}

/**
 * @desc 启用或禁用剪切功能
 * @param { bool } isEnable 
 */
function enableCut(isEnable) {
	try {
		var ret = wpsFrame.pluginObj.enableCut(isEnable);
		if (ret) {
			if (isEnable) {
				alert('启用剪切功能成功');
			} else {
				alert('禁用剪切功能成功');
			}
		}
	} catch (error) {
		console.log(error.message);
	}
}