/** 
 * @file scenes05.js
 * @desc 处理公文域的一些逻辑
 * @author wps-developer
 */


/**
 * @desc 打开服务器上带公文域文档
 */
function openServerDocWithField() {
	var fileURL = getPathByFileName('letter.wps', 'file');
    if (checkLocalMode()) {
        wpsFrame.openDocumentByLocal(fileURL, false);
    } else {
        wpsFrame.openDocumentBySever(fileURL, false);
    }
}

/**
 * @desc 获取正文公文域内容
 */
function getMainDocFieldText() {
	try {
		var app, field;

		app = wpsFrame.Application;
		field = app.ActiveDocument.DocumentFields.Item('正文');
		alert(field.Range.Text);
	} catch (error) {
		
	}
}

/**
 * @desc 替换正文公文域内容
 */
function replaceMainDocFieldText() {
	try {
		var app, field, insertText;

		insertText = prompt("请输入要替换的正文公文域内容", "你好呀！我是新的公文域内容");
		app = wpsFrame.Application;
		field = app.ActiveDocument.DocumentFields.Item('正文');
		field.Range.Text = insertText;
		if (!ret) {
			alert('替换公文域内容失败！');
		}	
	} catch (error) {
		console.log(error.message);
	}
}

/**
 * @desc 删除正文公文域（注意：只是删除公文域，不影响文档的文字内容）
 */
function deleteMainDocField() {
	try {
		var app, field;

		app = wpsFrame.Application;
		field = app.ActiveDocument.DocumentFields.Item('正文');
		field.Delete();
	} catch (error) {
		console.log(error.message);
	}
}

/**
 * @desc 插入一个新的公文域(注意：是在当选光标选中的位置插入)
 */
function insertTestField() {
	try {
		var app, fields, insertFieldName;

		insertFieldName = prompt("请输入要插入的公文域名称", "新公文域");
		app = wpsFrame.Application;
		fields = app.ActiveDocument.DocumentFields;
		fields.Add(insertFieldName, app.Selection.Range);
	} catch (error) {
		console.log(error.message);
	}
}

/**
 * @desc 切换公文域显示形式(0不显示，1.显示全部公文域标记，2.只显示公文域两边红线不显示文字红色底纹)
 */
function changeFieldStyle() {
	try {
		var setState;

		var fieldState = [0, 1, 2];
		var curState = wpsFrame.pluginObj.ShowDocumentFieldTarget;
		if (curState < fieldState.length - 1) {
			setState = fieldState[curState + 1];
		} else {
			setState = fieldState[0];
		}
		wpsFrame.pluginObj.ShowDocumentFieldTarget = setState;
	} catch (error) {
		console.log(error.message);
	}
}