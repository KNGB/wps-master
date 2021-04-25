/** 
 * @file scenes05.js
 * @desc 插入图片相关处理
 * @author wps-developer
 */

 /**
  * @desc 判断当前文档是否被修改过，如果修改过则直接保存
  */
function saveToDirty() {
	var app, isSaved;
	
	app = wpsFrame.Application;
	var isSaved = app.ActiveDocument.Saved;
	if (isSaved) {
		alert('文档在上一次保存之后再无变动，无需执行保存操作。');
	} else {
		var uploadURL = prompt("请输入要上传的服务器地址");
		if (0 == uploadURL.length || null == uploadURL) {
			alert('服务器地址不能为空');
		} else {
			wpsFrame.saveFileToSever(uploadURL);
		}
	}
}