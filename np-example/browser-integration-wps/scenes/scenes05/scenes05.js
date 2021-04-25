/** 
 * @file scenes05.js
 * @desc 打开服务器文件，设置修订模式相关的信息
 * @author wps-developer
 */


/**
 * @desc 根据文件后缀名，来保存对应格式的文档到服务器
 */
function saveFileServerForTargetFormat() {
	var fileName = prompt("请输入保存到远程的文件名称和后缀名，支持的后缀名包括doc,docx,wps,uot,pdf,ofd等:", "scenes05.ofd");
	if (fileName != null && fileName != "") {
		var uploadURL = prompt("请输入要保存文件的服务器地址！");
		if (uploadURL == null || uploadURL == '') {
			alert('服务器地址不能为空！');
		} else {
			wpsFrame.saveFileToSever(uploadURL);
		}
	}
}