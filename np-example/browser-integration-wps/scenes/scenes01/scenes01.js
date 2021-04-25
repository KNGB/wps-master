/** 
 * @file scenes01.js
 * @desc 打开和保存服务器文件相关函数
 * @author wps-developer
 */

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