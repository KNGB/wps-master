/** 
 * @file scenes15.js
 * @desc 文档比较
 * @author wps-developer
 */

 /**
  * @description 两个文档比较
  */
function compareDocs(){
    try {
        var app;
        app = wpsFrame.Application;
        //实现目标文档与当前活动文档的对比
        app.ActiveDocument.Compare(getPathByFileName('1.002.wps', 'file'));
        // app.ActiveDocument.Compare(getPathByFileName('1.002.wps', 'file'),"authorname",getPathByFileName('3.wps', 'file'),true,true,false,true,true);
        //方法说明：这是个在Document对象下的方法
        // Compare(Name, AuthorName, CompareTarget, DetectFormatChanges, IgnoreAllComparisonWarnings, AddToRecentFiles, RemovePersonalInformation, RemoveDateAndTime)
        // Name 必选 String 与指定文档作比较的文档的名称（可是网络文档，可是本地绝对路径）。
        // AuthorName 可选 Variant 与比较生成的区别相关联的审阅者姓名。 如果没有指定， 则默认值为修订的文档的作者姓名或字符串“ Comparison”（ 如果没有提供作者信息）。
        // CompareTarget 可选 Variant 用于比较的目标文档。 可以是任意 WdCompareTarget 常量。
        // DetectFormatChanges 可选 Boolean 如果该属性值为 True（ 默认值）， 则比较结果包括格式变化检测。
        // IgnoreAllComparisonWarnings 可选 Variant 如果该属性值为 True， 则比较文档， 而不通知用户问题。 默认值为 False。
        // AddToRecentFiles 可选 Variant 如果该属性值为 True， 则将文档添加到“ 文件” 菜单上最近使用的文件列表中。
        // RemovePersonalInformation 可选 Boolean 如果该属性值为 True， 则从返回的 Document 对象的备注、 修订和属性对话框中删除所有用户信息。 默认值为 False。
        // RemoveDateAndTime 可选 Boolean 如果该属性值为 True， 则从返回的 Document 对象的跟踪更改中删除日期和时间戳记。 默认值为 False。
    } catch (error) {
        alert(error.message);
    }
}

/**
 * @desc 打开原始版本的文档
 */
function openTemplateFile() {
    var fileURL = getPathByFileName('1.001.wps', 'file');
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
 * @desc 初始化插件代码
 */
function initCompareDocs() {
    wpsFrame.initPlugin('wps', openTemplateFile);
}