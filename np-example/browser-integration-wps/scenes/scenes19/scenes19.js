/** 
 * @file scenes19.js
 * @desc 视图操作
 * @author wps-developer
 */

 /**
  * @description 大纲视图
  */
function OutlineView(){
    try {
        var app = wpsFrame.Application;
        app.ActiveWindow.ActivePane.View.Type = 2;// WdViewType.wdOutlineView
       } catch (error) {
        alert(error.message);
    }
}

 /**
  * @description 显示标尺
  */
 function DisplayRulers(){
    try {
        var app = wpsFrame.Application;
        app.ActiveWindow.DisplayRulers = true;
       } catch (error) {
        alert(error.message);
    }
}

 /**
  * @description 打开文档结构图
  */
 function DocumentMap(){
    try {
        var app = wpsFrame.Application;
        app.ActiveWindow.DocumentMap = true;
       } catch (error) {
        alert(error.message);
    }
}
