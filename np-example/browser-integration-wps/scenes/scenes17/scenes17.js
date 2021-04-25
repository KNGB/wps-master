/** 
 * @file scenes17.js
 * @desc 文档比较
 * @author wps-developer
 */

 /**
  * @description 滚动界面
  */
function VerticalScrolled(){
    try {
        var app = wpsFrame.Application;
        app.ActiveWindow.ActivePane.VerticalPercentScrolled = 50;
    } catch (error) {
        alert(error.message);
    }
}

 /**
  * @description 插入分页符
  */
 function InsertBreak(){
    try {
        var app;
        app = wpsFrame.Application;
        //分页符枚举 WdBreakType.wdPageBreak值为7
        app.Selection.InsertBreak(7);
    } catch (error) {
        alert(error.message);
    }
}

 /**
  * @description 插入目录
  */
 function InsertArchives(){
    try {
        var app = wpsFrame.Application;
        var selection = app.Selection;
        for (i = 1; i < 4; i++)
        {
            var title = i + "级标题";
            var lever = "标题 " + i;
            selection.TypeText(title);
            selection.Style = lever;
            selection.TypeParagraph();
        }
        selection.SetRange(0,0);
        var rang = selection.Range;
        app.ActiveDocument.TablesOfContents.Add(rang, true, 1, 3);
    } catch (error) {
        alert(error.message);
    }
}

 /**
  * @description 设置纸张方向
  */
 function PageOrientation(){
    try {
        var app = wpsFrame.Application;
        // 页面方向枚举 WdOrientation.wdOrientLandscape值为1
        app.Selection.PageSetup.Orientation = 1;
    } catch (error) {
        alert(error.message);
    }
}

 /**
  * @description 设置页面边距
  */
 function PageMargin(){
    try {
        var app = wpsFrame.Application;
        var page = app.ActiveDocument.PageSetup;
        page.TopMargin = 36;
        page.BottomMargin = 36;
        page.LeftMargin = 36;
        page.RightMargin = 36;
    } catch (error) {
        alert(error.message);
    }
}

 /**
  * @description 显示行号
  */
 function ShowLineNum(){
    try {
        var app = wpsFrame.Application;
        var lineNum = app.ActiveDocument.Range(0,0).PageSetup.LineNumbering;
        lineNum.Active = -1;
        lineNum.RestartMode = 0;
    } catch (error) {
        alert(error.message);
    }
}





