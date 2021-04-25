/** 
 * @file scenes05.js
 * @desc 插入图片相关处理
 * @author wps-developer
 */


 /**
 * @desc 打开默认的模板文件
 */
function openTemplateFile() {
    var fileURL = getPathByFileName('header_footer.wps', 'file');
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
 * @desc 初始化插件
 */
function initHeader() {
    wpsFrame.initPlugin('wps', openTemplateFile);
}

/**
 * @desc 删除页眉横线
 */
function deleteHeaderLine() {
    try {
        var app, border, section;

        app = wpsFrame.Application;
        section = app.ActiveDocument.Sections.Item(1);

        border = section.Headers.Item(1).Range.Paragraphs.Item(1).Format.Borders.Item(-3);  /* -3为WdBorderType枚举 wdBorderBottom 底边框线 */
        border.LineStyle = 0;   /* 0为WdLineStyle枚举 wdLineStyleNone 无边框 */
        border.Color = 0;       /* 0为RGB颜色值 */

        border = section.Headers.Item(2).Range.Paragraphs.Item(1).Format.Borders.Item(-3);
        border.LineStyle = 0;
        border.Color = 0;
    } catch (error) {
        alert(error.message);
    }
}

/**
 * @desc 删除页码
 */
function deletePageNumber() {
    try {
        var app, section;

        app = wpsFrame.Application;
        section = app.ActiveDocument.Sections.Item(2);
        section.Footers.Item(1).Range.Delete();
        section.Footers.Item(2).Range.Delete();
    } catch (error) {
        alert(error.message);
    }
}

/**
 * @desc 插入页码首页不同
 */
function insertPageNumber() {
    try {
        var app, section, footer, shape;

        deletePageNumber();
        app = wpsFrame.Application;
        section = app.ActiveDocument.Sections.Item(2);
        footer = section.Footers.Item(1);
        shape = footer.Shapes.AddTextbox(1, 0, 0, 144, 144, footer.Range);
        shape.Fill.Visible = false;
        shape.Line.Visible = false;

        var textFrame;
        textFrame = shape.TextFrame;
        textFrame.AutoSize = 1;
        textFrame.WordWrap = 0;
        textFrame.MarginLeft = 0;
        textFrame.MarginRight = 0;
        textFrame.MarginTop = 0;
        textFrame.MarginBottom = 0;

        shape.RelativeHorizontalPosition = 0;
        shape.Left = '-999995';
        shape.RelativeVerticalPosition = 2;
        shape.Top = 0;
        shape.WrapFormat.Type = 3;

        textFrame.TextRange.Text = "X";
        textFrame.TextRange.Fields.Add(textFrame.TextRange, 33, "", true);

        var pageNumbers;
        pageNumbers = footer.PageNumbers;
        pageNumbers.NumberStyle = 0;
        pageNumbers.RestartNumberingAtSection = true;
        pageNumbers.StartingNumber = 1;

    } catch (error) {
        alert(error.message);
    }
}