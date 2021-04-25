/** 
 * @file scenes05.js
 * @desc 插入图片相关处理
 * @author wps-developer
 */


/**
 * @desc 光标出插入表格
 */
function insertTable() {
    try {
        var app, table;

        app = wpsFrame.Application;
        table = app.ActiveDocument.Tables.Add(app.Selection.Range, 10, 5);
    } catch (error) {
        alert(error.message);
    }
}

/**
 * @desc 设置表格边框样式
 */
function setTableStyle() {
    try {
        var app, table = null, borders, border;

        app = wpsFrame.Application;
        table = app.ActiveDocument.Tables.Item(1);
        if (null == table) {
            alert('请先插入表格');
            return;
        }

        borders = table.Borders;
        border = borders.Item(-1);  /* -1为WdBorderType枚举 wdBorderTop 上框线，以下都是改类型枚举值*/
        border.LineStyle = 1;       /* 1为WdLineStyle枚举 wdLineStyleSingle 单实线 */
        border.LineWidth = 4;       /* 4为WdLineWidt枚举 wdLineWidth050pt 0.50磅 */
        border.Color = 255;         /* 255为RGB颜色值 */

        border = borders.Item(-3);
        border.LineStyle = 1;
        border.LineWidth = 4;
        border.Color = 255;

        border = borders.Item(-5);
        border.LineStyle = 1;
        border.LineWidth = 4;
        border.Color = 255;

        border = borders.Item(-6);
        border.LineStyle = 1;
        border.LineWidth = 4;
        border.Color = 255;

        border = borders.Item(-2);
        border.LineStyle = 1;
        border.LineWidth = 4;
        border.Color = 255;

        border = borders.Item(-4);
        border.LineStyle = 1;
        border.LineWidth = 4;
        border.Color = 255;
    } catch (error) {
        alert(error.message);
    }
}

/**
 * @desc 合并单元格
 */
function mergeCell() {
    try {
        var app, table = null;

        app = wpsFrame.Application;
        table = app.ActiveDocument.Tables.Item(1);
        if (null == table) {
            alert('请先插入表格');
            return;
        }
        table.Cell(1,1).Merge(table.Cell(1,2));
    } catch (error) {
        alert(error.message);
    }
}