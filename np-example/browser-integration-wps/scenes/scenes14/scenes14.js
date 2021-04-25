/** 
 * @file scenes05.js
 * @desc 插入图片相关处理
 * @author wps-developer
 */


/**
 * @desc 打开默认的模板文件
 */
function openTemplateFile() {
    var fileURL = getPathByFileName('watermark.wps', 'file');
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
function initWaterMark() {
    wpsFrame.initPlugin('wps', openTemplateFile);
}

/**
 * @desc 删除水印
 */
function deleteWaterMark() {
    try {
        var app, shapes, shaperange;

        app = wpsFrame.Application;
        shapes = app.ActiveDocument.Shapes;
        while (shapes.Count) {
            shapes.Item(1).Delete();
        }

        shaperange = app.ActiveDocument.Sections.Item(1).Headers.Item(1).Range.ShapeRange;
        shaperange.Delete();
    } catch (error) {
        alert(error.message);
    }
}

/**
 * @desc 插入水印
 */
function insertNewWaterMark() {
    try {
        var app, shapeRange;

        app = wpsFrame.Application;
        app.ActiveDocument.Sections.Item(1).Headers.Item(1).Shapes.AddTextEffect(0, "公司绝密", "华文新魏", 36, false, false, 0, 0).Select();
        shapeRange = app.Selection.ShapeRange;
        shapeRange.TextEffect.NormalizedHeight = false;
        shapeRange.Line.Visible = false;
        shapeRange.Fill.Visible = true;
        shapeRange.Fill.Solid();
        shapeRange.Fill.ForeColor.RGB = 12632256;       /* WdColor枚举 wdColorGray25 代表颜色值 */
        shapeRange.Fill.Transparency = 0.5;             /* 填充透明度，值为0.0~1.0 */

        shapeRange.LockAspectRatio = true;
        shapeRange.Height = 4.58 * 28.346;
        shapeRange.Width = 28.07 * 28.346;
        shapeRange.Rotation = 315;                      /* 图形按照Z轴旋转度数，正值为顺时针旋转，负值为逆时针旋转 */
        shapeRange.WrapFormat.AllowOverlap = true;
        shapeRange.WrapFormat.Side = 3;                 /* WdWrapSideType枚举 wdWrapLargest 形状距离页边距最远的一侧 */
        shapeRange.WrapFormat.Type = 3;                 /* WdWarpType枚举 wdWrapNone 形状放到文字前面 */
        shapeRange.RelativeHorizontalPosition = 0;      /* WdRelativeHorizontalPosition枚举 wdRelativeHorizontalPositionMargin 相对于边距 */
        shapeRange.RelativeVerticalPosition = 0;        /* WdRelativeVerticalPosition枚举 wdRelativeVerticalPositionMargin 相对于边距 */
        shapeRange.Left = '-999995';                    /* WdShapePosition枚举 wdShapeCenter 形状的位置在中央 */
        shapeRange.Top = '-999995';                     /* WdShapePosition枚举 wdShapeCenter 形状的位置在中央 */
    } catch (error) {
        alert(error.message);
    }
}