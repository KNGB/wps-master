/** 
 * @file scenes05.js
 * @desc 插入图片相关处理
 * @author wps-developer
 */

/**
 * @desc 在当前光标处插入选定格式的图片
 */
function insertPicture() {
	var targetRange, inlineShape, shape, app, index, formatType, picturePath;

	app = wpsFrame.Application;
	targetRange = app.Selection.Range;
	picturePath = getPathByFileName('wps_tamplate.png', 'picture');
	inlineShape = app.ActiveDocument.InlineShapes.AddPicture(picturePath, false, true, targetRange);
	shape = inlineShape.ConvertToShape();
	shape.Select();
	index = document.getElementById('AreaId').selectedIndex;
	if (index <= 0) {
		index += 1;
	}
	formatType = document.getElementById('AreaId').children[index].value;
	shape.WrapFormat.Type = formatType;
}