/** 
 * @file scenes18.js
 * @desc 图形操作
 * @author wps-developer
 */

 /**
  * @description 插入椭圆
  */
function InsertEllipse(){
    try {
        var app = wpsFrame.Application;
        var shapes = app.ActiveDocument.Shapes;
        var Type = 9;   // MsoAutoShapeType msoShapeOval
        var Left = 100;
        var Top = 100;
        var Width = 320;
        var Height = 128;
        var shape = shapes.AddShape(Type, Left, Top, Width, Height);
        shape.TextFrame.TextRange.Text = "插入椭圆";
    } catch (error) {
        alert(error.message);
    }
}

 /**
  * @description 插入符号
  */
 function InsertSymbol(){
    try {
        var app = wpsFrame.Application;
        var CharacterNumber = 9632;
        var Font = "东文宋体";
        var Unicode =true;
        app.Selection.InsertSymbol(CharacterNumber, Font, Unicode);
    } catch (error) {
        alert(error.message);
    }
}

 /**
  * @description 插入艺术字
  */
 function InsertWordArt(){
    try {
        var app = wpsFrame.Application;
        var PresetTextEffect = 3; //预设的文本效果, 常数的值对应于 艺术字库对话框 (从左到右，从上到下编号) 中列出的格式
        var Text = "插入艺术字";
        var FontName = "宋体";
        var FontSize = 36;
        var FontBold = 1;
        var FontItalic = 1;
        var Left = 50;
        var Top = 50;
        app.ActiveDocument.Shapes.AddTextEffect(PresetTextEffect, Text, FontName, FontSize, FontBold, FontItalic, Left, Top);
    } catch (error) {
        alert(error.message);
    }
}

 /**
  * @description 旋转图形
  */
 function IncrementRotation(){
    try {
        var app = wpsFrame.Application;
        var shapes = app.ActiveDocument.Shapes;
        var Type = 13;   //MsoAutoShapeType msoShapeCan
        var Left = 100;
        var Top = 100;
        var Width = 50;
        var Height = 128;
        var shape = shapes.AddShape(Type, Left, Top, Width, Height);
        shape.IncrementRotation(-90);
    } catch (error) {
        alert(error.message);
    }
}
