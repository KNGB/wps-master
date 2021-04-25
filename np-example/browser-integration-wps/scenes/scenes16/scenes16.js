/** 
 * @file scenes16.js
 * @desc 常用的字体操作
 * @author wps-developer
 */

/**
 * @desc 设置字体
 */
function FontName(){
    try {
        var app = wpsFrame.Application;
        var selection = app.Selection;
        selection.Font.Name = "Dotum"
        selection.TypeText("设置字体为Dotum");
    } catch (error) {
        alert(error.message);
    }
}
/**
 * @desc 设置字体大小
 */
function FontSize(){
    try {
        var app = wpsFrame.Application;
        var selection = app.Selection;
        selection.Font.Size = 36;
        selection.TypeText("设置字体大小为36");
    } catch (error) {
        alert(error.message);
    }
}
/**
 * @desc 字体增大减小
 */
function FontGrowShrink(){
    try {
        var app = wpsFrame.Application;
        for(i = 0; i < 3 ; i++)
        {
            var selection = app.Selection;
            selection.Font.Grow();
            selection.TypeText("大");
        }
        for(i = 0; i < 3 ; i++)
        {
            var selection = app.Selection;
            selection.Font.Shrink();
            selection.TypeText("小");
        }    
    } catch (error) {
        alert(error.message);
    }
}
/**
 * @desc 上下标
 */
function SubscriptSuperscript(){
    try {
        var app = wpsFrame.Application;
        var selection = app.Selection
        selection.TypeText("上标示例");
        selection.Font.Superscript = -1;
        selection.TypeText("上");
        selection.Font.Superscript = 0;

        selection.TypeParagraph();
        selection.TypeText("下标示例");
        selection.Font.Subscript = -1;
        selection.TypeText("下");
        selection.Font.Subscript = 0;
    } catch (error) {
        alert(error.message);
    }
}
/**
 * @desc 字体加粗
 */
function FontBold(){
    try {
        var app = wpsFrame.Application;
        var selection = app.Selection;
        var font = selection.Font;
        font.Bold =-1;
        font.BoldBi =-1;
        selection.TypeText("字体加粗");
    } catch (error) {
        alert(error.message);
    }
}
/**
 * @desc 字体倾斜
 */
function FontItalic(){
    try {
        var app = wpsFrame.Application;
        var selection = app.Selection;
        var font = selection.Font;
        font.Italic =-1;
        font.ItalicBi =-1;
        selection.TypeText("字体倾斜");
    } catch (error) {
        alert(error.message);
    }
}
/**
 * @desc 字体下划线
 */
function FontUnderline(){
    try {
        var app = wpsFrame.Application;
        var selection = app.Selection;
        var font = selection.Font;
        font.Underline = 3;
        font.UnderlineColor = 255;
    
        selection.TypeText("字体下划线");
    } catch (error) {
        alert(error.message);
    }
}

