/** 
 * @file scenes05.js
 * @desc 插入图片相关处理
 * @author wps-developer
 */


/**
 * @desc 弹出打印对话框
 */
function popupPrintDialog() {
    try {
        var app, ret = 0;

        app = wpsFrame.Application;
        ret = app.Dialogs.Item(88).Show(); /* 88为WdWordDialog枚举 wdDialogFilePrint 打印对话框 */
        if (ret) {
            alert('执行打印操作');
        } else {
            alert('打印操作已被取消');
        }
    } catch (error) {
        alert(error.message);
    }
}

/**
 * @desc 直接打印文档，不弹出打印交互对话框
 */
function printOut() {
    try {
        var app;

        app = wpsFrame.Application;
        app.ActiveDocument.PrintOut();
    } catch (error) {
        alert(error.message);
    }
}