/** 
 * @file scenes05.js
 * @desc 插入图片相关处理
 * @author wps-developer
 */


/**
 * @desc 替换书签区域的文字内容
 */
function replaceBookMarkContent() {
    try {
        var app, doc, bookmark;

        app = wpsFrame.Application;
        doc = app.ActiveDocument;
        bookmark = doc.BookMarks.Item('example');
        bookmark.Range.Text = '这是替换后的书签内容';
    } catch (error) {
        alert(error.message);
    }
}

/**
 * @desc 书签光标定位
 */
function cursorBookMark() {
    try {
        var app, doc, bookmark, start;

        app = wpsFrame.Application;
        doc = app.ActiveDocument;
        bookmark = doc.BookMarks.Item('example1');
        start = bookmark.Range.Start;
        app.Selection.SetRange(start, start);
    } catch (error) {
        alert(error.message);
    }
}

/**
 * @desc 插入多个书签
 */
function insertBookMarks() {
    try {
        var app, doc, bookmarks;

        app = wpsFrame.Application;
        doc = app.ActiveDocument;
        bookmarks = doc.BookMarks;
        bookmarks.Add('a');
        bookmarks.Add('b');
        bookmarks.Add('c');
    } catch (error) {
        alert(error.message);
    }
}

/**
 * @desc 获取所有书签名字
 */
function getBookMarksName() {
    try {
        var app, doc, bookmarks, ret = '';

        app = wpsFrame.Application;
        doc = app.ActiveDocument;
        bookmarks = doc.BookMarks;
        for (var i = 1; i <= bookmarks.Count; ++i) {
            if (i != bookmarks.Count) {
                ret += bookmarks.Item(i).Name + "，";
            } else {
                ret += bookmarks.Item(i).Name;
            }
        }
        alert(ret);
    } catch (error) {
        alert(error.message);
    }
}