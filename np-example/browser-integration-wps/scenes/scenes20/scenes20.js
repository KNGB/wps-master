/** 
 * @file scenes20.js
 * @desc 查找替换文件中的内容
 * @author wps-developer
 */



/**
 * @desc 初始化插件代码
 */
function FindReplacement() {
    try {
        var app = wpsFrame.Application;
        var find = app.Selection.Find;
        find.Text = "b";
        find.Forward = true;
        find.Wrap = 1;//WdFindWrap.wdFindContinue; 
        find.MatchCase = false;
        find.MatchByte = true;
        find.MatchWildcards = false;
        find.MatchWholeWord = false;
        find.MatchFuzzy = false;
        find.Replacement.Text = "d";
        find.Style = "";
        find.Highlight = 9999999;//WdConstants.wdUndefined;
    
        var replacement = find.Replacement;
        replacement.Style = "";
        replacement.Highlight = 9999999;//WdConstants.wdUndefined;
        var empty = undefined; 
        var replace = 1;//WdReplace.wdReplaceAll
        app.Selection.Find.Execute(empty, empty, empty, empty, empty, empty, empty, empty, empty, empty, replace);
        app.Selection.Find.Replacement.Text = "d";
    } catch (error) {
        alert(error.message);
    }
}