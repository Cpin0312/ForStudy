// メニューの作成
var menu = "";
menu = menu + "[S]変換(&1)";
menu = menu + ",空白行削除(&1)";
menu = menu + ",[E]作成中(&2)";

var selectedItem = CreateMenu( 1, menu);

var indexMacro = 1

switch(selectedItem){
  case indexMacro:
    Editor.SelectAll(0);
    Editor.ReplaceAll('^[\\r\\n]+', '', 772);
    Editor.ReDraw(0);

}