// ���j���[�̍쐬
var menu = "";
menu = menu + "[S]�ϊ�(&1)";
menu = menu + ",�󔒍s�폜(&1)";
menu = menu + ",[E]�쐬��(&2)";

var selectedItem = CreateMenu( 1, menu);

var indexMacro = 1

switch(selectedItem){
  case indexMacro:
    Editor.SelectAll(0);
    Editor.ReplaceAll('^[\\r\\n]+', '', 772);
    Editor.ReDraw(0);
  case indexMacro + 1 :
  	var regexWork = "^((?!@REPLACE@).)*$";
  	var remain = InputBox("Remain Word","");
  	regexWork = regexWork.replace("@REPLACE@", remain);
    Editor.ReplaceAll(regexWork, '', 772);
    Editor.ReDraw(0);

}

Editor.GoFiletop(0);
