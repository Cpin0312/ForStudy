// ���j���[�̍쐬
var menu = "";
menu = menu + "[S]�폜(&R)";
menu = menu + ",�󔒍s�폜(&1)";
menu = menu + ",�ΏۊO�폜(&2)";
menu = menu + ",�ΏۊO�폜�y�P��z(&3)";
menu = menu + ",[EE]�쐬��(&4)";
menu = menu + ",[S]Read(&T)";

var selectedItem = CreateMenu( 1, menu);

var indexMacro = 1

switch(selectedItem){
  case indexMacro:
    Editor.SelectAll(0);
    Editor.ReplaceAll('^[\\r\\n]+', '', 772);
    Editor.ReDraw(0);
    break;
  case indexMacro + 1 :
  	var regexWork = "^((?!@REPLACE@).)*$";
  	var remain = InputBox("Remain Word","");
  	regexWork = regexWork.replace("@REPLACE@", remain);
    Editor.ReplaceAll(regexWork, '', 772);
    Editor.ReDraw(0);
    break;
  case indexMacro + 2 :
  	var regexWork = "^((?!\\b(@REPLACE@)\\b).)*$";
  	var remain = InputBox("Remain Word","");
  	regexWork = regexWork.replace("@REPLACE@", remain);
    Editor.ReplaceAll(regexWork, '', 772);
    Editor.ReDraw(0);
    break;
  case indexMacro + 3 :
    break;
}

Editor.GoFiletop(0);
