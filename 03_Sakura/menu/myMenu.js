String.prototype.padStart = function (length, char) {
  if (this.length > length) return this; // �K�v�Ȃ�ǉ�
  var left = '';
  for (;left.length < length; left += char);
  return (left+this.toString()).slice(-length);
}

function getTimestamp() {
  return getYYYYMMDD() + " "+ getHHMMSS()  ;
}

function getYYYYMMDDYouBi() {
  return getYYYYMMDD() + "�i" + getYoubi() + "�j";
}

function getYoubi() {
  return [ "��", "��", "��", "��", "��", "��", "�y" ][(new Date()).getDay()] ;
}

function getYYYYMMDD() {
  var currentDate = new Date();
  var year  = currentDate.getFullYear();
  var month = (currentDate.getMonth() + 1).toString().padStart(2 , '0');
  var day   = currentDate.getDate().toString().padStart(2, '0');
  return year + "-" + month + "-" + day;
}

function getHHMMSS() {
  var currentDate = new Date();
  var hours   = currentDate.getHours().toString().padStart(2, '0');
  var minutes = currentDate.getMinutes().toString().padStart(2, '0');
  var seconds = currentDate.getSeconds().toString().padStart(2, '0');
  return hours + ":" + minutes + ":"  + seconds ;
}

// ���j���[�̍쐬
var menu = "";
menu = menu + "[S]�폜(&R)";
menu = menu + ",�󔒍s�폜(&1)";
menu = menu + ",�ΏۊO�폜(&2)";
menu = menu + ",[EE]�ΏۊO�폜�y�P��z(&3)";
menu = menu + ",[S]�C�x���g(&E)";
menu = menu + ",[EE]�ꗗ�쐬(&1)";
menu = menu + ",GoogleCalender�ꗗ�쐬(&G)";
menu = menu + ",[S]GetDate(&D)";
menu = menu + ",GetTimeStamp(&1)";
menu = menu + ",GetDate(&2)";
menu = menu + ",[EE]GetTime(&3)";

var selectedItem = CreateMenu( 1, menu);

var indexMacro = 1
var indexEvent = 4
var indexGoogle = 5
var indexGetDate = 6

switch(selectedItem){
  case indexMacro:
    Editor.SelectAll(0);
    Editor.ReplaceAll('^[\\r\\n]+', '', 772);
    Editor.ReDraw(0);
    Editor.GoFiletop(0);
    break;
  case indexMacro + 1 :
    var regexWork = "^((?!@REPLACE@).)*$";
    var remain = InputBox("Remain Word","");
    regexWork = regexWork.replace("@REPLACE@", remain);
    Editor.ReplaceAll(regexWork, '', 772);
    Editor.ReDraw(0);
    Editor.GoFiletop(0);
    break;
  case indexMacro + 2 :
    var regexWork = "^((?!\\b(@REPLACE@)\\b).)*$";
    var remain = InputBox("Remain Word","");
    regexWork = regexWork.replace("@REPLACE@", remain);
    Editor.ReplaceAll(regexWork, '', 772);
    Editor.ReDraw(0);
    Editor.GoFiletop(0);
    break;
  case indexEvent :
    ExecExternalMacro("D:/0_MyFolder/07_�ėp/02_Excel/GetEvent/setEvent.mac");
    Editor.GoFiletop(0);
    break;
  case indexGoogle :
    ExecExternalMacro("D:/0_MyFolder/07_�ėp/03_Sakura/setToGoogleCalender.mac");
    Editor.GoFiletop(0);
    break;
  case indexGetDate :
    InsText(getTimestamp());
    InsText("\r\n");
    break;
  case indexGetDate + 1 :
    InsText(getYYYYMMDDYouBi());
    InsText("\r\n");
    break;
  case indexGetDate + 2:
    InsText(getHHMMSS());
    InsText("\r\n");
    break;
}

