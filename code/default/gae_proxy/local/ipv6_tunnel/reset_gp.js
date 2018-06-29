// coding:ANSI
// 检测并重置 Teredo 组策略
// 此文件为微软 JScript 源码，需编译成二进制运行
// 使用 .Net Framwork 的 JScript 编译器 JScriptCompiler 编译，不支持 1.0 版本
// jsc.exe /t:winexe /platform:anycpu /fast /out:x:\reset_gp.exe x:\reset_gp.js
// reset_gp.exe.config 为 .Net Framwork 版本兼容性配置文件
// Windows 2000 必须安装 .Net Framwork 1.0 以上版本方可运行此程序

function BinaryFile(filepath){
    this.path = filepath;
    this._initNewStream = function(){
        var Stream = new ActiveXObject('ADODB.Stream');
        Stream.Type = 2;
        Stream.CharSet = 'iso-8859-1';
        Stream.Open();
        return Stream;
    }
    this.WriteAll = function(content){
        var Stream = this._initNewStream();
        Stream.WriteText(content);
        Stream.SaveToFile(this.path, 2);
        Stream.Close();
    }
    this.ReadAll = function(){
        var Stream = this._initNewStream();
        Stream.LoadFromFile(this.path);
        var content = Stream.ReadText();
        Stream.Close();
        return content;
    }
}

import System.Windows.Forms;

var Wsr = new ActiveXObject('WScript.Shell');
var gp_split = '[\x00';
var gp_teredo = 'v\x006\x00T\x00r\x00a\x00n\x00s\x00i\x00t\x00i\x00o\x00n\x00\x00\x00;\x00T\x00e\x00r\x00e\x00d\x00o\x00';
var gp_regpol_filename = Wsr.ExpandEnvironmentStrings('%windir%') + '\\System32\\GroupPolicy\\Machine\\Registry.pol';
var gp_regpol_file = new BinaryFile(gp_regpol_filename);
var gp_regpol_old = gp_regpol_file.ReadAll().split(gp_split);
var gp_regpol_new = new Array();

for (var i=0; i<gp_regpol_old.length; i++) {
    var gp = gp_regpol_old[i];
    if (gp.indexOf(gp_teredo) == -1) {
        gp_regpol_new.push(gp);
    }
}

var result;
if (gp_regpol_new.length != gp_regpol_old.length) {
    var LangId = Wsr.RegRead('HKEY_CURRENT_USER\\Control Panel\\International\\Locale');
    if (LangId == '00000804') {
        var message = '\u53d1\u73b0 Teredo \u7ec4\u7b56\u7565\u8bbe\u7f6e\uff0c\u662f\u5426\u91cd\u7f6e\uff1f';
        var title = '\u63d0\u793a';
    } else {
        var message = 'Found Group Policy Teredo settings, do you want to reset it?';
        var title = 'Notice';
    }
    result = MessageBox.Show(
        message,
        title,
        MessageBoxButtons.OKCancel,
        MessageBoxIcon.Warning,
        MessageBoxDefaultButton.Button1
    );
}

var cmd = 'for /f "tokens=2 delims=[" %a in (\'ver\') do (' +
'for /f "tokens=2 delims= " %b in ("%a") do (' +
'for /f "tokens=1 delims=]" %c in ("%b") do (' +
'if %c lss 5.1 (' +
'secedit /refreshpolicy machine_policy /enforce' +
') else (' +
'gpupdate /target:computer /force))))';

if (result == DialogResult.OK) {
    gp_regpol_file.WriteAll(gp_regpol_new.join(gp_split));
    Wsr.Run("cmd.exe /c " + cmd, 0, true);
}

