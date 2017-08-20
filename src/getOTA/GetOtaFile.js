/*function setDefaultPath() {
    var OutSoftwarePath = getCurrentOutSoftwarePath();
    var objFile = document.getElementById("ota_file_path");
    var WshShell = new ActiveXObject("WScript.Shell");
    if (OutSoftwarePath != "") {
        objFile.focus();
        WshShell.SendKeys(OutSoftwarePath);
        WshShell.SendKeys("{Enter}");
    } else {
        objFile.focus();
        WshShell.SendKeys("{TAB}");
        WshShell.SendKeys("{TAB}");
        WshShell.SendKeys("{TAB}");
        WshShell.SendKeys("{Enter}");
    }
}*/

function jsAddOption(SelectId, OptionName) 
{
    var option = option_creat(OptionName, OptionName);
    parentNode_appendChild(SelectId, option);
}

function jsAddOptionValueAndName(SelectId, OptionValue, OptionName) 
{
    var option = option_creat(OptionValue, OptionName);
    parentNode_appendChild(SelectId, option);
}

function jsRemoveAllOption(SelectId)
{
    var select = document.getElementById(SelectId)
    var count = select.getElementsByTagName("option").length;
    if (count > 0) {
        for (var i = 1; i <= count; i++) {
            parentNode_removeChild(SelectId, 1)
        }
    }
}
