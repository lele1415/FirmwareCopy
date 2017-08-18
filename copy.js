/*function addVersionButton() {
    var obj = document.getElementById("version_button");
    var count = obj.getElementsByTagName("input").length;
    if(!(count > 0)){
        var div1 = document.createElement("div");
        div1.id = "getVerName"
        div1.innerHTML = "获取版本号作为文件夹名"

        var input_bt1 = document.createElement("input");
        input_bt1.id = "getVersion1";
        input_bt1.type = "button";
        input_bt1.onclick = function(){getCustomBuildVersion("display.id")};
        input_bt1.value = "display.id"

        var input_bt2 = document.createElement("input");
        input_bt2.id = "getVersion2";
        input_bt2.type = "button";
        input_bt2.onclick = function(){getCustomBuildVersion("build.version")};
        input_bt2.value = "build.version"
     
        obj.appendChild(div1);
        obj.appendChild(input_bt1);
        obj.appendChild(input_bt2);
    }    
}

function removeVersionButton() {
    var obj = document.getElementById("version_button");
    var count = obj.getElementsByTagName("input").length;
    if(count > 0){
        var child1 = document.getElementById("getVerName");
        var child2 = document.getElementById("getVersion1");
        var child3 = document.getElementById("getVersion2");
        obj.removeChild(child1); 
        obj.removeChild(child2); 
        obj.removeChild(child3); 
    }     
}*/

function dbCount(){
    var i = document.getElementById("db_count").innerHTML;
    i++;
    document.getElementById("db_count").innerHTML = i; 
}

function softwareCount(){
    var i = document.getElementById("software_count").innerHTML;
    i++;
    document.getElementById("software_count").innerHTML =  i; 
}

function otaCount(){
    var i = document.getElementById("ota_count").innerHTML;
    i++;
    document.getElementById("ota_count").innerHTML =  i; 
}

function cleanCount(){
    document.getElementById("db_count").innerHTML =  0; 
    document.getElementById("software_count").innerHTML =  0; 
    document.getElementById("ota_count").innerHTML =  0; 
}

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

function showAndHide(listId, types)
{ 
    var Layer=window.document.getElementById(listId); 
    switch(types){ 
        case "show": 
            Layer.style.display="block"; 
            break; 
        case "hide": 
            Layer.style.display="none"; 
    } 
}

function addBeforeLi(str, inputId, listId, ulId)
{
    var obj = document.getElementById(ulId);
    var li = document.createElement("li");
    li.onmousedown = function(){setValueOfTargetFolder(inputId, listId, str)};
    li.innerHTML = str;
    li.style.fontSize = "x-small";
    
    if (obj.childNodes.length > 0) {
        var node = obj.childNodes[0];
        obj.insertBefore(li,node);
    } else {
        obj.appendChild(li);
    }
}

function addAfterLi(str, inputId, listId, ulId)
{
    var obj = document.getElementById(ulId);
    var li = document.createElement("li");
    li.onmousedown = function(){setValueOfTargetFolder(inputId, listId, str)};
    li.innerHTML = str;
    li.style.fontSize = "x-small";
    
    obj.appendChild(li);
}

function removeLi(ulId, seq) {
    parentNode_removeChild(ulId, seq);
}