function addInputBox() {
    var obj = document.getElementById("inputbox1");
    var count = obj.getElementsByTagName("input").length;
    if(!(count > 0)){
        var input = document.createElement("input");
        input.id = "folder_name";
        input.type = "text";
        input.size = 50;

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
     
        obj.appendChild(input);
        obj.appendChild(input_bt1);
        obj.appendChild(input_bt2);
    }    
}

function removeInputBox() {
    var obj = document.getElementById("inputbox1");
    var count = obj.getElementsByTagName("input").length;
    if(count > 0){
        var child1 = document.getElementById("folder_name");
        var child2 = document.getElementById("getVersion1");
        var child3 = document.getElementById("getVersion2");
        obj.removeChild(child1); 
        obj.removeChild(child2); 
        obj.removeChild(child3); 
    }     
}

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

function option_creat(optionValue, optionInnerHTML)
{
    var option = document.createElement("option");
    option.value = optionValue;
    option.innerHTML = optionInnerHTML;

    return option;
}

function parentNode_appendChild(parentNodeId, node)
{
    var parentNode = document.getElementById(parentNodeId);
    parentNode.appendChild(node);
}


function addOption(SelectId, OptionName) 
{
    var option = option_creat(OptionName, OptionName);
    parentNode_appendChild(SelectId, option);
}