function input_creat(inputType,inputId,inputValue,inputOnclickFun,inputSize)
{
    var input = document.createElement("input");
    input.type = inputType;
    input.id = inputId;
    if (inputValue) {
        input.value = inputValue;
    }
    if (inputOnclickFun) {
        input.onclick = function(){inputOnclickFun()};
    }
    if (inputSize){
        input.style.width = inputSize;
    }

    return input;
}

function input_changeInfo(inputId,inputValue,inputOnclickFun)
{
    var input = document.getElementById(inputId);
    if(inputValue){
        input.value = inputValue;
    }
    if(inputOnclickFun){
        input.onclick = function(){inputOnclickFun()};
    }
}

function select_creat(selectId,selectOnchangeFun,selectWidth)
{
    var select = document.createElement("select");
    select.id = selectId;
    if(selectOnchangeFun){
        select.onchange = function(){selectOnchangeFun()};
    }
    if(selectWidth){
        select.style.width = selectWidth;
    }

    return select;
}

function option_creat(optionValue,optionInnerHTML)
{
    var option = document.createElement("option");
    option.value = optionValue;
    option.innerHTML = optionInnerHTML;

    return option;
}

function parentNode_appendChild(parentNodeId,node)
{
    var parentNode = document.getElementById(parentNodeId);
    parentNode.appendChild(node);
}

function parentNode_insertBefore(newNode,oldNodeId)
{
    var oldNode = document.getElementById(oldNodeId);
    oldNode.parentNode.insertBefore(newNode,oldNode);
}

function parentNode_getChildNodesLength(parentNodeId)
{
    var parentNode = document.getElementById(parentNodeId);
    return parentNode.childNodes.length;
}

function parentNode_removeChild(parentNodeId,childNodeCount)
{
    var parentNode = document.getElementById(parentNodeId);
    if(parentNode.childNodes[childNodeCount]){
        parentNode.removeChild(parentNode.childNodes[childNodeCount]);
    }
}

function node_removeNode(nodeId)
{
    var node = document.getElementById(nodeId);
    if(node){
        node.removeNode(true);
    }
}

function element_getValue(elementId)
{
    var elementValue = document.getElementById(elementId).value;
    return elementValue;
}

function element_setValue(elementId,elementValue)
{
    document.getElementById(elementId).value = elementValue;
}

function element_isChecked(elementId)
{
    var isChecked = document.getElementById(elementId).checked;
    return isChecked;
}
