function showAndHide(listId, types) {
    var Layer = window.document.getElementById(listId);
    switch (types) {
        case "show":
            Layer.style.display="block";
            break;
        case "hide":
            Layer.style.display="none";
    }
}

function addAfterLi(str, inputId, listId, ulId) {
    var obj = document.getElementById(ulId);
    var li = document.createElement("li");
    li.onmousedown = function(){setListValue(inputId, listId, str)};
    li.innerHTML = str;
    li.style.fontSize = "x-small";
    
    obj.appendChild(li);
}

function addBeforeLi(str, inputId, listId, ulId)
{
    var obj = document.getElementById(ulId);
    var li = document.createElement("li");
    li.onmousedown = function(){setListValue(inputId, listId, str)};
    li.innerHTML = str;
    li.style.fontSize = "x-small";
    
    if (obj.childNodes.length > 0) {
        var node = obj.childNodes[0];
        obj.insertBefore(li,node);
    } else {
        obj.appendChild(li);
    }
}

function removeLiByIndex(ulId, index) {
    parentNode_removeChild(ulId, index);
}

function disableElement(inputId) {
    document.getElementById(inputId).disabled = "disabled";
}

function enableElement(inputId) {
    document.getElementById(inputId).disabled = "";
}

function elementIsChecked(elementId) {
    var isChecked = document.getElementById(elementId).checked;
    return isChecked;
}
