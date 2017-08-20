Const ID_LIST_TARGET_HISTORY = "list_target_history"
Const ID_UL_TARGET_HISTORY = "ul_target_history"

Dim pTargetPathHistory
pTargetPathHistory = oWs.CurrentDirectory & "\tmp\targetPathHistory.txt"
Dim vaHistoryTargetFolder : Set vaHistoryTargetFolder = New VariableArray

Call readHistory(pTargetPathHistory, ID_INPUT_TARGET_PATH, ID_LIST_TARGET_HISTORY, ID_UL_TARGET_HISTORY)

Sub readHistory(DictPath, inputId, listId, ulId)
    Call readTextAndDoSomething(DictPath, _
            "If Len(Trim(sReadLine)) > 0 Then" &_
            " Call addBeforeLi(sReadLine, """&inputId&""", """&listId&""", """&ulId&""")" &_
            " : vaHistoryTargetFolder.Append(sReadLine)")
End Sub

Sub writeHistory(DictPath, inputId, listId, ulId, str)
    Dim seqInArray : seqInArray = vaHistoryTargetFolder.GetIndexIfExist(str)
    If seqInArray <> -1 Then
        Call removeLiByIndex(ulId, vaHistoryTargetFolder.Bound - seqInArray)
        Call addBeforeLi(str, inputId, listId, ulId)

        vaHistoryTargetFolder.MoveToEnd(seqInArray)

        Call writeNewHistoryTxt(DictPath)
        Exit Sub
    End If


    If vaHistoryTargetFolder.Bound < 7 Then
        Call addBeforeLi(str, inputId, listId, ulId)

        vaHistoryTargetFolder.Append(str)
    Else
        Call removeLiByIndex(ulId, vaHistoryTargetFolder.Bound)
        Call addBeforeLi(str, inputId, listId, ulId)

        vaHistoryTargetFolder.PopBySeq(0)
        vaHistoryTargetFolder.Append(str)
    End If
    
    Call writeNewHistoryTxt(DictPath)
End Sub

Sub writeNewHistoryTxt(DictPath)
    Call initTxtFile(DictPath)
    Call writeTextAndDoSomething(DictPath, _
            "Dim i : For i = 0 To vaHistoryTargetFolder.Bound" &_
            " : oText.WriteLine(vaHistoryTargetFolder.V(i))" &_
            " : Next")
End Sub
