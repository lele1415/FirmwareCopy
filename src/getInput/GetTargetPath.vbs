Const ID_INPUT_TARGET_PATH = "input_target_path"
Const ID_BUTTON_OPEN_TARGET_PATH = "open_target_path"

Dim mTargetPath

Sub onChangeTargetPath(when)
    If when = CHECK_WHEN_INPUT Then Call doSomethingBeforeCheckTargetPath()

    idTimer = window.setTimeout("checkTargetPath(" & when & ")", 0, "VBScript")
End Sub

Sub checkTargetPath(when)
    window.clearTimeout(idTimer)

    Dim pathInput : pathInput = getTargetPath()

    If pathInput = "" Then
        Call doSomethingIfTargetPathIsInvalid(when, "")
        Exit Sub
    End If

    If Not checkFolderExists(pathInput) Then
        Call doSomethingIfTargetPathIsInvalid(when, "目标文件夹路径不存在")
        Exit Sub
    End If

    mTargetPath = pathInput
    
    If when = CHECK_WHEN_INPUT Then Call doSomethingAfterGetTargetPath()
End Sub

        Function getTargetPath()
            getTargetPath = getElementValue(ID_INPUT_TARGET_PATH)
        End Function

        Sub doSomethingIfTargetPathIsInvalid(when, msg)
            mTargetPath = ""
            If msg <> "" Then MsgBox(msg)
            If when = CHECK_WHEN_INPUT Then Call disableElement(ID_BUTTON_START_COPY)
            If when = CHECK_WHEN_INPUT Then Call setProcessStatus(STATUS_WAIT)
        End Sub

        Sub doSomethingBeforeCheckTargetPath()
            Call resetAllProcessStatus()
            Call setProcessStatus(STATUS_CHECK)
        End Sub

        Sub doSomethingAfterGetTargetPath()
            Call checkCopyInfoIsEnough()
            Call setProcessStatus(STATUS_WAIT)
        End Sub
