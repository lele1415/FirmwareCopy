Const ID_INPUT_FOLDER_NAME = "input_folder_name"
Const ID_BUTTON_DISPLAY_ID = "get_display_id"
Const ID_BUTTON_BUILD_VERSION = "get_build_version"
Const INVALID_CHAR_OF_FOlDER_NAME = "/\:*?""<>|"

Dim mNewFolderName

Sub onChangeFolderName(when)
    If when = CHECK_WHEN_INPUT Then Call doSomethingBeforeCheckNewFolderName()

    idTimer = window.setTimeout("checkNewFolderName(" & when & ")", 0, "VBScript")
End Sub

Sub checkNewFolderName(when)
    window.clearTimeout(idTimer)

    Dim strInput : strInput = getNewFolderName()

    If strInput = "" Then
        Call doSomethingIfNewFolderNameIsInvalid(when, "")
        Exit Sub
    End If

    If Not isValidFolderName(strInput) Then
        Call doSomethingIfNewFolderNameIsInvalid(when, _
                "文件名不能包含下列任何字符：" & Vblf & INVALID_CHAR_OF_FOlDER_NAME)
        Exit Sub
    End If

    mNewFolderName = strInput
    
    If when = CHECK_WHEN_INPUT Then Call doSomethingAfterGetNewFolderName()
End Sub

        Function getNewFolderName()
            getNewFolderName = getElementValue(ID_INPUT_FOLDER_NAME)
        End Function

        Function isValidFolderName(name)
            Dim i, char
            For i = 1 To Len(name)
                char = Mid(name, i, 1)
                If InStr(INVALID_CHAR_OF_FOlDER_NAME, char) Then
                    isValidFolderName = False
                    Exit Function
                End If
            Next

            isValidFolderName = True
        End Function

        Sub doSomethingIfNewFolderNameIsInvalid(when, msg)
            mNewFolderName = ""
            If msg <> "" Then MsgBox(msg)
            If when = CHECK_WHEN_INPUT Then Call disableElement(ID_BUTTON_START_COPY)
            If when = CHECK_WHEN_INPUT Then Call setProcessStatus(STATUS_WAIT)
        End Sub

        Sub doSomethingBeforeCheckNewFolderName()
            Call resetAllProcessStatus()
            Call setProcessStatus(STATUS_CHECK)
        End Sub

        Sub doSomethingAfterGetNewFolderName()
            Call checkCopyInfoIsEnough()
            Call setProcessStatus(STATUS_WAIT)
        End Sub

Sub getCustomBuildVersion(which)
    Call setProcessStatus(STATUS_CHECK)
    idTimer = window.setTimeout("checkBuildProp(""" & which & """)", 0, "VBScript")

End Sub

Sub checkBuildProp(which)
    window.clearTimeout(idTimer)

    '//check build.prop
    Dim pBuildProp : pBuildProp = mOutSoftwarePath & "\system\build.prop"
    If Not oFso.FileExists(pBuildProp) Then
        MsgBox("文件不存在: " & Vblf & pBuildProp)
        Call setProcessStatus(STATUS_WAIT)
        Exit Sub
    End If

    '//get version
    Call readTextAndDoSomething(pBuildProp, _
            "If InStr(sReadLine, """&which&""") > 0 Then" &_
                " Call setElementValue(ID_INPUT_FOLDER_NAME, Trim(Mid(sReadLine,InStr(sReadLine,""="")+1)))" &_
                " : Call onChangeFolderName(0)" &_
                " : exitFlag = True")

    Call setProcessStatus(STATUS_WAIT)
End Sub
