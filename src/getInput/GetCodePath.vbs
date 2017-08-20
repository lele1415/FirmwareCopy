Const ID_BUTTON_OPEN_CODE_PATH = "open_code_path"

Dim mCodePath, mOutSoftwarePath

Sub onChangeCodePath(when)
    If when = CHECK_WHEN_INPUT Then Call doSomethingBeforeCheckCodePath()
        
    idTimer = window.setTimeout("checkCodePath(" & when & ")", 0, "VBScript")
End Sub

Sub checkCodePath(when)
    window.clearTimeout(idTimer)

    Dim pathInput
    pathInput = getCodePath()

    If pathInput = "" Then
        Call doSomethingIfCodePathIsInvalid(when, "")
        Exit Sub
    End If

    If Not checkFolderExists(pathInput) Then
        Call doSomethingIfCodePathIsInvalid(when, "代码out路径不存在")
        Exit Sub
    End If

    If Not checkFolderExists(pathInput & "\target\product") Then
        Call doSomethingIfCodePathIsInvalid(when, "不存在" & pathInput & "\target\product")
        Exit Sub
    End If

    Dim outPathSearch : outPathSearch = getOutSoftwarePath(pathInput & "\target\product")

    If outPathSearch = "" Then
        Call doSomethingIfCodePathIsInvalid(when, "out下的system.img不存在")
        Exit Sub
    End If

    mCodePath = pathInput
    mOutSoftwarePath = outPathSearch

    If when = CHECK_WHEN_INPUT Then Call doSomethingAfterGetCodePath()
End Sub

        Function getCodePath()
            getCodePath = getElementValue(ID_INPUT_CODE_PATH)
        End Function

        Function getOutSoftwarePath(path)
            Dim str, pSystemimg
            str = ""
            pSystemimg = searchFolder(path, "system.img", _
                    SEARCH_FILE, SEARCH_SUB, SEARCH_WHOLE_NAME, SEARCH_ONE, SEARCH_RETURN_PATH)

            If pSystemimg <> "" Then
                str = Replace(pSystemimg, "\system.img", "")
            End If

            getOutSoftwarePath = str
        End Function

        Sub doSomethingIfCodePathIsInvalid(when, msg)
            mCodePath = ""
            mOutSoftwarePath = ""
            If msg <> "" Then MsgBox(msg)
            If when = CHECK_WHEN_INPUT Then Call freezeElementsDependCodePath()
            If when = CHECK_WHEN_INPUT Then Call setProcessStatus(STATUS_WAIT)
        End Sub

        Sub doSomethingBeforeCheckCodePath()
            Call resetAllProcessStatus()
            Call resetOtaSelectInfo()
            Call setProcessStatus(STATUS_CHECK)
        End Sub

        Sub doSomethingAfterGetCodePath()
            Call unfreezeElementsDependCodePath()
            Call checkCopyInfoIsEnough()
            Call setProcessStatus(STATUS_WAIT)
        End Sub

        Sub unfreezeElementsDependCodePath()
            Call enableElement(ID_BUTTON_DISPLAY_ID)
            Call enableElement(ID_BUTTON_BUILD_VERSION)
            Call enableElement(ID_SELECT_OTA_FILE)
            Call enableElement(ID_BUTTON_CHECK_OTA_FILES)
        End Sub

        Sub freezeElementsDependCodePath()
            Call disableElement(ID_BUTTON_DISPLAY_ID)
            Call disableElement(ID_BUTTON_BUILD_VERSION)
            Call disableElement(ID_SELECT_OTA_FILE)
            Call disableElement(ID_BUTTON_CHECK_OTA_FILES)
            Call disableElement(ID_BUTTON_START_COPY)
        End Sub
