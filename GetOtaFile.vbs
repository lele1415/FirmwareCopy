'Const ID_INPUT_OTA_FILE_PATH = "ota_file_path"
Const ID_SELECT_OTA_FILE = "select_ota_file"
Const ID_BUTTON_CHECK_OTA_FILES = "check_ota_files"
Const ID_DIV_OTA_WAIT = "ota_wait"

Dim mOtaFilePath

Sub onClickCheckOtaFiles()
    Call resetAllProcessStatus()
    Call jsRemoveAllOption(ID_SELECT_OTA_FILE)
    Call setProcessStatus(STATUS_CHECK)
    idTimer = window.setTimeout("checkOtaFiles()", 0, "VBScript")
End Sub

Sub checkOtaFiles()
    window.clearTimeout(idTimer)

	If mOutSoftwarePath = "" Then MsgBox("out下的system.img不存在") : Exit Sub

    Dim pOta_1 : pOta_1 = searchFolder(mOutSoftwarePath, "target_files-package.zip" _
            , SEARCH_FILE, SEARCH_ROOT, SEARCH_WHOLE_NAME, SEARCH_ONE, SEARCH_RETURN_PATH)
    Dim pOta_2
    If checkFolderExists(mOutSoftwarePath) Then
        Set pOta_2 = searchFolder(mOutSoftwarePath, "-ota-"_
                , SEARCH_FILE, SEARCH_ROOT, SEARCH_PART_NAME, SEARCH_ALL, SEARCH_RETURN_PATH)
    ELse
        Set pOta_2 = New VariableArray
    End If

    Dim pOta_3
    If checkFolderExists(mOutSoftwarePath & "\obj\PACKAGING\target_files_intermediates") Then
        Set pOta_3 = searchFolder(mOutSoftwarePath & "\obj\PACKAGING\target_files_intermediates", "-target_files-" _
                , SEARCH_FILE, SEARCH_ROOT, SEARCH_PART_NAME, SEARCH_ALL, SEARCH_RETURN_PATH)
    ELse
        Set pOta_3 = New VariableArray
    End If

    If pOta_3.Bound > 0 Then Call pOta_3.SortArray()

    If pOta_1 = "" And pOta_2.Bound = -1 And pOta_3.Bound = -1 Then
        Call jsAddOptionValueAndName(ID_SELECT_OTA_FILE, "", "无OTA文件")
        Exit Sub
    End If

    Call jsAddOptionValueAndName(ID_SELECT_OTA_FILE, "", "不拷贝OTA")
    If pOta_1 <> "" Then Call addSingleOption(pOta_1)
    If pOta_2.Bound <> -1 Then Call addMultiOptions(pOta_2)
    If pOta_3.Bound <> -1 Then Call addMultiOptions(pOta_3)

    Call setProcessStatus(STATUS_WAIT)
End Sub

        Sub addSingleOption(value)
            Call jsAddOption(ID_SELECT_OTA_FILE, Replace(value, mOutSoftwarePath, ""))
        End Sub

        Sub addMultiOptions(vaObj)
            Dim i
            For i = 0 To vaObj.Bound
                Call jsAddOption(ID_SELECT_OTA_FILE, Replace(vaObj.V(i), mOutSoftwarePath, ""))
            Next
        End Sub

Sub onSelectOtaFile()
    If mOutSoftwarePath = "" Then
        mOtaFilePath = ""
        MsgBox("out下的system.img不存在")
        Exit Sub
    End If

    Dim mSelectOtaFilePath
    mSelectOtaFilePath = getElementValue(ID_SELECT_OTA_FILE)

    If mSelectOtaFilePath = "不拷贝OTA" Then
        mOtaFilePath = ""
        Exit Sub
    End If

    Call setElementInnerHTML(ID_DIV_OTA_WAIT, 1)
    mOtaFilePath = mOutSoftwarePath & getElementValue(ID_SELECT_OTA_FILE)
End Sub

Sub resetOtaSelectInfo()
    Call jsRemoveAllOption(ID_SELECT_OTA_FILE)
    Call jsAddOptionValueAndName(ID_SELECT_OTA_FILE, "", "如果编了OTA，请点击下面的按钮")
    mOtaFilePath = ""
End Sub
