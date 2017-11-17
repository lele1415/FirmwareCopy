Const ID_BUTTON_START_COPY = "start_copy"
Const ID_DIV_DB_COPYED = "db_copyed"
Const ID_DIV_SOFTWARE_COPYED = "software_copyed"
Const ID_DIV_OTA_COPYED = "ota_copyed"

Dim pCopyFolder_software
Dim pCopyFolder_db
Dim pCopyFolder_ota

Sub onClickCopy()
    Call resetAllCopyedProcessStatus()
    Call setProcessStatus(STATUS_CHECK)
    Call freezeAllInput()
    idTimer = window.setTimeout("checkInfoWhenCopy()", 0, "VBScript")
End Sub

Sub checkInfoWhenCopy()
    window.clearTimeout(idTimer)
    If Not checkInfo() Then
        Call doSomethingIfCheckError()
        Exit Sub
    End If

    Call writeHistory(pTargetPathHistory, ID_INPUT_TARGET_PATH, ID_LIST_TARGET_HISTORY, ID_UL_TARGET_HISTORY, mTargetPath)

    Call setProcessStatus(STATUS_COPY)
    idTimer = window.setTimeout("startCopy()", 0, "VBScript")
End Sub

        Function checkInfo()
            '//check holded value
            Call onChangeCodePath(CHECK_WHEN_COPY)
            If mCodePath = "" Then checkInfo = False : Exit Function
            Call onChangeTargetPath(CHECK_WHEN_COPY)
            If mTargetPath = "" Then checkInfo = False : Exit Function
            Call onChangeFolderName(CHECK_WHEN_COPY)
            If mNewFolderName = "" Then checkInfo = False : Exit Function

            '//dim folder path for copy
            pCopyFolder_software = mTargetPath & "\" & mNewFolderName
            pCopyFolder_db = pCopyFolder_software & "\DB"
            pCopyFolder_ota = pCopyFolder_software & "\OTA"

            '//check folder path for copy
            If oFso.FolderExists(pCopyFolder_software) Then
                MsgBox("目标路径已存在""" & mNewFolderName & """文件夹！")
                checkInfo = False
                Exit Function
            End If

            '//check files for copy, and set element str of copy process.
            Call getDbFiles()
            Call getSoftwareFiles()
            Call modifyScatterFile()

            checkInfo = True
        End Function

Sub doSomethingIfCheckError()
    Call unfreezeAllInput()
    Call setProcessStatus(STATUS_WAIT)
End Sub

Sub startCopy()
    window.clearTimeout(idTimer)

    Call createFolderForCopy()
    
    Call copyDbFiles()

    Call copySoftware()

    Call copyOtaFile()

    Call unfreezeAllInput()
    Call setProcessStatus(STATUS_DONE)
End Sub

        Sub createFolderForCopy()
            '//create folder
            oFso.CreateFolder(pCopyFolder_software)
            oFso.CreateFolder(pCopyFolder_db)
        End Sub

        Sub copyDbFiles()
            '//start copy AP files
            If pFile_AP <> "" Then
                Call copyFile(pFile_AP, pCopyFolder_db)
                Call setElementInnerHTML(ID_DIV_DB_COPYED, Cint(getElementInnerHTML(ID_DIV_DB_COPYED)) + 1)
            End If

            '//start copy BP files
            If vaFilePath_BP.Bound > -1 Then
                Dim i
                For i = 0 To vaFilePath_BP.Bound
                    Call copyFile(vaFilePath_BP.V(i), pCopyFolder_db)
                    Call setElementInnerHTML(ID_DIV_DB_COPYED, Cint(getElementInnerHTML(ID_DIV_DB_COPYED)) + 1)
                Next
            End If
        End Sub

        Sub copySoftware()
            '//start copy SOFTWARE files
            If vaFileNamesForCopy.Bound > -1 Then
                Dim j
                For j = 0 To vaFileNamesForCopy.Bound
                    Call copyFile(mOutSoftwarePath & "\" & vaFileNamesForCopy.V(j), pCopyFolder_software)
                    Call setElementInnerHTML(ID_DIV_SOFTWARE_COPYED, Cint(getElementInnerHTML(ID_DIV_SOFTWARE_COPYED)) + 1)
                Next
                Call copyFile(mOutSoftwarePath & "\system.img", pCopyFolder_software)
                Call setElementInnerHTML(ID_DIV_SOFTWARE_COPYED, Cint(getElementInnerHTML(ID_DIV_SOFTWARE_COPYED)) + 1)
            End If
        End Sub

        Sub copyOtaFile()
            '//start copy OTA files
            If mOtaFilePath <> "" Then
                oFso.CreateFolder(pCopyFolder_ota)
                Call copyFile(mOtaFilePath, pCopyFolder_ota)
                Call setElementInnerHTML(ID_DIV_OTA_COPYED, Cint(getElementInnerHTML(ID_DIV_OTA_COPYED)) + 1)
            End If
        End Sub

Sub copyFile(uCopyFilePath, uTargetFolderPath)
    If oFso.FileExists(uCopyFilePath) Then
        Dim filePath : filePath = """" & uCopyFilePath & """"
        Dim folderPath : folderPath = """" & uTargetFolderPath & """"
        
        oWs.Run "src\copyFiles\FsoCopyFile.vbs " & filePath & " " & folderPath & "\", , True
    Else
       MsgBox(uCopyFilePath & " is not exist!")
    End If
End Sub