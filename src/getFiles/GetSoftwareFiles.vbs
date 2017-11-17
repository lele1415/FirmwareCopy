Const ID_DIV_SOFTWARE_WAIT = "software_wait"
Const ID_CHECKBOX_COPY_VERIFIED = "copy_verified"

Dim vaFileNamesForCopy : Set vaFileNamesForCopy = New VariableArray
Dim uScatterFilePath

Sub getSoftwareFiles()
    vaFileNamesForCopy.ResetArray()

    uScatterFilePath = searchFolder(mOutSoftwarePath, "_Android_scatter.txt" _
            , SEARCH_FILE, SEARCH_ROOT, SEARCH_PART_NAME, SEARCH_ONE, SEARCH_RETURN_PATH)

    If uScatterFilePath = "" Then
        MsgBox("""Android_scatter.txt"" is not exists in: " & Vblf & mOutSoftwarePath)
        Exit Sub
    End If

    Call getNeedFilesName()

    If elementIsChecked(ID_CHECKBOX_COPY_VERIFIED) Then Call getVerifiedFiles()

    If vaFileNamesForCopy.Bound > -1 Then
        Call setElementInnerHTML(ID_DIV_SOFTWARE_WAIT, vaFileNamesForCopy.Bound + 2)
    End If
End Sub

Sub getNeedFilesName()
    '//add Android_scatter.txt first
    vaFileNamesForCopy.Append(getFileNameOfPath(uScatterFilePath))

    Call readTextAndDoSomething(uScatterFilePath, _
            "Call checkAndAddFileNameForCopy(sReadLine)")

    vaFileNamesForCopy.SortArray()
End Sub

Sub getVerifiedFiles()
    Dim vaVerifiedFileNames
    Set vaVerifiedFileNames = searchFolder(mOutSoftwarePath, "-verified" _
            , SEARCH_FILE, SEARCH_ROOT, SEARCH_PART_NAME, SEARCH_ALL, SEARCH_RETURN_NAME)

    If vaVerifiedFileNames.Bound = -1 Then
        MsgBox("不存在verified文件")
        Exit Sub
    End If

    Dim i
    For i = 0 To vaVerifiedFileNames.Bound
        vaFileNamesForCopy.Append(vaVerifiedFileNames.V(i))
    Next
End Sub

Sub checkAndAddFileNameForCopy(sReadLine)
    Dim iInFileName : iInFileName = InStr(sReadLine, "file_name")
    If iInFileName > 0 Then
        Dim sTmpFileName : sTmpFileName = Mid(sReadLine, iInFileName + 11)
        If sTmpFileName <> "NONE" And _
                sTmpFileName <> "system.img" And _
                vaFileNamesForCopy.GetIndexIfExist(sTmpFileName) = -1 Then
            vaFileNamesForCopy.Append(sTmpFileName)
        End If
    End If
End Sub

Function getFileNameOfPath(path)
    getFileNameOfPath = Mid(path, InStrRev(path, "\") + 1)
End Function
