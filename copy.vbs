Option Explicit

Const WINDOW_WIDTH = 400
Const WINDOW_HEIGHT = 750
Sub Window_OnLoad
    Dim ScreenWidth : ScreenWidth = CreateObject("HtmlFile").ParentWindow.Screen.AvailWidth
    Dim ScreenHeight : ScreenHeight = CreateObject("HtmlFile").ParentWindow.Screen.AvailHeight
    Window.MoveTo ScreenWidth - WINDOW_WIDTH ,(ScreenHeight - WINDOW_HEIGHT) \ 3
    Window.ResizeTo WINDOW_WIDTH, WINDOW_HEIGHT
End Sub

Dim oWs : set oWs = CreateObject("Wscript.Shell")
Dim oFso : Set oFso = CreateObject("Scripting.FileSystemObject")

Dim pAllCodePath : pAllCodePath = oWs.CurrentDirectory & "\allCodePath.txt"
Dim pInputHistory : pInputHistory = oWs.CurrentDirectory & "\InputHistory.txt"

Dim bDebug : bDebug = False

Const FOR_READING = 1
Const FOR_APPENDING = 8
Const FROM_SELECT = 0
Const FROM_INPUT = 1
Const ID_SELECT_CODE_PATH = "select_code_path"
Const ID_INPUT_CODE_PATH = "input_code_path"
Const ID_INPUT_TARGET_FOLDER_PATH = "target_folder_path"
Const ID_UL_INPUT_HISTORY = "input_history"
Const ID_RADIO_CREATE_NEW_FOLDER = "radio_create_new_folder"
Const ID_INPUT_NEW_FOLDER_NAME = "new_folder_name"
Const ID_SELECT_OTA_FILE = "select_ota_file"
Const ID_DIV_DB_COPYED = "db_copyed"
Const ID_DIV_SOFTWARE_COPYED = "software_copyed"
Const ID_DIV_OTA_COPYED = "ota_copyed"
Const ID_DIV_DB_WAIT = "db_wait"
Const ID_DIV_SOFTWARE_WAIT = "software_wait"
Const ID_DIV_OTA_WAIT = "ota_wait"
Const ID_DIV_COPY_STATUS = "copy_status"

Const INVALID_STR_OF_FOlDER_NAME = "/\:*?""<>|"

Dim pOutFolder, pOutProjectFolder, pTargetFolder, sNewFolderName, pFile_AP, pOtaFile
pOutFolder = ""
pOutProjectFolder = ""
pTargetFolder = ""
sNewFolderName = ""
pOtaFile = ""
Dim vaFolderPath_BP : Set vaFolderPath_BP = New VariableArray
Dim vaFilePath_BP : Set vaFilePath_BP = New VariableArray
Dim vaFileNamesForCopy : Set vaFileNamesForCopy = New VariableArray
Dim vaHistoryTargetFolder : Set vaHistoryTargetFolder = New VariableArray

Call ReadCodePath()
Call getOutPath(FROM_SELECT)
Call readHistory(pInputHistory, ID_UL_INPUT_HISTORY)

Sub ReadCodePath()
    If Not oFso.FileExists(pAllCodePath) Then
        MsgBox("代码路径文件不存在！")
        Exit Sub
    End If

    Dim oTxt : Set oTxt = oFso.OpenTextFile(pAllCodePath, FOR_READING)

    Dim sTmp
    Do Until oTxt.AtEndOfStream
        sTmp = oTxt.ReadLine
        If Trim(sTmp) <> "" Then Call jsAddOption(ID_SELECT_CODE_PATH, sTmp)
    Loop
    oTxt.Close
    Set oTxt = Nothing
End Sub

Sub readHistory(DictPath, ulId)
    If Not OFso.FileExists(DictPath) Then
        Exit Sub
    End If    
    
    Dim oDict : Set oDict = oFso.OpenTextFile(DictPath, FOR_READING)

    Do Until oDict.AtEndOfStream
        Dim sTmp : sTmp = oDict.ReadLine
        If sTmp <> "" Then
            Call addBeforeLi(sTmp, ulId)
            vaHistoryTargetFolder.Append(sTmp)
        End If
    Loop

    Set oDict = Nothing
End Sub

Sub writeHistory(DictPath, ulId, str)
    Dim seqInArray : seqInArray = vaHistoryTargetFolder.IsExistInArray(str)
    If seqInArray > -1 Then
        Call removeLi(ulId, vaHistoryTargetFolder.Length - seqInArray)
        Call addBeforeLi(str, ulId)

        vaHistoryTargetFolder.MoveToEnd(seqInArray)

        Call writeNewHistoryTxt(DictPath)
        Exit Sub
    End If


    If vaHistoryTargetFolder.Length < 7 Then
        Call addBeforeLi(str, ulId)

        vaHistoryTargetFolder.Append(str)

        Call appendStrToHistoryTxt(DictPath, str)
    Else
        Call removeLi(ulId, 0)
        Call addBeforeLi(str, ulId)

        vaHistoryTargetFolder.PopBySeq(0)
        vaHistoryTargetFolder.Append(str)

        Call writeNewHistoryTxt(DictPath)
    End If
End Sub

Sub appendStrToHistoryTxt(DictPath, str)
    Dim oDict
    If Not OFso.FileExists(DictPath) Then
        Call oFso.CreateTextFile(DictPath, True)
    End If
    
    Set oDict = oFso.OpenTextFile(DictPath, FOR_APPENDING)
    oDict.WriteLine(str)
    oDict.Close
    Set oDict = Nothing
End Sub

Sub writeNewHistoryTxt(DictPath)
    Call initTxtFile(DictPath)
    Dim oDict
    Set oDict = oFso.OpenTextFile(DictPath, FOR_APPENDING)
    Dim i
    For i = 0 To vaHistoryTargetFolder.Length
        oDict.WriteLine(vaHistoryTargetFolder.Value(i))
    Next
    oDict.Close
    Set oDict = Nothing
End Sub

Sub initTxtFile(FilePath)
    If oFso.FileExists(FilePath) Then
        Dim TxtFile
        Set TxtFile = oFso.getFile(FilePath)
        TxtFile.Delete
        Set TxtFile = Nothing
    End If
    oFso.CreateTextFile FilePath, True
End Sub

Sub setValueOfTargetFolder(value)
    Call showAndHide("List","hide")
    Call setElementValue(ID_INPUT_TARGET_FOLDER_PATH, value)
    Call getTargetPath()
End Sub

Sub openTargetFolder()
    Dim sTmp : sTmp = getElementValue(ID_INPUT_TARGET_FOLDER_PATH)
    If sTmp = "" Then Exit Sub
    If oFso.FolderExists(sTmp) Then oWs.run "explorer.exe " & sTmp
End Sub

Class VariableArray
    Private mLength, mArray()

    Private Sub Class_Initialize
        mLength = -1
    End Sub

    Public Property Get Length
        Length = mLength
    End Property

    Public Property Get Value(seq)
        If Not isNumeric(seq) Then
            MsgBox("Error: Get Value(seq) seq is not a number")
            Exit Property
        ELse
            seq = Cint(seq)
        End If

        'MsgBox("seq="&seq&" mLength="&mLength)
        If seq < 0 Or seq > mLength Then
            MsgBox("Error: Get Value(seq) seq out of bound")
            Exit Property
        End If

        If isObject(mArray(seq)) Then
            Set Value = mArray(seq)
        Else
            Value = mArray(seq)
        End If
    End Property

    Public Property Let Value(seq, sValue)
        If Not isNumeric(seq) Then
            MsgBox("Error: Let Value(seq) seq is not a number")
            Exit Property
        ELse
            seq = Cint(seq)
        End If

        If seq < 0 Or seq > mLength Then
            MsgBox("Error: Let Value(seq) seq out of bound")
            Exit Property
        End If

        mArray(seq) = sValue
    End Property

    Public Function Append(value)
        mLength = mLength + 1
        ReDim Preserve mArray(mLength)

        If isObject(value) Then
            Set mArray(mLength) = value
        ELse
            mArray(mLength) = value
        End If
    End Function

    Public Function ResetArray()
        mLength = -1
    End Function

    Public Property Let InnerArray(newArray)
        If Not isArray(newArray) Then
            MsgBox("Error: Set InnerArray(newArray) newArray is not array")
            Exit Property
        End If

        Dim i
        For i = 0 To UBound(newArray)
            mLength = mLength + 1
            ReDim Preserve mArray(mLength)
            mArray(mLength) = newArray(i)
        Next
    End Property

    Public Function PopBySeq(seq)
        If Not isNumeric(seq) Then
            MsgBox("Error: PopBySeq(seq) seq is not a number")
            Exit Function
        ELse
            seq = Cint(seq)
        End If

        If seq < 0 Or seq > mLength Then
            MsgBox("Error: PopBySeq(seq) seq out of bound")
            Exit Function
        End If

        If seq <> mLength Then
            Dim i
            For i = seq To mLength - 1
                mArray(i) = mArray(i + 1)
            Next
        End If

        mLength = mLength - 1
        ReDim Preserve mArray(mLength)
    End Function

    Public Function MoveToTop(seq)
        If Not isNumeric(seq) Then
            MsgBox("Error: MoveToTop(seq) seq is not a number")
            Exit Function
        ELse
            seq = Cint(seq)
        End If

        If seq < 0 Or seq > mLength Then
            MsgBox("Error: MoveToTop(seq) seq out of bound")
            Exit Function
        End If

        If seq = 0 Then Exit Function

        Dim i, sValueToBeMove
        If isObject(mArray(seq)) Then
            Set sValueToBeMove = mArray(seq)
            For i = seq To 1 Step -1
                Set mArray(i) = mArray(i - 1)
                Set mArray(0) = sValueToBeMove
            Next
        Else
            sValueToBeMove = mArray(seq)
            For i = seq To 1 Step -1
                mArray(i) = mArray(i - 1)
                mArray(0) = sValueToBeMove
            Next
        End If
    End Function

    Public Function MoveToEnd(seq)
        If Not isNumeric(seq) Then
            MsgBox("Error: MoveToEnd(seq) seq is not a number")
            Exit Function
        ELse
            seq = Cint(seq)
        End If

        If seq < 0 Or seq > mLength Then
            MsgBox("Error: MoveToEnd(seq) seq out of bound")
            Exit Function
        End If

        If seq = 0 Then Exit Function

        Dim i, sValueToBeMove
        If isObject(mArray(seq)) Then
            Set sValueToBeMove = mArray(seq)
            For i = seq To mLength - 1
                Set mArray(i) = mArray(i + 1)
            Next
            Set mArray(mLength) = sValueToBeMove
        Else
            sValueToBeMove = mArray(seq)
            For i = seq To mLength - 1
                mArray(i) = mArray(i + 1)
            Next
            mArray(mLength) = sValueToBeMove
        End If
    End Function

    Public Function IsExistInArray(value)
        If mLength = -1 Then
            IsExistInArray = -1
            Exit Function
        End If

        Dim i
        For i = 0 To mLength
            If StrComp(mArray(i), value) = 0 Then
                IsExistInArray = i
                Exit Function
            End If
        Next
        IsExistInArray = -1
    End Function

    Public Function SortArray()
        If mLength = -1 Then
            MsgBox("Error: SortArray() mLength <= 0, no need to sort")
            Exit Function
        End If

        Dim i, j
        For i = 0 To mLength - 1
            For j = i + 1 To mLength
                If StrComp(mArray(i), mArray(j)) > 0 Then
                    Dim sTmp : sTmp = mArray(i) : mArray(i) = mArray(j) : mArray(j) = sTmp
                End If
            Next
        Next
    End Function

    Public Function ToString()
        If mLength <> -1 Then
            Dim i, sTmp
            sTmp = "v(0) = " & mArray(0)
            If mLength > 0 Then
                For i = 1 To mLength
                    If isArray(mArray(i)) Then
                        sTmp = sTmp & Vblf & "v(" & i & ") = " & join(mArray(i))
                    ElseIf isObject(mArray(i)) Then
                        sTmp = sTmp & Vblf & "v(" & i & ") = [Object]"
                    Else
                        sTmp = sTmp & Vblf & "v(" & i & ") = " & mArray(i)
                    End If
                Next
            End If
            ToString = sTmp
        Else
            MsgBox("Error: ToString() mArray has no element")
        End If
    End Function
End Class

Const SEARCH_FILE = 0
Const SEARCH_FOLDER = 1
Const SEARCH_ROOT = 0
Const SEARCH_SUB = 1
Const SEARCH_WHOLE_NAME = 0
Const SEARCH_PART_NAME = 1

Function searchFolder(pRootFolder, str, searchType, searchWhere, searchMode)
    If Not oFso.FolderExists(pRootFolder) Then searchFolder = "" : Exit Function

    Dim oRootFolder : Set oRootFolder = oFso.GetFolder(pRootFolder)

    If searchType = SEARCH_FILE And searchWhere = SEARCH_ROOT Then
        '//SEARCH_FILE and SEARCH_ROOT
        If searchMode = SEARCH_WHOLE_NAME Then
            searchFolder = startSearch(oRootFolder.Files, pRootFolder, str, searchMode)
            Exit Function
        ElseIf searchMode = SEARCH_PART_NAME Then
            searchFolder = startSearch(oRootFolder.Files, pRootFolder, str, searchMode)
            Exit Function
        End If
    Else
        Dim oSubFolders : Set oSubFolders = oRootFolder.SubFolders

        If searchType = SEARCH_FOLDER And searchWhere = SEARCH_ROOT Then
            If searchMode = SEARCH_WHOLE_NAME Then
                searchFolder = startSearch(oSubFolders, pRootFolder, str, searchMode)
                Exit Function
            ElseIf searchMode = SEARCH_PART_NAME Then
                searchFolder = startSearch(oSubFolders, pRootFolder, str, searchMode)
                Exit Function
            End If
        End If

        '//searchWhere = SEARCH_SUB
        Dim Folder, sTmp
        If searchType = SEARCH_FILE Then
            If searchMode = SEARCH_WHOLE_NAME Then
                For Each Folder In oSubFolders
                    sTmp = startSearch(Folder.Files, pRootFolder & "\" & Folder.Name, str, searchMode)
                    If sTmp <> "" Then
                        searchFolder = sTmp
                        Exit Function
                    End If
                Next
            ElseIf searchMode = SEARCH_PART_NAME Then
                For Each Folder In oSubFolders
                    sTmp = startSearch(Folder.Files, pRootFolder & "\" & Folder.Name, str, searchMode)
                    If sTmp <> "" Then
                        searchFolder = sTmp
                        Exit Function
                    End If
                Next
            End If
        ElseIf searchType = SEARCH_FOLDER Then
            If searchMode = SEARCH_WHOLE_NAME Then
                For Each Folder In oSubFolders
                    sTmp = startSearch(Folder.SubFolders, pRootFolder & "\" & Folder.Name, str, searchMode)
                    If sTmp <> "" Then
                        searchFolder = sTmp
                        Exit Function
                    End If
                Next
            ElseIf searchMode = SEARCH_PART_NAME Then
                For Each Folder In oSubFolders
                    sTmp = startSearch(Folder.SubFolders, pRootFolder & "\" & Folder.Name, str, searchMode)
                    If sTmp <> "" Then
                        searchFolder = sTmp
                        Exit Function
                    End If
                Next
            End If
        End If
    End If
    searchFolder = ""
End Function

        Function startSearch(oAll, pRootFolder, str, searchMode)
            Dim oSingle
            For Each oSingle In oAll
                If checkSearchName(oSingle.Name, str, searchMode) Then
                    startSearch = pRootFolder & "\" & oSingle.Name
                    Exit Function
                End If
            Next
            startSearch = ""
        End Function

        Function checkSearchName(name, str, searchMode)
            If searchMode = SEARCH_WHOLE_NAME Then
                If StrComp(name ,str) = 0 Then
                    checkSearchName = True
                Else
                    checkSearchName = False
                End If
            ELseIf searchMode = SEARCH_PART_NAME Then
                If InStr(name ,str) > 0 Then
                    checkSearchName = True
                Else
                    checkSearchName = False
                End If
            End If
        End Function

Function getFileNameOfPath(path)
    getFileNameOfPath = Mid(path, InStrRev(path, "\") + 1)
End Function

Sub Sleep(MSecs)  
    If Not oFso.FileExists("sleeper.vbs") Then
        Dim objOutputFile : Set objOutputFile = oFso.CreateTextFile("sleeper.vbs", True)
        objOutputFile.Write "wscript.sleep WScript.Arguments(0)"
        objOutputFile.Close
        Set objOutputFile = Nothing
    End If
    CreateObject("WScript.Shell").Run "sleeper.vbs " & MSecs,1 , True
End Sub

Function getCheckedRadio(name)
    Dim radioObj, i
    Set radioObj = document.getElementsByName(name)
    For i = 0 To radioObj.length
        If radioObj(i).checked Then
            getCheckedRadio =  radioObj(i).value
            Exit For
        End If
    Next
End Function

Function getElementValue(elementId)
    getElementValue = document.getElementById(elementId).value
End Function

Sub setElementValue(elementId, str)
    document.getElementById(elementId).value = str
End Sub

Function getElementInnerHTML(elementId)
    getElementInnerHTML = document.getElementById(elementId).innerHTML
End Function

Sub setElementInnerHTML(elementId, str)
    document.getElementById(elementId).innerHTML = str
End Sub

Sub setElementColor(elementId, colorStr)
    document.getElementById(elementId).style.color = colorStr
End Sub

Const MY_COMPUTER = &H11&
Const WINDOW_HANDLE = 0 
Const OPTIONS = 0
Sub getBrowseValue()
    Dim objShell : Set objShell = CreateObject("Shell.Application") 
    Dim objFolder : Set objFolder = objShell.Namespace(MY_COMPUTER) 
    Dim objFolderItem : Set objFolderItem = objFolder.Self 
    Dim strPath : strPath = objFolderItem.Path

    Set objShell = Nothing
    Set objFolder = Nothing

    Set objShell = CreateObject("Shell.Application") 
    Set objFolder = objShell.BrowseForFolder(WINDOW_HANDLE, "Select a folder:", OPTIONS, strPath) 
    If Not objFolder Is Nothing Then 
        Set objFolderItem = objFolder.Self
        Call setElementValue(ID_INPUT_TARGET_FOLDER_PATH, objFolderItem.Path)
        Call getTargetPath()
    End If 

    Set objShell = Nothing
    Set objFolder = Nothing
End Sub

Sub clearProcess()
    Call setElementInnerHTML(ID_DIV_DB_WAIT, 0)
    Call setElementInnerHTML(ID_DIV_SOFTWARE_WAIT, 0)
    Call setElementInnerHTML(ID_DIV_OTA_WAIT, 0)
    Call setElementInnerHTML(ID_DIV_DB_COPYED, 0)
    Call setElementInnerHTML(ID_DIV_SOFTWARE_COPYED, 0)
    Call setElementInnerHTML(ID_DIV_OTA_COPYED, 0)
    Call setElementInnerHTML(ID_DIV_COPY_STATUS, "未开始")
    Call setElementColor(ID_DIV_COPY_STATUS, "black")
End Sub

Sub clearOtherElements()
    Call clearProcess()
    Call setElementValue(ID_INPUT_TARGET_FOLDER_PATH, "")
    Call setElementValue(ID_INPUT_NEW_FOLDER_NAME, "")
    Call jsRemoveAllOption(ID_SELECT_OTA_FILE)
    Call jsAddOption(ID_SELECT_OTA_FILE, "如果编了OTA，请点击下面的按钮")
End Sub

Sub clearHoldedValues()
    pOutFolder = ""
    pOutProjectFolder = ""
    pOtaFile = ""
End Sub



'///////////////////////////////////////////////////////'
'check input
'///////////////////////////////////////////////////////'
Sub getOutPath(where)
    If bDebug Then MsgBox("getOutPath 111")

    Call clearOtherElements()
    Call clearHoldedValues()

    Dim sTmp
    If StrComp(where, FROM_SELECT) = 0 Then
        Call setElementValue(ID_INPUT_CODE_PATH, "")

         sTmp = getElementValue(ID_SELECT_CODE_PATH)

        If Not oFso.FolderExists(sTmp) Then MsgBox("路径不存在:" & Vblf & sTmp) : Exit Sub
        If Not oFso.FolderExists(sTmp & "\out\target\product") Then MsgBox("路径下不存在\out\target\product") : Exit Sub

        pOutFolder = sTmp & "\out\target\product"
    ElseIf StrComp(where, FROM_INPUT) = 0 Then
        sTmp = getElementValue(ID_INPUT_CODE_PATH)

        If sTmp = "" Then pOutFolder = getElementValue(ID_SELECT_CODE_PATH) & "\out\target\product" : Exit Sub
        If Not oFso.FolderExists(sTmp) Then MsgBox("路径不存在:" & Vblf & sTmp) : Exit Sub
        If Not oFso.FolderExists(sTmp & "\target\product") Then MsgBox("路径下不存在\target\product") : Exit Sub

        pOutFolder = sTmp & "\target\product"
    End If
End Sub

Sub getOutProjectPath()
    If Not oFso.FolderExists(pOutFolder) Then
        MsgBox("路径不存在: " & Vblf & pOutFolder)
        pOutProjectFolder = ""
        Exit Sub
    End If

    Dim pSystemimg : pSystemimg = searchFolder(pOutFolder, "system.img", SEARCH_FILE, SEARCH_SUB, SEARCH_WHOLE_NAME)

    If pSystemimg = "" Then
        MsgBox("""system.img"" is not exists in: " & Vblf & pOutFolder)
        pOutProjectFolder = ""
        Exit Sub
    End If

    pOutProjectFolder = Replace(pSystemimg, "\system.img", "")
End Sub

Sub checkSoftwareFiles()
    If bDebug Then MsgBox("checkSoftwareFiles")
    vaFileNamesForCopy.ResetArray()

    Dim uScatterFilePath : uScatterFilePath = searchFolder(pOutProjectFolder, "_Android_scatter.txt" _
            , SEARCH_FILE, SEARCH_ROOT, SEARCH_PART_NAME)

    If uScatterFilePath = "" Then
        MsgBox("""Android_scatter.txt"" is not exists in: " & Vblf & pOutProjectFolder)
        Exit Sub
    End If

    Call getNeedFilesName(uScatterFilePath)

    If vaFileNamesForCopy.Length > -1 Then
        Call setElementInnerHTML(ID_DIV_SOFTWARE_WAIT, vaFileNamesForCopy.Length + 2)
    End If
End Sub

Sub getNeedFilesName(uPath)
    If bDebug Then MsgBox("getNeedFilesName")

    vaFileNamesForCopy.Append(getFileNameOfPath(uPath))

    Dim oTxt, sReadLine, sTmpFileName, iInFileName, iInNone
    Set oTxt = oFso.OpenTextFile(uPath, FOR_READING)
    Do Until oTxt.AtEndOfStream
        sReadLine = oTxt.ReadLine
        iInFileName = InStr(sReadLine, "file_name")
        iInNone = InStr(sReadLine, "NONE")
        If iInFileName > 0 And iInNone = 0 Then
            sTmpFileName = Mid(sReadLine, iInFileName + 11)
            If vaFileNamesForCopy.IsExistInArray(sTmpFileName) = -1 And sTmpFileName <> "system.img" Then
                vaFileNamesForCopy.Append(Mid(sReadLine, iInFileName + 11))
            End If
        End If
    Loop
    oTxt.Close
    Set oTxt = Nothing
    vaFileNamesForCopy.SortArray()
End Sub

Sub checkDbFiles()
    If bDebug Then MsgBox("checkDbFiles")
    pFile_AP = ""
    vaFolderPath_BP.ResetArray()
    vaFilePath_BP.ResetArray()

    Dim sKK_AP : sKK_AP = "\obj\CODEGEN\cgen"
    Dim sKK_BP : sKK_BP = "\obj\CUSTGEN\custom\modem"
    Dim sL1_AP : sL1_AP = "\obj\CGEN"
    Dim sL1_BP : sL1_BP = "\obj\ETC"

    Dim pFolder_AP

    '//get AP file path
    Select Case True
        Case oFso.FolderExists(pOutProjectFolder & sKK_AP)
            pFolder_AP = pOutProjectFolder & sKK_AP
        Case oFso.FolderExists(pOutProjectFolder & sL1_AP)
            pFolder_AP = pOutProjectFolder & sL1_AP
        Case Else
            pFolder_AP = ""
    End Select

    If pFolder_AP <> "" Then pFile_AP = searchFolder(pFolder_AP, "_ENUM", SEARCH_FILE, SEARCH_ROOT, SEARCH_PART_NAME)
    If pFile_AP <> "" Then pFile_AP = Replace(pFile_AP, "_ENUM", "")

    If pFile_AP = "" Then MsgBox("AP file is not exists!")

    '//get BP file path
    Dim sTmp
    Select Case True
        Case oFso.FolderExists(pOutProjectFolder & sKK_BP)
            vaFolderPath_BP.Append(pOutProjectFolder & sKK_BP)
        Case oFso.FolderExists(pOutProjectFolder & sL1_BP)
            sTmp = searchFolder(pOutProjectFolder & sL1_BP, "BPLGU", SEARCH_FOLDER, SEARCH_ROOT, SEARCH_PART_NAME)
            If sTmp <> "" Then vaFolderPath_BP.Append(sTmp)
    End Select

    If vaFolderPath_BP.Length > -1 Then
        Dim i 
        For i = 0 To vaFolderPath_BP.Length
            sTmp = searchFolder(vaFolderPath_BP.Value(i), "BPLGU", SEARCH_FILE, SEARCH_ROOT, SEARCH_PART_NAME)
            If sTmp <> "" Then vaFilePath_BP.Append(sTmp)
        Next
    End If

    'If vaFilePath_BP.Length = -1 Then MsgBox("BP file is not exists!")

    If pFile_AP <> "" Then
        Call setElementInnerHTML(ID_DIV_DB_WAIT, vaFilePath_BP.Length + 2)
    ELse
        Call setElementInnerHTML(ID_DIV_DB_WAIT, vaFilePath_BP.Length + 1)
    End If
End Sub

Sub checkOtaFiles()
    If bDebug Then MsgBox("checkOtaFiles")
    If getElementValue(ID_SELECT_OTA_FILE) <> "如果编了OTA，请点击下面的按钮" Then Exit Sub

    If pOutProjectFolder = "" Then Call getOutProjectPath()

    Dim pOta_1 : pOta_1 = searchFolder(pOutProjectFolder, "target_files-package.zip" _
            , SEARCH_FILE, SEARCH_ROOT, SEARCH_WHOLE_NAME)
    Dim pOta_2 : pOta_2 = searchFolder(pOutProjectFolder, "-ota-"_
            , SEARCH_FILE, SEARCH_ROOT, SEARCH_PART_NAME)
    Dim pOta_3 : pOta_3 = searchFolder(pOutProjectFolder & "\obj\PACKAGING\target_files_intermediates", "-target_files-" _
            , SEARCH_FILE, SEARCH_ROOT, SEARCH_PART_NAME)

    If pOta_1 = "" And pOta_2 = "" And pOta_3 = "" Then
        Call jsRemoveAllOption(ID_SELECT_OTA_FILE)
        Call jsAddOption(ID_SELECT_OTA_FILE, "无OTA文件")
        Exit Sub
    End If

    Call jsRemoveAllOption(ID_SELECT_OTA_FILE)

    Call jsAddOption(ID_SELECT_OTA_FILE, "不拷贝OTA")
    If pOta_1 <> "" Then Call jsAddOption(ID_SELECT_OTA_FILE, Replace(pOta_1, pOutProjectFolder, ""))
    If pOta_2 <> "" Then Call jsAddOption(ID_SELECT_OTA_FILE, Replace(pOta_2, pOutProjectFolder, ""))
    If pOta_3 <> "" Then Call jsAddOption(ID_SELECT_OTA_FILE, Replace(pOta_3, pOutProjectFolder, ""))
End Sub

Sub getOtaFiles()
    Dim pPartOtaFile : pPartOtaFile = getElementValue(ID_SELECT_OTA_FILE)
    'If pPartOtaFile = "无OTA文件" Then Exit Sub
    If pPartOtaFile = "不拷贝OTA" Then
        pOtaFile = ""
        Call setElementInnerHTML(ID_DIV_OTA_WAIT, 0)
    Else
        pOtaFile = pOutProjectFolder & pPartOtaFile
        Call setElementInnerHTML(ID_DIV_OTA_WAIT, 1)
    End If
End Sub

Sub getTargetPath()
    Call clearProcess()

    pTargetFolder = getElementValue(ID_INPUT_TARGET_FOLDER_PATH)
    If pTargetFolder = "" Then Exit Sub

    If Not oFso.FolderExists(pTargetFolder) Then
        If MsgBox(pTargetFolder & "不存在，是否创建该目录？", 4) = 6 Then
            oFso.CreateFolder(pTargetFolder)
        Else
            Call setElementValue(ID_INPUT_TARGET_FOLDER_PATH, "")
            pTargetFolder = ""
        End If
    End If
End Sub

Sub getNewFolderName()
    Call clearProcess()

    sNewFolderName = getElementValue(ID_INPUT_NEW_FOLDER_NAME)

    If Trim(sNewFolderName) = "" Then
        MsgBox("请输入新文件夹名")
        sNewFolderName = ""
        Exit Sub
    End If

    Dim i, sTmp
    For i = 1 To Len(sNewFolderName)
        sTmp = Mid(sNewFolderName, i, 1)
        If InStr(INVALID_STR_OF_FOlDER_NAME, sTmp) Then
            MsgBox("文件名不能包含下列任何字符：" & Vblf & INVALID_STR_OF_FOlDER_NAME)
            sNewFolderName = ""
            Exit Sub
        End If
    Next
End Sub

Sub getCustomBuildVersion(which)
    If pOutProjectFolder = "" Then Call getOutProjectPath()

    Dim sCheckStr
    If which = "display.id" Then
        sCheckStr = "ro.build.display.id"
    Else
        sCheckStr = "ro.custom.build.version"
    End If

    Dim pBuildProp : pBuildProp = pOutProjectFolder & "\system\build.prop"
    If Not oFso.FileExists(pBuildProp) Then
        MsgBox("文件不存在: " & Vblf & pBuildProp)
        Exit Sub
    End If

    Dim oTxt : Set oTxt = oFso.OpenTextFile(pBuildProp, FOR_READING)
    Dim sTmp
    Do Until oTxt.AtEndOfStream
        sTmp = oTxt.ReadLine
        If InStr(sTmp, sCheckStr) > 0 Then
            Call setElementValue(ID_INPUT_NEW_FOLDER_NAME, Trim(Mid(sTmp,InStr(sTmp,"=")+1)))
            Call getNewFolderName()
            Exit Do
        End If
    Loop
    oTxt.Close
    Set oTxt = Nothing
End Sub



Sub runCopy()
    '//check holded value
    If pOutFolder = "" Then MsgBox("未找到out目录") : Exit Sub
    If pOutProjectFolder = "" Then Call getOutProjectPath()
    If pOutProjectFolder = "" Then MsgBox("未找到out下的system.img") : Exit Sub

    If pTargetFolder = "" Then MsgBox("请输入目标路径") : Exit Sub
    If sNewFolderName = "" Then MsgBox("请输入新文件夹名") : Exit Sub


    '//set running status of copy process
    Call clearProcess()
    Call setElementInnerHTML(ID_DIV_COPY_STATUS, "拷贝中...")
    Call setElementColor(ID_DIV_COPY_STATUS, "red")
    
    '//dim folder path for copy
    Dim pCopyFolder_software : pCopyFolder_software = pTargetFolder & "\" & sNewFolderName
    Dim pCopyFolder_db : pCopyFolder_db = pCopyFolder_software & "\DB"
    Dim pCopyFolder_ota : pCopyFolder_ota = pCopyFolder_software & "\OTA"

    '//check folder path for copy
    If oFso.FolderExists(pCopyFolder_software) Then MsgBox("目标路径已存在""" & sNewFolderName & """文件夹！") : Exit Sub

    '//create folder
    oFso.CreateFolder(pCopyFolder_software)
    oFso.CreateFolder(pCopyFolder_db)


    '//check files for copy, and set element str of copy process.
    Call checkDbFiles()
    Call checkSoftwareFiles()
    If pOtaFile <> "" Then Call setElementInnerHTML(ID_DIV_OTA_WAIT, 1)


    '//start copy AP files
    If pFile_AP <> "" Then
        Call copyFile(pFile_AP, pCopyFolder_db)
        Call setElementInnerHTML(ID_DIV_DB_COPYED, Cint(getElementInnerHTML(ID_DIV_DB_COPYED)) + 1)
    End If

    '//start copy BP files
    If vaFilePath_BP.Length > -1 Then
        Dim i
        For i = 0 To vaFilePath_BP.Length
            Call copyFile(vaFilePath_BP.Value(i), pCopyFolder_db)
            Call setElementInnerHTML(ID_DIV_DB_COPYED, Cint(getElementInnerHTML(ID_DIV_DB_COPYED)) + 1)
        Next
    End If

    '//start copy SOFTWARE files
    If vaFileNamesForCopy.Length > -1 Then
        Dim j
        For j = 0 To vaFileNamesForCopy.Length
            Call copyFile(pOutProjectFolder & "\" & vaFileNamesForCopy.Value(j), pCopyFolder_software)
            Call setElementInnerHTML(ID_DIV_SOFTWARE_COPYED, Cint(getElementInnerHTML(ID_DIV_SOFTWARE_COPYED)) + 1)
        Next
        Call copyFile(pOutProjectFolder & "\system.img", pCopyFolder_software)
        Call setElementInnerHTML(ID_DIV_SOFTWARE_COPYED, Cint(getElementInnerHTML(ID_DIV_SOFTWARE_COPYED)) + 1)
    End If

    '//start copy OTA files
    If pOtaFile <> "" Then
        oFso.CreateFolder(pCopyFolder_ota)
        Call copyFile(pOtaFile, pCopyFolder_ota)
        Call setElementInnerHTML(ID_DIV_OTA_COPYED, Cint(getElementInnerHTML(ID_DIV_OTA_COPYED)) + 1)
    End If

    '//set finish status of copy process
    Call setElementInnerHTML(ID_DIV_COPY_STATUS, "拷贝完成！")
    Call setElementColor(ID_DIV_COPY_STATUS, "green")

    '//save history of target folder path
    Call writeHistory(pInputHistory, ID_UL_INPUT_HISTORY, pTargetFolder)
End Sub

Sub copyFile(uCopyFilePath, uTargetFolderPath)
    If oFso.FileExists(uCopyFilePath) Then
        Dim filePath : filePath = """" & uCopyFilePath & """"
        Dim folderPath : folderPath = """" & uTargetFolderPath & """"
        
        oWs.Run "FsoCopyFile.vbs " & filePath & " " & folderPath & "\", , True
    Else
       MsgBox(uCopyFilePath & " is not exist!")
    End If
End Sub