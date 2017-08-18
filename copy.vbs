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
Const FROM_INPUT = 0
Const FROM_CHECK = 1
Const PROCESS_WAIT = 0
Const PROCESS_COPY = 1
'Const ID_SELECT_CODE_PATH = "select_code_path"
Const ID_INPUT_CODE_PATH = "input_code_path"
Const ID_LIST_CODE_PATH = "list_code_path"
Const ID_UL_CODE_PATH = "ul_code_path"
Const ID_INPUT_TARGET_FOLDER_PATH = "target_folder_path"
Const ID_LIST_TARGET_HISTORY = "list_target_history"
Const ID_UL_TARGET_HISTORY = "ul_target_history"
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

Const INVALID_CHAR_OF_FOlDER_NAME = "/\:*?""<>|"

Dim pFile_AP

Dim stOutFolder : Set stOutFolder = New StatusHolder
Dim stOutProjectFolder : Set stOutProjectFolder = New StatusHolder
Dim stTargetFolder : Set stTargetFolder = New StatusHolder
Dim stNewFolderName : Set stNewFolderName = New StatusHolder
Dim stOtaFilePath : Set stOtaFilePath = New StatusHolder

Dim statusProcess : statusProcess = PROCESS_WAIT

Dim vaFolderPath_BP : Set vaFolderPath_BP = New VariableArray
Dim vaFilePath_BP : Set vaFilePath_BP = New VariableArray
Dim vaFileNamesForCopy : Set vaFileNamesForCopy = New VariableArray
Dim vaHistoryTargetFolder : Set vaHistoryTargetFolder = New VariableArray

Call readCodePath(pAllCodePath, ID_INPUT_CODE_PATH, ID_LIST_CODE_PATH, ID_UL_CODE_PATH)
Call readHistory(pInputHistory, ID_INPUT_TARGET_FOLDER_PATH, ID_LIST_TARGET_HISTORY, ID_UL_TARGET_HISTORY)

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


Const STATUS_INVALID = 0
Const STATUS_VALID = 1
Const STATUS_WAIT = 2

Class StatusHolder
    Private mValue, mStatus, mInvalidMsg

    Private Sub Class_Initialize
        mValue = ""
        mStatus = STATUS_WAIT
    End Sub

    Public Property Get Value
        Value = mValue
    End Property

    Public Property Get Status
        Status = mStatus
    End Property

    Public Property Let Status(value)
        mStatus = value
    End Property

    Public Property Get InvalidMsg
        InvalidMsg = mInvalidMsg
    End Property

    Public Property Let InvalidMsg(value)
        mInvalidMsg = value
    End Property

    Public Sub Reset()
        mValue = ""
        mStatus = STATUS_WAIT
    End Sub

    Public Sub SetValue(status, value)
        mStatus = status
        mValue = value
    End Sub

    Public Sub SetStatusAndMsg(status, msg, showMsg)
        mStatus = status
        mInvalidMsg = msg
        If showMsg Then MsgBox(mInvalidMsg)
    End Sub

    Public Function checkInvalidAndShowMsg()
        If mStatus = STATUS_INVALID Then
            MsgBox(mInvalidMsg)
            checkInvalidAndShowMsg = True
            Exit Function
        End If
        checkInvalidAndShowMsg = False
    End Function

    Public Function checkStatusAndDoSomething(waitFun, invalidFun)
        If mStatus = STATUS_WAIT Then
            Call Execute(waitFun)
            If mStatus = STATUS_INVALID Then
                Call Execute(invalidFun)
                checkStatusAndDoSomething = True
            End If
        ElseIf checkInvalidAndShowMsg() Then
            checkStatusAndDoSomething = True
        End If
    End Function
End Class

Const SEARCH_FILE = 0
Const SEARCH_FOLDER = 1
Const SEARCH_ROOT = 0
Const SEARCH_SUB = 1
Const SEARCH_WHOLE_NAME = 0
Const SEARCH_PART_NAME = 1

Function searchFolder(pRootFolder, str, searchType, searchWhere, searchMode, findAll)
    If Not oFso.FolderExists(pRootFolder) Then searchFolder = "" : Exit Function
    If searchMode = SEARCH_WHOLE_NAME Then findAll = False

    Dim oRootFolder : Set oRootFolder = oFso.GetFolder(pRootFolder)

    Dim Folder, sTmp
    Select Case True
        Case searchType = SEARCH_FILE And searchWhere = SEARCH_ROOT
            If findAll Then
                Set searchFolder = startSearch(oRootFolder.Files, pRootFolder, str, searchMode, True)
            Else
                searchFolder = startSearch(oRootFolder.Files, pRootFolder, str, searchMode, False)
            End If

        Case searchType = SEARCH_FOLDER And searchWhere = SEARCH_ROOT
            If findAll Then
                Set searchFolder = startSearch(oRootFolder.SubFolders, pRootFolder, str, searchMode, True)
            Else
                searchFolder = startSearch(oRootFolder.SubFolders, pRootFolder, str, searchMode, False)
            End If

        Case searchType = SEARCH_FILE And searchWhere = SEARCH_SUB
            For Each Folder In oRootFolder.SubFolders
                sTmp = startSearch(Folder.Files, pRootFolder & "\" & Folder.Name, str, searchMode, False)
                If sTmp <> "" Then
                    searchFolder = sTmp
                    Exit Function
                End If
            Next
            searchFolder = ""

        Case searchType = SEARCH_FILE And searchWhere = SEARCH_SUB
            For Each Folder In oRootFolder.SubFolders
                sTmp = startSearch(Folder.SubFolders, pRootFolder & "\" & Folder.Name, str, searchMode, False)
                If sTmp <> "" Then
                    searchFolder = sTmp
                    Exit Function
                End If
            Next
            searchFolder = ""
    End Select
End Function

        Function startSearch(oAll, pRootFolder, str, searchMode, findAll)
            Dim oSingle

            If findAll Then
                Dim vaStr : Set vaStr = New VariableArray
                For Each oSingle In oAll
                    If checkSearchName(oSingle.Name, str, searchMode) Then
                        vaStr.Append(pRootFolder & "\" & oSingle.Name)
                    End If
                Next
                Set startSearch = vaStr
                Exit Function
            Else
                For Each oSingle In oAll
                    If checkSearchName(oSingle.Name, str, searchMode) Then
                        startSearch = pRootFolder & "\" & oSingle.Name
                        Exit Function
                End If
                Next
            End If
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

Sub readCodePath(DictPath, inputId, listId, ulId)
    Call readTextAndDoSomething(DictPath, _
            "If Len(Trim(sReadLine)) > 0 Then" &_
            " Call addAfterLi(sReadLine, """&inputId&""", """&listId&""", """&ulId&""")")
End Sub

Sub readHistory(DictPath, inputId, listId, ulId)
    Call readTextAndDoSomething(DictPath, _
            "If Len(Trim(sReadLine)) > 0 Then" &_
            " Call addBeforeLi(sReadLine, """&inputId&""", """&listId&""", """&ulId&""")" &_
            " : vaHistoryTargetFolder.Append(sReadLine)")
End Sub

Sub writeHistory(DictPath, inputId, listId, ulId, str)
    Dim seqInArray : seqInArray = vaHistoryTargetFolder.IsExistInArray(str)
    If seqInArray > -1 Then
        Call removeLi(ulId, vaHistoryTargetFolder.Length - seqInArray)
        Call addBeforeLi(str, inputId, listId, ulId)

        vaHistoryTargetFolder.MoveToEnd(seqInArray)

        Call writeNewHistoryTxt(DictPath)
        Exit Sub
    End If


    If vaHistoryTargetFolder.Length < 7 Then
        Call addBeforeLi(str, inputId, listId, ulId)

        vaHistoryTargetFolder.Append(str)

        Call appendStrToHistoryTxt(DictPath, str)
    Else
        Call removeLi(ulId, 0)
        Call addBeforeLi(str, inputId, listId, ulId)

        vaHistoryTargetFolder.PopBySeq(0)
        vaHistoryTargetFolder.Append(str)

        Call writeNewHistoryTxt(DictPath)
    End If
End Sub

Sub appendStrToHistoryTxt(DictPath, str)
    Call initTxtFile(DictPath)
    Call writeTextAndDoSomething(DictPath, _
            "oText.WriteLine("""&str&""")")
End Sub

Sub writeNewHistoryTxt(DictPath)
    Call initTxtFile(DictPath)
    Call writeTextAndDoSomething(DictPath, _
            "Dim i : For i = 0 To vaHistoryTargetFolder.Length" &_
            " : oText.WriteLine(vaHistoryTargetFolder.Value(i))" &_
            " : Next")
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

Sub setValueOfTargetFolder(inputId, listId, value)
    Call showAndHide(listId, "hide")
    Call setElementValue(inputId, value)

    Call doAfterSetInputValue(inputId)
End Sub

Sub openTargetFolder(inputId)
    Dim sTmp : sTmp = getElementValue(inputId)
    If sTmp = "" Then Exit Sub
    If oFso.FolderExists(sTmp) Then oWs.run "explorer.exe " & sTmp
End Sub

Sub readTextAndDoSomething(path, strFun)
    If Not oFso.FileExists(path) Then Exit Sub
    
    Dim oText, sReadLine, exitFlag
    Set oText = oFso.OpenTextFile(path, FOR_READING)
    exitFlag = False

    Do Until oText.AtEndOfStream
        sReadLine = oText.ReadLine
        Execute strFun
        If exitFlag Then Exit Do
    Loop

    oText.Close
    Set oText = Nothing
End Sub

Sub writeTextAndDoSomething(path, strFun)
    If Not oFso.FileExists(path) Then Exit Sub
    
    Dim oText : Set oText = oFso.OpenTextFile(path, FOR_APPENDING)

    Execute strFun

    oText.Close
    Set oText = Nothing
End Sub

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
Sub getBrowseValue(inputId)
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
        Call setElementValue(inputId, objFolderItem.Path)
        Call doAfterSetInputValue(inputId)
    End If 

    Set objShell = Nothing
    Set objFolder = Nothing
End Sub

Sub doAfterSetInputValue(inputId)
    Select Case inputId
        Case ID_INPUT_CODE_PATH
            Call getOutPath()
        Case ID_INPUT_TARGET_FOLDER_PATH
            Call getTargetPath(FROM_INPUT)
    End Select
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
    statusProcess = PROCESS_WAIT
End Sub

Sub clearOtherElements()
    Call setElementValue(ID_INPUT_TARGET_FOLDER_PATH, "")
    Call setElementValue(ID_INPUT_NEW_FOLDER_NAME, "")
    Call jsRemoveAllOption(ID_SELECT_OTA_FILE)
    Call jsAddOption(ID_SELECT_OTA_FILE, "如果编了OTA，请点击下面的按钮")
End Sub

Sub clearHoldedValues()
    Call stOutFolder.Reset()
    Call stOutProjectFolder.Reset()
    Call stTargetFolder.Reset()
    Call stNewFolderName.Reset()
    Call stOtaFilePath.Reset()
End Sub



'///////////////////////////////////////////////////////'
'check input
'///////////////////////////////////////////////////////'
Sub getOutPath()
    If bDebug Then MsgBox("getOutPath 111")

    If statusProcess = PROCESS_COPY Then Call clearProcess()
    Call clearOtherElements()
    Call clearHoldedValues()

    Call stOutFolder.SetValue(STATUS_WAIT, _
            getElementValue(ID_INPUT_CODE_PATH))
    

    If stOutFolder.Value = "" Then _
            Call stOutFolder.SetStatusAndMsg(STATUS_INVALID, _
                    "请选择或输入代码的out路径", True) : Exit Sub

    If Not oFso.FolderExists(stOutFolder.Value) Then _
            Call stOutFolder.SetStatusAndMsg(STATUS_INVALID, _
                    "当前代码的out路径不存在，请重新输入", True) : Exit Sub

    If Not oFso.FolderExists(stOutFolder.Value & "\target\product") Then _
            Call stOutFolder.SetStatusAndMsg(STATUS_INVALID, _
                    "当前代码的out路径下不存在\target\product，请重新输入", True) : Exit Sub

    Call stOutFolder.SetValue(STATUS_VALID, _
            stOutFolder.Value & "\target\product")
End Sub

Sub getOutProjectPath()
    If stOutFolder.checkInvalidAndShowMsg() Then Exit Sub

    Dim pSystemimg : pSystemimg = searchFolder(stOutFolder.Value, "system.img", SEARCH_FILE, SEARCH_SUB, SEARCH_WHOLE_NAME, False)

    If pSystemimg = "" Then _
            Call stOutProjectFolder.SetStatusAndMsg(STATUS_INVALID, _
                    "当前out目录下不存在system.img", True) : Exit Sub

    Call stOutProjectFolder.SetValue(STATUS_VALID, _
            Replace(pSystemimg, "\system.img", ""))
End Sub

Sub getTargetPath(where)
    If statusProcess = PROCESS_COPY Then Call clearProcess()

    Call stTargetFolder.SetValue(STATUS_WAIT, _
            getElementValue(ID_INPUT_TARGET_FOLDER_PATH))

    If stTargetFolder.Value = "" Then _
            Call stTargetFolder.SetStatusAndMsg(STATUS_WAIT, _
                    "请选择或输入目标文件夹路径", Eval(where = FROM_CHECK)) : Exit Sub

    If Not oFso.FolderExists(stTargetFolder.Value) Then _
            Call stTargetFolder.SetStatusAndMsg(STATUS_INVALID, _
                    "目标文件夹路径不存在，请重新输入", True) : Exit Sub

    stTargetFolder.Status = STATUS_VALID
End Sub

Sub getNewFolderName(where)
    If statusProcess = PROCESS_COPY Then Call clearProcess()

    Call stNewFolderName.SetValue(STATUS_WAIT, _
            getElementValue(ID_INPUT_NEW_FOLDER_NAME))

    If stNewFolderName.Value = "" Then _
        Call stNewFolderName.SetStatusAndMsg(STATUS_WAIT, _
                "请选择或输入新文件夹名", Eval(where = FROM_CHECK)) : Exit Sub

    Call checkCharOfFolderName(stNewFolderName.Value)
End Sub

        Sub checkCharOfFolderName(name)
            Dim i, char
            For i = 1 To Len(name)
                char = Mid(name, i, 1)
                If InStr(INVALID_CHAR_OF_FOlDER_NAME, char) Then
                    Call stNewFolderName.SetStatusAndMsg(STATUS_INVALID, _
                            "文件名不能包含下列任何字符：" & Vblf & INVALID_CHAR_OF_FOlDER_NAME, True)
                    Exit Sub
                End If
            Next
            stNewFolderName.Status = STATUS_VALID
        End Sub

Sub getCustomBuildVersion(which)
    '//check out path
    If stOutFolder.checkStatusAndDoSomething("Call getOutPath()", "") Then Exit Sub
    If stOutProjectFolder.checkStatusAndDoSomething("Call getOutProjectPath()", "") Then Exit Sub

    '//check build.prop
    Dim pBuildProp : pBuildProp = stOutProjectFolder.Value & "\system\build.prop"
    If Not oFso.FileExists(pBuildProp) Then MsgBox("文件不存在: " & Vblf & pBuildProp) : Exit Sub

    '//get version
    Call readTextAndDoSomething(pBuildProp, _
            "If InStr(sReadLine, """&which&""") > 0 Then" &_
                " Call setElementValue(ID_INPUT_NEW_FOLDER_NAME, Trim(Mid(sReadLine,InStr(sReadLine,""="")+1)))" &_
                " : Call getNewFolderName(FROM_INPUT)" &_
                " : exitFlag = True")
End Sub

Sub checkOtaFiles()
    If bDebug Then MsgBox("checkOtaFiles")
    If stOtaFilePath.Status <> STATUS_WAIT Then Exit Sub

    If stOutFolder.checkStatusAndDoSomething("Call getOutPath()", "") Then Exit Sub
    If stOutProjectFolder.checkStatusAndDoSomething("Call getOutProjectPath()", "") Then Exit Sub

    Dim pOta_1 : pOta_1 = searchFolder(stOutProjectFolder.Value, "target_files-package.zip" _
            , SEARCH_FILE, SEARCH_ROOT, SEARCH_WHOLE_NAME, False)
    Dim pOta_2 : pOta_2 = searchFolder(stOutProjectFolder.Value, "-ota-"_
            , SEARCH_FILE, SEARCH_ROOT, SEARCH_PART_NAME, False)
    Dim pOta_3 : pOta_3 = searchFolder(stOutProjectFolder.Value & "\obj\PACKAGING\target_files_intermediates", "-target_files-" _
            , SEARCH_FILE, SEARCH_ROOT, SEARCH_PART_NAME, False)

    If pOta_1 = "" And pOta_2 = "" And pOta_3 = "" Then
        Call jsRemoveAllOption(ID_SELECT_OTA_FILE)
        Call jsAddOptionValueAndName(ID_SELECT_OTA_FILE, "", "无OTA文件")
        Call stOtaFilePath.SetValue(STATUS_VALID, "")
        Exit Sub
    End If

    Call jsRemoveAllOption(ID_SELECT_OTA_FILE)

    Call jsAddOptionValueAndName(ID_SELECT_OTA_FILE, "", "不拷贝OTA")
    If pOta_1 <> "" Then Call jsAddOption(ID_SELECT_OTA_FILE, Replace(pOta_1, stOutProjectFolder.Value, ""))
    If pOta_2 <> "" Then Call jsAddOption(ID_SELECT_OTA_FILE, Replace(pOta_2, stOutProjectFolder.Value, ""))
    If pOta_3 <> "" Then Call jsAddOption(ID_SELECT_OTA_FILE, Replace(pOta_3, stOutProjectFolder.Value, ""))

    Call stOtaFilePath.SetValue(STATUS_VALID, "")
End Sub

Sub getOtaFiles()
    Call stOtaFilePath.SetValue(STATUS_WAIT, _
            getElementValue(ID_SELECT_OTA_FILE))

    If stOtaFilePath.Value <> "" Then
        Call stOtaFilePath.SetValue(STATUS_VALID , _
                stOutProjectFolder.Value & stOtaFilePath.Value)
    End If
End Sub

Sub runCopy()
    '//init copy process
    If statusProcess = PROCESS_COPY Then Call clearProcess()
    Call setProcessStatus("checking")
    statusProcess = PROCESS_COPY

    '//check holded value
    If stOutFolder.checkStatusAndDoSomething("Call getOutPath()", "Call setProcessStatus(""wait"")") Then Exit Sub
    If stOutProjectFolder.checkStatusAndDoSomething("Call getOutProjectPath()", "Call setProcessStatus(""wait"")") Then Exit Sub
    If stTargetFolder.checkStatusAndDoSomething("Call getTargetPath(FROM_CHECK)", "Call setProcessStatus(""wait"")") Then Exit Sub
    If stNewFolderName.checkStatusAndDoSomething("Call getNewFolderName(FROM_CHECK)", "Call setProcessStatus(""wait"")") Then Exit Sub

    '//dim folder path for copy
    Dim pCopyFolder_software : pCopyFolder_software = stTargetFolder.Value & "\" & stNewFolderName.Value
    Dim pCopyFolder_db : pCopyFolder_db = pCopyFolder_software & "\DB"
    Dim pCopyFolder_ota : pCopyFolder_ota = pCopyFolder_software & "\OTA"

    '//check folder path for copy
    If oFso.FolderExists(pCopyFolder_software) Then MsgBox("目标路径已存在""" & stNewFolderName.Value & """文件夹！") : Call setProcessStatus("wait") : Exit Sub

    '//set running status of copy process
    Call setProcessStatus("copying")

    '//create folder
    oFso.CreateFolder(pCopyFolder_software)
    oFso.CreateFolder(pCopyFolder_db)

    '//check files for copy, and set element str of copy process.
    Call checkDbFiles()
    Call checkSoftwareFiles()
    If stOtaFilePath.Value <> "" Then Call setElementInnerHTML(ID_DIV_OTA_WAIT, 1)

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
            Call copyFile(stOutProjectFolder.Value & "\" & vaFileNamesForCopy.Value(j), pCopyFolder_software)
            Call setElementInnerHTML(ID_DIV_SOFTWARE_COPYED, Cint(getElementInnerHTML(ID_DIV_SOFTWARE_COPYED)) + 1)
        Next
        Call copyFile(stOutProjectFolder.Value & "\system.img", pCopyFolder_software)
        Call setElementInnerHTML(ID_DIV_SOFTWARE_COPYED, Cint(getElementInnerHTML(ID_DIV_SOFTWARE_COPYED)) + 1)
    End If

    '//start copy OTA files
    If stOtaFilePath.Value <> "" Then
        oFso.CreateFolder(pCopyFolder_ota)
        Call copyFile(stOtaFilePath.Value, pCopyFolder_ota)
        Call setElementInnerHTML(ID_DIV_OTA_COPYED, Cint(getElementInnerHTML(ID_DIV_OTA_COPYED)) + 1)
    End If

    '//set finish status of copy process
    Call setProcessStatus("done")

    '//save history of target folder path
    Call writeHistory(pInputHistory, ID_INPUT_TARGET_FOLDER_PATH, ID_LIST_TARGET_HISTORY, ID_UL_TARGET_HISTORY, stTargetFolder.Value)
End Sub

        Sub setProcessStatus(status)
            Select Case status
                Case "wait"
                    Call setElementInnerHTML(ID_DIV_COPY_STATUS, "未开始")
                    Call setElementColor(ID_DIV_COPY_STATUS, "black")
                Case "checking"
                    Call setElementInnerHTML(ID_DIV_COPY_STATUS, "检查文件中...")
                    Call setElementColor(ID_DIV_COPY_STATUS, "blue")
                Case "copying"
                    Call setElementInnerHTML(ID_DIV_COPY_STATUS, "拷贝中...")
                    Call setElementColor(ID_DIV_COPY_STATUS, "red")
                Case "done"
                    Call setElementColor(ID_DIV_COPY_STATUS, "green")
                    Call setElementInnerHTML(ID_DIV_COPY_STATUS, "拷贝完成！")
            End Select
        End Sub

        Sub checkSoftwareFiles()
            If bDebug Then MsgBox("checkSoftwareFiles")
            vaFileNamesForCopy.ResetArray()

            Dim uScatterFilePath : uScatterFilePath = searchFolder(stOutProjectFolder.Value, "_Android_scatter.txt" _
                    , SEARCH_FILE, SEARCH_ROOT, SEARCH_PART_NAME, False)

            If uScatterFilePath = "" Then
                MsgBox("""Android_scatter.txt"" is not exists in: " & Vblf & stOutProjectFolder.Value)
                Exit Sub
            End If

            Call getNeedFilesName(uScatterFilePath)

            If vaFileNamesForCopy.Length > -1 Then
                Call setElementInnerHTML(ID_DIV_SOFTWARE_WAIT, vaFileNamesForCopy.Length + 2)
            End If
        End Sub

        Sub getNeedFilesName(uPath)
            If bDebug Then MsgBox("getNeedFilesName")

            '//add Android_scatter.txt first
            vaFileNamesForCopy.Append(getFileNameOfPath(uPath))

            Call readTextAndDoSomething(uPath, _
                    "Call checkAndAddFileNameForCopy(sReadLine)")

            vaFileNamesForCopy.SortArray()
        End Sub

        Sub checkAndAddFileNameForCopy(sReadLine)
            Dim iInFileName : iInFileName = InStr(sReadLine, "file_name")
            If iInFileName > 0 Then
                Dim sTmpFileName : sTmpFileName = Mid(sReadLine, iInFileName + 11)
                If sTmpFileName <> "NONE" And _
                        sTmpFileName <> "system.img" And _
                        vaFileNamesForCopy.IsExistInArray(sTmpFileName) = -1 Then
                    vaFileNamesForCopy.Append(sTmpFileName)
                End If
            End If
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
                Case oFso.FolderExists(stOutProjectFolder.Value & sKK_AP)
                    pFolder_AP = stOutProjectFolder.Value & sKK_AP
                Case oFso.FolderExists(stOutProjectFolder.Value & sL1_AP)
                    pFolder_AP = stOutProjectFolder.Value & sL1_AP
                Case Else
                    pFolder_AP = ""
            End Select

            If pFolder_AP <> "" Then pFile_AP = searchFolder(pFolder_AP, "_ENUM", SEARCH_FILE, SEARCH_ROOT, SEARCH_PART_NAME, False)
            If pFile_AP <> "" Then pFile_AP = Replace(pFile_AP, "_ENUM", "")

            If pFile_AP = "" Then MsgBox("AP file is not exists!")

            '//get BP file path
            Dim vaTmp, sTmp
            Select Case True
                Case oFso.FolderExists(stOutProjectFolder.Value & sKK_BP)
                    vaFolderPath_BP.Append(stOutProjectFolder.Value & sKK_BP)
                Case oFso.FolderExists(stOutProjectFolder.Value & sL1_BP)
                    Set vaTmp = searchFolder(stOutProjectFolder.Value & sL1_BP, "BPLGU", SEARCH_FOLDER, SEARCH_ROOT, SEARCH_PART_NAME, True)
                    If vaTmp.Length > -1 Then Set vaFolderPath_BP = vaTmp
            End Select

            If vaFolderPath_BP.Length > -1 Then
                Dim i 
                For i = 0 To vaFolderPath_BP.Length
                    sTmp = searchFolder(vaFolderPath_BP.Value(i), "BPLGU", SEARCH_FILE, SEARCH_ROOT, SEARCH_PART_NAME, False)
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

        Sub copyFile(uCopyFilePath, uTargetFolderPath)
            If oFso.FileExists(uCopyFilePath) Then
                Dim filePath : filePath = """" & uCopyFilePath & """"
                Dim folderPath : folderPath = """" & uTargetFolderPath & """"
                
                oWs.Run "FsoCopyFile.vbs " & filePath & " " & folderPath & "\", , True
            Else
               MsgBox(uCopyFilePath & " is not exist!")
            End If
        End Sub