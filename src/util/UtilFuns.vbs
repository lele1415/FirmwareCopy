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

Const FOR_READING = 1
Const FOR_APPENDING = 8

Const ID_DIV_COPY_STATUS = "copy_status"

Const STATUS_WAIT = "wait"
Const STATUS_CHECK = "check"
Const STATUS_COPY = "copy"
Const STATUS_DONE = "done"

Const CHECK_WHEN_INPUT = 0
Const CHECK_WHEN_COPY = 1

Function initTxtFile(FilePath)
    If oFso.FileExists(FilePath) Then
        Dim TxtFile
        Set TxtFile = oFso.getFile(FilePath)
        TxtFile.Delete
        Set TxtFile = Nothing
    End If    
    oFso.CreateTextFile FilePath, True
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

Sub openFolderByInput(inputId)
    Dim sTmp : sTmp = getElementValue(inputId)
    If sTmp = "" Then Exit Sub
    If oFso.FolderExists(sTmp) Then oWs.run "explorer.exe " & sTmp
End Sub

Sub checkCopyInfoIsEnough()
    If mCodePath <> "" And _
            mOutSoftwarePath <> "" And _
            mTargetPath <> "" And _
            mNewFolderName <> "" Then
        Call enableElement(ID_BUTTON_START_COPY)
    Else
        Call disableElement(ID_BUTTON_START_COPY)
    End If
End Sub

Sub setProcessStatus(status)
    Select Case status
        Case STATUS_WAIT
            Call setElementInnerHTML(ID_DIV_COPY_STATUS, "空闲")
            Call setElementColor(ID_DIV_COPY_STATUS, "black")
        Case STATUS_CHECK
            Call setElementInnerHTML(ID_DIV_COPY_STATUS, "检查文件中...")
            Call setElementColor(ID_DIV_COPY_STATUS, "blue")
        Case STATUS_COPY
            Call setElementInnerHTML(ID_DIV_COPY_STATUS, "拷贝中...")
            Call setElementColor(ID_DIV_COPY_STATUS, "red")
        Case STATUS_DONE
            Call setElementColor(ID_DIV_COPY_STATUS, "green")
            Call setElementInnerHTML(ID_DIV_COPY_STATUS, "拷贝完成！")
    End Select
End Sub

Sub resetAllProcessStatus()
    Call setElementInnerHTML(ID_DIV_DB_WAIT, 0)
    Call setElementInnerHTML(ID_DIV_SOFTWARE_WAIT, 0)
    Call setElementInnerHTML(ID_DIV_OTA_WAIT, 0)
    Call setElementInnerHTML(ID_DIV_DB_COPYED, 0)
    Call setElementInnerHTML(ID_DIV_SOFTWARE_COPYED, 0)
    Call setElementInnerHTML(ID_DIV_OTA_COPYED, 0)
End Sub

Sub resetAllCopyedProcessStatus()
    Call setElementInnerHTML(ID_DIV_DB_COPYED, 0)
    Call setElementInnerHTML(ID_DIV_SOFTWARE_COPYED, 0)
    Call setElementInnerHTML(ID_DIV_OTA_COPYED, 0)
End Sub

Sub freezeAllInput()
    Call disableElement(ID_INPUT_CODE_PATH)
    'Call disableElement(ID_BUTTON_OPEN_CODE_PATH)

    Call disableElement(ID_INPUT_TARGET_PATH)
    'Call disableElement(ID_BUTTON_OPEN_TARGET_PATH)

    Call disableElement(ID_INPUT_FOLDER_NAME)
    Call disableElement(ID_BUTTON_DISPLAY_ID)
    Call disableElement(ID_BUTTON_BUILD_VERSION)

    Call disableElement(ID_SELECT_OTA_FILE)
    Call disableElement(ID_BUTTON_CHECK_OTA_FILES)

    Call disableElement(ID_BUTTON_START_COPY)
End Sub

Sub unfreezeAllInput()
    Call enableElement(ID_INPUT_CODE_PATH)
    'Call enableElement(ID_BUTTON_OPEN_CODE_PATH)

    Call enableElement(ID_INPUT_TARGET_PATH)
    'Call enableElement(ID_BUTTON_OPEN_TARGET_PATH)

    Call enableElement(ID_INPUT_FOLDER_NAME)
    Call enableElement(ID_BUTTON_DISPLAY_ID)
    Call enableElement(ID_BUTTON_BUILD_VERSION)

    Call enableElement(ID_SELECT_OTA_FILE)
    Call enableElement(ID_BUTTON_CHECK_OTA_FILES)
    
    Call enableElement(ID_BUTTON_START_COPY)
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

Function checkFolderExists(path)
    If path = "" Then
        checkFolderExists = False
        Exit Function
    End If

    If Not oFso.FolderExists(path) Then
        checkFolderExists = False
        Exit Function
    End If

    checkFolderExists = True
End Function

Sub clearInputValue(inputId)
    Call setElementValue(inputId, "")
End Sub
